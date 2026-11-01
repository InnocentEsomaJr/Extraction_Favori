"""Agrégation des datasets de reporting sans multiplier les comptes par enfant."""

import numpy as np
import pandas as pd


def _normalise_name(value):
    if value is None or pd.isna(value):
        return ""
    return " ".join(str(value).casefold().split())


def clean_ou_id(value):
    """Nettoie un identifiant sans transformer les valeurs absentes en texte."""
    if value is None or pd.isna(value):
        return ""
    cleaned = str(value).strip()
    return "" if cleaned.casefold() in {"", "nan", "none", "null", "<na>", "nat"} else cleaned


def attach_reporting_scope(df_source, scope_meta):
    """Associe les unités par ID et n'utilise que les noms non ambigus.

    Une unité dont la hiérarchie est absente conserve son propre ID. Les
    noms homonymes ne sont jamais utilisés pour inventer une correspondance.
    Dans chaque groupe, une ligne de l'unité cible prime sur ses enfants.
    """
    result = df_source.copy() if df_source is not None else pd.DataFrame()
    for column in ["Organisation unit", "ou_id", "Zone_ID", "Zone"]:
        if column not in result.columns:
            result[column] = ""
    if result.empty:
        return result

    metadata = scope_meta.copy() if scope_meta is not None else pd.DataFrame()
    for column in ["Organisation unit", "Organisation unit ID", "Zone_ID", "Zone"]:
        if column not in metadata.columns:
            metadata[column] = ""

    by_id = {}
    by_name = {}
    for name, own_id, zone_id, zone_name in metadata[
        ["Organisation unit", "Organisation unit ID", "Zone_ID", "Zone"]
    ].itertuples(index=False, name=None):
        own_id, zone_id = clean_ou_id(own_id), clean_ou_id(zone_id)
        normalized_name = _normalise_name(name)
        normalized_zone = _normalise_name(zone_name)
        record = (own_id, zone_id, str(zone_name).strip() if normalized_zone else "")
        identity = (own_id, zone_id, normalized_zone)
        if own_id:
            by_id.setdefault(own_id, {})[identity] = record
        if normalized_name:
            by_name.setdefault(normalized_name, {})[identity] = record

    source_ids = (
        result["Organisation unit ID"].map(clean_ou_id)
        if "Organisation unit ID" in result.columns
        else pd.Series("", index=result.index)
    )
    attached = []
    for name, own_id in zip(result["Organisation unit"], source_ids):
        candidates = by_id.get(own_id, {}) if own_id else {}
        if not candidates:
            named = by_name.get(_normalise_name(name), {})
            if len(named) == 1:
                candidate = next(iter(named.values()))
                if not own_id or not candidate[0] or candidate[0] == own_id:
                    candidates = named
        match = next(iter(candidates.values())) if len(candidates) == 1 else None
        if match:
            own_id = own_id or match[0]
            zone_id = match[1] or own_id
            zone_name = match[2] or (str(name).strip() if _normalise_name(name) else "")
        else:
            zone_id = own_id
            zone_name = str(name).strip() if _normalise_name(name) else ""
        attached.append((own_id, zone_id, zone_name))
    result[["ou_id", "Zone_ID", "Zone"]] = attached

    direct = result["ou_id"].ne("") & result["ou_id"].eq(result["Zone_ID"])
    has_direct = direct.groupby([result["Zone_ID"], result["Zone"]], dropna=False).transform("any")
    return result.loc[direct | ~has_direct].copy()


def _dataset_key(column, suffix):
    name = _normalise_name(column)
    if "on time" in name or not name.endswith(suffix):
        return None
    return name[:-len(suffix)].strip(" -_")


def count_datasets_95(df, rate_columns, target_id=None, id_column="ou_id"):
    """Compte une fois chaque dataset dont le taux agrégé atteint 95 %.

    Une ligne de l'unité cible prime sur ses descendants. Sinon, les taux
    sont recalculés depuis les rapports réels et attendus, jamais moyennés.
    Un résultat manquant signifie que le compte complet n'est pas calculable.
    Les lignes avec zéro rapport attendu et zéro rapport réel sont exclues.
    """
    if df is None or df.empty or not rate_columns:
        return float("nan")

    selected = df
    direct = False
    target_key = clean_ou_id(target_id)
    if target_key and id_column in df.columns:
        exact = df[id_column].map(clean_ou_id).eq(target_key)
        if exact.any():
            selected = df.loc[exact]
            direct = True

    datasets = {}
    for column in rate_columns:
        if column not in selected.columns:
            return float("nan")
        key = _dataset_key(column, "reporting rate")
        if key is not None:
            if key in datasets and datasets[key] != column:
                return float("nan")
            datasets[key] = column
    if not datasets:
        return float("nan")

    metrics = {"actual reports": {}, "expected reports": {}}
    for suffix, mapping in metrics.items():
        for column in selected.columns:
            key = _dataset_key(column, suffix)
            if key is not None:
                mapping.setdefault(key, []).append(column)

    count = 0
    applicable = 0
    for key, rate_column in datasets.items():
        actual_columns = metrics["actual reports"].get(key, [])
        expected_columns = metrics["expected reports"].get(key, [])
        if len(actual_columns) > 1 or len(expected_columns) > 1:
            return float("nan")
        if actual_columns and expected_columns:
            actual = pd.to_numeric(selected[actual_columns[0]], errors="coerce")
            expected = pd.to_numeric(selected[expected_columns[0]], errors="coerce")
            invalid = (
                actual.isna() | expected.isna()
                | ~np.isfinite(actual) | ~np.isfinite(expected)
                | (actual < 0) | (expected < 0)
                | ((expected == 0) & (actual > 0))
            )
            if invalid.any():
                return float("nan")
            expected_total = expected.sum()
            if expected_total == 0:
                continue
            rate = actual.sum() / expected_total * 100
        elif direct and len(selected) == 1:
            rate = pd.to_numeric(selected[rate_column], errors="coerce").iloc[0]
            if pd.isna(rate) or not np.isfinite(rate) or rate < 0:
                return float("nan")
        else:
            return float("nan")
        applicable += 1
        count += int(rate >= 95)

    return float(count) if applicable else float("nan")


def aggregate_validation_metrics(reporting_df, targets):
    """Agrège les indicateurs mensuels des seules zones ciblées par leur ID.

    Les données Analytics sont attendues au niveau ZS. Une ligne sans ID,
    d'un autre niveau ou d'une autre zone n'est jamais associée par son nom.
    Les zones sans données et les métriques incomplètes restent indisponibles.
    Les mois restent entiers : aucun prorata journalier n'est appliqué.
    """
    output_columns = [
        "Zone_ID", "Complétude_Globale", "Promptitude_Globale",
        "Actual_Reports", "Nombre_Datasets_95",
    ]
    if targets is None or targets.empty or "Zone_ID" not in targets:
        return pd.DataFrame(columns=output_columns)

    selected = targets.copy()
    selected["Zone_ID"] = selected["Zone_ID"].map(clean_ou_id)
    selected = selected.loc[selected["Zone_ID"].ne("")].drop_duplicates("Zone_ID")
    if selected.empty:
        return pd.DataFrame(columns=output_columns)

    source = reporting_df.copy() if reporting_df is not None else pd.DataFrame()
    if "Organisation unit ID" not in source:
        source = source.iloc[:0].copy()
        source["Organisation unit ID"] = pd.Series(dtype=object)
    source["Organisation unit ID"] = source["Organisation unit ID"].map(clean_ou_id)
    source = source.loc[source["Organisation unit ID"].isin(selected["Zone_ID"])].copy()
    if "Organisation unit" not in source:
        source["Organisation unit"] = source["Organisation unit ID"]
    # Une répétition de la même ligne n'est pas un second rapport. Une vraie
    # période différente est conservée grâce à ses champs de période distincts.
    source = source.drop_duplicates()

    scope = pd.DataFrame({
        "Organisation unit ID": selected["Zone_ID"],
        "Organisation unit": selected["Zone"] if "Zone" in selected else selected["Zone_ID"],
        "Zone_ID": selected["Zone_ID"],
        "Zone": selected["Zone"] if "Zone" in selected else selected["Zone_ID"],
    })
    source = attach_reporting_scope(source, scope)
    actual_columns, expected_columns, on_time_columns, rate_columns = [], [], [], []
    for column in source.columns:
        name = _normalise_name(column)
        if "actual reports" in name:
            (on_time_columns if "on time" in name else actual_columns).append(column)
        elif "expected reports" in name and "on time" not in name:
            expected_columns.append(column)
        if "reporting rate" in name and "on time" not in name:
            rate_columns.append(column)

    def total(frame, columns):
        if not columns:
            return float("nan")
        numeric = frame[columns].apply(pd.to_numeric, errors="coerce").astype(float)
        if numeric.isna().any().any() or (~np.isfinite(numeric)).any().any() or numeric.lt(0).any().any():
            return float("nan")
        return round(float(numeric.to_numpy().sum()), 2)

    records = []
    for target_id in selected["Zone_ID"]:
        rows = source.loc[source["ou_id"].eq(target_id)]
        metrics = {column: float("nan") for column in output_columns[1:]}
        if not rows.empty:
            actual = total(rows, actual_columns)
            expected = total(rows, expected_columns)
            on_time = total(rows, on_time_columns)
            metrics["Actual_Reports"] = actual
            if pd.notna(expected) and expected > 0:
                if pd.notna(actual):
                    metrics["Complétude_Globale"] = round(actual / expected * 100, 2)
                if pd.notna(on_time):
                    metrics["Promptitude_Globale"] = round(on_time / expected * 100, 2)
            metrics["Nombre_Datasets_95"] = count_datasets_95(rows, rate_columns, target_id=target_id)
        records.append({"Zone_ID": target_id, **metrics})
    return pd.DataFrame(records, columns=output_columns)


def _quality_count_columns(frame):
    """Repère les trois métriques de comptes, avec leurs préfixes de dataset."""
    families = {
        "rapports réels": "actual reports",
        "rapports attendus": "expected reports",
        "rapports à temps": "actual reports on time",
    }
    columns = {label: {} for label in families}
    if not isinstance(frame, pd.DataFrame):
        return columns
    for column in frame.columns:
        name = _normalise_name(column)
        for label, suffix in families.items():
            if name == suffix or name.endswith(" " + suffix):
                dataset = name[:-len(suffix)].strip(" -_")
                columns[label].setdefault(dataset, []).append(column)
    return columns


def has_complete_reporting_metrics(frame):
    """Indique si chaque dataset possède les trois comptes nécessaires et uniques.

    Ce contrôle porte sur les colonnes, pas sur les valeurs ni le périmètre.
    Une cellule absente reste à diagnostiquer avec ``reporting_quality_warnings``.
    """
    families = _quality_count_columns(frame)
    if not all(bool(columns) for columns in families.values()):
        return False
    dataset_keys = [set(columns) for columns in families.values()]
    if any(keys != dataset_keys[0] for keys in dataset_keys[1:]):
        return False
    if any(len(columns) != 1 for mapping in families.values() for columns in mapping.values()):
        return False
    # Un dataset seulement représenté par un taux doit aussi déclencher la
    # récupération de ses comptes, même si un autre dataset est déjà complet.
    for column in frame.columns:
        name = _normalise_name(column)
        for suffix in ("reporting rate on time", "reporting rate"):
            if name == suffix or name.endswith(" " + suffix):
                dataset = name[:-len(suffix)].strip(" -_")
                if dataset not in dataset_keys[0]:
                    return False
                break
    return True


def reporting_quality_warnings(reporting_df, metrics, targets):
    """Explique les indicateurs de déclaration manquants sans modifier leur calcul.

    Les zones sont recherchées exclusivement par leur ID. Les messages restent
    propres aux données de déclaration et ne dépendent pas des violations DHIS2.
    Une source incomplète n'est ni complétée par zéro ni remplacée par une moyenne.
    """
    if not isinstance(targets, pd.DataFrame) or targets.empty or "Zone_ID" not in targets:
        return []
    selected = targets.copy()
    selected["Zone_ID"] = selected["Zone_ID"].map(clean_ou_id)
    selected = selected.loc[selected["Zone_ID"].ne("")].drop_duplicates("Zone_ID")
    if selected.empty:
        return []
    source = reporting_df if isinstance(reporting_df, pd.DataFrame) else pd.DataFrame()
    families = _quality_count_columns(source)
    messages = []
    missing_families = [label for label, columns in families.items() if not columns]
    if missing_families:
        messages.append(
            "La source de déclaration ne fournit pas les métriques nécessaires : "
            + ", ".join(missing_families) + "."
        )
    if "Organisation unit ID" not in source:
        messages.append(
            "Les identifiants des unités sont absents de la source de déclaration ; "
            "les données ne peuvent pas être rattachées aux zones sélectionnées."
        )
        return messages
    source_ids = source["Organisation unit ID"].map(clean_ou_id)
    metric_frame = metrics if isinstance(metrics, pd.DataFrame) else pd.DataFrame()
    metric_ids = (
        metric_frame["Zone_ID"].map(clean_ou_id) if "Zone_ID" in metric_frame
        else pd.Series("", index=metric_frame.index, dtype=object)
    )
    metric_labels = {
        "Complétude_Globale": "complétude",
        "Promptitude_Globale": "promptitude",
        "Actual_Reports": "rapports réels",
        "Nombre_Datasets_95": "nombre de datasets à complétude ≥95 %",
    }
    all_datasets = set().union(*(set(columns) for columns in families.values()))
    incomplete_datasets = [
        dataset or "dataset sans libellé" for dataset in sorted(all_datasets)
        if any(dataset not in columns or len(columns[dataset]) != 1 for columns in families.values())
    ]
    if not missing_families and incomplete_datasets:
        messages.append(
            "Les trois métriques de comptes ne sont pas toutes présentes ou uniques pour les datasets : "
            + ", ".join(incomplete_datasets) + "."
        )
    for target in selected.to_dict("records"):
        identifier = target["Zone_ID"]
        name = target.get("Zone")
        label = str(name).strip() if _normalise_name(name) else identifier
        zone_label = f"{label} ({identifier})" if label != identifier else identifier
        rows = source.loc[source_ids.eq(identifier)]
        if rows.empty:
            messages.append(f"{zone_label} : aucune donnée de déclaration reçue pour cette zone et cette période.")
            continue
        reasons = []
        missing_values, invalid_values = [], []
        for family, dataset_columns in families.items():
            columns = [column for group in dataset_columns.values() for column in group]
            if not columns:
                continue
            raw = rows[columns]
            numeric = raw.apply(pd.to_numeric, errors="coerce").astype(float)
            if (raw.isna() | raw.astype(str).apply(lambda column: column.str.strip().eq(""))).any().any():
                missing_values.append(family)
            if (
                (numeric.isna() & raw.notna() & raw.astype(str).apply(lambda column: column.str.strip().ne(""))).any().any()
                or ((~np.isfinite(numeric)) & numeric.notna()).any().any()
                or numeric.lt(0).any().any()
            ):
                invalid_values.append(family)
        if missing_values:
            reasons.append("valeurs manquantes pour " + ", ".join(missing_values))
        if invalid_values:
            reasons.append("comptes invalides pour " + ", ".join(invalid_values))
        expected_columns = [column for group in families["rapports attendus"].values() for column in group]
        if expected_columns:
            expected = rows[expected_columns].apply(pd.to_numeric, errors="coerce").astype(float)
            if expected.notna().all().all() and np.isfinite(expected).all().all() and expected.eq(0).all().all():
                reasons.append("aucun rapport attendu ; les taux ne sont pas calculables")
        final_rows = metric_frame.loc[metric_ids.eq(identifier)]
        if len(final_rows) != 1:
            reasons.append("synthèse de déclaration absente ou ambiguë pour cet identifiant")
        elif not reasons:
            unavailable = []
            for column, metric_label in metric_labels.items():
                value = pd.to_numeric(pd.Series([final_rows.iloc[0].get(column)]), errors="coerce").iloc[0]
                if pd.isna(value) or not np.isfinite(value) or value < 0:
                    unavailable.append(metric_label)
            if unavailable:
                reasons.append("indicateurs non calculables : " + ", ".join(unavailable))
        if reasons:
            messages.append(f"{zone_label} : {' ; '.join(reasons)}.")
    return messages

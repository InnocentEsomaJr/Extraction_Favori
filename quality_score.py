"""Tableau de score qualité : reporting courant et violations rattachées par ID."""

import math

import pandas as pd

from geography import labels_for_units
from reporting_metrics import aggregate_validation_metrics, clean_ou_id


REPORTING_METRIC_COLUMNS = [
    'Zone_ID', 'Complétude_Globale', 'Promptitude_Globale',
    'Actual_Reports', 'Nombre_Datasets_95',
]


def _number(value):
    numeric = pd.to_numeric(value, errors='coerce')
    return float(numeric) if pd.notna(numeric) and math.isfinite(float(numeric)) else float('nan')


def _display_name(value):
    return '' if value is None or pd.isna(value) else str(value).strip()


def _explicit_unit_ids(frame):
    columns = [column for column in ['Organisation unit ID', 'Zone_ID'] if column in frame]
    if not columns:
        return None
    identifiers = []
    for row in frame[columns].to_dict('records'):
        candidates = {clean_ou_id(row[column]) for column in columns if clean_ou_id(row[column])}
        identifiers.append(next(iter(candidates)) if len(candidates) == 1 else '')
    return pd.Series(identifiers, index=frame.index, dtype=object)


def _count_evidence(rows, columns):
    """Tolère les cellules absentes, mais exige au moins un compte réel valide."""
    if not columns:
        return False, False
    values = rows[columns].apply(pd.to_numeric, errors='coerce').astype(float)
    provided = rows[columns].notna() & rows[columns].astype(str).apply(lambda col: col.str.strip().ne(''))
    malformed = values.isna() & provided
    invalid = malformed.any().any() or values.lt(0).any().any()
    invalid = invalid or any(
        not math.isfinite(float(value)) for value in values.to_numpy().ravel() if pd.notna(value)
    )
    has_values = bool(values.notna().any().any()) and not invalid
    return has_values, bool(has_values and values.sum().sum() > 0)


def metrics_from_dashboard_summary(summary_df, source_df, targets):
    """Réutilise les nombres de la synthèse affichée, avec rattachement sûr par ID.

    Le caller garantit les mêmes mois et le même favori. Les ID de la synthèse
    priment ; sans ID, le pont des libellés du tableau courant doit être bijectif.
    Les comptes bruts attestent la disponibilité des données afin qu'un zéro
    créé par un ancien ``fillna(0)`` ne soit pas pris pour une mesure réelle.
    Aucune moyenne ni pondération supplémentaire n'est appliquée à la synthèse.
    """
    if not isinstance(targets, pd.DataFrame) or 'Zone_ID' not in targets:
        return pd.DataFrame(columns=REPORTING_METRIC_COLUMNS)
    target_ids = list(dict.fromkeys(targets['Zone_ID'].map(clean_ou_id)))
    target_ids = [identifier for identifier in target_ids if identifier]
    records = [
        {'Zone_ID': identifier, **{column: float('nan') for column in REPORTING_METRIC_COLUMNS[1:]}}
        for identifier in target_ids
    ]
    result = pd.DataFrame(records, columns=REPORTING_METRIC_COLUMNS)
    if (
        result.empty or not isinstance(source_df, pd.DataFrame) or source_df.empty
        or not isinstance(summary_df, pd.DataFrame) or summary_df.empty
    ):
        return result
    source_ids = _explicit_unit_ids(source_df)
    if source_ids is None:
        return result
    count_columns = {'actual': [], 'expected': [], 'on_time': []}
    suffixes = {'actual': 'actual reports', 'expected': 'expected reports', 'on_time': 'actual reports on time'}
    rate_columns = []
    for column in source_df.columns:
        name = ' '.join(str(column).casefold().split())
        for family, suffix in suffixes.items():
            if name == suffix or name.endswith(' ' + suffix):
                count_columns[family].append(column)
        if name == 'reporting rate' or name.endswith(' reporting rate'):
            rate_columns.append(column)
    summary_ids = _explicit_unit_ids(summary_df)
    if summary_ids is None:
        if 'Organisation unit' not in source_df or 'Organisation unit' not in summary_df:
            return result
        pairs = pd.DataFrame({
            'id': source_ids, 'name': source_df['Organisation unit'].map(_display_name),
        }).drop_duplicates()
        pairs = pairs.loc[pairs['id'].ne('') & pairs['name'].ne('')]
        one_id = pairs.groupby('name')['id'].nunique().eq(1)
        one_name = pairs.groupby('id')['name'].nunique().eq(1)
        pairs = pairs.loc[pairs['name'].map(one_id) & pairs['id'].map(one_name)]
        bridge = dict(zip(pairs['name'], pairs['id']))
        summary_ids = summary_df['Organisation unit'].map(_display_name).map(bridge).fillna('')
    for position, identifier in enumerate(target_ids):
        raw = source_df.loc[source_ids.eq(identifier)]
        matching = summary_df.loc[summary_ids.eq(identifier)]
        if raw.empty or len(matching) != 1:
            continue
        summary = matching.iloc[0]
        actual_known, _ = _count_evidence(raw, count_columns['actual'])
        expected_known, expected_positive = _count_evidence(raw, count_columns['expected'])
        on_time_known, _ = _count_evidence(raw, count_columns['on_time'])
        for column, summary_column, available in [
            ('Actual_Reports', 'Reports_Actual', actual_known),
            ('Complétude_Globale', 'Complétude_Globale (%)', actual_known and expected_known and expected_positive),
            ('Promptitude_Globale', 'Promptitude_Globale (%)', on_time_known and expected_known and expected_positive),
        ]:
            value = _number(summary.get(summary_column))
            if available and pd.notna(value) and value >= 0:
                result.at[position, column] = value
        if rate_columns and expected_known and expected_positive:
            numeric_rates = raw[rate_columns].apply(pd.to_numeric, errors='coerce')
            valid_rates = numeric_rates.map(lambda value: pd.notna(value) and math.isfinite(float(value)) and value >= 0)
            known_datasets = int(valid_rates.any(axis=0).sum())
            count = _number(summary.get('Nombre des data set complétude >/=95%'))
            if (
                known_datasets > 0 and pd.notna(count) and count >= 0
                and count.is_integer() and count <= known_datasets
            ):
                result.at[position, 'Nombre_Datasets_95'] = count
    return result


def metrics_from_reporting_source(source_df, targets):
    """Calcule le reporting d'une cible, en excluant ses seuls datasets absents.

    Un dataset entièrement vide pour une DPS peut être actif dans une autre.
    Son exclusion se fait donc par ID, avant les calculs pondérés. Une cellule
    manquante dans un dataset actif reste indisponible ; elle ne devient pas
    zéro. Les familles de comptes sont évaluées séparément.
    """
    empty = metrics_from_dashboard_summary(pd.DataFrame(), source_df, targets)
    if not isinstance(source_df, pd.DataFrame) or source_df.empty:
        return empty
    source_ids = _explicit_unit_ids(source_df)
    if source_ids is None:
        return empty
    suffixes = [
        'actual reports on time', 'reporting rate on time',
        'actual reports', 'expected reports', 'reporting rate',
    ]
    datasets = {}
    for column in source_df.columns:
        name = ' '.join(str(column).casefold().split())
        for suffix in suffixes:
            if name == suffix or name.endswith(' ' + suffix):
                key = name[:-len(suffix)].strip(' -_')
                datasets.setdefault(key, []).append(column)
                break
    records = []
    for target in empty['Zone_ID']:
        rows = source_df.loc[source_ids.eq(target)].copy()
        # Ces calculs utilisent les cellules ; les diagnostics restent sur la source.
        rows.attrs = {}
        if rows.empty:
            records.append(empty.loc[empty['Zone_ID'].eq(target)].iloc[0].to_dict())
            continue
        inactive_columns = []
        for columns in datasets.values():
            provided = rows[columns].notna() & rows[columns].astype(str).apply(lambda col: col.str.strip().ne(''))
            if not provided.any().any():
                inactive_columns.extend(columns)
        rows = rows.drop(columns=inactive_columns)
        rows['Organisation unit ID'] = target
        selected = targets.loc[targets['Zone_ID'].map(clean_ou_id).eq(target)]
        metrics = aggregate_validation_metrics(rows, selected)
        summary = metrics.rename(columns={
            'Zone_ID': 'Organisation unit ID',
            'Complétude_Globale': 'Complétude_Globale (%)',
            'Promptitude_Globale': 'Promptitude_Globale (%)',
            'Actual_Reports': 'Reports_Actual',
            'Nombre_Datasets_95': 'Nombre des data set complétude >/=95%',
        })
        verified = metrics_from_dashboard_summary(summary, rows, selected)
        records.extend(verified.to_dict('records'))
    return pd.DataFrame(records, columns=REPORTING_METRIC_COLUMNS)


def build_quality_rows(metrics, zones, validation_data=None):
    """Prépare les données numériques du périmètre, sans réutiliser une autre ZS.

    Les indicateurs de déclaration viennent toujours des données courantes.
    Seules les colonnes de violations d'une analyse correspondante sont jointes.
    Une analyse absente ou incomplète laisse ces colonnes indisponibles.
    """
    labels = labels_for_units(zones, 'Zone_ID', 'Zone')
    metric_map = {
        clean_ou_id(row.get('Zone_ID')): row for row in metrics.to_dict('records')
    } if metrics is not None else {}
    validation_map = {
        clean_ou_id(row.get('Zone_ID')): row for row in validation_data.to_dict('records')
    } if validation_data is not None else {}
    profiles = {
        clean_ou_id(row.get('Zone_ID')): row.get('Profil_Bareme', 'ZS')
        for row in zones.to_dict('records')
    } if zones is not None else {}
    rows = []
    for zone_id, label in labels.items():
        source = metric_map.get(zone_id, {})
        validation = validation_map.get(zone_id, {})
        completeness = _number(source.get('Complétude_Globale'))
        timeliness = _number(source.get('Promptitude_Globale'))
        row = {
            'Zone_ID': zone_id, 'Zone de santé': label,
            'Complétude_Globale': completeness,
            'Promptitude_Globale': timeliness,
            'Actual_Reports': _number(source.get('Actual_Reports')),
            'Score_Qualite': round((completeness + timeliness) / 2, 1),
            'Nombre_Datasets_95': _number(source.get('Nombre_Datasets_95')),
            'Profil_Bareme': profiles.get(zone_id, 'ZS'),
            'Erreur_Validation': validation.get('Erreur_Validation') or '',
            'Erreur_Corrections': validation.get('Erreur_Corrections') or '',
            'Controle_Corrections_Effectue': bool(validation.get('Controle_Corrections_Effectue', False)),
        }
        for column in ['Violations_Actuelles', 'Violations_Mois_Precedent', 'Violations_Corrigees']:
            row[column] = _number(validation.get(column))
        rows.append(row)
    return pd.DataFrame(rows)

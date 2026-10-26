"""Hiérarchie géographique DHIS2 pour les filtres pays, province et zone de santé."""

from collections import Counter

import pandas as pd

from reporting_metrics import clean_ou_id


CATALOG_COLUMNS = [
    "Organisation unit", "Organisation unit ID", "Pays", "Pays_ID",
    "Province", "Province_ID", "Zone", "Zone_ID", "level", "is_zone_candidate",
]


def _name(value):
    return "" if value is None or pd.isna(value) else str(value).strip()


def _level(value):
    if isinstance(value, bool):
        return None
    try:
        level = int(value)
        return level if level >= 1 and float(value) == level else None
    except (TypeError, ValueError, OverflowError):
        return None


def _units(units):
    """Ignore les enregistrements sans identifiant et déduplique les IDs."""
    unique = {}
    for unit in units or []:
        if isinstance(unit, dict):
            unit_id = clean_ou_id(unit.get("id"))
            if unit_id:
                unique[unit_id] = unit
    return unique


def resolve_hierarchy(units):
    """Résout les ancêtres par niveau exact, sans inventer les niveaux absents.

    Les niveaux 1, 2 et 3 représentent respectivement pays, province et zone.
    Une unité de niveau 2 ne reçoit donc jamais son propre ID comme zone.
    """
    hierarchy = {}
    for unit_id, unit in _units(units).items():
        own_level = _level(unit.get("level"))
        levels = {1: {}, 2: {}, 3: {}}
        for ancestor in unit.get("ancestors") or []:
            if not isinstance(ancestor, dict):
                continue
            ancestor_id = clean_ou_id(ancestor.get("id"))
            ancestor_level = _level(ancestor.get("level"))
            if not ancestor_id or ancestor_level not in levels:
                continue
            if own_level is not None and ancestor_level >= own_level:
                continue
            levels[ancestor_level][ancestor_id] = ancestor
        if own_level in levels:
            levels[own_level] = {unit_id: unit}

        result = {"level": own_level}
        for level, prefix in [(1, "country"), (2, "province"), (3, "zone")]:
            candidates = levels[level]
            selected = next(iter(candidates.values())) if len(candidates) == 1 else {}
            result[f"{prefix}_name"] = _name(selected.get("displayName", selected.get("name")))
            result[f"{prefix}_id"] = clean_ou_id(selected.get("id"))
        hierarchy[unit_id] = result
    return hierarchy


def build_geography_catalog(units):
    """Crée le catalogue des unités avec leurs véritables ancêtres géographiques."""
    unique = _units(units)
    hierarchy = resolve_hierarchy(list(unique.values()))
    records = []
    for unit_id, unit in unique.items():
        metadata = hierarchy[unit_id]
        level = metadata["level"]
        records.append({
            "Organisation unit": _name(unit.get("displayName", unit.get("name"))),
            "Organisation unit ID": unit_id,
            "Pays": metadata["country_name"], "Pays_ID": metadata["country_id"],
            "Province": metadata["province_name"], "Province_ID": metadata["province_id"],
            "Zone": metadata["zone_name"], "Zone_ID": metadata["zone_id"],
            "level": level,
            "is_zone_candidate": level is not None and level >= 3 and bool(metadata["zone_id"]),
        })
    frame = pd.DataFrame(records, columns=CATALOG_COLUMNS)
    frame["level"] = pd.array(frame["level"], dtype="Int64")
    frame["is_zone_candidate"] = frame["is_zone_candidate"].astype(bool)
    return frame


def zone_catalog(catalog, country_id=None, province_id=None, zone_id=None):
    """Sélectionne uniquement les véritables zones de santé de niveau 3.

    ``None`` signifie l'absence de filtre. Un ID explicitement vide ou invalide
    ne sélectionne aucune unité, afin de ne pas élargir un périmètre par erreur.
    """
    if catalog is None or catalog.empty:
        return pd.DataFrame(columns=CATALOG_COLUMNS)
    result = catalog.copy()
    if "level" not in result or "Zone_ID" not in result:
        return result.iloc[0:0].copy()
    result = result.loc[pd.to_numeric(result["level"], errors="coerce").eq(3)]
    result = result.loc[result["Zone_ID"].map(clean_ou_id).ne("")]
    for column, value in [("Pays_ID", country_id), ("Province_ID", province_id), ("Zone_ID", zone_id)]:
        if value is None:
            continue
        selected_id = clean_ou_id(value)
        if not selected_id or column not in result:
            return result.iloc[0:0].copy()
        result = result.loc[result[column].map(clean_ou_id).eq(selected_id)]
    sort_columns = [column for column in ["Zone", "Zone_ID"] if column in result]
    return result.drop_duplicates("Zone_ID").sort_values(sort_columns).reset_index(drop=True)


def labels_for_units(frame, id_column, name_column):
    """Associe chaque ID à un libellé lisible, distinct en cas d'homonymes."""
    if frame is None or frame.empty or id_column not in frame:
        return {}
    records = {}
    for _, row in frame.iterrows():
        unit_id = clean_ou_id(row.get(id_column))
        if unit_id and unit_id not in records:
            records[unit_id] = (_name(row.get(name_column)) or unit_id, _name(row.get("Province")))
    normalized = lambda value: " ".join(value.casefold().split())
    name_counts = Counter(normalized(name) for name, _ in records.values())
    candidates = {}
    for unit_id, (name, province) in records.items():
        duplicate = name_counts[normalized(name)] > 1
        candidates[unit_id] = (
            f"{name} ({province})"
            if duplicate and province and normalized(province) != normalized(name)
            else f"{name} ({unit_id})" if duplicate else name
        )
    label_counts = Counter(normalized(label) for label in candidates.values())
    return {
        unit_id: f"{records[unit_id][0]} ({unit_id})" if label_counts[normalized(label)] > 1 else label
        for unit_id, label in candidates.items()
    }

"""Géographie et état des corrections pour les détails M et M-1."""

import pandas as pd

from geography import labels_for_units
from reporting_metrics import clean_ou_id


GEOGRAPHY_COLUMNS = {
    1: ('Pays_ID', '_country_id', 'Pays'),
    2: ('Province_ID', '_province_id', 'Province'),
    3: ('Zone_ID', '_health_zone_id', 'Zone de santé'),
}
CORRECTION_COL = 'État de correction'


def _object(value):
    return value if isinstance(value, dict) else {}


def _unit_id(row):
    return clean_ou_id(row.get('organisationUnitId')) or clean_ou_id(
        _object(row.get('organisationUnit')).get('id')
    ) or clean_ou_id(row.get('_ou_id'))


def _level(value):
    try:
        return int(value) if float(value) == int(value) else None
    except (ValueError, TypeError, OverflowError):
        return None


def enrich_detail_geography(details, raw, catalog, unit_metadata=None):
    """Rattache les détails par identifiants et ancêtres, jamais par leurs noms.

    Un chemin n'est accepté que s'il finit par l'ID de l'unité concernée. Les
    chemins et ancêtres contradictoires laissent le niveau concerné, ainsi que
    ses descendants, inconnus. ``_zone_id`` garde la cible d'analyse (DPS ou ZS) ;
    ``_health_zone_id`` représente exclusivement une véritable zone de santé.
    Les lignes, leur ordre et leur index restent ceux du tableau de détails.
    """
    result = details.copy()
    source = catalog if isinstance(catalog, pd.DataFrame) else pd.DataFrame()
    records_by_id = {}
    if 'Organisation unit ID' in source:
        for record in source.to_dict('records'):
            identifier = clean_ou_id(record.get('Organisation unit ID'))
            if identifier:
                records_by_id.setdefault(identifier, []).append(record)
    labels = {
        1: labels_for_units(source, 'Pays_ID', 'Pays'),
        2: labels_for_units(source, 'Province_ID', 'Province'),
        3: labels_for_units(source, 'Zone_ID', 'Zone'),
    }
    raw_by_id = {}
    if isinstance(raw, pd.DataFrame):
        for row in raw.to_dict('records'):
            uid = _unit_id(row)
            if uid:
                raw_by_id.setdefault(uid, []).append(row)
    metadata = unit_metadata or {}
    geography = []
    for detail in result.to_dict('records'):
        uid = clean_ou_id(detail.get('_ou_id'))
        candidates = {level: set() for level in GEOGRAPHY_COLUMNS}
        metadata_names = {level: {} for level in GEOGRAPHY_COLUMNS}

        def add_catalog(identifier):
            for record in records_by_id.get(identifier, []):
                for level, (catalog_column, _, _) in GEOGRAPHY_COLUMNS.items():
                    value = clean_ou_id(record.get(catalog_column))
                    if value:
                        candidates[level].add(value)

        def add_unit(unit):
            identifier = clean_ou_id(unit.get('id'))
            level = _level(unit.get('level'))
            if identifier and level in candidates:
                # Un ID dont le niveau est connu dans le catalogue doit être
                # cohérent avec le niveau transmis par les métadonnées.
                known_levels = {
                    _level(record.get('level')) for record in records_by_id.get(identifier, [])
                }
                if known_levels and level not in known_levels:
                    return
                candidates[level].add(identifier)
                name = unit.get('displayName') or unit.get('name')
                if isinstance(name, str) and name.strip():
                    metadata_names[level][identifier] = name.strip()
                add_catalog(identifier)

        if uid:
            add_catalog(uid)
        meta = _object(metadata.get(uid))
        rows = raw_by_id.get(uid, [])
        unit_sources = [meta]
        paths = [meta.get('path')]
        for row in rows:
            embedded = _object(row.get('organisationUnit'))
            unit_sources.append(embedded)
            paths.extend([row.get('organisationUnitPath'), embedded.get('path')])
            # La cible attribuée par l'analyse permet de connaître la DPS,
            # voire la ZS, même quand DHIS2 omet le chemin d'un centre.
            add_catalog(clean_ou_id(row.get('_zone_id')))
            # Une analyse pays peut regrouper par DPS des requêtes effectuées
            # séparément sur les ZS. L'orchestrateur conserve cette provenance.
            add_catalog(clean_ou_id(row.get('_request_unit_id')))
        for path in paths:
            if not isinstance(path, str):
                continue
            parts = [part for part in path.split('/') if part]
            if uid and parts and parts[-1] == uid:
                for identifier in parts:
                    add_catalog(identifier)
        for unit in unit_sources:
            if clean_ou_id(unit.get('id')) == uid:
                add_unit(unit)
            for ancestor in unit.get('ancestors') or []:
                if isinstance(ancestor, dict):
                    add_unit(ancestor)
        row_geography = {}
        ambiguous = False
        for level, (_, internal, display) in GEOGRAPHY_COLUMNS.items():
            ambiguous = ambiguous or len(candidates[level]) > 1
            identifier = next(iter(candidates[level])) if len(candidates[level]) == 1 and not ambiguous else ''
            row_geography[internal] = identifier
            row_geography[display] = (
                labels[level].get(identifier) or metadata_names[level].get(identifier) or identifier
            ) if identifier else 'Non renseignée'
        geography.append(row_geography)
    for _, internal, display in GEOGRAPHY_COLUMNS.values():
        result[internal] = pd.Series([row[internal] for row in geography], index=result.index, dtype=object)
        result[display] = pd.Series([row[display] for row in geography], index=result.index, dtype=object)
    result.attrs = dict(details.attrs)
    return result


def filter_detail_geography(details, country_id=None, province_id=None, zone_id=None):
    """Filtre les IDs exacts ; sans filtre, conserve aussi les lignes inconnues."""
    result = details.copy()
    for column, value in [
        ('_country_id', country_id), ('_province_id', province_id), ('_health_zone_id', zone_id),
    ]:
        if value is not None:
            identifier = clean_ou_id(value)
            if not identifier or column not in result:
                return result.iloc[:0].copy()
            result = result.loc[result[column].map(clean_ou_id).eq(identifier)].copy()
    return result


def mark_previous_corrections(previous_assigned, current_assigned, current_errors, signature_function):
    """Ajoute l'état de chaque ancienne ligne sans en supprimer les doublons.

    L'erreur M d'une cible empêche uniquement sa comparaison. Une réponse M vide
    mais réussie permet de marquer ses anciennes violations comme corrigées.
    Une ligne sans signature vérifiable reste explicitement non vérifiable.
    """
    result = previous_assigned.copy()
    current = current_assigned if isinstance(current_assigned, pd.DataFrame) else pd.DataFrame()
    errors = current_errors or {}
    signatures_by_target = {}
    for row in current.to_dict('records'):
        target = clean_ou_id(row.get('_request_unit_id')) or clean_ou_id(row.get('_zone_id'))
        if target:
            signatures_by_target.setdefault(target, set()).update(signature_function(pd.DataFrame([row])))
    statuses = []
    for row in result.to_dict('records'):
        target = clean_ou_id(row.get('_request_unit_id')) or clean_ou_id(row.get('_zone_id'))
        signatures = signature_function(pd.DataFrame([row]))
        if not target or errors.get(target) or len(signatures) != 1:
            statuses.append('Non vérifiable')
        else:
            statuses.append('Non corrigée' if signatures[0] in signatures_by_target.get(target, set()) else 'Corrigée')
    result[CORRECTION_COL] = pd.Series(statuses, index=result.index, dtype=object)
    result.attrs = dict(previous_assigned.attrs)
    return result

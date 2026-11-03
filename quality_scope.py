"""Cibles de catégorisation et attribution des lignes DHIS2 à leur niveau réel."""

import pandas as pd

from geography import labels_for_units, zone_catalog
from reporting_metrics import clean_ou_id
from validation_scope import assign_validation_zones


def quality_target_catalog(catalog, country_id, province_id=None, zone_id=None, target_level=None):
    """Affiche les DPS du pays, les ZS d'une DPS ou la ZS sélectionnée.

    Les noms Zone_ID/Zone servent de contrat interne avec les calculs existants.
    Le niveau et le profil distinguent toujours une véritable DPS d'une ZS.
    """
    if catalog is None or catalog.empty:
        return pd.DataFrame(columns=['Zone_ID', 'Zone', 'Profil_Bareme', 'level'])
    country = clean_ou_id(country_id)
    if not country or 'Pays_ID' not in catalog or 'level' not in catalog:
        return catalog.iloc[:0].assign(Profil_Bareme=pd.Series(dtype=object))
    if target_level not in (None, 2, 3):
        raise ValueError("Le niveau des scores doit être DPS (2) ou zone de santé (3).")
    if province_id is not None or zone_id is not None or target_level == 3:
        result = zone_catalog(catalog, country_id=country, province_id=province_id, zone_id=zone_id)
        result['Profil_Bareme'] = 'ZS'
        return result
    result = catalog.loc[
        catalog['Pays_ID'].map(clean_ou_id).eq(country)
        & pd.to_numeric(catalog['level'], errors='coerce').eq(2)
    ].copy()
    if 'Organisation unit ID' not in result:
        return result.iloc[:0].assign(Profil_Bareme=pd.Series(dtype=object))
    result['Zone_ID'] = result['Organisation unit ID'].map(clean_ou_id)
    result['Zone'] = result['Organisation unit'] if 'Organisation unit' in result else result['Zone_ID']
    result['Profil_Bareme'] = 'DPS'
    return result.loc[result['Zone_ID'].ne('')].drop_duplicates('Zone_ID').sort_values(
        ['Zone', 'Zone_ID'], kind='stable'
    ).reset_index(drop=True)


def assign_quality_units(results, targets, unit_metadata=None, parent_id=None):
    """Rattache chaque ligne à une cible sans changer valeurs ni multiplicité.

    Pour les ZS, conserve le comportement existant. Pour les DPS, un ancêtre ZS
    est normal : seuls les ancêtres de niveau DPS peuvent contredire la cible.
    Les lignes inconnues restent visibles et signalées plutôt qu'inventées.
    """
    levels = pd.to_numeric(targets.get('level', pd.Series(3, index=targets.index)), errors='coerce')
    if levels.eq(3).all():
        return assign_validation_zones(results, targets, unit_metadata, parent_id)
    frame = results.copy()
    labels = labels_for_units(targets, 'Zone_ID', 'Zone')
    level_by_id = dict(zip(targets['Zone_ID'].map(clean_ou_id), levels))
    target_levels = set(levels.dropna().astype(int))
    metadata = unit_metadata or {}
    assignments = []
    for row in frame.to_dict('records'):
        embedded = row.get('organisationUnit')
        embedded = embedded if isinstance(embedded, dict) else {}
        uid = clean_ou_id(row.get('organisationUnitId')) or clean_ou_id(embedded.get('id'))
        unit = metadata.get(uid)
        unit = unit if isinstance(unit, dict) else {}
        candidates = {uid} if uid in labels else set()
        known_path = False
        conflicting_ancestor = False
        for path in [row.get('organisationUnitPath'), embedded.get('path'), unit.get('path')]:
            if not isinstance(path, str) or not path.strip():
                continue
            segments = [part for part in path.split('/') if part]
            if not segments or segments[-1] != uid:
                continue
            known_path = True
            candidates.update(part for part in segments if part in labels)
        for source in [embedded, unit]:
            for ancestor in source.get('ancestors') or []:
                if not isinstance(ancestor, dict):
                    continue
                ancestor_id = clean_ou_id(ancestor.get('id'))
                level = pd.to_numeric(ancestor.get('level'), errors='coerce')
                if ancestor_id in labels:
                    if pd.isna(level) or level == level_by_id.get(ancestor_id):
                        candidates.add(ancestor_id)
                    else:
                        conflicting_ancestor = True
                elif ancestor_id and pd.notna(level) and level in target_levels:
                    conflicting_ancestor = True
        if conflicting_ancestor:
            candidates.clear()
        if not candidates and not known_path and not conflicting_ancestor and parent_id in labels:
            candidates.add(parent_id)
        assignments.append(next(iter(candidates)) if len(candidates) == 1 else '')
    frame['_zone_id'] = pd.Series(assignments, index=frame.index, dtype=object)
    frame['_zone_name'] = frame['_zone_id'].map(labels).fillna('Non renseignée')
    unresolved = int(frame['_zone_id'].eq('').sum())
    frame.attrs = dict(results.attrs)
    frame.attrs['unassigned_rows'] = unresolved
    warnings = [] if not unresolved else [
        f"{unresolved} ligne(s) DHIS2 sans rattachement vérifiable à une unité d'analyse. "
        "Elles restent dans le détail ; les compteurs concernés sont incomplets."
    ]
    return frame, warnings

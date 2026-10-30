"""Ventilation des lignes DHIS2 par zone, sans modifier les résultats d'analyse."""

import pandas as pd

from geography import labels_for_units
from reporting_metrics import clean_ou_id


def _object(value):
    return value if isinstance(value, dict) else {}


def _unit_id(row):
    return clean_ou_id(row.get('organisationUnitId')) or clean_ou_id(
        _object(row.get('organisationUnit')).get('id')
    )


def assign_validation_zones(results, zones, unit_metadata=None, parent_id=None):
    """Ajoute la zone de chaque ligne reçue, dans l'ordre et sans déduplication.

    L'API ne renvoie pas systématiquement le chemin des centres. Les métadonnées
    complètent ce rattachement. Une analyse d'une seule ZS permet aussi de
    rattacher ses centres au parent demandé. Une ligne non rattachable reste
    affichable ; elle rend seulement les comptes par ZS incomplets.
    """
    frame = results.copy()
    labels = labels_for_units(zones, 'Zone_ID', 'Zone')
    metadata = unit_metadata or {}
    assignments = []
    for row in frame.to_dict('records'):
        uid = _unit_id(row)
        embedded = _object(row.get('organisationUnit'))
        unit = _object(metadata.get(uid))
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
                ancestor_id = clean_ou_id(_object(ancestor).get('id'))
                if ancestor_id in labels:
                    candidates.add(ancestor_id)
                elif ancestor_id and str(_object(ancestor).get('level')) == '3':
                    conflicting_ancestor = True
        if conflicting_ancestor:
            candidates.clear()
        if not candidates and not known_path and not conflicting_ancestor and parent_id in labels:
            candidates.add(parent_id)
        assignments.append(next(iter(candidates)) if len(candidates) == 1 else '')
    frame['_zone_id'] = pd.Series(assignments, index=frame.index, dtype=object)
    frame['_zone_name'] = frame['_zone_id'].map(labels).fillna('Non renseignée')
    unresolved = int(frame['_zone_id'].eq('').sum())
    warnings = []
    if unresolved:
        warnings.append(
            f"{unresolved} ligne(s) DHIS2 sans rattachement vérifiable à une zone de santé. "
            "Elles restent dans le détail ; les compteurs par zone sont incomplets."
        )
    frame.attrs = dict(results.attrs)
    frame.attrs['unassigned_rows'] = unresolved
    return frame, warnings

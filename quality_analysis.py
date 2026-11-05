"""Prépare séparément M et M-1 pour les deux actions de validation du tableau."""

import pandas as pd

from geography import labels_for_units, zone_catalog
from quality_scope import assign_quality_units
from reporting_metrics import clean_ou_id


COMPARISON_COLUMNS = [
    'Zone_ID', 'Violations_Actuelles', 'Violations_Mois_Precedent',
    'Violations_Corrigees', 'Erreur_Validation', 'Erreur_Corrections',
    'Controle_Corrections_Effectue',
]


def quality_request_targets(score_targets, organisation_catalog):
    """Interroge les ZS plutôt qu'une province entière, même pour un score DPS.

    Une DPS sans ZS accessible ne crée aucune requête de substitution. Le lien
    explicite ``_score_target_id`` permettra de remonter les résultats reçus
    vers la DPS sans utiliser une correspondance approximative de noms.
    """
    columns = list(dict.fromkeys(
        list(score_targets.columns) + list(organisation_catalog.columns)
        + ['Zone_ID', 'Zone', 'level', 'Profil_Bareme', '_score_target_id']
    ))
    requests = []
    for target in score_targets.to_dict('records'):
        identifier = clean_ou_id(target.get('Zone_ID'))
        if not identifier:
            continue
        level = pd.to_numeric(target.get('level'), errors='coerce')
        is_dps = level == 2 or str(target.get('Profil_Bareme', '')).strip().upper() == 'DPS'
        if is_dps:
            country_id = clean_ou_id(target.get('Pays_ID')) or None
            leaves = zone_catalog(organisation_catalog, country_id=country_id, province_id=identifier)
            for leaf in leaves.to_dict('records'):
                requests.append({**leaf, 'Profil_Bareme': 'ZS', '_score_target_id': identifier})
        elif level == 3 or str(target.get('Profil_Bareme', '')).strip().upper() == 'ZS':
            requests.append({**target, 'level': 3, 'Profil_Bareme': 'ZS', '_score_target_id': identifier})
    result = pd.DataFrame(requests, columns=columns)
    if result.empty:
        return result
    return result.drop_duplicates(['Zone_ID', '_score_target_id']).reset_index(drop=True)


def rollup_period_to_scores(assigned, errors_by_request, request_targets, score_targets):
    """Remonte une période analysée par ZS aux cibles ZS ou DPS du tableau.

    Les lignes et leurs valeurs restent intactes ; seuls les champs internes de
    regroupement changent. Une ZS indisponible invalide le décompte de sa DPS,
    sans propager cet échec à une autre province. ``_request_unit_id`` conserve
    la véritable ZS interrogée, notamment pour les filtres et les diagnostics.
    """
    score_labels = labels_for_units(score_targets, 'Zone_ID', 'Zone')
    request_labels = labels_for_units(request_targets, 'Zone_ID', 'Zone')
    mapping = {}
    for target in request_targets.to_dict('records'):
        request_id = clean_ou_id(target.get('Zone_ID'))
        score_id = clean_ou_id(target.get('_score_target_id'))
        if request_id and score_id in score_labels:
            mapping.setdefault(request_id, set()).add(score_id)
    frame = assigned.copy() if isinstance(assigned, pd.DataFrame) else pd.DataFrame()
    source_column = '_analysis_target_id' if '_analysis_target_id' in frame else '_zone_id'
    if source_column not in frame:
        frame[source_column] = pd.Series('', index=frame.index, dtype=object)
    frame['_request_unit_id'] = frame[source_column].map(clean_ou_id)
    single_mapping = {request_id: next(iter(scores)) for request_id, scores in mapping.items() if len(scores) == 1}
    frame['_analysis_target_id'] = frame['_request_unit_id'].map(single_mapping).fillna('')
    frame['_zone_id'] = frame['_analysis_target_id']
    frame['_zone_name'] = frame['_zone_id'].map(score_labels).fillna('Non renseignée')
    errors_by_score = {identifier: [] for identifier in score_labels}
    for score_id in score_labels:
        leaf_ids = [request_id for request_id, scores in mapping.items() if score_id in scores]
        if not leaf_ids:
            errors_by_score[score_id].append("Aucune zone de santé accessible pour cette unité d’analyse")
        for request_id in leaf_ids:
            if len(mapping[request_id]) != 1:
                errors_by_score[score_id].append("Rattachement de la zone de santé à plusieurs unités d’analyse")
            for issue in _issues((errors_by_request or {}).get(request_id)):
                errors_by_score[score_id].append(
                    f"{request_labels.get(request_id, request_id)} : {issue}" if request_id != score_id else issue
                )
    unresolved = int(frame['_analysis_target_id'].eq('').sum())
    if unresolved:
        for issues in errors_by_score.values():
            issues.append(f"{unresolved} ligne(s) sans rattachement à une unité du tableau")
    warnings = []
    for score_id, issues in errors_by_score.items():
        errors_by_score[score_id] = _issues(issues)
        warnings.extend(f"{score_labels[score_id]} : {issue}" for issue in errors_by_score[score_id])
    frame.attrs = {}
    return frame, errors_by_score, warnings


def _issues(value):
    if isinstance(value, str):
        return [value] if value.strip() else []
    if isinstance(value, (list, tuple)):
        return list(dict.fromkeys(str(issue) for issue in value if str(issue).strip()))
    return []


def analyse_quality_period(batch, targets, unit_metadata=None):
    """Rattache chaque requête à sa cible, en isolant ses erreurs.

    La concaténation ne porte pas les attributs d'échec d'une seule cible aux
    autres unités. Les lignes non rattachables restent visibles et seules les
    statistiques de la cible interrogée deviennent indisponibles.
    """
    labels = labels_for_units(targets, 'Zone_ID', 'Zone')
    frames = []
    errors_by_id = {}
    warnings = []
    for identifier, name in labels.items():
        source = batch.get('frames_by_id', {}).get(identifier)
        issues = _issues(batch.get('errors_by_id', {}).get(identifier))
        if not isinstance(source, pd.DataFrame):
            source = pd.DataFrame()
            issues.append("Analyse DHIS2 non reçue pour cette unité")
        elif batch.get('succeeded_by_id', {}).get(identifier) is False and not issues:
            issues.append("Analyse DHIS2 indisponible pour cette unité")
        selected = targets.loc[targets['Zone_ID'].map(clean_ou_id).eq(identifier)]
        assigned, scope_issues = assign_quality_units(source, selected, unit_metadata, identifier)
        issues = _issues(issues + scope_issues)
        errors_by_id[identifier] = issues
        warnings.extend(f"{name} : {issue}" for issue in issues)
        assigned['_analysis_target_id'] = identifier
        assigned.attrs = {}
        frames.append(assigned)
    result = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(
        columns=['_zone_id', '_zone_name', '_analysis_target_id']
    )
    result.attrs = {}
    return result, errors_by_id, warnings


def _identifier(value):
    if isinstance(value, dict):
        return next((identifier for key in ('id', 'uid', 'displayName', 'name')
                     if (identifier := _identifier(value.get(key)))), '')
    if value is None or pd.isna(value):
        return ''
    return clean_ou_id(value)


def _default_signatures(frame):
    """Identité hors période et valeurs ; accepte champs plats et imbriqués."""
    signatures = []
    for row in frame.to_dict('records'):
        identifiers = [
            _identifier(row.get(key + 'Id')) or _identifier(row.get(key))
            for key in ('organisationUnit', 'validationRule', 'attributeOptionCombo', 'categoryOptionCombo', 'dataSet')
        ]
        if identifiers[0] or identifiers[1]:
            signatures.append(tuple(identifiers))
    return signatures


def _target_rows(frame, identifier):
    if not isinstance(frame, pd.DataFrame):
        return pd.DataFrame()
    key = '_analysis_target_id' if '_analysis_target_id' in frame else '_zone_id'
    if key not in frame:
        return frame.iloc[:0].copy()
    return frame.loc[frame[key].map(clean_ou_id).eq(identifier)]


def build_validation_comparison_rows(targets, current_assigned, current_errors,
                                     previous_assigned=None, previous_errors=None,
                                     signature_function=None):
    """Compte M, puis ajoute M-1 et les corrections sans remplacer M.

    L'absence d'exécution de M-1 reste distincte d'une analyse réussie sans
    violation. Chaque ligne reçue compte une fois ; seules les signatures
    stables servent au contrôle des corrections.
    """
    signature_function = signature_function or _default_signatures
    checked = previous_assigned is not None
    previous_errors = previous_errors or {}
    current_errors = current_errors or {}
    rows = []
    for identifier in labels_for_units(targets, 'Zone_ID', 'Zone'):
        now_issues = _issues(current_errors.get(identifier))
        before_issues = _issues(previous_errors.get(identifier))
        now_valid = isinstance(current_assigned, pd.DataFrame) and not now_issues
        before_valid = checked and not before_issues
        now = _target_rows(current_assigned, identifier)
        before = _target_rows(previous_assigned, identifier)
        corrected = float('nan')
        if now_valid and before_valid:
            current_signatures = set(signature_function(now))
            corrected = sum(signature not in current_signatures for signature in signature_function(before))
        rows.append({
            'Zone_ID': identifier,
            'Violations_Actuelles': len(now) if now_valid else float('nan'),
            'Violations_Mois_Precedent': len(before) if before_valid else float('nan'),
            'Violations_Corrigees': corrected,
            'Erreur_Validation': '; '.join(now_issues),
            'Erreur_Corrections': '; '.join(before_issues),
            'Controle_Corrections_Effectue': checked,
        })
    return pd.DataFrame(rows, columns=COMPARISON_COLUMNS)

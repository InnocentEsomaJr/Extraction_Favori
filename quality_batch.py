"""Exécution isolée des analyses de catégorisation, sans modifier l'API DHIS2."""

import pandas as pd

from geography import labels_for_units
from reporting_metrics import clean_ou_id


def run_quality_batch(targets, periods, group_id, fetch, cache_user=None,
                      start_date=None, end_date=None, label='M', on_progress=None):
    """Appelle une fois chaque cible et conserve succès, zéros et erreurs séparés.

    Aucun échec ne déclenche une relance automatique. Une réponse vide réussie
    vaut zéro violation ; une réponse échouée laisse le nombre indisponible.
    Les DataFrames originaux et leurs attributs sont conservés par cible.
    """
    labels = labels_for_units(targets, 'Zone_ID', 'Zone')
    frames, errors, successes, requests, raw, warnings = {}, {}, {}, [], [], []
    for position, (identifier, name) in enumerate(labels.items(), start=1):
        if on_progress is not None:
            on_progress(position, len(labels), identifier, name)
        kwargs = {'cache_user': cache_user}
        if start_date is not None:
            kwargs['start_date'] = start_date
        if end_date is not None:
            kwargs['end_date'] = end_date
        try:
            frame = fetch(identifier, periods, group_id, **kwargs)
            if not isinstance(frame, pd.DataFrame):
                raise ValueError("Réponse d'analyse DHIS2 inexploitable")
            frame = frame.copy()
        except Exception as error:
            frame = pd.DataFrame()
            frame.attrs = {
                'failed_periods': list(periods),
                'errors': {period: [f"Analyse DHIS2 indisponible ({type(error).__name__})"] for period in periods},
            }
        failed = list(frame.attrs.get('failed_periods', []))
        truncated = list(frame.attrs.get('truncated_periods', []))
        issues = []
        for period in dict.fromkeys(failed + truncated):
            reasons = frame.attrs.get('errors', {}).get(period) or [
                'Résultats limités par DHIS2' if period in truncated else 'Analyse DHIS2 indisponible'
            ]
            if isinstance(reasons, str):
                reasons = [reasons]
            issues.append(f"{label} ({period}) : {' ; '.join(map(str, reasons))}")
        frames[identifier] = frame
        errors[identifier] = issues
        successes[identifier] = not bool(failed or truncated)
        warnings.extend(f"{name} : {message}" for message in issues)
        params = frame.attrs.get('request_params', {})
        requests.append({
            'Période': label, 'Méthode': frame.attrs.get('request_method', 'POST'),
            'Unité parent': identifier, 'Unité': name,
            'Début': params.get('startDate', start_date.isoformat() if hasattr(start_date, 'isoformat') else start_date or ''),
            'Fin': params.get('endDate', end_date.isoformat() if hasattr(end_date, 'isoformat') else end_date or ''),
            'Statut': ('Indisponible' if frame.empty else 'Partiel') if issues else 'Terminé',
            'Lignes reçues': float('nan') if issues and frame.empty else len(frame),
            'Durée (s)': frame.attrs.get('elapsed_seconds', float('nan')),
            'Délai de lecture (s)': frame.attrs.get('validation_read_timeout_seconds', float('nan')),
        })
        tagged = frame.copy()
        tagged.attrs = {}
        tagged['_analysis_target_id'] = clean_ou_id(identifier)
        raw.append(tagged)
    return {
        'frames_by_id': frames,
        'raw': pd.concat(raw, ignore_index=True) if raw else pd.DataFrame(),
        'requests': requests, 'errors_by_id': errors,
        'succeeded_by_id': successes, 'warnings': warnings,
        'complete': bool(successes) and all(successes.values()),
    }

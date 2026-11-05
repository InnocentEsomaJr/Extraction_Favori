"""Exécution isolée des analyses de catégorisation, sans modifier l'API DHIS2."""

from copy import deepcopy

import pandas as pd

from geography import labels_for_units
from reporting_metrics import clean_ou_id


def _copy_frame(frame):
    """Copie aussi les objets JSON imbriqués que pandas conserve par référence."""
    copied = frame.copy(deep=True)
    for column in frame.select_dtypes(include='object').columns:
        copied[column] = pd.Series(
            [deepcopy(value) for value in frame[column]], index=frame.index, dtype=object,
        )
    copied.attrs = deepcopy(frame.attrs)
    return copied


def _date_text(value):
    return value.isoformat() if hasattr(value, 'isoformat') else value or ''


def _frame_issues(frame, periods, label):
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
    return issues


def run_quality_batch(targets, periods, group_id, fetch, cache_user=None,
                      start_date=None, end_date=None, label='M', on_progress=None,
                      initial_state=None, on_result=None, max_targets=None):
    """Appelle une fois chaque cible et conserve succès, zéros et erreurs séparés.

    Aucun échec ne déclenche une relance automatique. Une réponse vide réussie
    vaut zéro violation ; une réponse échouée laisse le nombre indisponible.
    Les DataFrames originaux et leurs attributs sont conservés par cible.
    Chaque réponse peut être sauvegardée avec ``on_result``. Une reprise ne
    relance aucune unité déjà traitée, y compris celles ayant échoué. Les
    unités restantes ne sont pas présentées comme des réponses vides réussies.
    ``max_targets`` borne le nombre de nouvelles requêtes de cette exécution.
    """
    if max_targets is not None and (
        isinstance(max_targets, bool) or not isinstance(max_targets, int) or max_targets < 1
    ):
        raise ValueError("Le nombre d’unités à traiter doit être un entier strictement positif.")
    periods = [periods] if isinstance(periods, str) else list(periods)
    labels = labels_for_units(targets, 'Zone_ID', 'Zone')
    context = {
        'periods': periods, 'group_id': group_id, 'label': label,
        'start_date': _date_text(start_date), 'end_date': _date_text(end_date),
    }
    frames, errors, successes = {}, {}, {}
    previous = initial_state if isinstance(initial_state, dict) else {}
    if previous.get('context', context) == context:
        previous_frames = previous.get('frames_by_id', {})
        previous_errors = previous.get('errors_by_id', {})
        previous_successes = previous.get('succeeded_by_id', {})
        processed = previous.get('processed_ids', list(previous_frames))
        for identifier in labels:
            source = previous_frames.get(identifier)
            if identifier not in processed or not isinstance(source, pd.DataFrame):
                continue
            frames[identifier] = _copy_frame(source)
            existing_issues = previous_errors.get(identifier, [])
            if isinstance(existing_issues, str):
                existing_issues = [existing_issues]
            issues = list(dict.fromkeys(list(existing_issues) + _frame_issues(source, periods, label)))
            if previous_successes.get(identifier) is False and not issues:
                issues.append("Analyse DHIS2 indisponible pour cette unité")
            errors[identifier] = issues
            successes[identifier] = not bool(issues)

    def snapshot():
        copied_frames, requests, raw, warnings = {}, [], [], []
        processed_ids = [identifier for identifier in labels if identifier in frames]
        pending_ids = [identifier for identifier in labels if identifier not in frames]
        for identifier in processed_ids:
            name = labels[identifier]
            frame = _copy_frame(frames[identifier])
            copied_frames[identifier] = frame
            issues = errors[identifier]
            warnings.extend(f"{name} : {message}" for message in issues)
            params = frame.attrs.get('request_params', {})
            requests.append({
                'Période': label, 'Méthode': frame.attrs.get('request_method', 'POST'),
                'Unité parent': identifier, 'Unité': name,
                'Début': params.get('startDate', _date_text(start_date)),
                'Fin': params.get('endDate', _date_text(end_date)),
                'Statut': 'Interrompu' if frame.attrs.get('interrupted') else (('Indisponible' if frame.empty else 'Partiel') if issues else 'Terminé'),
                'Lignes reçues': float('nan') if issues and frame.empty else len(frame),
                'Durée (s)': frame.attrs.get('elapsed_seconds', float('nan')),
                'Délai de lecture (s)': frame.attrs.get('validation_read_timeout_seconds', float('nan')),
            })
            tagged = _copy_frame(frame)
            tagged.attrs = {}
            tagged['_analysis_target_id'] = clean_ou_id(identifier)
            raw.append(tagged)
        return {
            'frames_by_id': copied_frames,
            'raw': pd.concat(raw, ignore_index=True) if raw else pd.DataFrame(),
            'requests': requests,
            'errors_by_id': deepcopy({identifier: errors[identifier] for identifier in processed_ids}),
            'succeeded_by_id': {identifier: successes[identifier] for identifier in processed_ids},
            'warnings': warnings, 'processed_ids': processed_ids, 'pending_ids': pending_ids,
            'finished': not pending_ids,
            'complete': bool(processed_ids) and not pending_ids and all(successes.values()),
            'context': deepcopy(context),
        }

    attempted = 0
    for position, (identifier, name) in enumerate(labels.items(), start=1):
        if identifier in frames:
            continue
        if max_targets is not None and attempted >= max_targets:
            break
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
            frame = _copy_frame(frame)
        except Exception as error:
            frame = pd.DataFrame()
            frame.attrs = {
                'failed_periods': list(periods),
                'errors': {period: [f"Analyse DHIS2 indisponible ({type(error).__name__})"] for period in periods},
            }
        issues = _frame_issues(frame, periods, label)
        frames[identifier] = frame
        errors[identifier] = issues
        successes[identifier] = not bool(issues)
        attempted += 1
        if on_result is not None:
            on_result(snapshot())
    return snapshot()

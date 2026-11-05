"""État reprenable d'une analyse de validation, indépendant de Streamlit."""

from copy import deepcopy

import pandas as pd

from geography import labels_for_units
from quality_batch import run_quality_batch
from reporting_metrics import clean_ou_id
from validation_filters import monthly_bounds


INTERRUPTED_MESSAGE = "Analyse interrompue : réponse non reçue"


def _copy(value):
    """Copie aussi les cellules JSON que pandas laisse normalement partagées."""
    if isinstance(value, pd.DataFrame):
        result = value.copy(deep=True)
        for column in value.select_dtypes(include='object').columns:
            result[column] = pd.Series(
                [deepcopy(item) for item in value[column]], index=value.index, dtype=object,
            )
        result.attrs = deepcopy(value.attrs)
        return result
    if isinstance(value, dict):
        return {key: _copy(item) for key, item in value.items()}
    if isinstance(value, list):
        return [_copy(item) for item in value]
    if isinstance(value, tuple):
        return tuple(_copy(item) for item in value)
    return deepcopy(value)


def _period_list(periods):
    return [periods] if isinstance(periods, str) else list(periods or [])


def _date_text(value):
    return value.isoformat() if hasattr(value, 'isoformat') else value or ''


def create_validation_job(context, action, cache_user, current_batch=None):
    """Prépare M, M-1 ou les deux périodes, sans lancer de requête."""
    if action not in {'M', 'M-1', 'both'}:
        raise ValueError("L'action doit être M, M-1 ou both.")
    if action == 'M-1' and not isinstance(current_batch, dict):
        raise ValueError("Le contrôle M-1 exige un état de la période M.")
    return {
        'context': _copy(context), 'action': action,
        'phase': 'M-1' if action == 'M-1' else 'M',
        'batches': {'M': _copy(current_batch), 'M-1': None},
        'status': 'running', 'inflight_id': None, 'auto_continue': False,
        'cache_user': _copy(cache_user),
    }


def advance_validation_job(job, targets, periods, reference_periods, group_id, fetch,
                           start_date=None, end_date=None, on_checkpoint=None,
                           on_progress=None):
    """Traite au plus une nouvelle zone et publie les états avant les rendus UI.

    Une interruption traverse cette fonction sans être absorbée. Le dernier
    checkpoint permet alors d'identifier la requête dont la réponse est inconnue.
    """
    result = _copy(job)
    if result.get('status') != 'running':
        return result
    phase = result.get('phase')
    if phase not in {'M', 'M-1'}:
        raise ValueError("La période active de l'analyse est invalide.")
    result['auto_continue'] = False
    result['group_id'] = group_id
    result['start_date'] = _date_text(start_date)
    result['end_date'] = _date_text(end_date)
    active_periods = _period_list(reference_periods if phase == 'M-1' else periods)

    def checkpoint():
        if on_checkpoint is not None:
            on_checkpoint(_copy(result))

    def progress(position, total, identifier, name):
        result['inflight_id'] = identifier
        checkpoint()
        if on_progress is not None:
            on_progress(position, total, identifier, name)

    def received(snapshot):
        result['batches'][phase] = _copy(snapshot)
        result['inflight_id'] = None
        checkpoint()

    kwargs = {}
    if phase == 'M':
        kwargs.update(start_date=start_date, end_date=end_date)
    snapshot = run_quality_batch(
        targets, active_periods, group_id, fetch,
        cache_user=result.get('cache_user'), label=phase,
        initial_state=result.get('batches', {}).get(phase),
        max_targets=1, on_progress=progress, on_result=received, **kwargs,
    )
    result['batches'][phase] = _copy(snapshot)
    result['inflight_id'] = None
    if snapshot.get('finished'):
        if phase == 'M' and result['action'] == 'both':
            result['phase'] = 'M-1'
        else:
            result['status'] = 'complete'
    checkpoint()
    return _copy(result)


def _recovered_snapshot(job, targets, periods, reference_periods):
    phase = job['phase']
    active_periods = _period_list(reference_periods if phase == 'M-1' else periods)
    labels = labels_for_units(targets, 'Zone_ID', 'Zone')
    existing = job.get('batches', {}).get(phase) or {}
    frames = {
        identifier: _copy(frame)
        for identifier, frame in existing.get('frames_by_id', {}).items()
        if identifier in labels and isinstance(frame, pd.DataFrame)
    }
    context = existing.get('context') or {
        'periods': active_periods, 'group_id': job.get('group_id'), 'label': phase,
        'start_date': job.get('start_date', '') if phase == 'M' else '',
        'end_date': job.get('end_date', '') if phase == 'M' else '',
    }
    inflight = clean_ou_id(job.get('inflight_id'))
    if inflight in labels and inflight not in frames:
        start = context.get('start_date', '')
        end = context.get('end_date', '')
        if active_periods and (not start or not end):
            first, last = monthly_bounds(active_periods)
            start, end = first.isoformat(), last.isoformat()
        frame = pd.DataFrame()
        frame.attrs = {
            'failed_periods': active_periods,
            'errors': {period: [INTERRUPTED_MESSAGE] for period in active_periods},
            'request_method': 'POST',
            'request_params': {'startDate': start, 'endDate': end},
            'interrupted': True,
        }
        frames[inflight] = frame

    processed = [identifier for identifier in labels if identifier in frames]
    pending = [identifier for identifier in labels if identifier not in frames]
    requests_by_id = {
        row.get('Unité parent'): _copy(row) for row in existing.get('requests', [])
        if row.get('Unité parent') in frames
    }
    errors, successes, warnings, requests, raw = {}, {}, [], [], []
    for identifier in processed:
        frame = frames[identifier]
        issues = existing.get('errors_by_id', {}).get(identifier, [])
        issues = [issues] if isinstance(issues, str) else list(issues)
        for period in dict.fromkeys(
            list(frame.attrs.get('failed_periods', [])) + list(frame.attrs.get('truncated_periods', []))
        ):
            reasons = frame.attrs.get('errors', {}).get(period) or ['Analyse DHIS2 indisponible']
            reasons = [reasons] if isinstance(reasons, str) else reasons
            issue = f"{phase} ({period}) : {' ; '.join(map(str, reasons))}"
            if issue not in issues:
                issues.append(issue)
        if existing.get('succeeded_by_id', {}).get(identifier) is False and not issues:
            issues.append("Analyse DHIS2 indisponible pour cette unité")
        errors[identifier] = issues
        successes[identifier] = not bool(issues)
        warnings.extend(f"{labels[identifier]} : {issue}" for issue in issues)
        request = requests_by_id.get(identifier)
        if request is None:
            params = frame.attrs.get('request_params', {})
            request = {
                'Période': phase, 'Méthode': frame.attrs.get('request_method', 'POST'),
                'Unité parent': identifier, 'Unité': labels[identifier],
                'Début': params.get('startDate', context.get('start_date', '')),
                'Fin': params.get('endDate', context.get('end_date', '')),
                'Statut': 'Interrompu' if frame.attrs.get('interrupted') else ('Indisponible' if issues else 'Terminé'),
                'Lignes reçues': float('nan') if issues and frame.empty else len(frame),
                'Durée (s)': frame.attrs.get('elapsed_seconds', float('nan')),
                'Délai de lecture (s)': frame.attrs.get('validation_read_timeout_seconds', float('nan')),
            }
        requests.append(request)
        tagged = _copy(frame)
        tagged.attrs = {}
        tagged['_analysis_target_id'] = identifier
        raw.append(tagged)
    return {
        'frames_by_id': frames,
        'raw': pd.concat(raw, ignore_index=True) if raw else pd.DataFrame(),
        'requests': requests, 'errors_by_id': errors, 'succeeded_by_id': successes,
        'warnings': warnings, 'processed_ids': processed, 'pending_ids': pending,
        'finished': not pending,
        'complete': bool(processed) and not pending and all(successes.values()),
        'context': _copy(context),
    }


def recover_validation_job(job, targets, periods, reference_periods):
    """Arrête une reprise ambiguë et ne renvoie jamais le POST en cours."""
    result = _copy(job)
    interrupted = bool(result.get('inflight_id')) or (
        result.get('status') == 'running' and not result.get('auto_continue', False)
    )
    if interrupted:
        phase = result['phase']
        result['batches'][phase] = _recovered_snapshot(result, targets, periods, reference_periods)
        result['status'] = 'interrupted'
        result['inflight_id'] = None
        result['auto_continue'] = False
    return result


def resume_validation_job(job):
    """Autorise les unités restantes ; les requêtes déjà tentées sont conservées."""
    result = _copy(job)
    if result.get('status') == 'interrupted':
        result['status'] = 'running'
        result['inflight_id'] = None
        result['auto_continue'] = False
    return result

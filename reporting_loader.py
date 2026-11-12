"""Chargement Analytics par lots, sans modifier les périodes ni agréger les taux."""

import re

import requests
from urllib3.exceptions import TimeoutError as UrllibTimeoutError

from reporting_metrics import clean_ou_id
from scoped_analytics import METRIC_LABELS, analytics_to_reporting_frame


def format_reporting_error(error):
    """Décrit aussi les timeouts encapsulés, sans exposer URL ou exception brute."""
    pending, visited = [error], set()
    timed_out = False
    while pending:
        current = pending.pop()
        if id(current) in visited:
            continue
        visited.add(id(current))
        if isinstance(current, (requests.exceptions.Timeout, UrllibTimeoutError)):
            timed_out = True
        for nested in [getattr(current, '__cause__', None), getattr(current, '__context__', None),
                       getattr(current, 'reason', None), *getattr(current, 'args', ())]:
            if isinstance(nested, BaseException):
                pending.append(nested)
    if timed_out:
        return "Délai de lecture Analytics dépassé sur DHIS2 pour ce lot."
    response = getattr(error, 'response', None)
    status = getattr(response, 'status_code', None)
    if status in (401, 403):
        return f"Accès aux indicateurs refusé par DHIS2 (HTTP {status})."
    if status in (400, 414):
        return "DHIS2 n'a pas accepté cette sélection d'indicateurs. Les autres réponses reçues sont conservées."
    if isinstance(status, int):
        return f"Les indicateurs DHIS2 sont indisponibles pour ce lot (HTTP {status})."
    if isinstance(error, requests.exceptions.SSLError):
        return "La connexion sécurisée aux indicateurs DHIS2 a échoué."
    if isinstance(error, requests.exceptions.ConnectionError):
        return "Connexion au serveur Analytics DHIS2 interrompue pour ce lot."
    if isinstance(error, ValueError):
        return "La réponse Analytics DHIS2 de ce lot est inexploitable."
    return "Impossible de charger les indicateurs DHIS2 de ce lot."


def _positive_limit(value, default):
    try:
        parsed = int(value)
        return parsed if parsed > 0 else default
    except (ValueError, TypeError, OverflowError):
        return default


def _item_families(identifiers):
    families = {}
    for identifier in identifiers:
        dataset, separator, metric = identifier.rpartition('.')
        key = ('dataset', dataset) if separator and metric in METRIC_LABELS else ('item', identifier)
        families.setdefault(key, []).append(identifier)
    return list(families.values())


def _item_batches(identifiers, limit):
    batches, current = [], []
    for family in _item_families(identifiers):
        if current and len(current) + len(family) > limit:
            batches.append(current)
            current = []
        current.extend(family)
    if current:
        batches.append(current)
    return batches


def fetch_reporting_chunks(client, params, labels, unit_ids, units_per_request=3,
                           items_per_request=15, on_progress=None):
    """Lit des lots disjoints d'unités et de datasets, puis pivote les valeurs.

    Les cinq métriques d'un dataset restent ensemble. Les filtres temporels et
    supplémentaires sont transmis à l'identique. Un refus HTTP 400 ou 414 d'un
    lot de plusieurs datasets est isolé par subdivisions disjointes, jusqu'au
    dataset ; ses cinq métriques restent ensemble. Aucun timeout, refus d'accès
    ni erreur serveur ne provoque de nouvelle tentative. ``unit_ids=None`` garde
    une requête initiale unique utilisant les paramètres originaux.
    Une liste explicite vide ne provoque jamais un appel de périmètre élargi.
    """
    query = list(params)
    labels = dict(labels)
    dx_dimensions = [value for key, value in query if key == 'dimension' and value.startswith('dx:')]
    ou_dimensions = [value for key, value in query if key == 'dimension' and value.startswith('ou:')]
    if len(dx_dimensions) != 1 or len(ou_dimensions) != 1:
        raise ValueError("Les dimensions Analytics ou et dx doivent être uniques.")
    identifiers = list(dict.fromkeys(dx_dimensions[0][3:].split(';')))
    if not identifiers or any(not identifier or identifier not in labels for identifier in identifiers):
        raise ValueError("Les éléments Analytics demandés ne sont pas identifiés.")
    if isinstance(unit_ids, (str, bytes)):
        raise ValueError("La sélection Analytics doit être une liste d'identifiants d'unités.")
    explicit_units = None if unit_ids is None else list(dict.fromkeys(clean_ou_id(value) for value in unit_ids))
    if explicit_units is not None and any(not re.fullmatch(r'[A-Za-z][A-Za-z0-9]{10}', value) for value in explicit_units):
        raise ValueError("Une unité Analytics explicitement demandée est invalide.")
    unit_limit = _positive_limit(units_per_request, 3)
    item_limit = _positive_limit(items_per_request, 15)
    if explicit_units is None:
        inferred = ou_dimensions[0][3:].split(';')
        allowed_units = None if any(value.startswith(('LEVEL-', 'USER_')) for value in inferred) else set(inferred)
        chunks = [(None, identifiers, query, allowed_units)]
    else:
        chunks = []
        for start in range(0, len(explicit_units), unit_limit):
            units = explicit_units[start:start + unit_limit]
            for items in _item_batches(identifiers, item_limit):
                chunk_query = [
                    (key, 'ou:' + ';'.join(units)) if key == 'dimension' and value.startswith('ou:') else
                    (key, 'dx:' + ';'.join(items)) if key == 'dimension' and value.startswith('dx:') else (key, value)
                    for key, value in query
                ]
                chunks.append((units, items, chunk_query, set(units)))
    rows, metadata, warnings, failed, successes, diagnostics = [], {}, [], {}, set(), []
    completed = 0
    successful_leaves = 0
    stop_error = None

    def fetch_chunk(units, items, chunk_query, allowed_units):
        nonlocal stop_error, successful_leaves
        try:
            if stop_error is not None:
                raise stop_error
            response = client.get('analytics.json', params=chunk_query)
            # Validation avant fusion : un lot invalide ne contamine pas ceux reçus.
            analytics_to_reporting_frame(response, {item: labels[item] for item in items})
            headers = [header.get('name', header.get('column')) for header in response.get('headers', [])]
            normalized = []
            for values in response['rows']:
                record = dict(zip(headers, values))
                if record['dx'] not in items or (allowed_units is not None and record['ou'] not in allowed_units):
                    raise ValueError("Une valeur Analytics ne correspond pas au lot demandé.")
                normalized.append([record['ou'], record['dx'], record['value']])
            rows.extend(normalized)
            metadata.update((response.get('metaData') or {}).get('items') or {})
            successful_leaves += 1
            successes.update(units or [row[0] for row in normalized])
            diagnostics.append({'Unités': ';'.join(units or ()), 'Métriques': len(items), 'Statut': 'Terminé'})
            return True
        except Exception as error:
            status = getattr(getattr(error, 'response', None), 'status_code', None)
            families = _item_families(items)
            if stop_error is None and status in (400, 414) and len(families) > 1:
                diagnostics.append({'Unités': ';'.join(units or ()), 'Métriques': len(items), 'Statut': 'Fractionné'})
                middle = len(families) // 2
                outcomes = []
                for family_group in (families[:middle], families[middle:]):
                    subset = [item for family in family_group for item in family]
                    subset_query = [
                        (key, 'dx:' + ';'.join(subset))
                        if key == 'dimension' and value.startswith('dx:') else (key, value)
                        for key, value in chunk_query
                    ]
                    outcomes.append(fetch_chunk(units, subset, subset_query, allowed_units))
                return all(outcomes)
            message = format_reporting_error(error)
            targets = units or list(allowed_units or ())
            for identifier in targets:
                failed.setdefault(identifier, []).extend(item for item in items if item not in failed.get(identifier, []))
            if stop_error is None or error is not stop_error:
                warnings.append((f"{', '.join(targets)} : " if targets else '') + message)
            diagnostics.append({'Unités': ';'.join(targets), 'Métriques': len(items), 'Statut': 'Indisponible'})
            if status in (401, 403):
                stop_error = error
            return False

    for position, (units, items, chunk_query, allowed_units) in enumerate(chunks, start=1):
        if on_progress is not None:
            on_progress(position, len(chunks), tuple(units or ()), tuple(items))
        if fetch_chunk(units, items, chunk_query, allowed_units):
            completed += 1
    result = analytics_to_reporting_frame({
        'headers': [{'name': name} for name in ['ou', 'dx', 'value']],
        'rows': rows, 'metaData': {'items': metadata},
    }, labels)
    result.attrs.update({
        'reporting_warnings': warnings,
        'failed_units': list(failed), 'failed_metrics_by_unit': failed,
        'successful_units': [identifier for identifier in (explicit_units or sorted(successes)) if identifier in successes],
        'reporting_status': 'complete' if completed == len(chunks) and chunks else ('partial' if successful_leaves else 'unavailable'),
        'reporting_requests': diagnostics,
    })
    return result

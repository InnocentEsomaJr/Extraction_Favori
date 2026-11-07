"""Sources de reporting qualité mensuelles et hebdomadaires, séparées par période."""

from datetime import date
import re

import pandas as pd

from reporting_loader import fetch_reporting_chunks
from reporting_metrics import clean_ou_id
from scoped_analytics import METRIC_LABELS


def _period_type(value):
    return str(value or '').strip().casefold()


def _weekly_periods(values):
    if isinstance(values, (str, bytes)):
        raise ValueError('Les semaines de surveillance doivent être une liste de périodes DHIS2.')
    result = []
    for value in values or []:
        match = re.fullmatch(r'(\d{4})W(\d{1,2})', str(value).strip())
        if not match:
            raise ValueError('Une semaine de surveillance ne possède pas un identifiant DHIS2 valide.')
        year, week = map(int, match.groups())
        try:
            date.fromisocalendar(year, week, 1)
        except ValueError as error:
            raise ValueError('Une semaine de surveillance est invalide.') from error
        identifier = f'{year}W{week}'
        if identifier not in result:
            result.append(identifier)
    return tuple(result)


def split_quality_reporting_query(params, labels, datasets, weekly_periods):
    """Sépare les métriques sélectionnées sans changer leurs filtres annexes.

    Les libellés proviennent de la requête complète, déjà désambiguïsée par UID.
    Seul le filtre ``pe`` des datasets hebdomadaires est remplacé. Les datasets
    non mensuels et non hebdomadaires sont refusés, jamais traités en mensuel.
    """
    query, labels = list(params), dict(labels)
    dx = [value[3:] for key, value in query if key == 'dimension' and value.startswith('dx:')]
    periods = [value[3:] for key, value in query if key == 'filter' and value.startswith('pe:')]
    if len(dx) != 1 or len(periods) != 1 or not periods[0]:
        raise ValueError('Les dimensions de données et la période qualité doivent être uniques.')
    if any(key == 'dimension' and value.startswith('pe:') for key, value in query):
        raise ValueError('La période qualité doit être un filtre agrégé, pas une dimension.')
    identifiers = list(dict.fromkeys(dx[0].split(';')))
    if not identifiers or any(not identifier or identifier not in labels for identifier in identifiers):
        raise ValueError('Les métriques de la sélection qualité ne sont pas identifiées.')
    if len({labels[identifier] for identifier in identifiers}) != len(identifiers):
        raise ValueError('Les libellés des métriques qualité doivent être uniques.')
    catalog = {clean_ou_id(record.get('id')): record for record in datasets or [] if isinstance(record, dict)}
    grouped = {'monthly': [], 'weekly': []}
    dataset_ids = {'monthly': [], 'weekly': []}
    for identifier in identifiers:
        dataset_id, separator, metric = identifier.rpartition('.')
        if not separator or metric not in METRIC_LABELS or dataset_id not in catalog:
            raise ValueError('Un dataset de la sélection qualité est absent des métadonnées disponibles.')
        record = catalog[dataset_id]
        period_type = _period_type(record.get('periodType'))
        if record.get('is_weekly') is True:
            if period_type not in ('', 'weekly'):
                raise ValueError('Le type de période du dataset hebdomadaire est incohérent.')
            kind = 'weekly'
        elif period_type in ('', 'monthly'):
            kind = 'monthly'
        else:
            raise ValueError('Le score qualité prend en charge uniquement les datasets mensuels et hebdomadaires.')
        grouped[kind].append(identifier)
        if dataset_id not in dataset_ids[kind]:
            dataset_ids[kind].append(dataset_id)
    weeks = _weekly_periods(weekly_periods) if grouped['weekly'] else ()
    if grouped['weekly'] and not weeks:
        raise ValueError('Sélectionnez les semaines de surveillance avant de charger le score qualité.')
    groups = []
    for kind, items in grouped.items():
        if not items:
            continue
        selected_periods = weeks if kind == 'weekly' else tuple(periods[0].split(';'))
        group_query = [
            (key, 'dx:' + ';'.join(items)) if key == 'dimension' and value.startswith('dx:') else
            (key, 'pe:' + ';'.join(selected_periods)) if key == 'filter' and value.startswith('pe:') else
            (key, value)
            for key, value in query
        ]
        groups.append({
            'kind': kind, 'params': group_query,
            'labels': {identifier: labels[identifier] for identifier in items},
            'periods': selected_periods, 'dataset_ids': tuple(dataset_ids[kind]),
        })
    return groups


def fetch_quality_reporting(client, params, labels, datasets, unit_ids, weekly_periods, on_progress=None):
    """Charge puis joint les sources par UID, sans moyenner les pourcentages.

    ``on_progress(kind, position, total, units, items)`` identifie la source.
    Les erreurs restent atomiques par unité dans ``failed_units`` : la source
    hebdomadaire en échec ou entièrement absente ne devient pas un dataset
    inactif et ne permet pas un score mensuel incomplet présenté comme complet.
    """
    groups = split_quality_reporting_query(params, labels, datasets, weekly_periods)
    selected_labels = {
        identifier: label for group in groups for identifier, label in group['labels'].items()
    }
    if isinstance(unit_ids, (str, bytes)):
        raise ValueError("La sélection qualité doit être une liste d'identifiants d'unités.")
    identifiers = None if unit_ids is None else list(dict.fromkeys(clean_ou_id(value) for value in unit_ids))
    sources = []
    for group in groups:
        progress = None
        if on_progress is not None:
            progress = lambda position, total, units, items, kind=group['kind']: on_progress(
                kind, position, total, units, items,
            )
        frame = fetch_reporting_chunks(
            client, group['params'], group['labels'], identifiers,
            units_per_request=1, items_per_request=max(1, len(group['labels'])), on_progress=progress,
        )
        sources.append((group, frame))

    if identifiers is None:
        identifiers = list(dict.fromkeys(
            identifier for _, frame in sources for identifier in frame['Organisation unit ID'].map(clean_ou_id)
        ))
        # Un appel direct sans LEVEL garde son unité même si DHIS2 ne renvoie aucune ligne.
        for key, value in params:
            if key == 'dimension' and value.startswith('ou:'):
                for identifier in value[3:].split(';'):
                    if re.fullmatch(r'[A-Za-z][A-Za-z0-9]{10}', identifier) and identifier not in identifiers:
                        identifiers.append(identifier)
    result = pd.DataFrame(index=pd.Index(identifiers, name='Organisation unit ID'))
    names = pd.Series(index=result.index, dtype=object)
    failed, warnings, diagnostics = {}, [], []
    received = set(identifiers)
    all_complete = bool(groups)
    for group, frame in sources:
        attrs = frame.attrs
        all_complete = all_complete and attrs.get('reporting_status') == 'complete'
        received.intersection_update(attrs.get('successful_units', []))
        for identifier, items in attrs.get('failed_metrics_by_unit', {}).items():
            failed.setdefault(identifier, set()).update(items)
        for identifier in attrs.get('failed_units', []):
            failed.setdefault(identifier, set()).update(group['labels'])
        warnings.extend(f"{group['kind']} : {message}" for message in attrs.get('reporting_warnings', []))
        diagnostics.extend({
            **entry, 'Source': group['kind'], 'Périodes': ';'.join(group['periods']),
        } for entry in attrs.get('reporting_requests', []))
        aligned = frame.set_index('Organisation unit ID').reindex(identifiers)
        incoming_names = aligned['Organisation unit']
        # Un vrai nom prime sur un UID de secours, quelle que soit la source.
        known_name = incoming_names.notna() & incoming_names.ne(incoming_names.index.to_series())
        names = names.where(names.notna(), incoming_names)
        names = names.where(~known_name | (names.notna() & names.ne(names.index.to_series())), incoming_names)
        for identifier, column in group['labels'].items():
            if column in result:
                raise ValueError('Deux datasets qualité partagent un libellé de métrique ambigu.')
            result[column] = pd.to_numeric(aligned[column], errors='coerce')
        if group['kind'] == 'weekly':
            for unit_id in identifiers:
                for dataset_id in group['dataset_ids']:
                    items = [item for item in group['labels'] if item.rpartition('.')[0] == dataset_id]
                    columns = [group['labels'][item] for item in items]
                    if aligned.loc[unit_id, columns].notna().any():
                        continue
                    if unit_id not in failed or not set(items).issubset(failed[unit_id]):
                        warnings.append(f"{unit_id} : surveillance hebdomadaire indisponible pour {dataset_id}.")
                        diagnostics.append({
                            'Unités': unit_id, 'Métriques': len(items), 'Statut': 'Indisponible',
                            'Source': 'weekly', 'Périodes': ';'.join(group['periods']),
                        })
                    failed.setdefault(unit_id, set()).update(items)
    result.insert(0, 'Organisation unit', names.fillna(names.index.to_series()))
    result = result.reset_index()
    values_available = bool(result[list(selected_labels.values())].notna().any().any()) if len(result) else False
    result.attrs.update({
        'reporting_columns': list(selected_labels.values()),
        'weekly_reporting_columns': [column for group in groups if group['kind'] == 'weekly' for column in group['labels'].values()],
        'monthly_reporting_columns': [column for group in groups if group['kind'] == 'monthly' for column in group['labels'].values()],
        'reporting_periods': {group['kind']: tuple(group['periods']) for group in groups},
        'failed_units': [identifier for identifier in identifiers if identifier in failed],
        'failed_metrics_by_unit': {
            identifier: [item for item in selected_labels if item in failed[identifier]]
            for identifier in identifiers if identifier in failed
        },
        'successful_units': [identifier for identifier in identifiers if identifier in received and identifier not in failed],
        'reporting_warnings': list(dict.fromkeys(warnings)),
        'reporting_requests': diagnostics,
        'reporting_status': 'complete' if all_complete and not failed else ('partial' if values_available else 'unavailable'),
    })
    return result

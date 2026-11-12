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
    Un groupe hebdomadaire sans semaines porte ``missing_selection=True`` : son
    schéma est conservé, mais aucune requête n'est envoyée pour ce groupe.
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
            'missing_selection': kind == 'weekly' and not selected_periods,
        })
    return groups


def _source_item_labels(group, frame):
    """Retrouve les colonnes par identifiant, même après enrichissement du nom."""
    resolved = frame.attrs.get('reporting_item_labels', {})
    return {item: resolved.get(item, label) for item, label in group['labels'].items()}


def _combined_item_labels(sources):
    """Conserve une famille de cinq métriques et des noms uniques par dataset."""
    preferred = {}
    for group, frame in sources:
        resolved = _source_item_labels(group, frame)
        for item, requested in group['labels'].items():
            dataset, metric = item.rsplit('.', 1)
            suffix = METRIC_LABELS[metric]
            requested_prefix = requested[:-len(suffix)].strip() if requested.endswith(suffix) else dataset
            actual = resolved[item]
            actual_prefix = actual[:-len(suffix)].strip() if actual.endswith(suffix) else dataset
            if requested_prefix and requested_prefix != dataset:
                candidate = (2, requested_prefix)
            elif actual_prefix and actual_prefix != dataset:
                candidate = (1, actual_prefix)
            else:
                candidate = (0, dataset)
            if dataset not in preferred or candidate[0] > preferred[dataset][0]:
                preferred[dataset] = candidate
    bases = {
        dataset: prefix[:-len(f' ({dataset})')] if prefix.endswith(f' ({dataset})') else prefix
        for dataset, (_, prefix) in preferred.items()
    }
    prefixes = {
        dataset: f'{base} ({dataset})' if list(bases.values()).count(base) > 1 else (
            preferred[dataset][1] if preferred[dataset][0] == 2 else base
        )
        for dataset, base in bases.items()
    }
    # Un nom réel peut déjà contenir le suffixe d'un autre dataset.
    while len(set(prefixes.values())) != len(prefixes):
        duplicate_prefixes = {prefix for prefix in prefixes.values() if list(prefixes.values()).count(prefix) > 1}
        prefixes = {
            dataset: f'{prefix} ({dataset})' if prefix in duplicate_prefixes else prefix
            for dataset, prefix in prefixes.items()
        }
    return {
        item: f'{prefixes[item.rsplit(".", 1)[0]]} {METRIC_LABELS[item.rsplit(".", 1)[1]]}'
        for group, _ in sources for item in group['labels']
    }


def _known_units(sources, params):
    identifiers = list(dict.fromkeys(
        identifier for _, frame in sources for identifier in frame['Organisation unit ID'].map(clean_ou_id)
    ))
    for key, value in params:
        if key == 'dimension' and value.startswith('ou:'):
            for identifier in value[3:].split(';'):
                if re.fullmatch(r'[A-Za-z][A-Za-z0-9]{10}', identifier) and identifier not in identifiers:
                    identifiers.append(identifier)
    return identifiers


def _missing_week_frame(group, identifiers):
    """Garde le schéma et les erreurs de la source non sélectionnée sans requête."""
    frame = pd.DataFrame({'Organisation unit ID': identifiers})
    frame.insert(0, 'Organisation unit', frame['Organisation unit ID'])
    for column in group['labels'].values():
        frame[column] = float('nan')
    frame.attrs.update({
        'reporting_columns': list(group['labels'].values()),
        'reporting_item_labels': dict(group['labels']),
        'failed_units': list(identifiers),
        'failed_metrics_by_unit': {identifier: list(group['labels']) for identifier in identifiers},
        'successful_units': [], 'reporting_status': 'unavailable',
        'reporting_warnings': ['Aucune semaine de surveillance sélectionnée.'],
        'reporting_requests': [{
            'Unités': ';'.join(identifiers), 'Métriques': len(group['labels']), 'Statut': 'Non sélectionné',
        }],
    })
    return frame


def fetch_quality_reporting(client, params, labels, datasets, unit_ids, weekly_periods, on_progress=None):
    """Charge puis joint les sources par UID, sans moyenner les pourcentages.

    ``on_progress(kind, position, total, units, items)`` identifie la source.
    Les erreurs restent atomiques par unité dans ``failed_units`` : la source
    hebdomadaire en échec ou entièrement absente ne devient pas un dataset
    inactif et ne permet pas un score mensuel incomplet présenté comme complet.
    Les libellés peuvent être enrichis par DHIS2 ; la jointure des métriques se
    fait exclusivement par leur identifiant, avant le nommage final des colonnes.
    """
    groups = split_quality_reporting_query(params, labels, datasets, weekly_periods)
    if isinstance(unit_ids, (str, bytes)):
        raise ValueError("La sélection qualité doit être une liste d'identifiants d'unités.")
    identifiers = None if unit_ids is None else list(dict.fromkeys(clean_ou_id(value) for value in unit_ids))
    sources = []
    for group in groups:
        if group['missing_selection']:
            targets = identifiers if identifiers is not None else _known_units(sources, params)
            sources.append((group, _missing_week_frame(group, targets)))
            continue
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

    selected_labels = _combined_item_labels(sources)

    if identifiers is None:
        identifiers = _known_units(sources, params)
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
        source_labels = _source_item_labels(group, frame)
        for identifier in group['labels']:
            column = selected_labels[identifier]
            source_column = source_labels[identifier]
            if column in result:
                raise ValueError('Deux datasets qualité partagent un libellé de métrique ambigu.')
            if source_column not in aligned:
                raise ValueError("Une métrique de reporting n'a pas de colonne correspondant à son identifiant.")
            result[column] = pd.to_numeric(aligned[source_column], errors='coerce')
        if group['kind'] == 'weekly':
            for unit_id in identifiers:
                for dataset_id in group['dataset_ids']:
                    items = [item for item in group['labels'] if item.rpartition('.')[0] == dataset_id]
                    columns = [source_labels[item] for item in items]
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
        'reporting_item_labels': dict(selected_labels),
        'weekly_reporting_columns': [selected_labels[item] for group in groups if group['kind'] == 'weekly' for item in group['labels']],
        'monthly_reporting_columns': [selected_labels[item] for group in groups if group['kind'] == 'monthly' for item in group['labels']],
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

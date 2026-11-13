"""Catalogue DHIS2 et sélection explicite des datasets du score qualité."""

from copy import deepcopy
import re

from scoped_analytics import METRIC_LABELS, _data_item, _metric_parts


_UID = re.compile(r"[A-Za-z][A-Za-z0-9]{10}")


def _valid_uid(value):
    return isinstance(value, str) and _UID.fullmatch(value) is not None


def _dataset_name(dataset):
    for field in ('displayName', 'name'):
        value = dataset.get(field)
        if isinstance(value, str) and value.strip():
            return value.strip()
    return ''


def build_dataset_catalog(response):
    """Normalise tous les datasets accessibles, sans sélection par leur nom.

    La réponse doit contenir une liste ``dataSets``. Les UID sont vérifiés et
    les doublons compatibles sont fusionnés ; des métadonnées contradictoires
    pour un même UID sont refusées. Un nom absent utilise l'UID comme libellé.
    Des datasets homonymes restent distincts par leur UID.
    """
    if not isinstance(response, dict) or not isinstance(response.get('dataSets'), list):
        raise ValueError('La réponse du catalogue DHIS2 doit contenir une liste dataSets.')
    entries = {}
    for dataset in response['dataSets']:
        if not isinstance(dataset, dict) or not _valid_uid(dataset.get('id')):
            raise ValueError('Un dataset du catalogue ne possède pas un UID DHIS2 valide.')
        identifier = dataset['id']
        name = _dataset_name(dataset)
        period_type = dataset.get('periodType')
        if period_type is not None and not isinstance(period_type, str):
            raise ValueError('Le type de période d’un dataset du catalogue est invalide.')
        period_type = period_type.strip() if period_type else ''
        previous = entries.setdefault(identifier, {'id': identifier, 'displayName': ''})
        if name and previous['displayName'] and name != previous['displayName']:
            raise ValueError('Un même UID de dataset possède plusieurs noms contradictoires.')
        if period_type and previous.get('periodType') and period_type != previous['periodType']:
            raise ValueError('Un même UID de dataset possède plusieurs types de période contradictoires.')
        if name:
            previous['displayName'] = name
        if period_type:
            previous['periodType'] = period_type
    for identifier, dataset in entries.items():
        dataset['displayName'] = dataset['displayName'] or identifier
    return sorted(entries.values(), key=lambda dataset: (dataset['displayName'].casefold(), dataset['id']))


def _dimensions(visualization):
    for location in ('columns', 'rows', 'filters'):
        dimensions = visualization.get(location)
        if dimensions is None:
            continue
        if not isinstance(dimensions, list):
            raise ValueError('Une disposition des dimensions du favori est invalide.')
        for dimension in dimensions:
            if not isinstance(dimension, dict):
                raise ValueError('Une dimension du favori est invalide.')
            yield dimension


def _is_dx(dimension):
    return str(dimension.get('dimension') or dimension.get('id') or '').strip() == 'dx'


def favorite_dataset_ids(visualization):
    """Retourne les UID de reporting uniques, dans l'ordre déclaré du favori."""
    if not isinstance(visualization, dict):
        raise ValueError('La définition du favori DHIS2 est invalide.')
    items = visualization.get('dataDimensionItems')
    if items is None:
        items = []
    if not isinstance(items, list):
        raise ValueError('Les éléments de données du favori doivent être une liste.')
    items = list(items)
    for dimension in _dimensions(visualization):
        if _is_dx(dimension):
            dimension_items = dimension.get('items')
            if dimension_items is None:
                dimension_items = []
            if not isinstance(dimension_items, list):
                raise ValueError('Les éléments de la dimension dx doivent être une liste.')
            items.extend(dimension_items)
    identifiers = {}
    for item in items:
        identifier, _ = _data_item(item)
        dataset_id, _ = _metric_parts(identifier)
        if dataset_id is not None:
            if not _valid_uid(dataset_id):
                raise ValueError('Un dataset de reporting du favori ne possède pas un UID DHIS2 valide.')
            identifiers.setdefault(dataset_id, None)
    return list(identifiers)


def select_quality_datasets(visualization, selected_ids, catalog):
    """Copie le favori avec exactement cinq mesures par dataset choisi.

    ``catalog`` est la liste canonique retournée par ``build_dataset_catalog``.
    La sélection doit être une liste ou un tuple d'UID connus. Ses doublons
    sont retirés en conservant l'ordre. Toutes les dimensions dx sont remplacées
    et les autres dimensions, filtres et options sont conservés sans mutation.
    """
    if not isinstance(visualization, dict):
        raise ValueError('La définition du favori DHIS2 est invalide.')
    if not isinstance(selected_ids, (list, tuple)):
        raise ValueError('La sélection des datasets doit être une liste d’UID DHIS2.')
    if not selected_ids:
        raise ValueError('Sélectionnez au moins un dataset pour calculer le score qualité.')
    if any(not _valid_uid(identifier) for identifier in selected_ids):
        raise ValueError('Un dataset sélectionné ne possède pas un UID DHIS2 valide.')
    selected = list(dict.fromkeys(selected_ids))
    if not isinstance(catalog, list):
        raise ValueError('Le catalogue des datasets doit être une liste.')
    names = {
        dataset['id']: dataset['displayName']
        for dataset in build_dataset_catalog({'dataSets': catalog})
    }
    if any(identifier not in names for identifier in selected):
        raise ValueError('Un dataset sélectionné est absent du catalogue DHIS2 disponible.')
    result = deepcopy(visualization)
    result['dataDimensionItems'] = [
        {'dataDimensionItemType': 'REPORTING_RATE', 'reportingRate': {
            'id': f'{identifier}.{metric}',
            'metric': metric,
            'dataSet': {'id': identifier, 'displayName': names[identifier]},
        }}
        for identifier in selected for metric in METRIC_LABELS
    ]
    for dimension in _dimensions(result):
        if _is_dx(dimension):
            dimension['items'] = [
                {'id': f'{identifier}.{metric}'}
                for identifier in selected for metric in METRIC_LABELS
            ]
    return result

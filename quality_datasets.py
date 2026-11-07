"""Sélection explicite des datasets du score qualité, rattachés par UID DHIS2."""

from copy import deepcopy
import re
import unicodedata

from scoped_analytics import METRIC_LABELS, _data_item, _metric_parts


_UID = re.compile(r"^[A-Za-z][A-Za-z0-9]{10}$")


def _name(dataset):
    return str(dataset.get('displayName') or dataset.get('name') or '').strip()


def _words(value):
    text = unicodedata.normalize('NFKD', str(value)).casefold()
    text = ''.join(character for character in text if not unicodedata.combining(character))
    return set(re.findall(r'[a-z]+', text))


def _normalized_name(value):
    text = unicodedata.normalize('NFKD', str(value)).casefold()
    text = ''.join(character for character in text if not unicodedata.combining(character))
    return ' '.join(re.findall(r'[a-z]+|[0-9]+', text))


def _visualization_items(visualization):
    if not isinstance(visualization, dict):
        raise ValueError('La définition du favori DHIS2 est invalide.')
    items = list(visualization.get('dataDimensionItems') or [])
    for location in ('columns', 'rows', 'filters'):
        for dimension in visualization.get(location) or []:
            if isinstance(dimension, dict) and (dimension.get('dimension') or dimension.get('id')) == 'dx':
                items.extend(dimension.get('items') or [])
    return items


def _period_type(value):
    if isinstance(value, dict):
        value = value.get('name') or value.get('id')
    value = str(value or '').strip()
    return {'weekly': 'Weekly', 'monthly': 'Monthly'}.get(value.casefold(), value)


def _favorite_period_types(visualization):
    """Conserve les types fournis sur les datasets imbriqués du favori."""
    result = {}
    for item in _visualization_items(visualization):
        if not isinstance(item, dict):
            continue
        reporting = item.get('reportingRate') or item.get('dataSetReportingRate') or item
        if not isinstance(reporting, dict):
            reporting = {}
        dataset = reporting.get('dataSet') or item.get('dataSet')
        if not isinstance(dataset, dict):
            continue
        identifier = str(dataset.get('id') or dataset.get('uid') or '').strip()
        period_type = _period_type(dataset.get('periodType'))
        if _UID.fullmatch(identifier) and period_type:
            result[identifier] = period_type
    return result


def _reporting_items(visualization):
    result = {}
    for item in _visualization_items(visualization):
        identifier, label = _data_item(item)
        dataset, metric = _metric_parts(identifier)
        if dataset is None:
            continue
        if not _UID.fullmatch(dataset):
            raise ValueError('Un dataset de reporting du favori ne possède pas un UID DHIS2 valide.')
        suffix = METRIC_LABELS[metric]
        name = label[:-len(suffix)].strip() if label.endswith(suffix) else ''
        if dataset not in result or result[dataset] in ('', dataset):
            result[dataset] = name or dataset
    return result


def reporting_dataset_ids(visualization):
    """Retourne les UID uniques de reporting du favori, dans leur ordre déclaré."""
    return tuple(_reporting_items(visualization))


def _catalog_entries(catalog):
    if isinstance(catalog, dict):
        catalog = catalog.get('dataSets', [])
    return [entry for entry in (catalog or []) if isinstance(entry, dict)]


def quality_dataset_selection(visualization, catalog=None, surveillance_name=None):
    """Construit les options et la sélection initiale du tableau qualité.

    ``catalog`` est la liste de métadonnées DHIS2 ou sa réponse ``dataSets``.
    Le favori constitue la sélection initiale, hors service tertiaire ; ce
    dernier est aussi retiré des options évaluables. Si le
    favori ne contient pas déjà la surveillance, un unique dataset de ce nom
    est ajouté. Plusieurs candidats restent à choisir explicitement : aucun
    UID n'est inventé et aucun résultat n'est sélectionné arbitrairement.
    Si ``surveillance_name`` est fourni, seul ce libellé normalisé peut être
    sélectionné par défaut : une autre surveillance ne le remplace jamais.
    """
    favorites = _reporting_items(visualization)
    period_types = _favorite_period_types(visualization)
    names = dict(favorites)
    for entry in _catalog_entries(catalog):
        identifier = str(entry.get('id') or '').strip()
        if not _UID.fullmatch(identifier):
            continue
        name = _name(entry)
        if name:
            names[identifier] = name
        else:
            names.setdefault(identifier, identifier)
        period_type = _period_type(entry.get('periodType'))
        if period_type:
            period_types[identifier] = period_type
    datasets = []
    for identifier, name in names.items():
        words = _words(name)
        period_type = period_types.get(identifier, '')
        is_weekly = (
            period_type.casefold().startswith('weekly') if period_type else
            bool({'esurveillance', 'hebdomadaire'} <= words)
        )
        datasets.append({
            'id': identifier, 'name': name,
            'periodType': period_type, 'is_weekly': is_weekly,
            'in_favorite': identifier in favorites,
            'is_surveillance': bool(words & {'surveillance', 'esurveillance'}),
            'is_tertiary': bool(words & {'tertiaire', 'tertiaires', 'tertiary', 'tertiaries'}),
        })
    by_id = {dataset['id']: dataset for dataset in datasets}
    defaults = [identifier for identifier in favorites if not by_id[identifier]['is_tertiary']]
    surveillance_ids = [dataset['id'] for dataset in datasets if dataset['is_surveillance'] and not dataset['is_tertiary']]
    ambiguous, warnings = [], []
    if surveillance_name is not None:
        expected_name = _normalized_name(surveillance_name)
        if not expected_name:
            raise ValueError('Le nom du dataset de surveillance attendu est vide.')
        matches = [dataset['id'] for dataset in datasets
                   if not dataset['is_tertiary'] and _normalized_name(dataset['name']) == expected_name]
        defaults = [identifier for identifier in defaults
                    if not by_id[identifier]['is_surveillance'] or (len(matches) == 1 and identifier in matches)]
        if len(matches) == 1:
            if matches[0] not in defaults:
                defaults.append(matches[0])
        elif len(matches) > 1:
            ambiguous = matches
            warnings.append(
                f'Plusieurs datasets portent le nom « {surveillance_name} » : sélectionnez son UID exact.'
            )
        else:
            warnings.append(
                f'Le dataset « {surveillance_name} » est absent des métadonnées disponibles ; '
                'aucun autre dataset de surveillance ne le remplace.'
            )
    elif not any(identifier in defaults for identifier in surveillance_ids):
        if len(surveillance_ids) == 1:
            defaults.append(surveillance_ids[0])
        elif len(surveillance_ids) > 1:
            ambiguous = surveillance_ids
            warnings.append('Plusieurs datasets de surveillance sont disponibles : sélectionnez celui à évaluer.')
        else:
            warnings.append('Aucun dataset de surveillance identifiable dans les métadonnées disponibles. Vérifiez la sélection des datasets.')
    unnamed_favorites = [identifier for identifier in favorites if names[identifier] == identifier]
    if unnamed_favorites:
        warnings.append(
            'Le libellé de certains datasets du favori est indisponible ; leur type ne peut pas être vérifié : '
            + ', '.join(unnamed_favorites) + '.'
        )
    duplicate_names = {name for name in names.values() if list(names.values()).count(name) > 1}
    labels = {
        identifier: f'{name} [{identifier}]' if name in duplicate_names else name
        for identifier, name in names.items()
    }
    options = sorted(
        (identifier for identifier in names if not by_id[identifier]['is_tertiary']),
        key=lambda identifier: (labels[identifier].casefold(), identifier),
    )
    return {
        'datasets': datasets, 'options': options, 'labels': labels, 'defaults': defaults,
        'favorite_ids': list(favorites), 'surveillance_ids': surveillance_ids,
        'ambiguous_surveillance_ids': ambiguous, 'warnings': warnings,
    }


def select_quality_datasets(visualization, selected_ids, datasets=None):
    """Copie le favori avec exactement cinq métriques par dataset sélectionné.

    ``datasets`` contient les métadonnées ``datasets`` du sélecteur, une liste
    DHIS2 ou sa réponse ``dataSets``. Les filtres hors dx et leurs sélections
    sont conservés. La liste vide, les UID invalides et le service tertiaire
    sont refusés ; aucun
    dataset retiré du sélecteur ne reste dans les éléments dx du favori.
    """
    favorites = _reporting_items(visualization)
    if isinstance(selected_ids, (str, bytes)):
        raise ValueError('La sélection des datasets doit être une liste d’UID DHIS2.')
    selected = list(dict.fromkeys(str(identifier).strip() for identifier in (selected_ids or [])))
    if not selected:
        raise ValueError('Sélectionnez au moins un dataset pour calculer le score qualité.')
    if any(not _UID.fullmatch(identifier) for identifier in selected):
        raise ValueError('Un dataset sélectionné ne possède pas un UID DHIS2 valide.')
    names = dict(favorites)
    period_types = _favorite_period_types(visualization)
    for entry in _catalog_entries(datasets):
        identifier = str(entry.get('id') or '').strip()
        if _UID.fullmatch(identifier):
            names[identifier] = _name(entry) or names.get(identifier, identifier)
            period_type = _period_type(entry.get('periodType'))
            if period_type:
                period_types[identifier] = period_type
    missing = [identifier for identifier in selected if identifier not in names]
    if missing:
        raise ValueError('Un dataset sélectionné est absent des métadonnées disponibles.')
    if any(_words(names[identifier]) & {'tertiaire', 'tertiaires', 'tertiary', 'tertiaries'} for identifier in selected):
        raise ValueError('Le service tertiaire est exclu de l’évaluation du score qualité.')
    result = deepcopy(visualization)
    result['dataDimensionItems'] = [
        {'dataDimensionItemType': 'REPORTING_RATE', 'reportingRate': {
            'id': f'{identifier}.{metric}', 'metric': metric,
            'dataSet': {'id': identifier, 'displayName': names[identifier],
                        **({'periodType': period_types[identifier]} if identifier in period_types else {})},
        }}
        for identifier in selected for metric in METRIC_LABELS
    ]
    for location in ('columns', 'rows', 'filters'):
        for dimension in result.get(location) or []:
            if isinstance(dimension, dict) and (dimension.get('dimension') or dimension.get('id')) == 'dx':
                dimension['items'] = [
                    {'id': f'{identifier}.{metric}'}
                    for identifier in selected for metric in METRIC_LABELS
                ]
    return result

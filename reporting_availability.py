"""Empêche une requête Analytics échouée de devenir un indicateur numérique."""

import pandas as pd

from reporting_metrics import clean_ou_id
from scoped_analytics import METRIC_LABELS


_ID_COLUMNS = ['Organisation unit ID', 'Zone_ID']
_LABEL_COLUMNS = {'Organisation unit', 'Zone de santé', 'Zone', 'Province', 'Pays'}


def select_reporting_source(frame):
    """Projette les seules métriques certifiées sans lever un refus de reporting.

    Le contrat Analytics identifie les colonnes de reporting et leurs DX. Un
    échec explicitement rattaché à un indicateur générique peut être écarté de
    cette source ; un échec de dataset ou non identifiable reste indisponible.
    Sans ce contrat, la copie conserve le comportement historique prudent.
    Les valeurs reçues, y compris les zéros et cellules absentes, restent telles
    quelles. Les avertissements et diagnostics de chargement sont conservés.
    """
    if not isinstance(frame, pd.DataFrame):
        return pd.DataFrame()
    certified = frame.attrs.get('reporting_columns')
    if not isinstance(certified, (list, tuple, set)):
        return frame.copy()
    certified = {column for column in certified if isinstance(column, str)}
    retained = {*_ID_COLUMNS, *_LABEL_COLUMNS, *certified}
    result = frame.loc[:, [column for column in frame if column in retained]].copy()
    result.attrs['reporting_columns'] = [column for column in result if column in certified]

    labels = frame.attrs.get('reporting_item_labels')
    labels = labels if isinstance(labels, dict) else {}
    per_unit = frame.attrs.get('failed_metrics_by_unit')
    per_unit = per_unit if isinstance(per_unit, dict) else {}
    grouped_failures = {}
    for unit, items in per_unit.items():
        identifier = clean_ou_id(unit)
        if identifier:
            grouped_failures.setdefault(identifier, []).append(items)
    failed = frame.attrs.get('failed_units') or []
    if not isinstance(failed, (list, tuple, set)):
        failed = [failed]
    failed_ids = list(dict.fromkeys(
        [identifier for unit in failed if (identifier := clean_ou_id(unit))]
        + list(grouped_failures)
    ))
    reporting_failures, reporting_failed_units = {}, []
    for identifier in failed_ids:
        groups = grouped_failures.get(identifier)
        unknown = not groups
        relevant = []
        for items in groups or []:
            if isinstance(items, str):
                items = [items]
            if not isinstance(items, (list, tuple, set)) or not items:
                unknown = True
                continue
            for item in items:
                if not isinstance(item, str) or not item:
                    unknown = True
                    continue
                label = labels.get(item)
                _, separator, metric = item.rpartition('.')
                is_reporting = bool(separator and metric in METRIC_LABELS)
                known_label = isinstance(label, str) and bool(label.strip())
                if is_reporting and known_label and label in certified:
                    relevant.append(item)
                elif is_reporting or not known_label or label in certified:
                    unknown = True
        if unknown or relevant:
            reporting_failed_units.append(identifier)
            if relevant:
                reporting_failures[identifier] = list(dict.fromkeys(relevant))
    result.attrs['failed_units'] = reporting_failed_units
    result.attrs['failed_metrics_by_unit'] = reporting_failures
    return result


def _label(value):
    return '' if value is None or pd.isna(value) else str(value).strip()


def _identifiers(frame):
    for column in _ID_COLUMNS:
        if column in frame:
            return frame[column].map(clean_ou_id)
    return None


def _cell_identifiers(frame, source_df):
    """Privilégie les UID ; un pont de noms doit être strictement bijectif."""
    identifiers = _identifiers(frame)
    if identifiers is not None:
        return identifiers
    source_ids = _identifiers(source_df)
    if source_ids is None or 'Organisation unit' not in source_df or 'Organisation unit' not in frame:
        return None
    pairs = pd.DataFrame({
        'id': source_ids, 'label': source_df['Organisation unit'].map(_label),
    }).drop_duplicates()
    pairs = pairs.loc[pairs['id'].ne('') & pairs['label'].ne('')]
    one_id = pairs.groupby('label')['id'].nunique().eq(1)
    one_label = pairs.groupby('id')['label'].nunique().eq(1)
    pairs = pairs.loc[pairs['label'].map(one_id) & pairs['id'].map(one_label)]
    bridge = dict(zip(pairs['label'], pairs['id']))
    return frame['Organisation unit'].map(_label).map(bridge).fillna('')


def mask_failed_reporting_cells(frame, source_df):
    """Masque les métriques refusées sans effacer les autres datasets reçus.

    Les totaux combinés restent soumis au masque strict par unité. Ce masque
    concerne uniquement les cellules du reporting identifié par DXID. Sans
    correspondance exploitable, le comportement conservateur historique est
    maintenu sur les colonnes certifiées, ou numériques à défaut.
    """
    result = frame.copy()
    if result.empty or not isinstance(source_df, pd.DataFrame):
        return result
    per_unit = source_df.attrs.get('failed_metrics_by_unit') or {}
    if not isinstance(per_unit, dict):
        per_unit = {}
    per_unit = {clean_ou_id(uid): items for uid, items in per_unit.items() if clean_ou_id(uid)}
    failed = source_df.attrs.get('failed_units') or []
    if isinstance(failed, str):
        failed = [failed]
    failed_ids = {clean_ou_id(uid) for uid in failed if clean_ou_id(uid)} | set(per_unit)
    if not failed_ids:
        return result
    identifiers = _cell_identifiers(result, source_df)
    if identifiers is None:
        return result
    labels = source_df.attrs.get('reporting_item_labels') or {}
    if not isinstance(labels, dict):
        labels = {}
    certified = source_df.attrs.get('reporting_columns')
    fallback = (
        list(certified) if isinstance(certified, (list, tuple, set)) else
        [column for column in result if pd.api.types.is_numeric_dtype(result[column])]
    )
    excluded = {*_ID_COLUMNS, *_LABEL_COLUMNS}
    for uid in failed_ids:
        affected = identifiers.eq(uid)
        if not affected.any():
            continue
        items = per_unit.get(uid)
        if isinstance(items, str):
            items = [items]
        mapped = isinstance(items, (list, tuple, set)) and bool(items) and all(
            isinstance(item, str) and isinstance(labels.get(item), str) and labels[item]
            for item in items
        )
        columns = [labels[item] for item in items] if mapped else fallback
        columns = [column for column in dict.fromkeys(columns) if column in result and column not in excluded]
        if columns:
            result.loc[affected, columns] = float('nan')
    return result


def mask_failed_reporting_units(frame, source_df, metric_columns=None):
    """Masque les mesures d'unités dont le chargement Analytics est incomplet.

    Une unité est atomique : ses comptes ou taux ne sont pas recalculés sur un
    sous-ensemble des datasets reçus. Les autres unités, identifiants et libellés
    restent intacts. Les synthèses sans ID utilisent uniquement un pont strict
    et bijectif entre leurs libellés et les IDs de leur source courante.
    """
    result = frame.copy()
    if result.empty or not isinstance(source_df, pd.DataFrame):
        return result
    failed = source_df.attrs.get('failed_units') or []
    if isinstance(failed, str):
        failed = [failed]
    failed_ids = {identifier for value in failed if (identifier := clean_ou_id(value))}
    if not failed_ids:
        return result

    identifiers = _identifiers(result)
    if identifiers is None:
        source_ids = _identifiers(source_df)
        if source_ids is None or 'Organisation unit' not in source_df or 'Organisation unit' not in result:
            return result
        pairs = pd.DataFrame({
            'id': source_ids, 'label': source_df['Organisation unit'].map(_label),
        }).drop_duplicates()
        pairs = pairs.loc[pairs['id'].ne('') & pairs['label'].ne('')]
        unique_id_by_label = pairs.groupby('label')['id'].nunique().eq(1)
        unique_label_by_id = pairs.groupby('id')['label'].nunique().eq(1)
        pairs = pairs.loc[pairs['label'].map(unique_id_by_label) & pairs['id'].map(unique_label_by_id)]
        bridge = dict(zip(pairs['label'], pairs['id']))
        identifiers = result['Organisation unit'].map(_label).map(bridge).fillna('')

    affected = identifiers.isin(failed_ids)
    if not affected.any():
        return result
    if metric_columns is None:
        metric_columns = [column for column in result if pd.api.types.is_numeric_dtype(result[column])]
    elif isinstance(metric_columns, str):
        metric_columns = [metric_columns]
    columns = [
        column for column in metric_columns
        if column in result and column not in _ID_COLUMNS and column not in _LABEL_COLUMNS
    ]
    if columns:
        result.loc[affected, columns] = float('nan')
    return result

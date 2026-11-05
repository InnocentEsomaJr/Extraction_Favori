"""Empêche une requête Analytics échouée de devenir un indicateur numérique."""

import pandas as pd

from reporting_metrics import clean_ou_id


_ID_COLUMNS = ['Organisation unit ID', 'Zone_ID']
_LABEL_COLUMNS = {'Organisation unit', 'Zone de santé', 'Zone', 'Province', 'Pays'}


def _label(value):
    return '' if value is None or pd.isna(value) else str(value).strip()


def _identifiers(frame):
    for column in _ID_COLUMNS:
        if column in frame:
            return frame[column].map(clean_ou_id)
    return None


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

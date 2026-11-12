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


def _attached_identifiers(frame, source_df):
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
    unique_id_by_label = pairs.groupby('label')['id'].nunique().eq(1)
    unique_label_by_id = pairs.groupby('id')['label'].nunique().eq(1)
    pairs = pairs.loc[pairs['label'].map(unique_id_by_label) & pairs['id'].map(unique_label_by_id)]
    bridge = dict(zip(pairs['label'], pairs['id']))
    return frame['Organisation unit'].map(_label).map(bridge).fillna('')


def mask_failed_reporting_cells(frame, source_df):
    """Masque les seules cellules de reporting dont la requête a échoué.

    Les DX sont associés aux colonnes par ``reporting_item_labels``. Les
    valeurs reçues des autres datasets restent visibles ; cette fonction ne
    rend pas les totaux partiels valides. Les totaux et scores restent à
    masquer avec ``mask_failed_reporting_units``.

    Quand ``failed_cells_by_unit`` est présent, ce détail remplace les échecs
    globaux pour la vue brute. Un dictionnaire vide conserve les valeurs de
    secours reçues même si les totaux combinés restent indisponibles.

    Sans détail exploitable pour un échec, toutes les colonnes de reporting
    certifiées sont masquées pour l'unité. Les anciennes sources dépourvues
    de certification conservent le masquage prudent des colonnes numériques.
    """
    result = frame.copy()
    if result.empty or not isinstance(source_df, pd.DataFrame):
        return result
    identifiers = _attached_identifiers(result, source_df)
    if identifiers is None:
        return result

    cell_override = 'failed_cells_by_unit' in source_df.attrs
    failed_units = [] if cell_override else source_df.attrs.get('failed_units') or []
    if isinstance(failed_units, str):
        failed_units = [failed_units]
    failed_metrics = source_df.attrs.get(
        'failed_cells_by_unit' if cell_override else 'failed_metrics_by_unit'
    ) or {}
    if not isinstance(failed_metrics, dict):
        failed_metrics = {}
    per_unit = {
        identifier: items for value, items in failed_metrics.items()
        if (identifier := clean_ou_id(value))
    }
    failed_ids = {identifier for value in failed_units if (identifier := clean_ou_id(value))}
    failed_ids.update(per_unit)
    if not failed_ids:
        return result

    item_labels = source_df.attrs.get('reporting_item_labels') or {}
    if not isinstance(item_labels, dict):
        item_labels = {}
    certified = source_df.attrs.get('reporting_columns')
    fallback = (
        list(certified) if isinstance(certified, (list, tuple, set)) else
        [column for column in result if pd.api.types.is_numeric_dtype(result[column])]
    )
    excluded = {*_ID_COLUMNS, *_LABEL_COLUMNS}
    for identifier in failed_ids:
        affected = identifiers.eq(identifier)
        if not affected.any():
            continue
        items = per_unit.get(identifier)
        if isinstance(items, str):
            items = [items]
        if isinstance(items, (list, tuple, set)) and items and all(
            isinstance(item, str) and isinstance(item_labels.get(item), str) and item_labels[item] for item in items
        ):
            columns = [item_labels[item] for item in items]
        else:
            columns = fallback
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

    identifiers = _attached_identifiers(result, source_df)
    if identifiers is None:
        return result

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

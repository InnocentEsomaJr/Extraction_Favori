"""Remplacement du reporting du favori par la source mensuelle et hebdomadaire."""

from copy import deepcopy

import pandas as pd

from reporting_metrics import clean_ou_id


UNIT_ID = 'Organisation unit ID'
UNIT_NAME = 'Organisation unit'
_METRIC_SUFFIXES = (
    'actual reports', 'expected reports', 'actual reports on time',
    'reporting rate', 'reporting rate on time',
)


def _frame(source, label):
    if source is None:
        return pd.DataFrame(columns=[UNIT_ID])
    if not isinstance(source, pd.DataFrame):
        raise ValueError(f'La source {label} doit être un tableau de données.')
    if source.columns.duplicated().any():
        raise ValueError(f'Les colonnes de la source {label} sont ambiguës.')
    result = source.copy(deep=True)
    result.attrs = deepcopy(source.attrs)
    if UNIT_ID not in result:
        if not result.empty:
            raise ValueError(f'Les identifiants des unités sont absents de la source {label}.')
        result[UNIT_ID] = pd.Series(dtype=object)
    result[UNIT_ID] = result[UNIT_ID].map(clean_ou_id)
    if result[UNIT_ID].eq('').any() or result[UNIT_ID].duplicated().any():
        raise ValueError(f'Les identifiants des unités de la source {label} sont absents ou répétés.')
    return result


def _reporting_columns(frame):
    if 'reporting_columns' in frame.attrs:
        columns = frame.attrs['reporting_columns']
        if not isinstance(columns, (list, tuple)):
            raise ValueError('La liste des colonnes de reporting est invalide.')
        return [column for column in columns if column in frame and column not in (UNIT_ID, UNIT_NAME)]
    result = []
    for column in frame:
        name = ' '.join(str(column).casefold().split())
        if any(name == suffix or name.endswith(' ' + suffix) for suffix in _METRIC_SUFFIXES):
            result.append(column)
    return result


def _warnings(attrs):
    values = []
    for key in ('favorite_warnings', 'reporting_warnings', 'quality_warnings'):
        entries = attrs.get(key) or []
        if isinstance(entries, str):
            entries = [entries]
        values.extend(str(value) for value in entries)
    return list(dict.fromkeys(values))


def replace_reporting_source(original_df, combined_df):
    """Joint le reporting combiné aux autres indicateurs du favori par UID.

    Les seules unités reçues dans ``combined_df`` constituent le périmètre et
    déterminent l'ordre de sortie. Les anciennes métriques certifiées sont
    retirées avant la jointure ; sans certification, seuls les suffixes de
    reporting reconnus sont retirés. Une colonne générique ressemblante reste
    distincte lorsque le favori fournit explicitement sa certification.

    Les erreurs et périodes du reporting combiné sont conservées. Les erreurs
    du favori restent disponibles dans ``favorite_warnings`` et
    ``favorite_failed_units``, sans masquer des mesures combinées réussies.
    Les doublons ou collisions ambiguës sont refusés au lieu d'être agrégés.
    """
    combined = _frame(combined_df, 'combinée')
    if original_df is None:
        return combined
    original = _frame(original_df, 'du favori')
    old_reporting = set(_reporting_columns(original))
    generic_columns = [
        column for column in original
        if column not in old_reporting and column not in (UNIT_ID, UNIT_NAME)
    ]
    if any(column in combined for column in generic_columns):
        raise ValueError(
            "Un indicateur du favori porte le même libellé qu'une colonne de la source combinée."
        )
    identifiers = combined[UNIT_ID].tolist()
    aligned = original.set_index(UNIT_ID).reindex(identifiers)
    result = combined.copy(deep=True).reset_index(drop=True)
    if UNIT_NAME not in result:
        result[UNIT_NAME] = result[UNIT_ID]
    if UNIT_NAME in aligned:
        original_names = aligned[UNIT_NAME].reset_index(drop=True)
        names = result[UNIT_NAME]
        missing_name = names.isna() | names.astype(str).str.strip().eq('') | names.eq(result[UNIT_ID])
        known_original = original_names.notna() & original_names.astype(str).str.strip().ne('')
        result[UNIT_NAME] = names.where(~(missing_name & known_original), original_names)
    for column in generic_columns:
        result[column] = aligned[column].reset_index(drop=True)
    result.attrs = deepcopy(combined.attrs)
    result.attrs['reporting_columns'] = _reporting_columns(combined)
    favorite_warnings = _warnings(original.attrs)
    if favorite_warnings:
        result.attrs['favorite_warnings'] = list(dict.fromkeys(
            [*result.attrs.get('favorite_warnings', []), *favorite_warnings]
        ))
    original_failed = original.attrs.get('failed_units') or []
    if isinstance(original_failed, str):
        original_failed = [original_failed]
    failed_ids = {clean_ou_id(identifier) for identifier in original_failed}
    if failed_ids:
        result.attrs['favorite_failed_units'] = [identifier for identifier in identifiers if identifier in failed_ids]
    return result

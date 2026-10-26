"""Recherche, filtres et tris de la vue de catégorisation, sans modifier les scores."""

import unicodedata

import pandas as pd

from categorisation import TOTAL_COL


VIEW_STATUS_OPTIONS = [
    "Toutes",
    "Scores complets",
    "Scores incomplets",
    "Score inférieur à 80",
]
VIEW_SORT_OPTIONS = [
    "Ordre initial",
    "Score croissant",
    "Score décroissant",
    "Nom de zone",
]
SEARCH_COLUMNS = ["Zone de santé", "Barème datasets", "Statut catégorisation", "Province"]


def _normalize_text(value):
    if value is None or pd.isna(value):
        return ""
    text = " ".join(str(value).casefold().split())
    return "".join(
        character for character in unicodedata.normalize("NFKD", text)
        if not unicodedata.combining(character)
    )


def filter_categorisation(df, query="", status="Toutes", sort_by="Ordre initial"):
    """Retourne une copie filtrée et triée, en gardant colonnes, types et index.

    La recherche ignore les accents, la casse et les espaces supplémentaires.
    Une absence de score reste distincte d'une véritable note de zéro.
    """
    if status not in VIEW_STATUS_OPTIONS:
        raise ValueError("Le filtre de statut de catégorisation est inconnu.")
    if sort_by not in VIEW_SORT_OPTIONS:
        raise ValueError("Le tri de catégorisation est inconnu.")

    result = df.copy()
    if result.empty:
        return result

    normalized_query = _normalize_text(query)
    if normalized_query:
        searchable = pd.Series("", index=result.index, dtype=object)
        for column in SEARCH_COLUMNS:
            if column in result.columns:
                searchable = searchable + " " + result[column].map(_normalize_text)
        result = result.loc[searchable.str.contains(normalized_query, regex=False)].copy()

    scores = pd.to_numeric(
        result[TOTAL_COL] if TOTAL_COL in result.columns else pd.Series(float("nan"), index=result.index),
        errors="coerce",
    )
    if status == "Scores complets":
        result = result.loc[scores.notna()].copy()
    elif status == "Scores incomplets":
        result = result.loc[scores.isna()].copy()
    elif status == "Score inférieur à 80":
        result = result.loc[scores.lt(80)].copy()

    if sort_by in {"Score croissant", "Score décroissant"} and TOTAL_COL in result.columns:
        result = result.sort_values(
            TOTAL_COL, ascending=sort_by == "Score croissant", kind="stable",
            na_position="last", key=lambda values: pd.to_numeric(values, errors="coerce"),
        )
    elif sort_by == "Nom de zone" and "Zone de santé" in result.columns:
        result = result.sort_values(
            "Zone de santé", kind="stable", na_position="last",
            key=lambda values: values.map(
                lambda value: None if pd.isna(value) else _normalize_text(value)
            ),
        )
    return result.copy()

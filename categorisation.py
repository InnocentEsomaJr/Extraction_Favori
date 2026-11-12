"""Calculs de catégorisation SNIS, indépendants de Streamlit et de DHIS2."""

import numpy as np
import pandas as pd


QUALITY_NOTE_COL = "Score de qualité /20"
COMPONENT_COLS = [
    "Score complétude /20",
    "Score promptitude /20",
    "Score datasets /20",
    "Score violations /20",
    "Score corrections /20",
]
NOTE_COLS = [QUALITY_NOTE_COL, *COMPONENT_COLS]
TOTAL_COL = "Score catégorisation /100"
DATASET_COUNT_COL = "Nombre de datasets complétude ≥95 %"
REMAINING_COL = "Règles restantes (M-1 -> M)"
INTEGER_COLS = [
    "Règles violées (M-1)",
    "Règles corrigées (M-1 -> M)",
    "Règles violées (M)",
    DATASET_COUNT_COL,
    REMAINING_COL,
    *NOTE_COLS,
    TOTAL_COL,
]


def _numeric(values):
    """Conserve les valeurs numériques finies et positives ou nulles."""
    series = values.copy() if isinstance(values, pd.Series) else pd.Series(values)
    numeric = pd.to_numeric(series, errors="coerce").astype(float)
    return numeric.where(np.isfinite(numeric) & numeric.ge(0))


def _column(frame, name):
    return frame[name] if name in frame.columns else pd.Series(np.nan, index=frame.index)


def _profile(value):
    if pd.isna(value):
        return None
    normalized = str(value).strip().upper()
    return normalized if normalized in {"DPS", "ZS"} else None


def score_taux(series):
    """Note un taux /100 sur 20, sans convertir une absence en zéro."""
    values = _numeric(series)
    scores = np.select(
        [values.ge(95), values.ge(80), values.ge(60), values.ge(50)],
        [20, 15, 10, 5],
        default=0,
    )
    return pd.Series(scores, index=values.index, dtype=float).where(values.notna())


def score_datasets(series, profil):
    """Note le nombre de datasets >=95 % selon le barème DPS ou ZS."""
    values = _numeric(series)
    if isinstance(profil, pd.Series):
        profiles = profil.reindex(values.index).map(_profile)
    else:
        profiles = pd.Series(_profile(profil), index=values.index, dtype=object)
    scores = pd.Series(np.nan, index=values.index, dtype=float)
    for name, thresholds in {"DPS": [12, 10, 8, 6], "ZS": [10, 8, 6, 5]}.items():
        mask = profiles.eq(name) & values.notna()
        points = np.select([values.ge(t) for t in thresholds], [20, 15, 10, 5], default=0)
        scores.loc[mask] = points[mask.to_numpy()]
    return scores


def score_violations(series):
    """Note le ratio de violations /100 rapports, avec une frontière 20 explicite."""
    values = _numeric(series)
    scores = np.select(
        [values.le(5), values.le(10), values.le(15), values.le(20)],
        [20, 15, 10, 5],
        default=0,
    )
    return pd.Series(scores, index=values.index, dtype=float).where(values.notna())


def calculer_categorisation(df, profil="Automatique"):
    """Ajoute les cinq composantes et leur somme aux colonnes originales.

    Le score historique de qualité reste un repère distinct, exclu du total.
    En mode automatique, le profil vient exclusivement de ``Profil_Bareme``.
    Une analyse indisponible ou une composante absente empêche le total /100.
    """
    if profil != "Automatique" and _profile(profil) is None:
        raise ValueError("Le profil doit être Automatique, DPS ou ZS.")

    result = df.copy()
    profiles = (
        _column(result, "Profil_Bareme").map(_profile)
        if profil == "Automatique"
        else pd.Series(_profile(profil), index=result.index, dtype=object)
    )
    counts = _numeric(_column(result, "Nombre_Datasets_95"))
    current = _numeric(_column(result, "Violations_Actuelles"))
    previous = _numeric(_column(result, "Violations_Mois_Precedent"))
    corrected = _numeric(_column(result, "Violations_Corrigees"))
    actual = _numeric(_column(result, "Actual_Reports"))
    errors = _column(result, "Erreur_Validation").fillna("").astype(str).str.strip().ne("")
    correction_errors = _column(result, "Erreur_Corrections").fillna("").astype(str).str.strip().ne("")
    # Les anciens résultats contenaient déjà M et M-1 dans une seule analyse.
    # Dans le nouveau flux, seul le second bouton confirme le contrôle M-1.
    correction_checked = (
        result["Controle_Corrections_Effectue"].eq(True).fillna(False).astype(bool)
        if "Controle_Corrections_Effectue" in result.columns
        else pd.Series(True, index=result.index, dtype=bool)
    )
    correction_counts_valid = previous.notna() & corrected.notna() & corrected.le(previous)
    correction_valid = correction_counts_valid & correction_checked & ~errors & ~correction_errors
    remaining = (previous - corrected).clip(lower=0).where(correction_valid)

    result["Ratio_Violations_sur_100"] = current / actual.where(actual.gt(0)) * 100
    result[QUALITY_NOTE_COL] = score_taux(_column(result, "Score_Qualite"))
    result["Barème datasets"] = profiles.fillna("À préciser")
    result[DATASET_COUNT_COL] = counts
    result[REMAINING_COL] = remaining
    result[COMPONENT_COLS[0]] = score_taux(_column(result, "Complétude_Globale"))
    result[COMPONENT_COLS[1]] = score_taux(_column(result, "Promptitude_Globale"))
    result[COMPONENT_COLS[2]] = score_datasets(counts, profiles)
    result[COMPONENT_COLS[3]] = score_violations(result["Ratio_Violations_sur_100"]).where(~errors)
    correction_score = pd.Series(np.where(remaining.eq(0), 20, 0), index=result.index, dtype=float)
    result[COMPONENT_COLS[4]] = correction_score.where(remaining.notna())
    result[TOTAL_COL] = result[COMPONENT_COLS].sum(axis=1, min_count=len(COMPONENT_COLS))

    reasons = [
        (profiles.isna(), "barème datasets à préciser"),
        (counts.isna(), "nombre de datasets indisponible"),
        (result[COMPONENT_COLS[0]].isna(), "complétude indisponible"),
        (result[COMPONENT_COLS[1]].isna(), "promptitude indisponible"),
        (actual.isna() | actual.le(0), "nombre de rapports nul ou indisponible"),
        (errors, "analyse des violations indisponible"),
        (~errors & current.isna(), "violations actuelles indisponibles"),
        (~correction_checked, "contrôle des corrections non effectué"),
        (correction_checked & correction_errors, "contrôle des corrections indisponible"),
        (
            correction_checked & ~errors & ~correction_errors & ~correction_counts_valid,
            "comparaison des corrections indisponible ou incohérente",
        ),
    ]
    statuses = []
    for position in range(len(result)):
        missing = [label for mask, label in reasons if bool(mask.iloc[position])]
        statuses.append("Incomplète : " + " ; ".join(missing) if missing else "Complète")
    result["Statut catégorisation"] = pd.Series(statuses, index=result.index, dtype=object)
    return result


def build_baremes_df():
    """Retourne une légende tabulaire des règles actives, sans données métier."""
    rows = []

    def add(criterion, profile, condition, points):
        rows.append({"Critère": criterion, "Barème": profile, "Condition": condition, "Points /20": points})

    for criterion in ["Complétude", "Promptitude"]:
        for condition, points in [
            ("Moins de 50 %", 0), ("50 à moins de 60 %", 5),
            ("60 à moins de 80 %", 10), ("80 à moins de 95 %", 15),
            ("95 % et plus", 20),
        ]:
            add(criterion, "Tous", condition, points)
    for profile, bands in {
        "DPS": [("Moins de 6", 0), ("6 à moins de 8", 5), ("8 à moins de 10", 10), ("10 à moins de 12", 15), ("12 et plus", 20)],
        "ZS": [("Moins de 5", 0), ("5 à moins de 6", 5), ("6 à moins de 8", 10), ("8 à moins de 10", 15), ("10 et plus", 20)],
    }.items():
        for condition, points in bands:
            add("Nombre de datasets complétude ≥95 %", profile, condition, points)
    for condition, points in [
        ("0 à 5", 20), ("Plus de 5 à 10", 15), ("Plus de 10 à 15", 10),
        ("Plus de 15 à 20", 5), ("Plus de 20", 0),
    ]:
        add("Violations /100 rapports", "Tous", condition, points)
    add("Corrections", "Tous", "Aucune règle restante", 20)
    add("Corrections", "Tous", "Au moins une règle restante", 0)
    return pd.DataFrame(rows)


def style_note_20(value):
    """Cinq couleurs distinctes ; une note indisponible reste neutre."""
    try:
        value = float(value)
    except (TypeError, ValueError):
        return ""
    colors = {
        0: ("white", "#B91C1C"),
        5: ("black", "#FDBA74"),
        10: ("black", "#FDE047"),
        15: ("black", "#86EFAC"),
        20: ("white", "#15803D"),
    }
    if value not in colors:
        return ""
    foreground, background = colors[value]
    return f"color: {foreground}; background-color: {background}"


def style_context_tab5():
    return {
        "taux_cols": ["Complétude globale (%)", "Promptitude globale (%)", "Score de qualité", TOTAL_COL],
        "note_cols": NOTE_COLS.copy(),
        "int_cols": INTEGER_COLS.copy(),
        "highlight_max_cols": ["Règles violées (M)"],
        "ratio_alert_col": "Ratio / 100 rapports",
        "ratio_alert_threshold": 10.0,
        "decimals": 2,
    }

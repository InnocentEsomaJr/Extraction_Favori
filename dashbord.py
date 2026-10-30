import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import io
import os
from datetime import date, datetime, timedelta, timezone
from uuid import uuid4
from categorisation import (
    NOTE_COLS, TOTAL_COL, build_baremes_df, calculer_categorisation,
    style_context_tab5, style_note_20,
)
from reporting_metrics import aggregate_validation_metrics, clean_ou_id
from geography import build_geography_catalog, labels_for_units, resolve_hierarchy, zone_catalog
from scoped_analytics import analytics_to_reporting_frame, build_analytics_query
from categorisation_view import VIEW_SORT_OPTIONS, VIEW_STATUS_OPTIONS, filter_categorisation
from validation_api import fetch_validation_analysis
from validation_filters import date_range_months, monthly_bounds, reference_month
from validation_scope import assign_validation_zones
from validation_details import (
    DISPLAY_COLUMNS as VIOLATION_COLUMNS, build_validation_details, detail_export_frame,
    filter_validation_details, style_importance, validation_detail_identifiers,
)

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="SNIS RDC - Dashboard Performance", layout="wide")
KINSHASA_TZ = timezone(timedelta(hours=1), "Africa/Kinshasa")

# Fonctions de coloration (Scripts originaux)
def style_taux(val):
    try: val = float(val)
    except: return ''
    if val < 50: return 'color: white; background-color: #FF0000'
    elif 50 <= val <= 69: return 'color: white; background-color: #800000'
    elif 70 <= val <= 79: return 'color: black; background-color: #FFFF00'
    elif 80 <= val <= 95: return 'color: black; background-color: #32CD32'
    elif val > 95: return 'color: white; background-color: #008000'
    return ''

def style_score(val):
    try: val = int(val)
    except: return ''
    if val < 5: return 'color: white; background-color: #800000'
    elif val == 5: return 'color: black; background-color: #FFC0CB'
    elif 6 <= val <= 9: return 'color: black; background-color: #32CD32'
    elif val == 10: return 'color: white; background-color: #008000'
    elif val > 10: return 'color: white; background-color: #004d00'
    return ''

def normalize_org_name(name):
    """Normalise les noms d'unités pour les comparaisons robustes."""
    if pd.isna(name):
        return ""
    return " ".join(str(name).strip().lower().split())

def _find_best_column(columns, exact_names=None, include_terms=None, exclude_terms=None):
    """
    Trouve la meilleure colonne selon noms exacts ou termes inclus/exclus.
    Retourne None si non trouvée.
    """
    cols = list(columns)
    exact_names = exact_names or []
    include_terms = include_terms or []
    exclude_terms = exclude_terms or []

    for name in exact_names:
        if name in cols:
            return name

    for col in cols:
        col_l = str(col).strip().lower()
        if include_terms and not all(term in col_l for term in include_terms):
            continue
        if exclude_terms and any(term in col_l for term in exclude_terms):
            continue
        return col
    return None

def normalize_orgunit_columns(df):
    """
    Normalise les colonnes organisationnelles vers:
    - Organisation unit
    - Organisation unit ID (si disponible)
    - Organisation unit is parent (si disponible)
    """
    if df is None or not isinstance(df, pd.DataFrame):
        return df

    out = df.copy()
    cols = out.columns.tolist()

    # Colonne nom d'unité
    org_col = _find_best_column(
        cols,
        exact_names=["Organisation unit", "Organisation Unit", "Org unit", "Org Unit"],
        include_terms=["organisation", "unit"],
        exclude_terms=["id", "uid", "code", "description", "parameter", "is parent"]
    )
    if org_col is None:
        org_col = _find_best_column(
            cols,
            include_terms=["org", "unit"],
            exclude_terms=["id", "uid", "code", "description", "parameter", "is parent"]
        )
    if org_col is None and len(cols) > 0:
        # Fallback: première colonne non métrique
        metric_terms = ["reporting rate", "actual reports", "expected reports", "ratio", "score"]
        fallback_cols = [c for c in cols if not any(t in str(c).lower() for t in metric_terms)]
        org_col = fallback_cols[0] if fallback_cols else cols[0]

    # Colonne ID d'unité
    id_col = _find_best_column(
        cols,
        exact_names=["Organisation unit ID", "Organisation Unit ID", "Org unit ID", "Organisation unit UID", "ou"],
        include_terms=["organisation", "unit", "id"]
    )
    if id_col is None:
        id_col = _find_best_column(
            cols,
            include_terms=["org", "unit", "id"]
        )
    if id_col is None:
        id_col = _find_best_column(
            cols,
            include_terms=["organisation", "unit", "uid"]
        )
    if id_col is None:
        id_col = _find_best_column(
            cols,
            include_terms=["org", "unit", "uid"]
        )
    if id_col is None and "ou" in cols:
        id_col = "ou"

    # Colonne parent (optionnelle)
    parent_col = _find_best_column(
        cols,
        exact_names=["Organisation unit is parent"],
        include_terms=["is parent"]
    )
    if parent_col is None:
        parent_col = _find_best_column(
            cols,
            include_terms=["organisation", "unit", "parent"]
        )

    rename_map = {}
    if org_col and org_col != "Organisation unit":
        rename_map[org_col] = "Organisation unit"
    if id_col and id_col != "Organisation unit ID":
        rename_map[id_col] = "Organisation unit ID"
    if parent_col and parent_col != "Organisation unit is parent":
        rename_map[parent_col] = "Organisation unit is parent"
    if rename_map:
        out = out.rename(columns=rename_map)

    # Garantir l'existence de la colonne organisation unit pour éviter les KeyError
    if "Organisation unit" not in out.columns:
        out["Organisation unit"] = "Unité non renseignée"

    return out

def _extract_identifier(value):
    """Extrait un identifiant stable depuis un champ DHIS2 (dict ou string)."""
    if isinstance(value, dict):
        for key in ('id', 'uid', 'displayName', 'name'):
            identifier = _extract_identifier(value.get(key))
            if identifier:
                return identifier
        return ''
    if value is None or pd.isna(value):
        return ''
    text = str(value).strip()
    return '' if text.casefold() in {'nan', 'none', 'null', '<na>', 'nat'} else text

def build_violation_signatures(df_val):
    """
    Construit les signatures de violations pour comparer T-1 vs T.
    On exclut volontairement la période pour détecter les erreurs réellement corrigées.
    """
    if df_val is None or df_val.empty:
        return []

    signatures = []
    for row in df_val.to_dict("records"):
        ou = _extract_identifier(row.get("organisationUnit"))
        rule = _extract_identifier(row.get("validationRule"))
        aoc = _extract_identifier(row.get("attributeOptionCombo"))
        coc = _extract_identifier(row.get("categoryOptionCombo"))
        ds = _extract_identifier(row.get("dataSet"))

        # Une violation doit au minimum avoir une règle ou une OU.
        if not ou and not rule:
            continue

        signatures.append((ou, rule, aoc, coc, ds))
    return signatures


def build_violation_signature_set(df_val):
    return set(build_violation_signatures(df_val))

# Fonction pour l'export Excel
def to_excel(df):
    output = io.BytesIO()
    # Utilisation de xlsxwriter (assurez-vous qu'il est installé via pip install xlsxwriter)
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Performance')
    return output.getvalue()

def _build_styler(
    df,
    taux_cols=None,
    score_cols=None,
    int_cols=None,
    decimals=2,
    highlight_max_cols=None,
    ratio_alert_col=None,
    ratio_alert_threshold=10.0,
    note_cols=None
):
    """Construit un Styler pandas avec les mêmes règles de couleurs que le dashboard."""
    if df is None or df.empty:
        return None

    taux_cols = [c for c in (taux_cols or []) if c in df.columns]
    score_cols = [c for c in (score_cols or []) if c in df.columns]
    note_cols = [c for c in (note_cols or []) if c in df.columns]
    int_cols = [c for c in (int_cols or []) if c in df.columns]
    highlight_max_cols = [c for c in (highlight_max_cols or []) if c in df.columns]

    styler = df.style

    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    format_map = {c: f"{{:.{int(decimals)}f}}" for c in numeric_cols}
    for c in int_cols:
        format_map[c] = "{:.0f}"
    if format_map:
        styler = styler.format(format_map, na_rep="")

    if taux_cols:
        styler = styler.map(style_taux, subset=taux_cols)
    if score_cols:
        styler = styler.map(style_score, subset=score_cols)
    if note_cols:
        styler = styler.map(style_note_20, subset=note_cols)
    if highlight_max_cols:
        styler = styler.highlight_max(subset=highlight_max_cols, color="#ffcccc")
    if ratio_alert_col and ratio_alert_col in df.columns:
        styler = styler.map(
            lambda x: (
                "color: red; font-weight: bold"
                if pd.notna(x) and pd.to_numeric(x, errors="coerce") > ratio_alert_threshold
                else ""
            ),
            subset=[ratio_alert_col]
        )

    return styler

def _extract_css_values(css_text):
    """Extrait couleur texte/fond/gras depuis une chaîne CSS."""
    default_text = "#000000"
    default_bg = "#FFFFFF"
    bold = False
    if not css_text:
        return default_text, default_bg, bold

    text_color = default_text
    bg_color = default_bg
    for item in str(css_text).split(";"):
        if ":" not in item:
            continue
        key, value = item.split(":", 1)
        key = key.strip().lower()
        value = value.strip()
        if key == "color" and value:
            text_color = value
        elif key == "background-color" and value:
            bg_color = value
        elif key == "font-weight" and "bold" in value.lower():
            bold = True
    return text_color, bg_color, bold

def build_dashboard_comments_df(
    df_final_c,
    df_final_p,
    df_synth,
    selected_zone,
    df_tab4_fusion=None,
    df_tab5=None,
    df_base=None,
    df_actual_detail=None,
    df_expected_detail=None,
    df_comparatif=None,
    df_top5=None,
    df_flop5=None,
):
    """Prépare des commentaires clairs pour graphiques et tableaux exportés."""
    zone_label = selected_zone if selected_zone != "Toutes les zones" else "Toutes les zones"
    comp_avg = float(df_synth['Complétude_Globale (%)'].mean()) if not df_synth.empty else 0.0
    prom_avg = float(df_synth['Promptitude_Globale (%)'].mean()) if not df_synth.empty else 0.0
    top_comp = df_synth.nlargest(1, 'Complétude_Globale (%)')['Organisation unit'].iloc[0] if not df_synth.empty else "N/A"
    flop_prom = df_synth.nsmallest(1, 'Promptitude_Globale (%)')['Organisation unit'].iloc[0] if not df_synth.empty else "N/A"

    rows = [
        {
            "Element": "Légende des couleurs",
            "Commentaire": (
                "Taux: <50 rouge vif, 50-69 bordeaux, 70-79 jaune, 80-95 vert clair, >95 vert foncé. "
                "Score datasets: <5 bordeaux, 5 rose, 6-9 vert clair, 10 vert, >10 vert foncé."
            )
        },
        {
            "Element": "Onglet 2 - Tableau Complétude",
            "Commentaire": (
                f"Complétude moyenne {comp_avg:.2f}%. Les cellules colorées reprennent strictement la logique du dashboard. "
                "Priorité: unités en rouge/bordeaux."
            )
        },
        {
            "Element": "Base de données",
            "Commentaire": (
                "Cette feuille reprend les données sources filtrées (niveau, période, zone). "
                "Elle sert de référence aux tableaux de performance. Les dates et le parent "
                "de l'analyse des règles figurent dans Contexte_Validation."
            )
        },
        {
            "Element": "Rapports détaillés - Données réelles",
            "Commentaire": (
                "Ce tableau détaille les 'Actual reports' par indicateur et la somme Reports_Actual. "
                "Il reflète le volume réel des notifications."
            )
        },
        {
            "Element": "Rapports détaillés - Participants/Attendus",
            "Commentaire": (
                "Ce tableau présente les 'Expected reports' et la somme Reports_Attendu. "
                "Il représente le dénominateur attendu pour les taux."
            )
        },
        {
            "Element": "Onglet 2 - Graphique Classement",
            "Commentaire": (
                f"Classement horizontal par complétude globale. L'unité la plus performante observée est '{top_comp}'. "
                "Lire du bas vers le haut pour la progression."
            )
        },
        {
            "Element": "Onglet 3 - Tableau Promptitude",
            "Commentaire": (
                f"Promptitude moyenne {prom_avg:.2f}%. Les couleurs suivent la même échelle que l'onglet 3 "
                "pour identifier rapidement les retards de notification."
            )
        },
        {
            "Element": "Onglet 4 - Quadrant Comparatif",
            "Commentaire": (
                "Le quadrant croise complétude (axe X) et promptitude (axe Y). "
                "Les lignes à 80% séparent les unités à renforcer des unités performantes."
            )
        },
        {
            "Element": "Onglet 4 - Tableau comparatif",
            "Commentaire": (
                "Le tableau comparatif présente, par unité, la complétude et la promptitude "
                "avec les compteurs de datasets >=95% pour prioriser les actions."
            )
        },
        {
            "Element": "Onglet 4 - Top/Flop",
            "Commentaire": (
                f"Top complétude: {top_comp}. Flop promptitude: {flop_prom}. "
                "Action recommandée: cibler d'abord les unités du flop promptitude."
            )
        },
        {
            "Element": "Top 5 Complétude",
            "Commentaire": "Classement des 5 unités avec les meilleurs taux de complétude."
        },
        {
            "Element": "Flop 5 Promptitude",
            "Commentaire": "Classement des 5 unités avec les plus faibles taux de promptitude."
        },
        {
            "Element": "Onglet 5 - Tableau Catégorisation",
            "Commentaire": (
                "Le tableau compare M-1 et M: règles violées, règles corrigées et ratio /100 rapports. "
                "Le score de catégorisation /100 additionne les notes de complétude, promptitude, "
                "datasets à complétude >=95%, violations et corrections. "
                "Le score de qualité historique reste distinct de ce total. "
                "Les barèmes DPS/ZS sont détaillés dans la feuille Baremes_Categorisation."
            )
        },
    ]

    if df_tab4_fusion is not None and not df_tab4_fusion.empty:
        rows.append(
            {
                "Element": "Onglet 4 - Tableau fusionné (zone filtrée)",
                "Commentaire": (
                    f"Pour '{zone_label}', ce tableau aligne indicateurs dataset Reporting/Actual avec "
                    "complétude, promptitude et compteurs >=95%."
                )
            }
        )
    if df_tab5 is not None and not df_tab5.empty:
        rows.append(
            {
                "Element": "Onglet 5 - Lecture décisionnelle",
                "Commentaire": (
                    "Comparer d'abord 'Règles violées (M)' puis 'Règles corrigées (M-1 -> M)'. "
                    "La note des corrections vaut 20 si aucune violation de M-1 ne reste à corriger. "
                    "Une note ou un total vide indique une donnée indisponible, précisée dans le statut."
                )
            }
        )

    return pd.DataFrame(rows)

def to_excel_dashboard_report(
    df_base,
    df_actual_detail,
    df_expected_detail,
    df_final_c,
    df_final_p,
    df_synth,
    comments_df,
    df_tab4_fusion=None,
    df_comparatif=None,
    df_tab5=None,
    df_top5=None,
    df_flop5=None,
    style_context=None,
    df_violations=None,
    df_validation_context=None
):
    """Exporte un rapport Excel commenté avec colorations conditionnelles."""
    style_context = style_context or {}
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        sty_base = _build_styler(df_base, **style_context.get("base", {}))
        (sty_base if sty_base is not None else df_base).to_excel(
            writer, index=False, sheet_name='Base_Donnees'
        )

        sty_actual = _build_styler(df_actual_detail, **style_context.get("actual_detail", {}))
        (sty_actual if sty_actual is not None else df_actual_detail).to_excel(
            writer, index=False, sheet_name='Detail_Donnees_Reelles'
        )

        sty_expected = _build_styler(df_expected_detail, **style_context.get("expected_detail", {}))
        (sty_expected if sty_expected is not None else df_expected_detail).to_excel(
            writer, index=False, sheet_name='Detail_Reports_Attendus'
        )

        sty_comp = _build_styler(df_final_c, **style_context.get("completude", {}))
        (sty_comp if sty_comp is not None else df_final_c).to_excel(
            writer, index=False, sheet_name='Perf_Final_Completude'
        )

        sty_prom = _build_styler(df_final_p, **style_context.get("promptitude", {}))
        (sty_prom if sty_prom is not None else df_final_p).to_excel(
            writer, index=False, sheet_name='Analyse_Promptitude'
        )

        sty_synth = _build_styler(df_synth, **style_context.get("synthese", {}))
        (sty_synth if sty_synth is not None else df_synth).to_excel(
            writer, index=False, sheet_name='Synthese_Comparative'
        )

        if df_comparatif is not None and not df_comparatif.empty:
            sty_compa = _build_styler(df_comparatif, **style_context.get("comparatif", {}))
            (sty_compa if sty_compa is not None else df_comparatif).to_excel(
                writer, index=False, sheet_name='Comparatif_C_P'
            )

        if df_tab4_fusion is not None and not df_tab4_fusion.empty:
            sty_fusion = _build_styler(df_tab4_fusion, **style_context.get("fusion", {}))
            (sty_fusion if sty_fusion is not None else df_tab4_fusion).to_excel(
                writer, index=False, sheet_name='Perf_Zone_Filtree'
            )

        if df_top5 is not None and not df_top5.empty:
            sty_top5 = _build_styler(df_top5, **style_context.get("top5", {}))
            (sty_top5 if sty_top5 is not None else df_top5).to_excel(
                writer, index=False, sheet_name='Top5_Completude'
            )

        if df_flop5 is not None and not df_flop5.empty:
            sty_flop5 = _build_styler(df_flop5, **style_context.get("flop5", {}))
            (sty_flop5 if sty_flop5 is not None else df_flop5).to_excel(
                writer, index=False, sheet_name='Flop5_Promptitude'
            )

        if df_tab5 is not None and not df_tab5.empty:
            sty_tab5 = _build_styler(df_tab5, **style_context.get("tab5", {}))
            (sty_tab5 if sty_tab5 is not None else df_tab5).to_excel(
                writer, index=False, sheet_name='Resultats_Regles'
            )
        else:
            pd.DataFrame(
                [{"Information": "Aucun résultat de règles. Lancez d'abord l'analyse des violations (onglet 5)."}]
            ).to_excel(writer, index=False, sheet_name='Resultats_Regles')

        build_baremes_df().to_excel(writer, index=False, sheet_name='Baremes_Categorisation')
        if df_violations is not None:
            details_styler = df_violations.style.format(
                {'Valeur gauche': '{:.2f}', 'Valeur droite': '{:.2f}'}, na_rep=''
            )
            if 'Importance' in df_violations.columns:
                details_styler = details_styler.map(style_importance, subset=['Importance'])
            details_styler.to_excel(writer, index=False, sheet_name='Details_Violations')
        else:
            pd.DataFrame([{'Information': "Lancez l'analyse des violations pour afficher leur détail."}]).to_excel(
                writer, index=False, sheet_name='Details_Violations'
            )
        detail_sheet = writer.sheets['Details_Violations']
        detail_sheet.freeze_panes(1, 2)
        detail_sheet.set_column(0, 1, 30)
        detail_sheet.set_column(2, 3, 20)
        detail_sheet.set_column(4, 4, 65)
        detail_sheet.set_column(5, 7, 18)
        detail_sheet.set_column(8, 8, 90)
        if df_validation_context is not None:
            df_validation_context.to_excel(writer, index=False, sheet_name='Contexte_Validation')
            writer.sheets['Contexte_Validation'].set_column(0, 0, 34)
            writer.sheets['Contexte_Validation'].set_column(1, 1, 95)
            writer.sheets['Contexte_Validation'].freeze_panes(1, 0)
        for name in ['Resultats_Regles', 'Baremes_Categorisation']:
            sheet = writer.sheets[name]
            sheet.freeze_panes(1, 1)
            sheet.set_column(0, 0, 28)
            if name == 'Baremes_Categorisation':
                sheet.set_column(1, 1, 12)
                sheet.set_column(2, 2, 32)
                sheet.set_column(3, 3, 12)
            elif df_tab5 is not None and not df_tab5.empty:
                sheet.set_column(1, len(df_tab5.columns) - 1, 22)
                sheet.set_column(len(df_tab5.columns) - 1, len(df_tab5.columns) - 1, 65)

        comments_df.to_excel(writer, index=False, sheet_name='Commentaires')
    return output.getvalue()

def _add_dataframe_to_docx(doc, title, df, max_rows=20):
    doc.add_heading(title, level=2)
    if df is None or df.empty:
        doc.add_paragraph("Aucune donnée disponible.")
        return
    df_show = df.head(max_rows).copy()
    table = doc.add_table(rows=1, cols=len(df_show.columns))
    table.style = 'Table Grid'
    for i, col in enumerate(df_show.columns):
        table.rows[0].cells[i].text = str(col)
    for _, row in df_show.iterrows():
        cells = table.add_row().cells
        for i, col in enumerate(df_show.columns):
            val = row[col]
            if isinstance(val, float):
                cells[i].text = f"{val:.2f}"
            else:
                cells[i].text = str(val)

def _add_dataframe_image_to_docx(doc, title, df, style_ctx=None, comment=None, max_rows=20):
    """Ajoute un tableau en image (avec couleurs) dans Word."""
    from docx.shared import Inches

    doc.add_heading(title, level=2)
    if comment:
        doc.add_paragraph(comment)

    img_bytes, img_error = dataframe_to_png_bytes(
        df,
        title=title,
        max_rows=max_rows,
        style_context=style_ctx
    )
    if img_bytes is not None:
        doc.add_picture(io.BytesIO(img_bytes), width=Inches(6.7))
    else:
        doc.add_paragraph(f"Image indisponible: {img_error}")
        _add_dataframe_to_docx(doc, title, df, max_rows=max_rows)

def to_word_dashboard_report(
    df_final_c,
    df_final_p,
    df_synth,
    comments_df,
    df_tab4_fusion=None,
    fig_comp=None,
    fig_quad=None,
    df_comparatif=None,
    df_tab5=None,
    style_context=None
):
    """Exporte un rapport Word commenté avec tableaux colorés et explications."""
    try:
        from docx import Document
        from docx.shared import Inches
    except Exception:
        return None, "python-docx non installé (pip install python-docx)."

    style_context = style_context or {}
    doc = Document()
    doc.add_heading("Rapport Dashboard SNIS RDC", level=1)
    doc.add_paragraph(f"Généré le {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    doc.add_heading("Commentaires d'interprétation", level=2)
    for _, row in comments_df.iterrows():
        doc.add_paragraph(f"{row['Element']}: {row['Commentaire']}")

    img_comp, _ = figure_to_png_bytes(fig_comp)
    if img_comp is not None:
        doc.add_heading("Graphique - Classement Complétude", level=2)
        doc.add_paragraph(_comment_for(comments_df, "Graphique Classement", "Classement par complétude globale."))
        doc.add_picture(io.BytesIO(img_comp), width=Inches(6.7))

    img_quad, _ = figure_to_png_bytes(fig_quad)
    if img_quad is not None:
        doc.add_heading("Graphique - Quadrant Comparatif", level=2)
        doc.add_paragraph(_comment_for(comments_df, "Quadrant", "Comparaison complétude/promptitude."))
        doc.add_picture(io.BytesIO(img_quad), width=Inches(6.7))

    _add_dataframe_image_to_docx(
        doc,
        "Tableau - Complétude",
        df_final_c,
        style_ctx=style_context.get("completude", {}),
        comment=_comment_for(comments_df, "Tableau Complétude", "")
    )
    _add_dataframe_image_to_docx(
        doc,
        "Tableau - Promptitude",
        df_final_p,
        style_ctx=style_context.get("promptitude", {}),
        comment=_comment_for(comments_df, "Tableau Promptitude", "")
    )
    _add_dataframe_image_to_docx(
        doc,
        "Tableau - Comparatif",
        df_synth,
        style_ctx=style_context.get("synthese", {}),
        comment=_comment_for(comments_df, "Top/Flop", "")
    )
    if df_comparatif is not None and not df_comparatif.empty:
        _add_dataframe_image_to_docx(
            doc,
            "Tableau comparatif Complétude / Promptitude",
            df_comparatif,
            style_ctx=style_context.get("comparatif", {}),
            comment=_comment_for(comments_df, "Tableau comparatif", "")
        )
    if df_tab4_fusion is not None and not df_tab4_fusion.empty:
        _add_dataframe_image_to_docx(
            doc,
            "Tableau fusionné zone filtrée",
            df_tab4_fusion,
            style_ctx=style_context.get("fusion", {}),
            comment=_comment_for(comments_df, "fusionné", "")
        )
    if df_tab5 is not None and not df_tab5.empty:
        _add_dataframe_image_to_docx(
            doc,
            "Tableau catégorisation (M-1 vs M)",
            df_tab5,
            style_ctx=style_context.get("tab5", {}),
            comment=_comment_for(comments_df, "Catégorisation", "")
        )

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue(), None

def dataframe_to_png_bytes(df, title="Tableau", max_rows=20, style_context=None):
    """Convertit un DataFrame en image PNG avec colorations conditionnelles."""
    try:
        import matplotlib.pyplot as plt
    except Exception:
        return None, "matplotlib non installé (pip install matplotlib)."

    if df is None or df.empty:
        return None, "Aucune donnée à convertir en image."

    style_context = style_context or {}
    taux_cols = [c for c in style_context.get("taux_cols", []) if c in df.columns]
    score_cols = [c for c in style_context.get("score_cols", []) if c in df.columns]
    note_cols = [c for c in style_context.get("note_cols", []) if c in df.columns]
    int_cols = [c for c in style_context.get("int_cols", []) if c in df.columns]
    highlight_max_cols = [c for c in style_context.get("highlight_max_cols", []) if c in df.columns]
    ratio_alert_col = style_context.get("ratio_alert_col")
    ratio_alert_threshold = float(style_context.get("ratio_alert_threshold", 10.0))
    decimals = int(style_context.get("decimals", 2))

    df_show = df.head(max_rows).copy()
    max_values = {}
    for c in highlight_max_cols:
        s = pd.to_numeric(df_show[c], errors='coerce')
        max_values[c] = s.max() if not s.isna().all() else np.nan

    cell_text = []
    for _, row in df_show.iterrows():
        row_text = []
        for col in df_show.columns:
            val = row[col]
            if pd.isna(val):
                row_text.append("")
            elif col in int_cols:
                row_text.append(f"{pd.to_numeric(val, errors='coerce'):.0f}")
            elif isinstance(val, (float, np.floating, int, np.integer)):
                row_text.append(f"{float(val):.{decimals}f}")
            else:
                row_text.append(str(val))
        cell_text.append(row_text)

    fig_h = max(2.8, 0.45 * (len(df_show) + 2))
    fig_w = max(10, 1.5 * len(df_show.columns))
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.axis('off')
    ax.set_title(title, fontsize=11, pad=8)

    table = ax.table(
        cellText=cell_text,
        colLabels=df_show.columns,
        cellLoc='center',
        loc='center'
    )
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.scale(1, 1.2)

    # Style entête
    for j in range(len(df_show.columns)):
        header_cell = table[(0, j)]
        header_cell.set_facecolor("#1F4E78")
        header_cell.set_edgecolor("#FFFFFF")
        header_cell.get_text().set_color("white")
        header_cell.get_text().set_weight("bold")

    # Style cellules data
    for i in range(len(df_show)):
        for j, col in enumerate(df_show.columns):
            val = df_show.iloc[i, j]
            cell = table[(i + 1, j)]
            cell.set_edgecolor("#D9D9D9")
            text_color = "#000000"
            bg_color = "#FFFFFF"
            bold = False

            if col in taux_cols:
                text_color, bg_color, bold_taux = _extract_css_values(style_taux(val))
                bold = bold or bold_taux
            if col in score_cols:
                text_color, bg_color, bold_score = _extract_css_values(style_score(val))
                bold = bold or bold_score
            if col in note_cols:
                text_color, bg_color, bold_note = _extract_css_values(style_note_20(val))
                bold = bold or bold_note

            if col in highlight_max_cols and pd.notna(max_values.get(col)):
                val_num = pd.to_numeric(val, errors='coerce')
                if pd.notna(val_num) and np.isclose(float(val_num), float(max_values[col])):
                    bg_color = "#ffcccc"

            if ratio_alert_col and col == ratio_alert_col:
                val_num = pd.to_numeric(val, errors='coerce')
                if pd.notna(val_num) and float(val_num) > ratio_alert_threshold:
                    text_color = "red"
                    bold = True

            cell.set_facecolor(bg_color)
            cell.get_text().set_color(text_color)
            if bold:
                cell.get_text().set_weight("bold")

    fig.tight_layout()
    output = io.BytesIO()
    fig.savefig(output, format='png', dpi=180, bbox_inches='tight')
    plt.close(fig)
    return output.getvalue(), None

def figure_to_png_bytes(fig):
    """Convertit un graphique Plotly en PNG (retourne bytes, error)."""
    if fig is None:
        return None, "Figure indisponible."
    try:
        return fig.to_image(format="png", scale=2), None
    except Exception:
        return None, "Export image Plotly indisponible (installer kaleido: pip install kaleido)."

def _comment_for(comments_df, contains_text, default_text):
    if comments_df is None or comments_df.empty:
        return default_text
    mask = comments_df["Element"].astype(str).str.contains(contains_text, case=False, na=False)
    if mask.any():
        return str(comments_df.loc[mask, "Commentaire"].iloc[0])
    return default_text

def _add_image_slide(prs, title, comment, image_bytes):
    from pptx.util import Inches, Pt

    slide = prs.slides.add_slide(prs.slide_layouts[5])  # Title only
    slide.shapes.title.text = title

    tx_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.7), Inches(12.3), Inches(1.2))
    tx_frame = tx_box.text_frame
    tx_frame.text = comment
    tx_frame.word_wrap = True
    tx_frame.paragraphs[0].font.size = Pt(12)

    if image_bytes is not None:
        slide.shapes.add_picture(io.BytesIO(image_bytes), Inches(0.5), Inches(1.9), width=Inches(12.3))

def to_powerpoint_dashboard_report(
    df_synth,
    comments_df,
    fig_comp=None,
    fig_quad=None,
    df_comparatif=None,
    df_tab4_fusion=None,
    df_final_c=None,
    df_final_p=None,
    df_tab5=None,
    df_base=None,
    df_actual_detail=None,
    df_expected_detail=None,
    df_top5=None,
    df_flop5=None,
    style_context=None
):
    """Exporte un PowerPoint avec graphiques + tableaux colorés + commentaires clairs."""
    try:
        from pptx import Presentation
    except Exception:
        return None, "python-pptx non installé (pip install python-pptx)."

    style_context = style_context or {}
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = "Rapport Dashboard SNIS RDC"
    slide.placeholders[1].text = f"Généré le {datetime.now().strftime('%Y-%m-%d %H:%M')}"

    # 1) Graphique complétude
    img_comp, _ = figure_to_png_bytes(fig_comp)
    _add_image_slide(
        prs,
        "Graphique Classement Complétude",
        _comment_for(comments_df, "Graphique Classement", "Classement des unités par complétude globale."),
        img_comp
    )

    # 1-bis) Tableau complétude coloré
    if df_final_c is not None and not df_final_c.empty:
        img_tab_comp, _ = dataframe_to_png_bytes(
            df_final_c,
            title="Tableau Complétude",
            style_context=style_context.get("completude", {})
        )
        _add_image_slide(
            prs,
            "Tableau Complétude (coloré)",
            _comment_for(comments_df, "Tableau Complétude", "Lecture détaillée de la complétude par unité."),
            img_tab_comp
        )

    # 1-ter) Base de données
    if df_base is not None and not df_base.empty:
        img_base, _ = dataframe_to_png_bytes(
            df_base,
            title="Base de données",
            style_context=style_context.get("base", {})
        )
        _add_image_slide(
            prs,
            "Base de données",
            _comment_for(comments_df, "Base de données", "Données sources utilisées après filtres."),
            img_base
        )

    # 1-quater) Rapports détaillés données réelles
    if df_actual_detail is not None and not df_actual_detail.empty:
        img_actual, _ = dataframe_to_png_bytes(
            df_actual_detail,
            title="Rapports détaillés - Données réelles",
            style_context=style_context.get("actual_detail", {})
        )
        _add_image_slide(
            prs,
            "Rapports détaillés - Données réelles",
            _comment_for(comments_df, "Données réelles", "Détail des actual reports et total Reports_Actual."),
            img_actual
        )

    # 1-quint) Rapports détaillés participants/attendus
    if df_expected_detail is not None and not df_expected_detail.empty:
        img_expected, _ = dataframe_to_png_bytes(
            df_expected_detail,
            title="Rapports détaillés - Participants/Attendus",
            style_context=style_context.get("expected_detail", {})
        )
        _add_image_slide(
            prs,
            "Rapports détaillés - Participants/Attendus",
            _comment_for(comments_df, "Participants/Attendus", "Détail des expected reports et total Reports_Attendu."),
            img_expected
        )

    # 2) Quadrant comparatif
    img_quad, _ = figure_to_png_bytes(fig_quad)
    _add_image_slide(
        prs,
        "Quadrant Complétude / Promptitude",
        _comment_for(comments_df, "Quadrant", "Comparaison de la complétude et de la promptitude."),
        img_quad
    )

    # 2-bis) Tableau promptitude coloré
    if df_final_p is not None and not df_final_p.empty:
        img_tab_prom, _ = dataframe_to_png_bytes(
            df_final_p,
            title="Tableau Promptitude",
            style_context=style_context.get("promptitude", {})
        )
        _add_image_slide(
            prs,
            "Tableau Promptitude (coloré)",
            _comment_for(comments_df, "Tableau Promptitude", "Lecture détaillée de la promptitude par unité."),
            img_tab_prom
        )

    # 3) Tableau comparatif complétude/promptitude
    img_comp_table, _ = dataframe_to_png_bytes(
        df_comparatif,
        title="Tableau comparatif Complétude / Promptitude",
        style_context=style_context.get("comparatif", {})
    )
    _add_image_slide(
        prs,
        "Tableau comparatif",
        _comment_for(comments_df, "Tableau comparatif", "Lecture comparative par unité."),
        img_comp_table
    )

    # 4) Tableau fusionné zone filtrée (si disponible)
    if df_tab4_fusion is not None and not df_tab4_fusion.empty:
        img_fusion, _ = dataframe_to_png_bytes(
            df_tab4_fusion,
            title="Tableau fusionné indicateurs dataset",
            style_context=style_context.get("fusion", {})
        )
        _add_image_slide(
            prs,
            "Tableau fusionné zone filtrée",
            _comment_for(comments_df, "fusionné", "Vue détaillée des indicateurs dataset en zone filtrée."),
            img_fusion
        )

    # 4-bis) Top 5 complétude
    if df_top5 is not None and not df_top5.empty:
        img_top5, _ = dataframe_to_png_bytes(
            df_top5,
            title="Top 5 Complétude",
            style_context=style_context.get("top5", {})
        )
        _add_image_slide(
            prs,
            "Top 5 Complétude",
            _comment_for(comments_df, "Top 5 Complétude", "Unités les plus performantes en complétude."),
            img_top5
        )

    # 4-ter) Flop 5 promptitude
    if df_flop5 is not None and not df_flop5.empty:
        img_flop5, _ = dataframe_to_png_bytes(
            df_flop5,
            title="Flop 5 Promptitude",
            style_context=style_context.get("flop5", {})
        )
        _add_image_slide(
            prs,
            "Flop 5 Promptitude",
            _comment_for(comments_df, "Flop 5 Promptitude", "Unités prioritaires à renforcer en promptitude."),
            img_flop5
        )

    # 5) Tableau catégorisation (si disponible)
    if df_tab5 is not None and not df_tab5.empty:
        img_tab5, _ = dataframe_to_png_bytes(
            df_tab5,
            title="Tableau Catégorisation (M-1 vs M)",
            style_context=style_context.get("tab5", {})
        )
        _add_image_slide(
            prs,
            "Tableau Catégorisation (coloré)",
            _comment_for(comments_df, "Catégorisation", "Lecture des violations, corrections et ratio /100 rapports."),
            img_tab5
        )
    else:
        _add_image_slide(
            prs,
            "Résultats des règles de validation",
            "Aucun résultat à afficher. Lancez d'abord l'analyse des violations dans l'onglet 5.",
            None
        )

    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = "Synthèse"
    comp_avg = float(df_synth['Complétude_Globale (%)'].mean()) if not df_synth.empty else 0.0
    prom_avg = float(df_synth['Promptitude_Globale (%)'].mean()) if not df_synth.empty else 0.0
    s.placeholders[1].text = (
        f"Complétude moyenne: {comp_avg:.2f}%\n"
        f"Promptitude moyenne: {prom_avg:.2f}%\n"
        f"Unités analysées: {len(df_synth)}"
    )

    output = io.BytesIO()
    prs.save(output)
    return output.getvalue(), None

# Configuration pour téléchargement d'images Plotly
plotly_config = {
    'toImageButtonOptions': {
        'format': 'png',
        'filename': 'export_graphique_snis',
        'scale': 2
    }
}

# --- 2. FONCTIONS API DHIS2 ---
def _read_config_value(key):
    """Lit une valeur de config depuis Streamlit secrets, sinon variables d'environnement."""
    value = None
    try:
        value = st.secrets.get(key)
    except Exception:
        value = None
    if value is None or str(value).strip() == "":
        value = os.getenv(key)
    return str(value).strip() if value is not None else None

def _read_int_config_value(key, default_value):
    """Lit un entier de config; retourne default_value si absent/invalide."""
    raw_value = _read_config_value(key)
    if raw_value is None or str(raw_value).strip() == "":
        return default_value
    try:
        parsed = int(str(raw_value).strip())
        return parsed if parsed > 0 else default_value
    except Exception:
        return default_value

def test_dhis2_credentials(base_url, username, password):
    """Teste les identifiants utilisateur DHIS2 sur /api/me."""
    try:
        response = requests.get(
            f"{base_url.rstrip('/')}/api/me",
            auth=(username, password),
            timeout=10
        )
    except Exception as e:
        return False, f"Erreur de connexion : {e}"

    if response.status_code == 200:
        return True, None
    if response.status_code in (401, 403):
        return False, "Identifiants incorrects. Veuillez réessayer."
    return False, f"Connexion DHIS2 échouée (HTTP {response.status_code})."

@st.cache_resource
def _build_dhis2_client(url, username, password, timeout_connect=10, timeout_read=90, retries=2, backoff=1.0):
    class SimpleDHIS2Client:
        """Client DHIS2 minimal sans dépendance openhexa."""
        def __init__(self, base_url, user, pwd, timeout_connect_sec=10, timeout_read_sec=90, retries_count=2, backoff_factor=1.0):
            self.base_url = str(base_url).rstrip("/")
            self.timeout = (timeout_connect_sec, timeout_read_sec)
            self.session = requests.Session()
            self.session.auth = (user, pwd)
            self.session.headers.update({"Accept": "application/json"})

            retry_strategy = Retry(
                total=retries_count,
                connect=retries_count,
                read=retries_count,
                status=retries_count,
                backoff_factor=backoff_factor,
                status_forcelist=[429, 500, 502, 503, 504],
                allowed_methods=["GET", "HEAD", "OPTIONS"],
                raise_on_status=False,
            )
            adapter = HTTPAdapter(max_retries=retry_strategy)
            self.session.mount("https://", adapter)
            self.session.mount("http://", adapter)

        def _request(self, method, endpoint, params=None, data=None):
            clean_endpoint = str(endpoint).lstrip("/")
            if not clean_endpoint.startswith("api/"):
                clean_endpoint = f"api/{clean_endpoint}"
            url = f"{self.base_url}/{clean_endpoint}"
            response = self.session.request(method, url, params=params, json=data, timeout=self.timeout)
            response.raise_for_status()
            if not response.text:
                return {}
            return response.json()

        def get(self, endpoint, params=None):
            return self._request("GET", endpoint, params=params)

        def post(self, endpoint, data=None):
            return self._request("POST", endpoint, data=data)

    return SimpleDHIS2Client(url, username, password, timeout_connect, timeout_read, retries, backoff)

def get_dhis2_client():
    dhis2_url = _read_config_value("DHIS2_URL")
    dhis2_user = str(st.session_state.get("dhis2_user", "")).strip()
    dhis2_pass = str(st.session_state.get("dhis2_pass", "")).strip()
    dhis2_timeout_connect = _read_int_config_value("DHIS2_TIMEOUT_CONNECT", 10)
    dhis2_timeout_read = _read_int_config_value("DHIS2_TIMEOUT_READ", 90)
    dhis2_retries = _read_int_config_value("DHIS2_HTTP_RETRIES", 2)

    missing_keys = []
    if not dhis2_url:
        missing_keys.append("DHIS2_URL")
    if not dhis2_user:
        missing_keys.append("Nom d'utilisateur DHIS2")
    if not dhis2_pass:
        missing_keys.append("Mot de passe DHIS2")
    if missing_keys:
        missing_text = ", ".join(missing_keys)
        raise RuntimeError(
            f"Configuration DHIS2 manquante: {missing_text}. "
            "Ajoute DHIS2_URL dans Streamlit Cloud (App settings > Secrets), "
            "puis connecte-toi via la barre latérale."
        )

    return _build_dhis2_client(
        dhis2_url,
        dhis2_user,
        dhis2_pass,
        timeout_connect=dhis2_timeout_connect,
        timeout_read=dhis2_timeout_read,
        retries=dhis2_retries,
        backoff=1.0,
    )

@st.cache_data(ttl=3600)
def get_data(favori_id, period=None, cache_user=None, ou_dimension=None):
    _ = cache_user
    if not favori_id: return None
    try:
        dhis = get_dhis2_client()
        if ou_dimension:
            visualization = dhis.get(
                f"visualizations/{favori_id}",
                params={"fields": "*,dataDimensionItems[*,reportingRate[*,dataSet[id,displayName,name]]]"},
            )
            params, labels = build_analytics_query(visualization, period, ou_dimension)
            return analytics_to_reporting_frame(dhis.get("analytics.json", params=params), labels)
        endpoint = f"visualizations/{favori_id}/data.json"
        if period:
            endpoint += f"?dimension=pe:{period}"

        response = dhis.get(endpoint)
        headers, rows = response.get("headers", []), response.get("rows", [])
        columns = [h.get("name", h.get("column")) for h in headers]
        return pd.DataFrame(rows, columns=columns)
    except ValueError as e:
        st.error(f"Impossible de préparer les données du favori pour ce périmètre : {e}")
        return None
    except requests.exceptions.ReadTimeout:
        st.error(
            "Erreur de connexion : délai de lecture dépassé sur DHIS2. "
            "Essaie une période plus courte (ex. 1-2 mois) ou augmente "
            "DHIS2_TIMEOUT_READ dans les secrets Streamlit."
        )
        return None
    except requests.exceptions.Timeout:
        st.error(
            "Erreur de connexion : timeout réseau DHIS2. "
            "Vérifie la connectivité et réessaie."
        )
        return None
    except Exception as e:
        st.error(f"Erreur de connexion : {e}")
        return None

@st.cache_data(ttl=3600)
def get_organisation_catalog(cache_user=None):
    """Catalogue des pays, provinces et zones, indépendant des lignes du favori."""
    _ = cache_user
    try:
        response = get_dhis2_client().get(
            "organisationUnits",
            params={
                "fields": "id,displayName,level,ancestors[id,displayName,level]",
                "filter": "level:le:3",
                "paging": "false",
            },
        )
        units = response.get("organisationUnits")
        if not isinstance(units, list):
            return None
        return build_geography_catalog(units)
    except Exception:
        return None

@st.cache_data(ttl=3600)
def get_validation_groups(cache_user=None):
    _ = cache_user
    try:
        dhis = get_dhis2_client()
        response = dhis.get("validationRuleGroups", params={"fields": "id,displayName", "paging": "false"})
        groups = response.get('validationRuleGroups')
        if not isinstance(groups, list) or any(
            not isinstance(group, dict) or not group.get('id') or not group.get('displayName')
            for group in groups
        ):
            return None
        return groups
    except Exception:
        return None

@st.cache_data(ttl=3600)
def get_validation_results(ou_id, period_list, group_id, cache_user=None, start_date=None, end_date=None):
    _ = cache_user
    if not ou_id or ou_id == "USER_ORGUNIT" or len(str(ou_id)) < 5:
        result = pd.DataFrame()
        result.attrs['failed_periods'] = list(period_list)
        result.attrs['errors'] = {pe: ["Identifiant de l'unité parent indisponible"] for pe in period_list}
        return result
    try:
        dhis = get_dhis2_client()
    except Exception:
        result = pd.DataFrame()
        result.attrs['failed_periods'] = list(period_list)
        result.attrs['errors'] = {pe: ["Connexion DHIS2 non configurée pour cette session"] for pe in period_list}
        return result
    return fetch_validation_analysis(
        dhis, ou_id, period_list, group_id, start_date=start_date, end_date=end_date
    )

@st.cache_data(ttl=3600)
def get_validation_detail_metadata(rule_ids, org_ids, cache_user=None):
    """Complète les libellés et explications par lots, sans bloquer les résultats."""
    _ = cache_user
    rules, units, errors = {}, {}, []
    if not rule_ids and not org_ids:
        return rules, units, errors
    try:
        dhis = get_dhis2_client()
    except Exception:
        return rules, units, ["Les métadonnées des violations sont indisponibles."]
    requests_metadata = [
        ('validationRules', rule_ids, 'id,displayName,description,instruction,importance,operator,leftSide[description,expression],rightSide[description,expression]', rules),
        ('organisationUnits', org_ids, 'id,displayName,path,ancestors[id,displayName,level]', units),
    ]
    for endpoint, identifiers, fields, target in requests_metadata:
        label = 'règles de validation' if endpoint == 'validationRules' else 'unités d’organisation'
        identifiers = sorted({clean_ou_id(value) for value in (identifiers or []) if clean_ou_id(value)})
        for start in range(0, len(identifiers), 60):
            chunk = identifiers[start:start + 60]
            try:
                response = dhis.get(endpoint, params={
                    'fields': fields, 'filter': f"id:in:[{','.join(chunk)}]", 'paging': 'false'
                })
                items = response.get(endpoint)
                if not isinstance(items, list):
                    raise ValueError("Réponse de métadonnées invalide")
                for item in items:
                    if isinstance(item, dict) and clean_ou_id(item.get('id')):
                        target[clean_ou_id(item['id'])] = item
                if any(identifier not in target for identifier in chunk):
                    errors.append(f"Certaines informations sur les {label} sont indisponibles.")
            except Exception:
                errors.append(f"Les informations sur les {label} sont indisponibles pour {len(chunk)} élément(s).")
    return rules, units, list(dict.fromkeys(errors))

@st.cache_data(ttl=3600)
def get_children_org_units_details(parent_id, cache_user=None):
    _ = cache_user
    """Récupère les enfants directs (id + nom) d'une unité d'organisation."""
    if not parent_id:
        return []
    try:
        dhis = get_dhis2_client()
        response = dhis.get(
            f"organisationUnits/{parent_id}",
            params={"fields": "children[id,displayName]", "paging": "false"}
        )
        return response.get("children", []) or []
    except Exception:
        return []

@st.cache_data(ttl=3600)
def get_org_units_hierarchy(org_unit_ids, cache_user=None):
    """
    Récupère la hiérarchie organisationnelle (pays/province/zone) pour une liste d'IDs.
    Retour: dict[ou_id] -> {
        country_name, country_id, province_name, province_id, zone_name, zone_id, level
    }
    """
    _ = cache_user
    cleaned_ids = []
    for ou_id in (org_unit_ids or []):
        if ou_id is None:
            continue
        ou_value = str(ou_id).strip()
        if not ou_value or ou_value.lower() in {"nan", "none", "null"}:
            continue
        cleaned_ids.append(ou_value)
    unique_ids = sorted(set(cleaned_ids))
    if not unique_ids:
        return {}

    try:
        dhis = get_dhis2_client()
    except Exception:
        return {}

    hierarchy = {}
    chunk_size = 60
    for i in range(0, len(unique_ids), chunk_size):
        chunk = unique_ids[i:i + chunk_size]
        try:
            params = {
                "fields": "id,displayName,level,ancestors[id,displayName,level]",
                "filter": f"id:in:[{','.join(chunk)}]",
                "paging": "false",
            }
            response = dhis.get("organisationUnits", params=params)
            units = response.get("organisationUnits", [])
        except Exception:
            units = []

        hierarchy.update(resolve_hierarchy(units))

    return hierarchy

# --- 3. SIDEBAR & FILTRES ---
with st.sidebar:
    st.image("https://snisrdc.com/dhis-web-commons/security/logo_front.png", width=150)
    st.title("⚙️ Configuration")

    st.subheader("🔐 Connexion DHIS2")
    base_url = _read_config_value("DHIS2_URL")
    if not base_url:
        st.error("DHIS2_URL manquant. Ajoute-le dans les secrets Streamlit Cloud.")
        st.stop()

    user_input = st.text_input(
        "Nom d'utilisateur",
        value=str(st.session_state.get("dhis2_user", "")),
        key="dhis2_user_input"
    )
    password_input = st.text_input(
        "Mot de passe",
        type="password",
        key="dhis2_password_input"
    )

    col_conn, col_disc = st.columns(2)
    login_button = col_conn.button("Se connecter", use_container_width=True)
    logout_button = col_disc.button("Se déconnecter", use_container_width=True)

    if logout_button:
        for key in ["connected", "dhis2_user", "dhis2_pass", "dhis2_user_input", "dhis2_password_input", "categorisation_result", "data_revision", "show_report_preview", "validation_defaults_context"]:
            st.session_state.pop(key, None)
        st.cache_resource.clear()
        st.rerun()

    if login_button:
        if not user_input or not password_input:
            st.session_state["connected"] = False
            st.warning("Renseigne le nom d'utilisateur et le mot de passe.")
        else:
            ok, error_msg = test_dhis2_credentials(base_url, user_input, password_input)
            if ok:
                st.cache_resource.clear()
                st.session_state["connected"] = True
                st.session_state["dhis2_user"] = user_input
                st.session_state["dhis2_pass"] = password_input
                st.session_state['data_revision'] = uuid4().hex
                st.session_state.pop('categorisation_result', None)
                st.session_state.pop('show_report_preview', None)
                st.session_state.pop('validation_defaults_context', None)
                st.success(f"Connecté en tant que : {user_input}")
                st.rerun()
            else:
                st.session_state["connected"] = False
                st.error(error_msg)

    if st.session_state.get("connected"):
        st.success(f"Connecté en tant que : {st.session_state.get('dhis2_user')}")
        if 'data_revision' not in st.session_state:
            st.session_state['data_revision'] = uuid4().hex
        if st.button(
            "Actualiser les données", key="refresh_data",
            help="Recharge les données DHIS2 pour cette session. L'analyse des violations devra être relancée."
        ):
            st.session_state['data_revision'] = uuid4().hex
            st.session_state.pop('categorisation_result', None)
            st.session_state.pop('show_report_preview', None)
    else:
        st.info("Veuillez vous connecter via la barre latérale pour accéder aux données.")
        st.stop()

    st.divider()

    dict_favoris = {
        "Performance Globale (ROzCY14OLTE)": "ROzCY14OLTE",
        "Analyse Promptitude (mldsgxAvIIi)": "mldsgxAvIIi",
        "Rapport TyAJb0qitMz": "TyAJb0qitMz",
        "Autre ID personnalisé": "CUSTOM"
    }
    choix_fav = st.selectbox("Rapport DHIS2 :", list(dict_favoris.keys()))
    id_final = st.text_input("Saisissez l'ID :") if choix_fav == "Autre ID personnalisé" else dict_favoris[choix_fav]

    st.divider()
    st.subheader("📅 Sélection de la Période")
    annee_choisie = st.selectbox("Choisir l'année :", ["2026", "2025", "2024", "2023"], index=1)

    mois_liste = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre"]
    mois_dict = {m: str(i+1).zfill(2) for i, m in enumerate(mois_liste)}

    col_m1, col_m2 = st.columns(2)
    with col_m1: mois_debut = st.selectbox("Mois début :", mois_liste, index=0)
    with col_m2: mois_fin = st.selectbox("Mois fin :", mois_liste, index=0)

    idx_start, idx_end = mois_liste.index(mois_debut), mois_liste.index(mois_fin)

    # Valeurs par défaut pour éviter des variables non définies.
    selection_mois = [f"{annee_choisie}{mois_dict[mois_debut]}"]
    period_id = selection_mois[0]
    prev_idx = idx_start - 1 if idx_start > 0 else 11
    prev_year = annee_choisie if idx_start > 0 else str(int(annee_choisie)-1)
    selection_mois_prev = [f"{prev_year}{mois_dict[mois_liste[prev_idx]]}"]

    if idx_start <= idx_end:
        selection_mois = [f"{annee_choisie}{mois_dict[mois_liste[i]]}" for i in range(idx_start, idx_end + 1)]
        period_id = ";".join(selection_mois)
    else:
        st.error("Le mois de début doit être avant le mois de fin.")
        st.session_state.pop('categorisation_result', None)
        st.session_state.pop('show_report_preview', None)
        st.stop()

# --- 4. FILTRES PAYS / PROVINCE / ZONE ET CHARGEMENT AU NIVEAU ZS ---
current_user_for_cache = (
    base_url, st.session_state.get("dhis2_user", ""), st.session_state['data_revision']
)
organisation_catalog = get_organisation_catalog(cache_user=current_user_for_cache)
if organisation_catalog is None:
    st.error("Impossible de charger la liste des provinces et zones de santé depuis DHIS2. Réessayez la connexion.")
    st.stop()
if organisation_catalog.empty:
    st.info("Aucun pays, province ou zone de santé accessible avec cette connexion.")
    st.stop()

st.sidebar.subheader("🔍 Filtrage")

def _select_scope_id(label, options, labels, key, surface=None, on_change=None):
    if st.session_state.get(key) not in options:
        st.session_state.pop(key, None)
    surface = st.sidebar if surface is None else surface
    return surface.selectbox(label, options, format_func=lambda item: labels.get(item, item), key=key, on_change=on_change)

country_rows = organisation_catalog[organisation_catalog['Pays_ID'].map(clean_ou_id).ne('')]
country_labels = labels_for_units(country_rows, 'Pays_ID', 'Pays')
country_options = sorted(country_labels, key=lambda item: country_labels[item])
if not country_options:
    st.info("Le pays de rattachement des unités est indisponible dans la hiérarchie DHIS2.")
    st.stop()
selected_country_id = (
    _select_scope_id("Pays :", country_options, country_labels, "filter_country_main")
    if len(country_options) > 1 else country_options[0]
)
country_label = country_labels[selected_country_id]
if len(country_options) == 1:
    st.sidebar.caption(f"Pays : {country_label}")

if st.session_state.get('filter_level_main') == 'Synthèse pays':
    st.session_state['filter_level_main'] = 'Pays'
filter_level = st.sidebar.selectbox(
    "Niveau d'analyse :", ["Pays", "Province", "Zone de santé"], key="filter_level_main"
)
selected_province_id = None
selected_zone_id = None
selected_province = "Toutes les provinces"
selected_zone = "Toutes les zones"
selected_scope_label = f"Pays {country_label}"
current_id_systeme = selected_country_id
country_catalog = organisation_catalog[organisation_catalog['Pays_ID'].eq(selected_country_id)]

if filter_level in ('Province', 'Zone de santé'):
    province_rows = country_catalog[country_catalog['Province_ID'].map(clean_ou_id).ne('')]
    province_labels = labels_for_units(province_rows, 'Province_ID', 'Province')
    province_options = sorted(province_labels, key=lambda item: province_labels[item])
    if not province_options:
        st.info("Aucune province accessible pour ce pays.")
        st.stop()
    selected_province_id = _select_scope_id(
        "Province :", province_options, province_labels, "filter_province_main"
    )
    selected_province = province_labels[selected_province_id]
    selected_scope_label = f"Province {selected_province}"
    current_id_systeme = selected_province_id

selected_zones_df = zone_catalog(
    organisation_catalog, country_id=selected_country_id, province_id=selected_province_id
)
if selected_zones_df.empty:
    st.info("Aucune zone de santé accessible pour ce périmètre.")
    st.stop()

if filter_level in ('Province', 'Zone de santé'):
    zone_labels = labels_for_units(selected_zones_df, 'Zone_ID', 'Zone')
    zone_options = sorted(zone_labels, key=lambda item: zone_labels[item])
    if filter_level == 'Province':
        zone_options = [''] + zone_options
        zone_labels[''] = 'Toutes les zones'
    selected_zone_id = _select_scope_id("Zone de santé :", zone_options, zone_labels, "filter_zone_main") or None
    if selected_zone_id:
        selected_zone = zone_labels[selected_zone_id]
        current_id_systeme = selected_zone_id
        selected_scope_label = f"Zone de santé {selected_zone} ({selected_province})"
        selected_zones_df = zone_catalog(
            organisation_catalog, country_id=selected_country_id,
            province_id=selected_province_id, zone_id=selected_zone_id
        )

is_zone_focus = selected_zone_id is not None
ou_dimension = selected_zone_id or f"LEVEL-3;{selected_province_id or selected_country_id}"
st.sidebar.caption(f"{len(selected_zones_df)} zone(s) de santé dans le périmètre sélectionné.")

df_raw = get_data(id_final, period=period_id, cache_user=current_user_for_cache, ou_dimension=ou_dimension)
if df_raw is not None:
    df_raw = normalize_orgunit_columns(df_raw)
    if 'Organisation unit ID' not in df_raw.columns:
        st.error("La réponse DHIS2 ne contient pas les identifiants des zones de santé.")
        st.stop()
    df_raw['Organisation unit ID'] = df_raw['Organisation unit ID'].map(clean_ou_id)
    df_raw = df_raw[df_raw['Organisation unit ID'].isin(selected_zones_df['Zone_ID'])].copy()
    display_labels = labels_for_units(selected_zones_df, 'Zone_ID', 'Zone')
    df_raw['Organisation unit'] = df_raw['Organisation unit ID'].map(display_labels)

if df_raw is not None:
    target_df_all = df_raw.copy()
    default_cols = ['Organisation unit code', 'Organisation unit description',
                    'Reporting month', 'Organisation unit parameter', 'Organisation unit is parent']
    target_df_all = target_df_all.drop(columns=default_cols, errors='ignore')
    target_df = target_df_all.copy()
    mapping_ou_id = dict(zip(df_raw['Organisation unit'], df_raw['Organisation unit ID']))
    is_parent_by_ou = {}

    catalog_by_id = organisation_catalog.drop_duplicates('Organisation unit ID').set_index('Organisation unit ID')
    scope_df = catalog_by_id.reindex(df_raw['Organisation unit ID']).reset_index()
    scope_df.index = df_raw.index
    scope_df['Organisation unit'] = df_raw['Organisation unit']
    hierarchy_by_id = {
        row['Organisation unit ID']: {
            'country_name': row['Pays'], 'country_id': row['Pays_ID'],
            'province_name': row['Province'], 'province_id': row['Province_ID'],
            'zone_name': row['Zone'], 'zone_id': row['Zone_ID'], 'level': row['level'],
        }
        for row in organisation_catalog.to_dict('records')
    }

    # Garde-fou final: garantir la colonne 'Organisation unit' après tous les filtres.
    if "Organisation unit" not in target_df_all.columns:
        target_df_all = normalize_orgunit_columns(target_df_all)
    if "Organisation unit" not in target_df.columns:
        target_df = normalize_orgunit_columns(target_df)

    if "Organisation unit" not in target_df.columns:
        fallback_org_col = _find_best_column(
            target_df.columns.tolist(),
            include_terms=["organisation", "unit"]
        ) or _find_best_column(
            target_df.columns.tolist(),
            include_terms=["org", "unit"]
        )

        if fallback_org_col and fallback_org_col in target_df.columns:
            target_df = target_df.rename(columns={fallback_org_col: "Organisation unit"})
        elif target_df.shape[1] > 0:
            # Dernier recours: créer la colonne à partir de la première colonne disponible.
            target_df = target_df.copy()
            target_df["Organisation unit"] = target_df.iloc[:, 0].astype(str)
            st.warning("La colonne 'Organisation unit' n'était pas fournie par le favori. Un fallback a été appliqué.")
        else:
            st.error("Impossible d'identifier les unités d'organisation dans ce favori DHIS2.")
            st.stop()

    if target_df.empty:
        st.warning("Aucune donnée trouvée avec ce niveau de filtre.")


    st.title(f"📊 SNIS RDC - Performance {selected_scope_label}")

    # Variables partagées pour export hors onglets.
    fig_comp = None
    fig_quad = None
    df_tab4_fusion = pd.DataFrame()
    df_comparatif = pd.DataFrame()
    df_tab5_export = pd.DataFrame()
    df_violations_export = None
    df_validation_context_export = None
    validation_applied_label = None
    df_tab1_export = pd.DataFrame()
    df_actual_detail_export = pd.DataFrame()
    df_expected_detail_export = pd.DataFrame()
    df_top5_export = pd.DataFrame()
    df_flop5_export = pd.DataFrame()

    tab1, tab2, tab3, tab4, tab5 = st.tabs(["📁 Base de données", "✅ Complétude", "⏱️ Promptitude", "⚖️ Analyse Comparative", "🩺 Éléments de catégorisation"])

    # --- ONGLET 1 : BASE DE DONNÉES ---
    with tab1:
        st.header("Présentation des données brutes")
        st.write(f"Période affichée : **{mois_debut} à {mois_fin} {annee_choisie}**")
        df_tab1 = target_df.copy()
        if 'Organisation unit ID' in df_tab1.columns:
            df_tab1 = df_tab1.drop(columns=['Organisation unit ID'])
        if is_parent_by_ou and 'Organisation unit' in df_tab1.columns:
            df_tab1 = df_tab1[~df_tab1['Organisation unit'].map(is_parent_by_ou).fillna(False)]

        if df_tab1.empty:
            st.warning("Aucune zone de santé à afficher avec ce filtre.")
        else:
            st.dataframe(df_tab1.head(15), use_container_width=True)
        df_tab1_export = df_tab1.copy()

    # --- ONGLET 2 : COMPLÉTUDE ---
    with tab2:
        st.header("Analyse de la Complétude")
        col_actual_strict = [c for c in target_df.columns if 'actual reports' in c.lower() and 'time' not in c.lower()]
        if col_actual_strict:
            df_actual = target_df[['Organisation unit'] + col_actual_strict].copy()
            for col in col_actual_strict: df_actual[col] = pd.to_numeric(df_actual[col], errors='coerce').fillna(0)
            df_actual['Reports_Actual'] = df_actual[col_actual_strict].sum(axis=1).round(2)
            with st.expander("Détail Reports Actuals"): st.dataframe(df_actual.head(15))
            df_actual_detail_export = df_actual.copy()

        col_expected_strict = [c for c in target_df.columns if 'expected reports' in c.lower() and 'time' not in c.lower()]
        if col_expected_strict:
            df_expected = target_df[['Organisation unit'] + col_expected_strict].copy()
            for col in col_expected_strict: df_expected[col] = pd.to_numeric(df_expected[col], errors='coerce').fillna(0)
            df_expected['Reports_Attendu'] = df_expected[col_expected_strict].sum(axis=1).round(2)
            with st.expander("Détail Reports Attendus"): st.dataframe(df_expected.head(15))
            df_expected_detail_export = df_expected.copy()

        col_rate_uniquement = [c for c in target_df.columns if 'reporting rate' in c.lower() and 'on time' not in c.lower()]
        df_affichage = target_df[['Organisation unit']].copy()
        if col_rate_uniquement:
            df_affichage = target_df[['Organisation unit'] + col_rate_uniquement].copy()
            for col in col_rate_uniquement: df_affichage[col] = pd.to_numeric(df_affichage[col], errors='coerce').fillna(0).round(2)
            with st.expander(f"Affichage de {len(col_rate_uniquement)} indicateurs"): st.dataframe(df_affichage.head(15), use_container_width=True)

        df_synthese = target_df[['Organisation unit']].copy()
        df_synthese['Reports_Actual'] = target_df[col_actual_strict].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1).round(2) if col_actual_strict else 0
        df_synthese['Reports_Attendu'] = target_df[col_expected_strict].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1).round(2) if col_expected_strict else 0
        df_synthese['Complétude_Globale (%)'] = (df_synthese['Reports_Actual'] / df_synthese['Reports_Attendu'].replace(0, np.nan) * 100).fillna(0).round(2)

        st.subheader("Tableau de Performance Final (Stylisé)")
        df_final_c = pd.merge(df_affichage, df_synthese[['Organisation unit', 'Reports_Actual', 'Reports_Attendu', 'Complétude_Globale (%)']], on='Organisation unit')
        df_final_c['Nombre des data set complétude >/=95%'] = (df_final_c[col_rate_uniquement] >= 95).sum(axis=1) if col_rate_uniquement else 0

        st.dataframe(
            df_final_c.style.format({c: "{:.2f}" for c in df_final_c.columns if c not in ['Organisation unit', 'Nombre des data set complétude >/=95%']})
            .map(style_taux, subset=col_rate_uniquement + ['Complétude_Globale (%)'])
            .map(style_score, subset=['Nombre des data set complétude >/=95%']),
            use_container_width=True
        )

        st.divider()
        fig_comp = px.bar(df_final_c.sort_values('Complétude_Globale (%)'), x='Complétude_Globale (%)', y='Organisation unit', orientation='h', color='Complétude_Globale (%)', color_continuous_scale='RdYlGn', title="Classement des zones")
        st.plotly_chart(fig_comp, use_container_width=True, config=plotly_config, key="chart_comp_tab2")

    # --- ONGLET 3 : PROMPTITUDE ---
    with tab3:
        st.header("Analyse de la Promptitude")
        col_rate_ot = [c for c in target_df.columns if 'reporting rate' in c.lower() and 'on time' in c.lower()]
        actual_ot_cols = [c for c in target_df.columns if 'actual reports' in c.lower() and 'on time' in c.lower()]

        df_final_p = target_df[['Organisation unit']].copy()
        for col in col_rate_ot: df_final_p[col] = pd.to_numeric(target_df[col], errors='coerce').fillna(0).round(2)
        df_final_p['Reports_Actual_On_Time'] = target_df[actual_ot_cols].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1).round(2) if actual_ot_cols else 0
        df_final_p['Reports_Attendu'] = target_df[col_expected_strict].apply(pd.to_numeric, errors='coerce').fillna(0).sum(axis=1).round(2) if col_expected_strict else 0
        df_final_p['Promptitude_Globale (%)'] = (df_final_p['Reports_Actual_On_Time'] / df_final_p['Reports_Attendu'].replace(0, np.nan) * 100).fillna(0).round(2)

        nom_col_score_p = 'Nombre des data set promptitude >/=95%'
        df_final_p[nom_col_score_p] = (df_final_p[col_rate_ot] >= 95).sum(axis=1) if col_rate_ot else 0

        st.dataframe(
            df_final_p.style.format({c: "{:.2f}" for c in df_final_p.columns if c not in ['Organisation unit', nom_col_score_p]})
            .map(style_taux, subset=col_rate_ot + ['Promptitude_Globale (%)'])
            .map(style_score, subset=[nom_col_score_p]), use_container_width=True
        )

    # --- ONGLET 4 : ANALYSE COMPARATIVE ---
    with tab4:
        st.header("⚖️ Analyse Comparative et Performance")
        col_score_comp = 'Nombre des data set complétude >/=95%'
        col_score_prompt = nom_col_score_p
        df_synth = pd.merge(
            df_final_c[['Organisation unit', 'Complétude_Globale (%)', 'Reports_Actual', 'Reports_Attendu', col_score_comp]],
            df_final_p[['Organisation unit', 'Promptitude_Globale (%)', 'Reports_Actual_On_Time', col_score_prompt]],
            on='Organisation unit'
        )
        float_cols_synth = ['Complétude_Globale (%)', 'Promptitude_Globale (%)', 'Reports_Actual', 'Reports_Attendu', 'Reports_Actual_On_Time']
        for c in float_cols_synth:
            if c in df_synth.columns:
                df_synth[c] = pd.to_numeric(df_synth[c], errors='coerce').fillna(0).round(2)

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Complétude moyenne", f"{df_synth['Complétude_Globale (%)'].mean():.2f}%")
        m2.metric("Promptitude moyenne", f"{df_synth['Promptitude_Globale (%)'].mean():.2f}%")
        m3.metric("Total Attendu", f"{df_synth['Reports_Attendu'].sum():.2f}")
        m4.metric("Complétude >= 95%", (df_final_c[col_rate_uniquement] >= 95).sum(axis=1).sum() if col_rate_uniquement else 0)
        m5.metric("Promptitude >= 95%", (df_final_p[col_rate_ot] >= 95).sum(axis=1).sum() if col_rate_ot else 0)

        st.subheader("📊 Tableau comparatif Complétude / Promptitude")
        df_comparatif = (
            df_synth[
                [
                    'Organisation unit',
                    'Complétude_Globale (%)',
                    col_score_comp,
                    'Promptitude_Globale (%)',
                    col_score_prompt
                ]
            ]
            .groupby('Organisation unit', as_index=False)
            .agg({
                'Complétude_Globale (%)': 'mean',
                col_score_comp: 'mean',
                'Promptitude_Globale (%)': 'mean',
                col_score_prompt: 'mean'
            })
            .round(2)
        )
        df_comparatif = df_comparatif.rename(columns={
            'Organisation unit': 'Zone de santé',
            col_score_comp: 'Nombre de dataset complétude >/= 95%',
            col_score_prompt: 'Nombre de dataset promptitude >/= 95%'
        })
        st.dataframe(
            df_comparatif.style.format({
                'Complétude_Globale (%)': '{:.2f}',
                'Promptitude_Globale (%)': '{:.2f}',
                'Nombre de dataset complétude >/= 95%': '{:.2f}',
                'Nombre de dataset promptitude >/= 95%': '{:.2f}'
            }),
            use_container_width=True,
            key="grid_comp_prompt"
        )

        # TABLEAU FUSIONNE (zone filtrée): reporting + actual + scores dataset
        if is_zone_focus:
            st.divider()
            st.subheader(f"📋 Tableau fusionné des indicateurs dataset : {selected_zone}")

            row_zone = target_df.iloc[0] if not target_df.empty else pd.Series(dtype=object)
            fusion_rows = []

            # Lignes reporting rate (complétude/promptitude en % + score >=95)
            for c_col in col_rate_uniquement:
                base_name = c_col.replace('Reporting rate', '').strip()
                p_col = next((p for p in col_rate_ot if base_name in p), None)
                compl_val = pd.to_numeric(row_zone.get(c_col, np.nan), errors='coerce')
                promp_val = pd.to_numeric(row_zone.get(p_col, np.nan), errors='coerce') if p_col else np.nan
                fusion_rows.append({
                    "Indicateurs (dataset)": f"{base_name} - Reporting",
                    "Complétude": round(float(compl_val), 2) if pd.notna(compl_val) else np.nan,
                    "Nombre de dataset complétude >/= 95%": int(compl_val >= 95) if pd.notna(compl_val) else np.nan,
                    "Promptitude": round(float(promp_val), 2) if pd.notna(promp_val) else np.nan,
                    "Nombre de dataset promptitude >/= 95%": int(promp_val >= 95) if pd.notna(promp_val) else np.nan
                })

            # Lignes actual reports (afficher tous les actual)
            for a_col in col_actual_strict:
                base_name = a_col.replace('Actual reports', '').strip()
                aot_col = next((a for a in actual_ot_cols if base_name in a), None)
                actual_val = pd.to_numeric(row_zone.get(a_col, np.nan), errors='coerce')
                actual_ot_val = pd.to_numeric(row_zone.get(aot_col, np.nan), errors='coerce') if aot_col else np.nan
                fusion_rows.append({
                    "Indicateurs (dataset)": f"{base_name} - Actual",
                    "Complétude": round(float(actual_val), 2) if pd.notna(actual_val) else np.nan,
                    "Nombre de dataset complétude >/= 95%": np.nan,
                    "Promptitude": round(float(actual_ot_val), 2) if pd.notna(actual_ot_val) else np.nan,
                    "Nombre de dataset promptitude >/= 95%": np.nan
                })

            df_tab4_fusion = pd.DataFrame(fusion_rows)
            st.dataframe(
                df_tab4_fusion.style.format({
                    "Complétude": "{:.2f}",
                    "Promptitude": "{:.2f}",
                    "Nombre de dataset complétude >/= 95%": "{:.0f}",
                    "Nombre de dataset promptitude >/= 95%": "{:.0f}"
                }),
                use_container_width=True,
                key="grid_fusion_zone"
            )

        st.subheader("Quadrant de Performance")
        fig_quad = px.scatter(df_synth, x='Complétude_Globale (%)', y='Promptitude_Globale (%)', text='Organisation unit', range_x=[0, 110], range_y=[0, 110])
        fig_quad.add_hline(y=80, line_dash="dot", line_color="red")
        fig_quad.add_vline(x=80, line_dash="dot", line_color="red")
        st.plotly_chart(fig_quad, use_container_width=True, key="chart_quad_tab4")

        col_l, col_r = st.columns(2)
        with col_l:
            st.success("🏆 **Top 5 Complétude**")
            df_top5_export = df_synth.nlargest(5, 'Complétude_Globale (%)')[['Organisation unit', 'Complétude_Globale (%)']].round(1)
            st.table(df_top5_export)
        with col_r:
            st.error("⚠️ **Flop 5 Promptitude**")
            df_flop5_export = df_synth.nsmallest(5, 'Promptitude_Globale (%)')[['Organisation unit', 'Promptitude_Globale (%)']].round(1)
            st.table(df_flop5_export)

    # --- ONGLET 5 : ÉLÉMENTS DE CATÉGORISATION ---
    with tab5:
        st.header("🩺 Éléments de catégorisation interactifs")

        vr_groups = get_validation_groups(cache_user=current_user_for_cache)
        if vr_groups is None:
            st.warning("La liste des groupes est indisponible. L'analyse de toutes les règles reste accessible.")
        group_mapping = {g['displayName']: g['id'] for g in (vr_groups or [])}
        options_regles = ["Toutes les règles (Global)"] + list(group_mapping.keys())

        default_start, default_end = monthly_bounds(selection_mois)
        validation_defaults_context = (
            selected_country_id, selected_province_id, selected_zone_id, tuple(selection_mois)
        )
        if (
            st.session_state.get('validation_defaults_context') != validation_defaults_context
            or 'validation_start_date' not in st.session_state or 'validation_end_date' not in st.session_state
        ):
            st.session_state.update({
                'validation_country': selected_country_id,
                'validation_province': selected_province_id or '',
                'validation_zone': selected_zone_id or '',
                'validation_start_date': default_start,
                'validation_end_date': default_end,
                'validation_defaults_context': validation_defaults_context,
            })

        st.subheader("Filtres de l'analyse des règles")
        parent_column, dates_column = st.columns([1, 1.2])
        with parent_column:
            st.markdown("**Unité d’organisation parent**")
            if len(country_options) > 1:
                validation_country_id = _select_scope_id(
                    "Pays de l’analyse", country_options, country_labels, 'validation_country', surface=st,
                    on_change=lambda: st.session_state.update({'validation_province': '', 'validation_zone': ''})
                )
            else:
                validation_country_id = country_options[0]
                st.caption(country_labels[validation_country_id])
            validation_country_catalog = organisation_catalog[
                organisation_catalog['Pays_ID'].eq(validation_country_id)
            ]
            validation_province_labels = labels_for_units(
                validation_country_catalog, 'Province_ID', 'Province'
            )
            validation_province_options = [''] + sorted(
                validation_province_labels, key=lambda item: validation_province_labels[item]
            )
            validation_province_labels[''] = 'Toutes les provinces'
            validation_province_id = _select_scope_id(
                "Province de l’analyse", validation_province_options, validation_province_labels,
                'validation_province', surface=st,
                on_change=lambda: st.session_state.update({'validation_zone': ''})
            ) or None
            validation_zones_df = zone_catalog(
                organisation_catalog, country_id=validation_country_id, province_id=validation_province_id
            )
            validation_zone_labels = labels_for_units(validation_zones_df, 'Zone_ID', 'Zone')
            validation_zone_options = [''] + sorted(validation_zone_labels, key=lambda item: validation_zone_labels[item])
            validation_zone_labels[''] = 'Toutes les zones'
            validation_zone_id = _select_scope_id(
                "Zone de santé de l’analyse", validation_zone_options, validation_zone_labels,
                'validation_zone', surface=st
            ) or None
            if validation_zone_id:
                validation_zones_df = zone_catalog(
                    organisation_catalog, country_id=validation_country_id,
                    province_id=validation_province_id, zone_id=validation_zone_id
                )
            validation_parent_id = validation_zone_id or validation_province_id or validation_country_id
            validation_parent_label = (
                validation_zone_labels[validation_zone_id] if validation_zone_id else
                validation_province_labels[validation_province_id] if validation_province_id else
                country_labels[validation_country_id]
            )
            st.caption(f"Parent sélectionné : {validation_parent_label} — {len(validation_zones_df)} zone(s) de santé.")
        with dates_column:
            validation_start_date = st.date_input(
                "Date de début", value=None, format="YYYY-MM-DD", key='validation_start_date'
            )
            validation_end_date = st.date_input(
                "Date de fin", value=None, format="YYYY-MM-DD", key='validation_end_date'
            )
            selected_vr_name = st.selectbox("Sélectionner le Groupe de Règles de Validation :", options_regles)
        target_group_id = group_mapping.get(selected_vr_name)
        validation_dates_valid = (
            isinstance(validation_start_date, date) and isinstance(validation_end_date, date)
            and validation_start_date <= validation_end_date
        )
        if not validation_dates_valid:
            st.error("Renseignez les deux dates ; la date de début doit être antérieure ou égale à la date de fin.")
        if validation_zones_df.empty:
            st.warning("Aucune zone de santé accessible pour cette unité parent.")
        validation_months = date_range_months(validation_start_date, validation_end_date) if validation_dates_valid else []
        validation_reference_months = [reference_month(validation_start_date)] if validation_dates_valid else []
        validation_period_id = ';'.join(validation_months)
        validation_ou_dimension = validation_zone_id or f"LEVEL-3;{validation_parent_id}"
        profil_bareme = st.selectbox(
            "Barème du nombre de datasets à complétude ≥95 % :",
            ["Automatique", "DPS", "ZS"],
            help="Automatique utilise le niveau DHIS2 : province = DPS, zone de santé = ZS."
        )
        with st.expander("Consulter les barèmes de catégorisation"):
            st.dataframe(build_baremes_df(), hide_index=True, use_container_width=True)
            st.caption(
                "Le score de catégorisation /100 additionne cinq notes sur 20. "
                "Le score de qualité historique et sa note /20 restent basés sur la moyenne "
                "de la complétude et de la promptitude et ne sont pas ajoutés au total."
            )
        validation_context = (
            'dhis2-post-parent-v2', base_url, current_user_for_cache, id_final,
            validation_start_date, validation_end_date, tuple(validation_months),
            tuple(validation_reference_months), validation_parent_id, target_group_id,
            tuple(validation_zones_df['Zone_ID']),
            tuple(target_df.columns),
            int(pd.util.hash_pandas_object(target_df.astype(str), index=True).sum()),
        )
        saved_validation = st.session_state.get("categorisation_result")
        if saved_validation and saved_validation.get("context") != validation_context:
            st.session_state.pop("categorisation_result", None)
        st.caption(
            f"Mois couverts : {', '.join(validation_months) or '—'}. "
            f"Référence M-1 : {', '.join(validation_reference_months) or '—'}. "
            "L'analyse est envoyée à DHIS2 avec le parent et l'intervalle complets. Les indicateurs de déclaration restent mensuels "
            "et couvrent les mois de cet intervalle. Le calcul n'enregistre aucun résultat et n'envoie aucune notification."
        )

        if st.button("Lancer l'analyse des violations", disabled=not validation_dates_valid or validation_zones_df.empty):
            can_run_validation = bool(validation_parent_id)

            if can_run_validation:
                with st.spinner('Analyse des violations en cours...'):

                    # 1) Le catalogue sélectionné contient exclusivement des zones de santé.
                    validation_labels = labels_for_units(validation_zones_df, 'Zone_ID', 'Zone')
                    zones_resolues = [(validation_labels[ou_id], ou_id) for ou_id in validation_zones_df['Zone_ID']]

                    if not zones_resolues:
                        st.warning("Aucune zone de santé valide trouvée.")
                        st.stop()

                    # 2) Déclarations mensuelles du même périmètre et des mois choisis.
                    reporting_warnings = []
                    if validation_ou_dimension == ou_dimension and validation_period_id == period_id:
                        validation_reporting_source = target_df
                    else:
                        validation_reporting_source = get_data(
                            id_final, period=validation_period_id, cache_user=current_user_for_cache,
                            ou_dimension=validation_ou_dimension
                        )
                        if validation_reporting_source is not None:
                            validation_reporting_source = normalize_orgunit_columns(validation_reporting_source)
                    if validation_reporting_source is None or 'Organisation unit ID' not in validation_reporting_source:
                        reporting_warnings.append(
                            "Les indicateurs de déclaration sont indisponibles pour cet intervalle et ce parent. "
                            "Les violations restent consultables ; les scores incomplets restent vides."
                        )
                        validation_reporting_source = pd.DataFrame()
                    validation_metrics = aggregate_validation_metrics(validation_reporting_source, validation_zones_df)
                    metrics_by_id = validation_metrics.set_index('Zone_ID').to_dict('index')

                    profile_by_id = {}
                    for ou_id, metadata in hierarchy_by_id.items():
                        level = metadata.get('level')
                        if level is not None and level >= 2 and clean_ou_id(metadata.get('province_id')):
                            profile_by_id[clean_ou_id(metadata['province_id'])] = 'DPS'
                        if level is not None and level >= 3 and clean_ou_id(metadata.get('zone_id')):
                            profile_by_id[clean_ou_id(metadata['zone_id'])] = 'ZS'
                        if metadata.get('level') in (2, 3):
                            profile_by_id[str(ou_id)] = 'DPS' if metadata['level'] == 2 else 'ZS'

                    # 3) Même requête que l'application DHIS2 : parent et intervalle complets.
                    rows_cat = []
                    validation_warnings = []
                    validation_requests = []
                    # Une nouvelle analyse doit lire DHIS2, même avec les mêmes filtres.
                    validation_cache_user = (*current_user_for_cache, uuid4().hex)
                    analysis_progress = st.progress(0.05, text="Analyse DHIS2 de l'intervalle sélectionné")
                    current_violations = get_validation_results(
                        validation_parent_id, validation_months, target_group_id, cache_user=validation_cache_user,
                        start_date=validation_start_date, end_date=validation_end_date
                    )
                    analysis_progress.progress(0.4, text="Analyse DHIS2 de la référence M-1")
                    previous_violations = get_validation_results(
                        validation_parent_id, validation_reference_months, target_group_id,
                        cache_user=validation_cache_user
                    )
                    for label, frame in [('M', current_violations), ('M-1', previous_violations)]:
                        request_params = frame.attrs.get('request_params', {})
                        validation_requests.append({
                            'Période': label,
                            'Méthode': frame.attrs.get('request_method', 'POST'),
                            'Unité parent': validation_parent_id,
                            'Début': request_params.get('startDate', ''),
                            'Fin': request_params.get('endDate', ''),
                            'Lignes reçues': len(frame),
                        })
                        for period in frame.attrs.get('failed_periods', []):
                            reasons = frame.attrs.get('errors', {}).get(period) or ["Analyse DHIS2 indisponible"]
                            validation_warnings.append(f"{label} ({period}) : {' ; '.join(reasons)}")

                    analysis_progress.progress(0.65, text="Rattachement des résultats DHIS2 aux zones de santé")
                    identifiers = validation_detail_identifiers(current_violations)
                    previous_identifiers = validation_detail_identifiers(previous_violations)
                    all_unit_ids = tuple(dict.fromkeys(identifiers['org_ids'] + previous_identifiers['org_ids']))
                    rule_metadata, unit_metadata, detail_warnings = get_validation_detail_metadata(
                        tuple(identifiers['rule_ids']), all_unit_ids, cache_user=current_user_for_cache
                    )
                    all_violations, current_scope_warnings = assign_validation_zones(
                        current_violations, validation_zones_df, unit_metadata, validation_parent_id
                    )
                    previous_violations, previous_scope_warnings = assign_validation_zones(
                        previous_violations, validation_zones_df, unit_metadata, validation_parent_id
                    )
                    validation_warnings.extend(f"M : {message}" for message in current_scope_warnings)
                    validation_warnings.extend(f"M-1 : {message}" for message in previous_scope_warnings)
                    incomplete_now = bool(all_violations.attrs.get('failed_periods') or current_scope_warnings)
                    incomplete_prev = bool(previous_violations.attrs.get('failed_periods') or previous_scope_warnings)
                    for zone_index, (zone_name, zone_id) in enumerate(zones_resolues, start=1):
                        analysis_progress.progress(
                            0.7 + (zone_index - 1) / len(zones_resolues) * 0.25,
                            text=f"Zone {zone_index}/{len(zones_resolues)} : {zone_name}"
                        )
                        zone_metrics = metrics_by_id.get(zone_id, {})
                        comp_val = float(zone_metrics.get('Complétude_Globale', np.nan))
                        prompt_val = float(zone_metrics.get('Promptitude_Globale', np.nan))
                        actual_val = float(zone_metrics.get('Actual_Reports', np.nan))
                        score_qualite = round((comp_val + prompt_val) / 2, 1)
                        row_cat = {
                            'Zone de santé': zone_name,
                            'Complétude_Globale': comp_val,
                            'Promptitude_Globale': prompt_val,
                            'Actual_Reports': actual_val,
                            'Score_Qualite': score_qualite,
                            'Nombre_Datasets_95': zone_metrics.get('Nombre_Datasets_95', np.nan),
                            'Profil_Bareme': profile_by_id.get(str(zone_id)),
                        }
                        if not zone_id or len(str(zone_id)) < 5:
                            rows_cat.append({**row_cat,
                                'Violations_Actuelles': np.nan,
                                'Violations_Mois_Precedent': np.nan,
                                'Violations_Corrigees': np.nan,
                                'Erreur_Validation': "Identifiant DHIS2 indisponible",
                            })
                            continue

                        df_val_zone = all_violations.loc[all_violations['_zone_id'].eq(zone_id)]
                        df_val_prev_zone = previous_violations.loc[previous_violations['_zone_id'].eq(zone_id)]

                        # Compter chaque ligne serveur ; les signatures servent seulement aux corrections.
                        sig_now = build_violation_signature_set(df_val_zone)
                        previous_signatures = build_violation_signatures(df_val_prev_zone)
                        violations_actuelles = np.nan if incomplete_now else len(df_val_zone)
                        violations_mois_precedent = np.nan if incomplete_prev else len(df_val_prev_zone)
                        violations_corrigees = np.nan if incomplete_now or incomplete_prev else sum(
                            signature not in sig_now for signature in previous_signatures
                        )

                        rows_cat.append({**row_cat,
                            'Violations_Actuelles': violations_actuelles,
                            'Violations_Mois_Precedent': violations_mois_precedent,
                            'Violations_Corrigees': violations_corrigees,
                            'Erreur_Validation': '; '.join(validation_warnings),
                        })

                    violations_details = build_validation_details(all_violations, rule_metadata, unit_metadata)
                    analysis_progress.progress(1.0, text=f"{len(zones_resolues)} zone(s) traitée(s)")
                    st.session_state['categorisation_result'] = {
                        'context': validation_context,
                        'data': pd.DataFrame(rows_cat),
                        'violations': violations_details,
                        'validation_warnings': validation_warnings,
                        'validation_requests': validation_requests,
                        'detail_warnings': detail_warnings,
                        'reporting_warnings': reporting_warnings,
                        'filters': {
                            'Date de début': validation_start_date.isoformat(),
                            'Date de fin': validation_end_date.isoformat(),
                            'Unité d’organisation parent': validation_parent_label,
                            'ID unité parent': validation_parent_id,
                            'Groupe de règles': selected_vr_name,
                            'Compte DHIS2': st.session_state.get('dhis2_user', ''),
                            'Source des violations': 'DHIS2 /api/dataAnalysis/validationRules',
                            'Mois des indicateurs de déclaration': validation_period_id,
                            'Période de référence M-1': ';'.join(validation_reference_months),
                        },
                        'analysed_at': datetime.now(KINSHASA_TZ).isoformat(timespec='seconds'),
                    }

            else:
                st.error("L'ID de la zone de santé sélectionnée est manquant dans les données sources.")

        saved_validation = st.session_state.get('categorisation_result')
        if saved_validation and saved_validation.get('context') == validation_context:
            df_cat = calculer_categorisation(saved_validation['data'], profil=profil_bareme)
            st.subheader(f"Résultats : {selected_vr_name}")
            applied_filters = dict(saved_validation.get('filters', {}))
            applied_filters['Barème datasets'] = profil_bareme
            df_validation_context_export = pd.DataFrame(
                [{'Paramètre': key, 'Valeur': value} for key, value in applied_filters.items()]
            )
            validation_applied_label = (
                f"Parent : {applied_filters.get('Unité d’organisation parent', validation_parent_label)}. "
                f"Du {applied_filters.get('Date de début', '')} au {applied_filters.get('Date de fin', '')}."
            )
            st.caption(validation_applied_label)
            with st.expander("Vérifier la requête envoyée à DHIS2"):
                st.caption(
                    f"Compte DHIS2 : {applied_filters.get('Compte DHIS2', '')}. "
                    "Source : /api/dataAnalysis/validationRules. "
                    f"Groupe : {applied_filters.get('Groupe de règles', selected_vr_name)}."
                )
                st.dataframe(
                    pd.DataFrame(saved_validation.get('validation_requests', [])),
                    hide_index=True, use_container_width=True,
                )
                st.caption(
                    "Pour comparer avec SNIS, utilisez le même compte, le même parent, "
                    "les mêmes dates et le même groupe de règles."
                )
            for message in saved_validation.get('reporting_warnings', []):
                st.warning(message)
            analysed_at = saved_validation.get('analysed_at')
            if analysed_at:
                st.caption(f"Analyse effectuée le {datetime.fromisoformat(analysed_at).strftime('%d/%m/%Y à %H:%M:%S')} (Kinshasa).")
            st.caption(
                "Les nouvelles notes et le total /100 sont ajoutés à droite des colonnes existantes. "
                "Une cellule vide indique qu'une donnée nécessaire au calcul est indisponible."
            )
            historical_columns = {
                'Zone de santé': 'Zone de santé',
                'Complétude_Globale': 'Complétude globale (%)',
                'Promptitude_Globale': 'Promptitude globale (%)',
                'Actual_Reports': 'Reports_Actual',
                'Violations_Mois_Precedent': 'Règles violées (M-1)',
                'Violations_Corrigees': 'Règles corrigées (M-1 -> M)',
                'Violations_Actuelles': 'Règles violées (M)',
                'Ratio_Violations_sur_100': 'Ratio / 100 rapports',
                'Score_Qualite': 'Score de qualité',
            }
            added_columns = [
                'Score de qualité /20', 'Barème datasets',
                'Nombre de datasets complétude ≥95 %', 'Règles restantes (M-1 -> M)',
                *[c for c in NOTE_COLS if c != 'Score de qualité /20'],
                TOTAL_COL, 'Statut catégorisation',
            ]
            df_display = df_cat[list(historical_columns) + added_columns].rename(columns=historical_columns)
            df_tab5_export = df_display.copy()
            view_query = st.text_input("Rechercher une zone", key="categorisation_search")
            view_left, view_right = st.columns(2)
            with view_left:
                view_status = st.selectbox("Afficher les résultats", VIEW_STATUS_OPTIONS, key="categorisation_status")
            with view_right:
                view_sort = st.selectbox("Trier les résultats", VIEW_SORT_OPTIONS, key="categorisation_sort")
            df_visible = filter_categorisation(df_display, query=view_query, status=view_status, sort_by=view_sort)
            st.caption(
                f"{len(df_visible)} zone(s) affichée(s) sur {len(df_display)}. "
                "L'export Excel, l'aperçu et les indicateurs conservent l'analyse complète."
            )
            if df_visible.empty:
                st.info("Aucune zone ne correspond à cette recherche et à ces filtres.")
            st.dataframe(
                _build_styler(df_visible, **style_context_tab5()) if not df_visible.empty else df_visible,
                use_container_width=True,
                hide_index=True,
            )
            invalid_scores = df_cat[TOTAL_COL].isna()
            if invalid_scores.any():
                st.warning(
                    f"{int(invalid_scores.sum())} unité(s) sans score total : consultez la colonne "
                    "« Statut catégorisation ». Pour un barème non identifié, choisissez DPS ou ZS."
                )
            col_res1, col_res2, col_res3, col_res4 = st.columns(4)
            detected = df_cat['Violations_Actuelles'].sum(min_count=1)
            corrected = df_cat['Violations_Corrigees'].sum(min_count=1)
            col_res1.metric("Violations détectées (données disponibles)", '—' if pd.isna(detected) else int(detected))
            col_res2.metric("Règles corrigées (données disponibles)", '—' if pd.isna(corrected) else int(corrected))
            average_score = df_cat[TOTAL_COL].mean()
            col_res3.metric("Score moyen de catégorisation /100", '—' if pd.isna(average_score) else f"{average_score:.1f}")
            col_res4.metric("Unités avec score complet", f"{int((~invalid_scores).sum())} / {len(df_cat)}")

            st.divider()
            st.subheader("Détail des règles violées")
            violations_all = saved_validation.get('violations', build_validation_details(pd.DataFrame()))
            df_violations_export = detail_export_frame(violations_all)
            validation_warnings = saved_validation.get('validation_warnings', [])
            if validation_warnings:
                st.warning(
                    "Certaines analyses sont indisponibles ou limitées par DHIS2. "
                    "Les lignes reçues restent visibles, mais les scores concernés ne sont pas calculés."
                )
                with st.expander("Voir les périodes et zones concernées"):
                    for message in validation_warnings:
                        st.text(message)
            for message in saved_validation.get('detail_warnings', []):
                st.info(message)
            st.caption(
                "Les lignes, périodes et valeurs sont celles renvoyées par DHIS2 pour l'intervalle M. "
                "Chaque ligne est conservée et comptée ; les signatures servent uniquement à comparer les corrections avec M-1."
            )
            detail_query = st.text_input("Rechercher dans les violations", key="validation_detail_search")
            detail_importance = st.selectbox(
                "Filtrer les violations par importance", ['Toutes', 'HIGH', 'MEDIUM', 'LOW', 'Non renseignée'],
                key="validation_detail_importance"
            )
            violations_visible = filter_validation_details(violations_all, query=detail_query, importance=detail_importance)
            detail_view = detail_export_frame(violations_visible)
            st.caption(
                f"{len(detail_view)} violation(s) affichée(s) sur {len(df_violations_export)}. "
                "Le CSV contient les lignes affichées ; le rapport Excel et l'aperçu conservent le détail complet."
            )
            if violations_all.empty:
                if validation_warnings:
                    st.info("Aucune ligne de violation n'a été reçue. Consultez les messages d'analyse ci-dessus.")
                else:
                    st.success("Aucune règle violée détectée pour les filtres sélectionnés.")
            elif detail_view.empty:
                st.info("Aucune violation ne correspond à cette recherche et à ce filtre d'importance.")
            else:
                st.dataframe(
                    detail_view.style.format({'Valeur gauche': '{:g}', 'Valeur droite': '{:g}'}, na_rep='')
                    .map(style_importance, subset=['Importance']),
                    use_container_width=True, hide_index=True,
                    column_config={
                        'Règle de validation': st.column_config.TextColumn(width='large'),
                        'Détails': st.column_config.TextColumn(width='large'),
                    },
                )
                st.download_button(
                    "Télécharger les violations affichées en CSV",
                    data=detail_view.to_csv(index=False).encode('utf-8-sig'),
                    file_name=f"violations_{validation_start_date.isoformat()}_{validation_end_date.isoformat()}.csv", mime='text/csv',
                    key="download_validation_details_csv"
                )
                with st.expander("Consulter le détail d'une violation"):
                    row_options = list(range(len(violations_visible)))
                    if st.session_state.get('validation_selected_row') not in row_options:
                        st.session_state.pop('validation_selected_row', None)
                    selected_row = st.selectbox(
                        "Violation à consulter", row_options,
                        format_func=lambda position: (
                            f"{violations_visible.iloc[position][VIOLATION_COLUMNS[1]]} — "
                            f"{violations_visible.iloc[position]['Période']} — "
                            f"{violations_visible.iloc[position]['Règle de validation']}"
                        ),
                        key="validation_selected_row"
                    )
                    violation = violations_visible.iloc[selected_row]
                    st.text(
                        f"Unité : {violation[VIOLATION_COLUMNS[1]]}\n"
                        f"Période : {violation['Période']}\nImportance : {violation['Importance']}"
                    )
                    st.text(str(violation['Règle de validation']))
                    left_value = 'Non renseignée' if pd.isna(violation['Valeur gauche']) else f"{violation['Valeur gauche']:g}"
                    right_value = 'Non renseignée' if pd.isna(violation['Valeur droite']) else f"{violation['Valeur droite']:g}"
                    st.text(f"{left_value} {violation['Opérateur']} {right_value}")
                    st.text(str(violation['Détails']))

    # --- MENU SIDEBAR : EXTRACTION COMMENTEE ---
    with st.sidebar:
        st.divider()
        st.subheader("📤 Extraction du rapport")
        report_type = st.selectbox(
            "Type de téléchargement :",
            ["Excel"],
            key="report_type_selector"
        )

        export_style_context = {
            "base": {
                "decimals": 2,
            },
            "actual_detail": {
                "decimals": 2,
            },
            "expected_detail": {
                "decimals": 2,
            },
            "completude": {
                "taux_cols": (col_rate_uniquement if 'col_rate_uniquement' in locals() else []) + ['Complétude_Globale (%)'],
                "score_cols": ['Nombre des data set complétude >/=95%'],
                "int_cols": ['Nombre des data set complétude >/=95%'],
                "decimals": 2,
            },
            "promptitude": {
                "taux_cols": (col_rate_ot if 'col_rate_ot' in locals() else []) + ['Promptitude_Globale (%)'],
                "score_cols": [nom_col_score_p] if 'nom_col_score_p' in locals() else [],
                "int_cols": [nom_col_score_p] if 'nom_col_score_p' in locals() else [],
                "decimals": 2,
            },
            "synthese": {
                "taux_cols": ['Complétude_Globale (%)', 'Promptitude_Globale (%)'],
                "decimals": 2,
            },
            "comparatif": {
                "taux_cols": ['Complétude_Globale (%)', 'Promptitude_Globale (%)'],
                "score_cols": ['Nombre de dataset complétude >/= 95%', 'Nombre de dataset promptitude >/= 95%'],
                "int_cols": ['Nombre de dataset complétude >/= 95%', 'Nombre de dataset promptitude >/= 95%'],
                "decimals": 2,
            },
            "fusion": {
                "taux_cols": ['Complétude', 'Promptitude'],
                "score_cols": ['Nombre de dataset complétude >/= 95%', 'Nombre de dataset promptitude >/= 95%'],
                "int_cols": ['Nombre de dataset complétude >/= 95%', 'Nombre de dataset promptitude >/= 95%'],
                "decimals": 2,
            },
            "tab5": style_context_tab5(),
            "top5": {
                "taux_cols": ['Complétude_Globale (%)'],
                "decimals": 2,
            },
            "flop5": {
                "taux_cols": ['Promptitude_Globale (%)'],
                "decimals": 2,
            },
        }

        comments_df = build_dashboard_comments_df(
            df_final_c=df_final_c,
            df_final_p=df_final_p,
            df_synth=df_synth,
            selected_zone=selected_scope_label,
            df_tab4_fusion=df_tab4_fusion,
            df_tab5=df_tab5_export,
            df_base=df_tab1_export,
            df_actual_detail=df_actual_detail_export,
            df_expected_detail=df_expected_detail_export,
            df_comparatif=df_comparatif,
            df_top5=df_top5_export,
            df_flop5=df_flop5_export
        )

        export_bytes = None
        export_mime = None
        export_name = None
        export_error = None

        export_bytes = to_excel_dashboard_report(
            df_base=df_tab1_export,
            df_actual_detail=df_actual_detail_export,
            df_expected_detail=df_expected_detail_export,
            df_final_c=df_final_c,
            df_final_p=df_final_p,
            df_synth=df_synth,
            comments_df=comments_df,
            df_tab4_fusion=df_tab4_fusion,
            df_comparatif=df_comparatif,
            df_tab5=df_tab5_export,
            df_top5=df_top5_export,
            df_flop5=df_flop5_export,
            style_context=export_style_context,
            df_violations=df_violations_export,
            df_validation_context=df_validation_context_export,
        )
        export_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        export_name = f"rapport_dashboard_commente_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

        if export_bytes is not None:
            st.download_button(
                f"Télécharger {report_type}",
                data=export_bytes,
                file_name=export_name,
                mime=export_mime,
                use_container_width=True,
                key=f"btn_download_report_{report_type}"
            )
        else:
            st.info(export_error if export_error else "Export indisponible.")

        st.divider()
        if "show_report_preview" not in st.session_state:
            st.session_state["show_report_preview"] = False
        col_prev1, col_prev2 = st.columns(2)
        if col_prev1.button("Visualiser le rapport", use_container_width=True):
            st.session_state["show_report_preview"] = True
        if col_prev2.button("Masquer la vue", use_container_width=True):
            st.session_state["show_report_preview"] = False

    if st.session_state.get("show_report_preview", False):
        st.divider()
        st.subheader("👁️ Visualisation consolidée du rapport")
        st.caption("Aperçu interactif des manipulations en cours (filtres, période, niveau).")

        p_tabs = st.tabs([
            "Base de données",
            "Données réelles",
            "Participants/Attendus",
            "Perf finale",
            "Promptitude",
            "Comparatif",
            "Top/Flop",
            "Règles",
            "Commentaires",
            "Règles détaillées",
        ])

        with p_tabs[0]:
            st.dataframe(df_tab1_export, use_container_width=True)
        with p_tabs[1]:
            st.dataframe(df_actual_detail_export, use_container_width=True)
        with p_tabs[2]:
            st.dataframe(df_expected_detail_export, use_container_width=True)
        with p_tabs[3]:
            st.dataframe(df_final_c, use_container_width=True)
        with p_tabs[4]:
            st.dataframe(df_final_p, use_container_width=True)
        with p_tabs[5]:
            st.dataframe(df_comparatif, use_container_width=True)
            if not df_tab4_fusion.empty:
                st.write("Tableau fusionné zone filtrée")
                st.dataframe(df_tab4_fusion, use_container_width=True)
        with p_tabs[6]:
            cprev1, cprev2 = st.columns(2)
            with cprev1:
                st.write("Top 5 complétude")
                st.dataframe(df_top5_export, use_container_width=True)
            with cprev2:
                st.write("Flop 5 promptitude")
                st.dataframe(df_flop5_export, use_container_width=True)
        with p_tabs[7]:
            if validation_applied_label:
                st.caption(validation_applied_label)
            if df_tab5_export.empty:
                st.info("Lancez l'analyse des violations dans l'onglet 5 pour afficher la catégorisation.")
            else:
                st.dataframe(_build_styler(df_tab5_export, **style_context_tab5()), use_container_width=True, hide_index=True)
        with p_tabs[8]:
            st.dataframe(comments_df, use_container_width=True)
        with p_tabs[9]:
            if validation_applied_label:
                st.caption(validation_applied_label)
            if df_violations_export is None:
                st.info("Lancez l'analyse des violations dans l'onglet 5 pour afficher leur détail.")
            elif df_violations_export.empty:
                st.info("Aucune ligne de violation disponible pour cette analyse.")
            else:
                st.dataframe(
                    df_violations_export.style.format({'Valeur gauche': '{:g}', 'Valeur droite': '{:g}'}, na_rep='')
                    .map(style_importance, subset=['Importance']), use_container_width=True, hide_index=True
                )
else:
    st.error("❌ Données indisponibles.")

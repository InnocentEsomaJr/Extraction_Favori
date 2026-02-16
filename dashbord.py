import streamlit as st
import pandas as pd
import numpy as np
from openhexa.toolbox.dhis2 import DHIS2
import plotly.express as px
import io
import os
from datetime import datetime

# --- 1. CONFIGURATION ---
st.set_page_config(page_title="SNIS RDC - Dashboard Performance", layout="wide")

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

def _extract_identifier(value):
    """Extrait un identifiant stable depuis un champ DHIS2 (dict ou string)."""
    if isinstance(value, dict):
        return str(
            value.get("id")
            or value.get("uid")
            or value.get("displayName")
            or value.get("name")
            or ""
        ).strip()
    if value is None:
        return ""
    return str(value).strip()

def build_violation_signature_set(df_val):
    """
    Construit un ensemble de signatures de violations pour comparer T-1 vs T.
    On exclut volontairement la période pour détecter les erreurs réellement corrigées.
    """
    if df_val is None or df_val.empty:
        return set()

    signatures = set()
    for row in df_val.to_dict("records"):
        ou = _extract_identifier(row.get("organisationUnit"))
        rule = _extract_identifier(row.get("validationRule"))
        aoc = _extract_identifier(row.get("attributeOptionCombo"))
        coc = _extract_identifier(row.get("categoryOptionCombo"))
        ds = _extract_identifier(row.get("dataSet"))

        # Une violation doit au minimum avoir une règle ou une OU.
        if not ou and not rule:
            continue

        signatures.add((ou, rule, aoc, coc, ds))
    return signatures

# Fonction pour l'export Excel
def to_excel(df):
    output = io.BytesIO()
    # Utilisation de xlsxwriter (assurez-vous qu'il est installé via pip install xlsxwriter)
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False, sheet_name='Performance')
    return output.getvalue()

def build_dashboard_comments_df(df_final_c, df_final_p, df_synth, selected_zone, df_tab4_fusion=None):
    """Prépare des commentaires d'interprétation pour les tableaux et graphiques."""
    zone_label = selected_zone if selected_zone != "Toutes les zones" else "Toutes les zones"
    comp_avg = float(df_synth['Complétude_Globale (%)'].mean()) if not df_synth.empty else 0.0
    prom_avg = float(df_synth['Promptitude_Globale (%)'].mean()) if not df_synth.empty else 0.0
    top_comp = df_synth.nlargest(1, 'Complétude_Globale (%)')['Organisation unit'].iloc[0] if not df_synth.empty else "N/A"
    flop_prom = df_synth.nsmallest(1, 'Promptitude_Globale (%)')['Organisation unit'].iloc[0] if not df_synth.empty else "N/A"

    rows = [
        {
            "Element": "Onglet 2 - Tableau Complétude",
            "Commentaire": f"Complétude moyenne {comp_avg:.2f}%. Le tableau identifie les unités avec meilleure couverture des rapports."
        },
        {
            "Element": "Onglet 2 - Graphique Classement",
            "Commentaire": f"Classement visuel par complétude globale. Meilleure unité observée: {top_comp}."
        },
        {
            "Element": "Onglet 3 - Tableau Promptitude",
            "Commentaire": f"Promptitude moyenne {prom_avg:.2f}%. Le tableau met en évidence le respect des délais de soumission."
        },
        {
            "Element": "Onglet 4 - Quadrant Comparatif",
            "Commentaire": "Le quadrant compare simultanément complétude et promptitude autour du seuil de 80%."
        },
        {
            "Element": "Onglet 4 - Top/Flop",
            "Commentaire": f"Top complétude: {top_comp}. Flop promptitude: {flop_prom}. Prioriser les unités en bas de promptitude."
        },
        {
            "Element": "Onglet 5 - Tableau Catégorisation",
            "Commentaire": "Le tableau compare M-1 et M sur les violations, les règles corrigées et le ratio sur 100 rapports."
        },
    ]

    if df_tab4_fusion is not None and not df_tab4_fusion.empty:
        rows.append(
            {
                "Element": "Onglet 4 - Tableau fusionné (zone filtrée)",
                "Commentaire": f"Pour la zone '{zone_label}', ce tableau fusionne indicateurs dataset, complétude, promptitude et scores >=95%."
            }
        )

    return pd.DataFrame(rows)

def to_excel_dashboard_report(df_final_c, df_final_p, df_synth, comments_df, df_tab4_fusion=None):
    """Exporte un rapport Excel commenté."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_final_c.to_excel(writer, index=False, sheet_name='Onglet2_Completude')
        df_final_p.to_excel(writer, index=False, sheet_name='Onglet3_Promptitude')
        df_synth.to_excel(writer, index=False, sheet_name='Onglet4_Comparatif')
        if df_tab4_fusion is not None and not df_tab4_fusion.empty:
            df_tab4_fusion.to_excel(writer, index=False, sheet_name='Onglet4_Fusion_Zone')
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

def to_word_dashboard_report(df_final_c, df_final_p, df_synth, comments_df, df_tab4_fusion=None):
    """Exporte un rapport Word commenté (retourne bytes, error)."""
    try:
        from docx import Document
    except Exception:
        return None, "python-docx non installé (pip install python-docx)."

    doc = Document()
    doc.add_heading("Rapport Dashboard SNIS RDC", level=1)
    doc.add_paragraph(f"Généré le {datetime.now().strftime('%Y-%m-%d %H:%M')}")

    doc.add_heading("Commentaires", level=2)
    for _, row in comments_df.iterrows():
        doc.add_paragraph(f"{row['Element']}: {row['Commentaire']}")

    _add_dataframe_to_docx(doc, "Onglet 2 - Complétude", df_final_c)
    _add_dataframe_to_docx(doc, "Onglet 3 - Promptitude", df_final_p)
    _add_dataframe_to_docx(doc, "Onglet 4 - Comparatif", df_synth)
    if df_tab4_fusion is not None and not df_tab4_fusion.empty:
        _add_dataframe_to_docx(doc, "Onglet 4 - Fusion zone filtrée", df_tab4_fusion)

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue(), None

def dataframe_to_png_bytes(df, title="Tableau", max_rows=20):
    """Convertit un DataFrame en image PNG (retourne bytes, error)."""
    try:
        import matplotlib.pyplot as plt
    except Exception:
        return None, "matplotlib non installé (pip install matplotlib)."

    if df is None or df.empty:
        return None, "Aucune donnée à convertir en image."

    df_show = df.head(max_rows).copy()
    fig_h = max(2.8, 0.45 * (len(df_show) + 2))
    fig_w = max(10, 1.5 * len(df_show.columns))
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    ax.axis('off')
    ax.set_title(title, fontsize=11, pad=8)
    table = ax.table(
        cellText=df_show.astype(str).values,
        colLabels=df_show.columns,
        cellLoc='center',
        loc='center'
    )
    table.auto_set_font_size(False)
    table.set_fontsize(8)
    table.scale(1, 1.2)
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

    tx_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.7), Inches(12.3), Inches(0.8))
    tx_frame = tx_box.text_frame
    tx_frame.text = comment
    tx_frame.paragraphs[0].font.size = Pt(14)

    if image_bytes is not None:
        slide.shapes.add_picture(io.BytesIO(image_bytes), Inches(0.5), Inches(1.5), width=Inches(12.3))

def to_powerpoint_dashboard_report(df_synth, comments_df, fig_comp=None, fig_quad=None, df_comparatif=None, df_tab4_fusion=None):
    """Exporte un rapport PowerPoint commenté avec images (retourne bytes, error)."""
    try:
        from pptx import Presentation
    except Exception:
        return None, "python-pptx non installé (pip install python-pptx)."

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

    # 2) Quadrant comparatif
    img_quad, _ = figure_to_png_bytes(fig_quad)
    _add_image_slide(
        prs,
        "Quadrant Complétude / Promptitude",
        _comment_for(comments_df, "Quadrant", "Comparaison de la complétude et de la promptitude."),
        img_quad
    )

    # 3) Tableau comparatif complétude/promptitude
    img_comp_table, _ = dataframe_to_png_bytes(df_comparatif, title="Tableau comparatif Complétude / Promptitude")
    _add_image_slide(
        prs,
        "Tableau comparatif",
        _comment_for(comments_df, "Onglet 4 - Top/Flop", "Lecture comparative par unité."),
        img_comp_table
    )

    # 4) Tableau fusionné zone filtrée (si disponible)
    if df_tab4_fusion is not None and not df_tab4_fusion.empty:
        img_fusion, _ = dataframe_to_png_bytes(df_tab4_fusion, title="Tableau fusionné indicateurs dataset")
        _add_image_slide(
            prs,
            "Tableau fusionné zone filtrée",
            _comment_for(comments_df, "fusionné", "Vue détaillée des indicateurs dataset en zone filtrée."),
            img_fusion
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

@st.cache_resource
def get_dhis2_client():
    dhis2_url = _read_config_value("DHIS2_URL")
    dhis2_user = _read_config_value("DHIS2_USER")
    dhis2_pass = _read_config_value("DHIS2_PASS")

    missing_keys = [
        key for key, val in {
            "DHIS2_URL": dhis2_url,
            "DHIS2_USER": dhis2_user,
            "DHIS2_PASS": dhis2_pass
        }.items()
        if not val
    ]
    if missing_keys:
        missing_text = ", ".join(missing_keys)
        raise RuntimeError(
            f"Configuration DHIS2 manquante: {missing_text}. "
            "Ajoute ces clés dans Streamlit Cloud (App settings > Secrets) "
            "ou définis-les comme variables d'environnement."
        )

    return DHIS2(url=dhis2_url, username=dhis2_user, password=dhis2_pass)

@st.cache_data(ttl=3600)
def get_data(favori_id, period=None):
    if not favori_id: return None
    try:
        dhis = get_dhis2_client()
        endpoint = f"visualizations/{favori_id}/data.json"
        if period:
            endpoint += f"?dimension=pe:{period}"

        response = dhis.api.get(endpoint)
        headers, rows = response.get("headers", []), response.get("rows", [])
        columns = [h.get("name", h.get("column")) for h in headers]
        return pd.DataFrame(rows, columns=columns)
    except Exception as e:
        st.error(f"Erreur de connexion : {e}")
        return None

@st.cache_data(ttl=3600)
def get_validation_groups():
    try:
        dhis = get_dhis2_client()
        return dhis.api.get("validationRuleGroups", params={"fields": "id,displayName", "paging": "false"})['validationRuleGroups']
    except: return []

@st.cache_data(ttl=3600)
def get_validation_results(ou_id, period_list, group_id):
    # Sécurité anti-Erreur 409 : si l'ID n'est pas un ID système DHIS2, on bloque.
    if not ou_id or ou_id == "USER_ORGUNIT" or len(str(ou_id)) < 5:
        return pd.DataFrame()

    all_results = []
    try:
        dhis = get_dhis2_client()
    except Exception as e:
        st.error(f"Erreur de connexion DHIS2 : {e}")
        return pd.DataFrame()
    for pe in period_list:
        try:
            params = {"ou": ou_id, "pe": pe, "ouMode": "DESCENDANTS"}
            if group_id:
                params["vrg"] = group_id
            res = dhis.api.get("validationResults", params=params)
            if 'validationResults' in res:
                all_results.extend(res['validationResults'])
        except Exception:
            continue # Ignore les erreurs de périodes vides
    return pd.DataFrame(all_results)

@st.cache_data(ttl=3600)
def get_children_org_units_details(parent_id):
    """Récupère les enfants directs (id + nom) d'une unité d'organisation."""
    if not parent_id:
        return []
    try:
        dhis = get_dhis2_client()
        response = dhis.api.get(
            f"organisationUnits/{parent_id}",
            params={"fields": "children[id,displayName]", "paging": "false"}
        )
        return response.get("children", []) or []
    except Exception:
        return []

# --- 3. SIDEBAR & FILTRES ---
with st.sidebar:
    st.image("https://snisrdc.com/dhis-web-commons/security/logo_front.png", width=150)
    st.title("⚙️ Configuration")

    dict_favoris = {
        "Performance Globale (ROzCY14OLTE)": "ROzCY14OLTE",
        "Analyse Promptitude (mldsgxAvIIi)": "mldsgxAvIIi",
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

# --- 4. CHARGEMENT & FILTRAGE DES DONNÉES ---
df_raw = get_data(id_final, period=period_id)

if df_raw is not None:
    # Récupération dynamique de l'ID de l'unité d'organisation (OrgUnit ID)
    # On cherche l'ID dans les colonnes renvoyées par le favori
    mapping_ou_id = {}
    if 'Organisation unit' in df_raw.columns and 'Organisation unit ID' in df_raw.columns:
        mapping_ou_id = dict(zip(df_raw['Organisation unit'], df_raw['Organisation unit ID']))

    # Indicateur parent (zone/province) pour filtrer les aires de santé.
    is_parent_by_ou = {}
    if 'Organisation unit' in df_raw.columns and 'Organisation unit is parent' in df_raw.columns:
        temp_parent = df_raw[['Organisation unit', 'Organisation unit is parent']].drop_duplicates('Organisation unit')
        is_parent_by_ou = {
            row['Organisation unit']: str(row['Organisation unit is parent']).strip().lower() in ['true', '1']
            for _, row in temp_parent.iterrows()
        }

    target_df_all = df_raw.copy()
    default_cols = ['Organisation unit code', 'Organisation unit description',
                    'Reporting month', 'Organisation unit parameter', 'Organisation unit is parent']
    cols_existing = [c for c in default_cols if c in target_df_all.columns]
    if cols_existing: target_df_all.drop(columns=cols_existing, inplace=True)

    st.sidebar.subheader("🔍 Filtrage")
    zones = ["Toutes les zones"] + sorted(target_df_all['Organisation unit'].unique().tolist())
    selected_zone = st.sidebar.selectbox("Filtrer par Zone ou Aire de Santé :", zones)

    # Détermination de l'ID à utiliser pour l'onglet 5
    if selected_zone != "Toutes les zones":
        target_df = target_df_all[target_df_all['Organisation unit'] == selected_zone].copy()
        current_id_systeme = mapping_ou_id.get(selected_zone)
    else:
        target_df = target_df_all.copy()
        # Si possible on prend l'ID de la ligne parent (province), sinon premier ID.
        if 'Organisation unit is parent' in df_raw.columns:
            parent_mask = df_raw['Organisation unit is parent'].astype(str).str.lower().isin(['true', '1'])
            parent_rows = df_raw[parent_mask]
            if not parent_rows.empty and 'Organisation unit ID' in parent_rows.columns:
                current_id_systeme = parent_rows['Organisation unit ID'].iloc[0]
            else:
                current_id_systeme = df_raw['Organisation unit ID'].iloc[0] if 'Organisation unit ID' in df_raw.columns else None
        else:
            current_id_systeme = df_raw['Organisation unit ID'].iloc[0] if 'Organisation unit ID' in df_raw.columns else None

    st.title(f"📊 SNIS RDC - Performance {selected_zone if selected_zone != 'Toutes les zones' else 'Haut-Uele'}")

    # Variables partagées pour export hors onglets.
    fig_comp = None
    fig_quad = None
    df_tab4_fusion = pd.DataFrame()
    df_comparatif = pd.DataFrame()

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
            st.warning("Aucune aire de santé à afficher avec ce filtre.")
        else:
            st.dataframe(df_tab1.head(15), use_container_width=True)

    # --- ONGLET 2 : COMPLÉTUDE ---
    with tab2:
        st.header("Analyse de la Complétude")
        col_actual_strict = [c for c in target_df.columns if 'actual reports' in c.lower() and 'time' not in c.lower()]
        if col_actual_strict:
            df_actual = target_df[['Organisation unit'] + col_actual_strict].copy()
            for col in col_actual_strict: df_actual[col] = pd.to_numeric(df_actual[col], errors='coerce').fillna(0)
            df_actual['Reports_Actual'] = df_actual[col_actual_strict].sum(axis=1).round(2)
            with st.expander("Détail Reports Actuals"): st.dataframe(df_actual.head(15))

        col_expected_strict = [c for c in target_df.columns if 'expected reports' in c.lower() and 'time' not in c.lower()]
        if col_expected_strict:
            df_expected = target_df[['Organisation unit'] + col_expected_strict].copy()
            for col in col_expected_strict: df_expected[col] = pd.to_numeric(df_expected[col], errors='coerce').fillna(0)
            df_expected['Reports_Attendu'] = df_expected[col_expected_strict].sum(axis=1).round(2)
            with st.expander("Détail Reports Attendus"): st.dataframe(df_expected.head(15))

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
        st.plotly_chart(fig_comp, use_container_width=True, config=plotly_config)

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
            'Organisation unit': 'Aire de santé',
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
        if selected_zone != "Toutes les zones":
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
        st.plotly_chart(fig_quad, use_container_width=True)

        col_l, col_r = st.columns(2)
        with col_l:
            st.success("🏆 **Top 5 Complétude**")
            st.table(df_synth.nlargest(5, 'Complétude_Globale (%)')[['Organisation unit', 'Complétude_Globale (%)']].round(1))
        with col_r:
            st.error("⚠️ **Flop 5 Promptitude**")
            st.table(df_synth.nsmallest(5, 'Promptitude_Globale (%)')[['Organisation unit', 'Promptitude_Globale (%)']].round(1))

    # --- ONGLET 5 : ÉLÉMENTS DE CATÉGORISATION (CORRECTIF APPLIQUÉ SUR TON SCRIPT) ---
    with tab5:
        st.header("🩺 Éléments de catégorisation interactifs")

        vr_groups = get_validation_groups()
        if vr_groups:
            group_mapping = {g['displayName']: g['id'] for g in vr_groups}
            options_regles = ["Toutes les règles (Global)"] + list(group_mapping.keys())

            selected_vr_name = st.selectbox("Sélectionner le Groupe de Règles de Validation :", options_regles)

            if st.button("Lancer l'analyse des violations"):
                if current_id_systeme and len(str(current_id_systeme)) > 5:
                    with st.spinner('Analyse des violations en cours...'):

                        target_group_id = group_mapping.get(selected_vr_name) if selected_vr_name != "Toutes les règles (Global)" else None

                        # 1) Construire la liste des zones de santé (pas la province)
                        zone_id_mapping = mapping_ou_id.copy()
                        if selected_zone != "Toutes les zones":
                            zones_cibles = [selected_zone]
                        else:
                            children_info = get_children_org_units_details(current_id_systeme)
                            zones_cibles = [c['displayName'] for c in children_info if c.get('displayName') and c.get('id')]
                            for child in children_info:
                                zone_id_mapping[child['displayName']] = child['id']

                            # Fallback si l'API enfants ne renvoie rien
                            if not zones_cibles:
                                zones_cibles = sorted(target_df_all['Organisation unit'].dropna().unique().tolist())
                                province_name = next((nom for nom, ou_id in mapping_ou_id.items() if ou_id == current_id_systeme), None)
                                if province_name in zones_cibles and len(zones_cibles) > 1:
                                    zones_cibles = [z for z in zones_cibles if z != province_name]

                        zones_cibles = [z for z in zones_cibles if zone_id_mapping.get(z)]
                        zones_cibles = list(dict.fromkeys(zones_cibles))

                        # Mapping normalisé (fallback si les libellés diffèrent légèrement)
                        zone_id_mapping_norm = {
                            normalize_org_name(k): v
                            for k, v in zone_id_mapping.items()
                            if k and v
                        }

                        zones_resolues = []
                        for z in zones_cibles:
                            zid = zone_id_mapping.get(z) or zone_id_mapping_norm.get(normalize_org_name(z))
                            if zid:
                                zones_resolues.append((z, zid))
                        zones_cibles = [z for z, _ in zones_resolues]
                        zone_id_mapping = {z: zid for z, zid in zones_resolues}

                        if not zones_cibles:
                            st.warning("Aucune zone de santé valide trouvée.")
                            st.stop()

                        # 2) Reports_Actual provenant directement de l'onglet 2 (df_synthese)
                        actual_map_by_id = {}
                        actual_map_by_name = {}
                        if 'Organisation unit' in df_synthese.columns and 'Reports_Actual' in df_synthese.columns:
                            df_actual_map = df_synthese[['Organisation unit', 'Reports_Actual']].copy()
                            df_actual_map['Reports_Actual'] = pd.to_numeric(df_actual_map['Reports_Actual'], errors='coerce').fillna(0)
                            df_actual_map['ou_id'] = df_actual_map['Organisation unit'].map(mapping_ou_id)
                            df_actual_map['ou_id'] = df_actual_map['ou_id'].where(
                                df_actual_map['ou_id'].notna(),
                                df_actual_map['Organisation unit'].map(lambda x: zone_id_mapping_norm.get(normalize_org_name(x)))
                            )
                            actual_map_by_id = (
                                df_actual_map.dropna(subset=['ou_id'])
                                .groupby('ou_id')['Reports_Actual']
                                .mean()
                                .to_dict()
                            )
                            actual_map_by_name = (
                                df_actual_map.assign(_k=df_actual_map['Organisation unit'].map(normalize_org_name))
                                .groupby('_k')['Reports_Actual']
                                .mean()
                                .to_dict()
                            )

                        # 2-bis) Prendre les valeurs déjà calculées:
                        # Complétude depuis l'onglet 2 (df_synthese),
                        # Promptitude depuis l'onglet 3 (df_final_p)
                        comp_map_by_id = {}
                        comp_map_by_name = {}
                        prompt_map_by_id = {}
                        prompt_map_by_name = {}
                        if 'Organisation unit' in df_synthese.columns and 'Complétude_Globale (%)' in df_synthese.columns:
                            df_comp_map = df_synthese[['Organisation unit', 'Complétude_Globale (%)']].copy()
                            df_comp_map['ou_id'] = df_comp_map['Organisation unit'].map(mapping_ou_id)
                            df_comp_map['ou_id'] = df_comp_map['ou_id'].where(
                                df_comp_map['ou_id'].notna(),
                                df_comp_map['Organisation unit'].map(lambda x: zone_id_mapping_norm.get(normalize_org_name(x)))
                            )
                            comp_map_by_id = (
                                df_comp_map.dropna(subset=['ou_id'])
                                .groupby('ou_id')['Complétude_Globale (%)']
                                .mean()
                                .to_dict()
                            )
                            comp_map_by_name = (
                                df_comp_map.assign(_k=df_comp_map['Organisation unit'].map(normalize_org_name))
                                .groupby('_k')['Complétude_Globale (%)']
                                .mean()
                                .to_dict()
                            )
                        if 'Organisation unit' in df_final_p.columns and 'Promptitude_Globale (%)' in df_final_p.columns:
                            df_prompt_map = df_final_p[['Organisation unit', 'Promptitude_Globale (%)']].copy()
                            df_prompt_map['ou_id'] = df_prompt_map['Organisation unit'].map(mapping_ou_id)
                            df_prompt_map['ou_id'] = df_prompt_map['ou_id'].where(
                                df_prompt_map['ou_id'].notna(),
                                df_prompt_map['Organisation unit'].map(lambda x: zone_id_mapping_norm.get(normalize_org_name(x)))
                            )
                            prompt_map_by_id = (
                                df_prompt_map.dropna(subset=['ou_id'])
                                .groupby('ou_id')['Promptitude_Globale (%)']
                                .mean()
                                .to_dict()
                            )
                            prompt_map_by_name = (
                                df_prompt_map.assign(_k=df_prompt_map['Organisation unit'].map(normalize_org_name))
                                .groupby('_k')['Promptitude_Globale (%)']
                                .mean()
                                .to_dict()
                            )

                        # 3) Violations par zone + ratio /100 + corrigées T vs T-1
                        rows_cat = []
                        for zone_name in zones_cibles:
                            zone_id = zone_id_mapping.get(zone_name) or zone_id_mapping_norm.get(normalize_org_name(zone_name))
                            zone_key = normalize_org_name(zone_name)
                            comp_val = float(comp_map_by_id.get(zone_id, comp_map_by_name.get(zone_key, 0)))
                            prompt_val = float(prompt_map_by_id.get(zone_id, prompt_map_by_name.get(zone_key, 0)))
                            actual_val = float(actual_map_by_id.get(zone_id, actual_map_by_name.get(zone_key, 0)))
                            score_qualite = round((comp_val + prompt_val) / 2, 1)
                            if not zone_id or len(str(zone_id)) < 5:
                                rows_cat.append({
                                    'Zone de santé': zone_name,
                                    'Complétude_Globale': comp_val,
                                    'Promptitude_Globale': prompt_val,
                                    'Actual_Reports': actual_val,
                                    'Violations_Actuelles': 0,
                                    'Violations_Mois_Precedent': 0,
                                    'Violations_Corrigees': 0,
                                    'Score_Qualite': score_qualite
                                })
                                continue

                            df_val_zone = get_validation_results(zone_id, selection_mois, target_group_id)
                            df_val_prev_zone = get_validation_results(zone_id, selection_mois_prev, target_group_id)

                            # Comparaison des mêmes erreurs entre T-1 et T
                            sig_now = build_violation_signature_set(df_val_zone)
                            sig_prev = build_violation_signature_set(df_val_prev_zone)
                            violations_actuelles = len(sig_now)
                            violations_mois_precedent = len(sig_prev)
                            violations_corrigees = len(sig_prev - sig_now)

                            rows_cat.append({
                                'Zone de santé': zone_name,
                                'Complétude_Globale': comp_val,
                                'Promptitude_Globale': prompt_val,
                                'Actual_Reports': actual_val,
                                'Violations_Actuelles': int(violations_actuelles),
                                'Violations_Mois_Precedent': int(violations_mois_precedent),
                                'Violations_Corrigees': int(violations_corrigees),
                                'Score_Qualite': score_qualite
                            })

                        df_cat = pd.DataFrame(rows_cat)

                        # Formule demandée: nombre de règle violée / Actual * 100
                        df_cat['Ratio_Violations_sur_100'] = (
                            (df_cat['Violations_Actuelles'] / df_cat['Actual_Reports'].replace(0, np.nan)) * 100
                        ).fillna(0).round(2)

                        if 'Violations_Corrigees' not in df_cat.columns:
                            df_cat['Violations_Corrigees'] = 0
                        df_cat['Violations_Corrigees'] = pd.to_numeric(
                            df_cat['Violations_Corrigees'], errors='coerce'
                        ).fillna(0).astype(int)

                        st.subheader(f"Résultats : {selected_vr_name}")

                        df_display = df_cat[['Zone de santé', 'Complétude_Globale', 'Promptitude_Globale', 'Actual_Reports', 'Violations_Mois_Precedent', 'Violations_Corrigees', 'Violations_Actuelles', 'Ratio_Violations_sur_100', 'Score_Qualite']].rename(columns={
                            'Complétude_Globale': 'Complétude globale (%)',
                            'Promptitude_Globale': 'Promptitude globale (%)',
                            'Actual_Reports': 'Reports_Actual',
                            'Violations_Mois_Precedent': 'Règles violées (M-1)',
                            'Violations_Corrigees': 'Règles corrigées (M-1 -> M)',
                            'Violations_Actuelles': 'Règles violées (M)',
                            'Ratio_Violations_sur_100': 'Ratio / 100 rapports',
                            'Score_Qualite': 'Score de qualité'
                        })

                        st.dataframe(
                            df_display.style.format({
                                'Complétude globale (%)': '{:.2f}',
                                'Promptitude globale (%)': '{:.2f}',
                                'Reports_Actual': '{:.2f}',
                                'Ratio / 100 rapports': '{:.2f}',
                                'Score de qualité': '{:.2f}'
                            })
                            .map(style_taux, subset=['Complétude globale (%)', 'Promptitude globale (%)', 'Score de qualité'])
                            .highlight_max(subset=['Règles violées (M)'], color='#ffcccc')
                            .map(lambda x: 'color: red; font-weight: bold' if x > 10 else '', subset=['Ratio / 100 rapports']),
                            use_container_width=True
                        )

                        col_res1, col_res2 = st.columns(2)
                        col_res1.metric("Total violations détectées", int(df_cat['Violations_Actuelles'].sum()))
                        col_res2.metric("Total règles corrigées", int(df_cat['Violations_Corrigees'].sum()), delta=f"{int(df_cat['Violations_Corrigees'].sum())}")

                else:
                    st.error("L'ID de l'unité d'organisation est manquant dans les données sources.")
        else:
            st.warning("⚠️ Impossible de charger les groupes de règles.")

    # --- MENU SIDEBAR : EXTRACTION COMMENTEE ---
    with st.sidebar:
        st.divider()
        st.subheader("📤 Extraction du rapport")
        report_type = st.selectbox(
            "Type de téléchargement :",
            ["Excel", "Word", "PowerPoint"],
            key="report_type_selector"
        )

        comments_df = build_dashboard_comments_df(
            df_final_c=df_final_c,
            df_final_p=df_final_p,
            df_synth=df_synth,
            selected_zone=selected_zone,
            df_tab4_fusion=df_tab4_fusion
        )

        export_bytes = None
        export_mime = None
        export_name = None
        export_error = None

        if report_type == "Excel":
            export_bytes = to_excel_dashboard_report(
                df_final_c=df_final_c,
                df_final_p=df_final_p,
                df_synth=df_synth,
                comments_df=comments_df,
                df_tab4_fusion=df_tab4_fusion
            )
            export_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            export_name = f"rapport_dashboard_commente_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        elif report_type == "Word":
            export_bytes, export_error = to_word_dashboard_report(
                df_final_c=df_final_c,
                df_final_p=df_final_p,
                df_synth=df_synth,
                comments_df=comments_df,
                df_tab4_fusion=df_tab4_fusion
            )
            export_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            export_name = f"rapport_dashboard_commente_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
        else:
            export_bytes, export_error = to_powerpoint_dashboard_report(
                df_synth=df_synth,
                comments_df=comments_df,
                fig_comp=fig_comp,
                fig_quad=fig_quad,
                df_comparatif=df_comparatif,
                df_tab4_fusion=df_tab4_fusion
            )
            export_mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            export_name = f"rapport_dashboard_commente_{datetime.now().strftime('%Y%m%d_%H%M')}.pptx"

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
else:
    st.error("❌ Données indisponibles.")

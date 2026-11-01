"""Tableau de score qualité : reporting courant et violations rattachées par ID."""

import math

import pandas as pd

from geography import labels_for_units
from reporting_metrics import clean_ou_id


def _number(value):
    numeric = pd.to_numeric(value, errors='coerce')
    return float(numeric) if pd.notna(numeric) and math.isfinite(float(numeric)) else float('nan')


def build_quality_rows(metrics, zones, validation_data=None):
    """Prépare les données numériques du périmètre, sans réutiliser une autre ZS.

    Les indicateurs de déclaration viennent toujours des données courantes.
    Seules les colonnes de violations d'une analyse correspondante sont jointes.
    Une analyse absente ou incomplète laisse ces colonnes indisponibles.
    """
    labels = labels_for_units(zones, 'Zone_ID', 'Zone')
    metric_map = {
        clean_ou_id(row.get('Zone_ID')): row for row in metrics.to_dict('records')
    } if metrics is not None else {}
    validation_map = {
        clean_ou_id(row.get('Zone_ID')): row for row in validation_data.to_dict('records')
    } if validation_data is not None else {}
    rows = []
    for zone_id, label in labels.items():
        source = metric_map.get(zone_id, {})
        validation = validation_map.get(zone_id, {})
        completeness = _number(source.get('Complétude_Globale'))
        timeliness = _number(source.get('Promptitude_Globale'))
        row = {
            'Zone_ID': zone_id, 'Zone de santé': label,
            'Complétude_Globale': completeness,
            'Promptitude_Globale': timeliness,
            'Actual_Reports': _number(source.get('Actual_Reports')),
            'Score_Qualite': round((completeness + timeliness) / 2, 1),
            'Nombre_Datasets_95': _number(source.get('Nombre_Datasets_95')),
            'Profil_Bareme': 'ZS',
            'Erreur_Validation': validation.get('Erreur_Validation') or '',
            'Erreur_Corrections': validation.get('Erreur_Corrections') or '',
            'Controle_Corrections_Effectue': bool(validation.get('Controle_Corrections_Effectue', False)),
        }
        for column in ['Violations_Actuelles', 'Violations_Mois_Precedent', 'Violations_Corrigees']:
            row[column] = _number(validation.get(column))
        rows.append(row)
    return pd.DataFrame(rows)

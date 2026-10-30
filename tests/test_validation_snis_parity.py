"""Contrat de l'écran DHIS2 Data Quality, sans données ni connexion réelles.

Les noms et valeurs visibles dans la capture fournie servent de cas de
régression ; les identifiants sont fictifs. La réponse suit les champs plats
du convertisseur officiel de data-quality-app.
"""

from copy import deepcopy
from datetime import date
import unittest

from validation_api import fetch_validation_analysis
from validation_details import build_validation_details, detail_export_frame


SNIS_ROWS = [
    {
        "organisationUnitId": "DigbaAAAAAA",
        "organisationUnitDisplayName": "bu Digba Centre de Santé",
        "validationRuleId": "SuspectAAAA",
        "validationRuleDescription": (
            "1q-Cas suspect de Paludisme de moins de 5 ans doit être inférieur "
            "ou égal aux Nouveaux"
        ),
        "periodId": "202609", "periodDisplayName": "Septembre 2026",
        "importance": "MEDIUM", "leftSideValue": 84, "operator": "<=", "rightSideValue": 59,
        "attributeOptionComboId": "DefaultAAAA", "attributeOptionComboDisplayName": "default",
    },
    {
        "organisationUnitId": "MakpoyoAAAA",
        "organisationUnitDisplayName": "bu Makpoyo Centre de Santé",
        "validationRuleId": "TdrRealAAAA",
        "validationRuleDescription": (
            "1q-TDR réalisé de 2 - < 5 ans doit être inférieur ou égal au Cas "
            "suspect de Paludisme de 2 - < 5"
        ),
        "periodId": "202609", "periodDisplayName": "Septembre 2026",
        "importance": "MEDIUM", "leftSideValue": 20, "operator": "<=", "rightSideValue": 18,
        "attributeOptionComboId": "DefaultAAAA", "attributeOptionComboDisplayName": "default",
    },
    {
        "organisationUnitId": "MbibiliAAAA",
        "organisationUnitDisplayName": "bu Mbibili Centre de Santé",
        "validationRuleId": "EqualityAAA", "validationRuleDescription": "",
        "periodId": "202609", "periodDisplayName": "Septembre 2026",
        "importance": "HIGH", "leftSideValue": 567, "operator": "==", "rightSideValue": 566,
        "attributeOptionComboId": "DefaultAAAA", "attributeOptionComboDisplayName": "default",
    },
    {
        "organisationUnitId": "NgaluAAAAAA",
        "organisationUnitDisplayName": "bu Ngalu Poste de Santé",
        "validationRuleId": "ConfirmedAA",
        "validationRuleDescription": (
            "1q-Paludisme simple confirmé de 2 - < 5 ans doit être inférieur "
            "ou égal au TDR positif de 2"
        ),
        "periodId": "202609", "periodDisplayName": "Septembre 2026",
        "importance": "MEDIUM", "leftSideValue": 6, "operator": "<=", "rightSideValue": 5,
        "attributeOptionComboId": "DefaultAAAA", "attributeOptionComboDisplayName": "default",
    },
]


class SnisResponseClient:
    def __init__(self, rows=None):
        self.rows = deepcopy(SNIS_ROWS if rows is None else rows)
        self.calls = []

    def post(self, endpoint, data=None):
        self.calls.append((endpoint, deepcopy(data)))
        return deepcopy(self.rows)

    def get(self, endpoint, params=None):
        raise AssertionError("L'analyse manuelle DHIS2 utilise POST JSON.")


class SnisValidationParityTests(unittest.TestCase):
    def test_four_server_rows_without_paths_reach_the_table_unchanged(self):
        client = SnisResponseClient()
        source = fetch_validation_analysis(
            client, "AngoZoneAAA", None,
            start_date=date(2026, 9, 1), end_date=date(2026, 9, 30),
        )
        view = detail_export_frame(build_validation_details(source))

        self.assertEqual(source.attrs["failed_periods"], [])
        self.assertEqual(len(source), 4)
        self.assertEqual(len(view), 4)
        self.assertEqual(view["Unité d’organisation"].tolist(), [
            "bu Digba Centre de Santé", "bu Makpoyo Centre de Santé",
            "bu Mbibili Centre de Santé", "bu Ngalu Poste de Santé",
        ])
        self.assertEqual(view["Période"].tolist(), ["Septembre 2026"] * 4)
        self.assertEqual(view["Valeur gauche"].tolist(), [84, 20, 567, 6])
        self.assertEqual(view["Valeur droite"].tolist(), [59, 18, 566, 5])
        self.assertEqual(view["Importance"].tolist(), ["MEDIUM", "MEDIUM", "HIGH", "MEDIUM"])
        self.assertEqual(view["Opérateur"].tolist(), ["<=", "<=", "==", "<="])
        self.assertEqual(view["Règle de validation"].tolist(), [row["validationRuleDescription"] for row in SNIS_ROWS])

    def test_parent_and_whole_interval_are_sent_in_one_manual_analysis(self):
        client = SnisResponseClient()
        source = fetch_validation_analysis(
            client, "AngoZoneAAA", ["202608", "202609"], "MalariaAAAA",
            start_date="2026-08-27", end_date="2026-09-30",
        )
        self.assertEqual(len(client.calls), 1)
        endpoint, parameters = client.calls[0]
        self.assertEqual(endpoint, "dataAnalysis/validationRules")
        self.assertEqual(parameters["ou"], "AngoZoneAAA")
        self.assertEqual(parameters["startDate"], "2026-08-27")
        self.assertEqual(parameters["endDate"], "2026-09-30")
        self.assertEqual(parameters["vrg"], "MalariaAAAA")
        self.assertIs(parameters["persist"], False)
        self.assertIs(parameters["notification"], False)
        self.assertEqual(len(source), 4)


if __name__ == "__main__":
    unittest.main()

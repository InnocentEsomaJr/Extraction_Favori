"""Détails des violations, périmètres et exports sans connexion DHIS2 réelle."""

import csv
import io
import unittest
from unittest import mock
from xml.etree import ElementTree as ET
import zipfile

import pandas as pd
import streamlit as st
from streamlit.testing.v1 import AppTest

from categorisation import TOTAL_COL
from reporting_metrics import clean_ou_id
import test_dashboard_categorisation as dashboard_test
import test_dashboard_filters as filter_test


VALIDATION_CLIENT = '''
def get_dhis2_client():
    class Client:
        def post(self, endpoint, data=None):
            if endpoint != "dataAnalysis/validationRules":
                raise AssertionError("Appel fixture non prévu : " + endpoint)
            from validation_filters import date_range_months
            params = data
            queries = list(st.session_state.get("_fixture_detail_queries", []))
            queries.append(dict(params))
            st.session_state["_fixture_detail_queries"] = queries
            units = [
                ("CountryAAAA", "ProvinceAAA", "ZoneAAAAAAA", "CentreAAAAA", "RuleAAAAAAA", 0, 3),
                ("CountryAAAA", "ProvinceAAA", "ZoneBBBBBBB", "CentreBBBBB", "RuleBBBBBBB", 12, 10),
                ("CountryAAAA", "ProvinceBBB", "ZoneCCCCCCC", "CentreCCCCC", "RuleCCCCCCC", 25, 0),
                ("CountryBBBB", "ProvinceCCC", "ZoneDDDDDDD", "CentreDDDDD", "RuleDDDDDDD", 1, 0),
            ]
            results = []
            for period in date_range_months(params["startDate"], params["endDate"]):
                if period.endswith("12"):
                    continue
                for country, province, zone, centre, rule, left, right in units:
                    if params["ou"] not in (country, province, zone):
                        continue
                    if params.get("vrg") and rule != "RuleAAAAAAA":
                        continue
                    results.append({
                        "id": period + rule, "periodId": period,
                        "organisationUnitId": centre, "validationRuleId": rule,
                        "organisationUnitPath": f"/{country}/{province}/{zone}/{centre}",
                        "leftSideValue": left, "rightSideValue": right,
                        "attributeOptionComboId": "AttributeAA",
                    })
            return results
    return Client()
'''


DETAIL_METADATA = '''
def get_validation_detail_metadata(rule_ids, org_ids, cache_user=None):
    st.session_state["_fixture_detail_metadata_calls"] = st.session_state.get("_fixture_detail_metadata_calls", 0) + 1
    st.session_state["_fixture_detail_rule_ids"] = list(rule_ids)
    st.session_state["_fixture_detail_org_ids"] = list(org_ids)
    rules = {}
    for uid, label, importance, operator in [
        ("RuleAAAAAAA", "Vaccinations au-delà des consultations", "HIGH", "greater_than"),
        ("RuleBBBBBBB", "Consultations attendues", "MEDIUM", "equal_to"),
        ("RuleCCCCCCC", "Notification obligatoire", "LOW", "less_than_or_equal_to"),
    ]:
        rules[uid] = {
            "id": uid, "displayName": label, "importance": importance, "operator": operator,
            "leftSide": {"description": "Valeur déclarée", "expression": "#{LeftAAAAAAA}"},
            "rightSide": {"description": "Valeur attendue", "expression": "#{RightAAAAAA}"},
        }
    orgs = {
        "CentreAAAAA": {"id": "CentreAAAAA", "displayName": "Centre Étoile", "path": "/CountryAAAA/ProvinceAAA/ZoneAAAAAAA/CentreAAAAA"},
        "CentreBBBBB": {"id": "CentreBBBBB", "displayName": "Centre B", "path": "/CountryAAAA/ProvinceAAA/ZoneBBBBBBB/CentreBBBBB"},
        "CentreCCCCC": {"id": "CentreCCCCC", "displayName": "Centre C", "path": "/CountryAAAA/ProvinceBBB/ZoneCCCCCCC/CentreCCCCC"},
        "CentreDDDDD": {"id": "CentreDDDDD", "displayName": "Centre D", "path": "/CountryBBBB/ProvinceCCC/ZoneDDDDDDD/CentreDDDDD"},
    }
    return {uid: rules[uid] for uid in rule_ids if uid in rules}, {uid: orgs[uid] for uid in org_ids if uid in orgs}, []
'''


MISSING_METADATA = '''
def get_validation_detail_metadata(rule_ids, org_ids, cache_user=None):
    st.session_state["_fixture_detail_metadata_calls"] = st.session_state.get("_fixture_detail_metadata_calls", 0) + 1
    return {}, {}, ["Métadonnées de règles et d'unités indisponibles"]
'''


PATHLESS_VALIDATION_CLIENT = VALIDATION_CLIENT.replace(
    '                        "organisationUnitPath": f"/{country}/{province}/{zone}/{centre}",\n',
    '',
)


FORBIDDEN_VALIDATION_CLIENT = '''
def get_dhis2_client():
    class Client:
        def post(self, endpoint, data=None):
            import requests
            response = requests.Response()
            response.status_code = 403
            raise requests.HTTPError("Accès refusé", response=response)
    return Client()
'''


PREVIOUS_FAILURE_CLIENT = VALIDATION_CLIENT.replace(
    '                if period.endswith("12"):\n                    continue',
    '''                if period.endswith("12"):
                    import requests
                    response = requests.Response()
                    response.status_code = 403
                    raise requests.HTTPError("Accès refusé au mois précédent", response=response)''',
)


def excel_sheet(data, sheet_name):
    """Lit les cellules exportées sans dépendre d'un moteur Excel installé."""
    ns = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        shared_path = "xl/sharedStrings.xml"
        strings = ["".join(item.itertext()) for item in ET.fromstring(archive.read(shared_path))] if shared_path in archive.namelist() else []
        workbook = ET.fromstring(archive.read("xl/workbook.xml"))
        sheets = workbook.find("m:sheets", ns)
        index = next(i for i, sheet in enumerate(sheets, 1) if sheet.attrib["name"] == sheet_name)
        rows = ET.fromstring(archive.read(f"xl/worksheets/sheet{index}.xml")).find("m:sheetData", ns)

        def value(cell):
            item = cell.find("m:v", ns)
            if cell.attrib.get("t") == "s":
                return strings[int(item.text)]
            if cell.attrib.get("t") == "inlineStr":
                return "".join(cell.find("m:is", ns).itertext())
            return float(item.text) if item is not None else None

        header_cells = list(rows[0])
        headers = [value(cell) for cell in header_cells]
        column_keys = ["".join(char for char in cell.attrib["r"] if char.isalpha()) for cell in header_cells]
        output = []
        for row in rows[1:]:
            cells = {"".join(char for char in cell.attrib["r"] if char.isalpha()): value(cell) for cell in row}
            output.append({header: cells.get(key) for header, key in zip(headers, column_keys)})
        return headers, output


class DashboardValidationDetailTests(unittest.TestCase):
    def widget(self, app, kind, label):
        return next(widget for widget in getattr(app, kind) if widget.label == label)

    def run_app(self, app):
        app.run(timeout=30)
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])

    def new_app(self, metadata=DETAIL_METADATA, client=VALIDATION_CLIENT):
        source = dashboard_test.app_fixture_source({
            "get_organisation_catalog": filter_test.CATALOG_FIXTURE,
            "get_data": filter_test.DATA_FIXTURE,
            "get_dhis2_client": client,
            "get_validation_detail_metadata": metadata,
        })
        app = AppTest.from_string(source, default_timeout=30)
        app.session_state["connected"] = True
        app.session_state["dhis2_user"] = "fixture-details-" + self._testMethodName
        app.session_state["dhis2_pass"] = "fixture-password"
        return app

    def detail_view(self, app):
        candidates = [table.value for table in app.dataframe if {"Règle de validation", "Importance", "Période"}.issubset(table.value.columns)]
        return candidates[0] if candidates else pd.DataFrame()

    def analyse(self, app):
        self.widget(app, "button", "Lancer l'analyse des violations").click()
        self.run_app(app)
        return self.detail_view(app)

    def test_country_details_keep_zero_values_and_filtered_csv_leaves_excel_complete(self):
        app = self.new_app()
        downloads = {}
        real_download = st.download_button

        def capture(label, *args, **kwargs):
            downloads[label] = kwargs.get("data", args[0] if args else None)
            return real_download(label, *args, **kwargs)

        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")), \
                mock.patch("streamlit.download_button", side_effect=capture):
            self.run_app(app)
            frame = self.analyse(app)
            self.assertEqual(len(frame), 3)
            self.assertEqual(set(frame["Unité d’organisation"]), {"Centre Étoile", "Centre B", "Centre C"})
            high = frame.loc[frame["Importance"].eq("HIGH")].iloc[0]
            self.assertEqual(high["Valeur gauche"], 0)
            self.assertEqual(high["Opérateur"], ">")
            self.assertIn("déclarée", high["Détails"])
            self.assertEqual(frame.loc[frame["Importance"].eq("LOW"), "Valeur droite"].iloc[0], 0)
            queries = len(app.session_state["_fixture_detail_queries"])
            self.assertEqual(queries, 1)
            self.assertEqual(app.session_state["_fixture_detail_metadata_calls"], 1)

            self.widget(app, "text_input", "Rechercher dans les violations").input("etoile")
            self.widget(app, "selectbox", "Filtrer les violations par importance").select("HIGH")
            self.run_app(app)
            self.assertEqual(self.detail_view(app)["Unité d’organisation"].tolist(), ["Centre Étoile"])
            self.assertEqual(len(app.session_state["_fixture_detail_queries"]), queries)
            self.assertEqual(app.session_state["_fixture_detail_metadata_calls"], 1)

            csv_data = downloads["Télécharger les violations affichées en CSV"]
            if isinstance(csv_data, bytes):
                csv_data = csv_data.decode("utf-8-sig")
            dialect = csv.Sniffer().sniff(csv_data.splitlines()[0])
            csv_rows = list(csv.DictReader(io.StringIO(csv_data), dialect=dialect))
            self.assertEqual(len(csv_rows), 1)
            self.assertEqual(csv_rows[0]["Unité d’organisation"], "Centre Étoile")
            self.assertEqual(float(csv_rows[0]["Valeur gauche"]), 0)

            excel_data = downloads["Télécharger Excel"]
            headers, rows = excel_sheet(excel_data, "Details_Violations")
            self.assertIn("Règle de validation", headers)
            self.assertEqual(len(rows), 3)
            self.assertEqual(set(row["Unité d’organisation"] for row in rows), {"Centre Étoile", "Centre B", "Centre C"})
            self.assertEqual(next(row["Valeur gauche"] for row in rows if row["Importance"] == "HIGH"), 0)

            self.widget(app, "button", "Visualiser le rapport").click()
            self.run_app(app)
            self.assertTrue(any(tab.label == "Règles détaillées" for tab in app.tabs))
            self.assertEqual(len(app.session_state["_fixture_detail_queries"]), queries)

    def test_details_follow_province_then_one_selected_health_zone(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            self.widget(app, "selectbox", "Niveau d'analyse :").select("Province")
            self.run_app(app)
            frame = self.analyse(app)
            self.assertEqual(set(frame["Unité d’organisation"]), {"Centre Étoile", "Centre B"})
            self.assertEqual(set(query["ou"] for query in app.session_state["_fixture_detail_queries"]), {"ProvinceAAA"})

            self.widget(app, "selectbox", "Province :").select("ProvinceBBB")
            self.run_app(app)
            self.widget(app, "selectbox", "Zone de santé :").select("ZoneCCCCCCC")
            self.run_app(app)
            app.session_state["_fixture_detail_queries"] = []
            frame = self.analyse(app)
            self.assertEqual(frame["Unité d’organisation"].tolist(), ["Centre C"])
            self.assertEqual(set(query["ou"] for query in app.session_state["_fixture_detail_queries"]), {"ZoneCCCCCCC"})

    def test_rule_group_and_month_range_are_reflected_in_visible_details(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            year = str(self.widget(app, "selectbox", "Choisir l'année :").value)
            self.widget(app, "selectbox", "Sélectionner le Groupe de Règles de Validation :").select("Groupe test")
            self.widget(app, "selectbox", "Mois fin :").select("Février")
            self.run_app(app)
            frame = self.analyse(app)
            self.assertEqual(len(frame), 2)
            self.assertEqual(frame["Unité d’organisation"].tolist(), ["Centre Étoile", "Centre Étoile"])
            self.assertEqual(set(frame["Période"]), {f"Janvier {year}", f"Février {year}"})
            summary = app.session_state["_test_export_frame"]
            self.assertEqual(sorted(summary["Règles violées (M)"].tolist()), [0, 0, 2])
            queries = app.session_state["_fixture_detail_queries"]
            self.assertTrue(all(query.get("vrg") == "groupAAAAAA" for query in queries))
            self.assertTrue(all(query["persist"] is False and query["notification"] is False for query in queries))
            current_dates = {query["startDate"] for query in queries if query["startDate"].startswith(year)}
            self.assertEqual(current_dates, {f"{year}-01-01"})
            self.assertEqual(queries[0]["endDate"], f"{year}-02-28")

    def test_post_response_without_paths_is_displayed_and_assigned_using_metadata(self):
        self.assertNotIn("organisationUnitPath", PATHLESS_VALIDATION_CLIENT)
        app = self.new_app(client=PATHLESS_VALIDATION_CLIENT)
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            frame = self.analyse(app)
            self.assertEqual(len(frame), 3)
            self.assertEqual(set(frame["Unité d’organisation"]), {"Centre Étoile", "Centre B", "Centre C"})
            summary = app.session_state["_test_export_frame"]
            self.assertEqual(summary["Règles violées (M)"].tolist(), [1, 1, 1])
            self.assertEqual(app.session_state["categorisation_result"]["validation_warnings"], [])
            requests = app.session_state["categorisation_result"]["validation_requests"]
            self.assertEqual([request["Méthode"] for request in requests], ["POST"])
            self.assertEqual([request["Unité parent"] for request in requests], ["CountryAAAA"])
            self.assertEqual(requests[0]["Lignes reçues"], 3)
            self.assertTrue(any(
                item.label == "Vérifier la requête envoyée à DHIS2" for item in app.expander
            ))

            queries = list(app.session_state["_fixture_detail_queries"])
            repeated = self.analyse(app)
            self.assertEqual(len(app.session_state["_fixture_detail_queries"]), len(queries) + 1)
            pd.testing.assert_frame_equal(repeated, frame)

            queries = list(app.session_state["_fixture_detail_queries"])
            self.widget(app, "button", "Actualiser les données").click()
            self.run_app(app)
            self.assertNotIn("categorisation_result", app.session_state)
            self.analyse(app)
            self.assertEqual(len(app.session_state["_fixture_detail_queries"]), len(queries) + 1)

    def test_forbidden_analysis_explains_the_http_failure_and_never_reports_a_healthy_zero(self):
        app = self.new_app(client=FORBIDDEN_VALIDATION_CLIENT)
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            frame = self.analyse(app)
            self.assertTrue(frame.empty)
            summary = app.session_state["_test_export_frame"]
            self.assertTrue(summary["Règles violées (M)"].isna().all())
            saved = app.session_state["categorisation_result"]
            self.assertTrue(any("403" in str(message) for message in saved["validation_warnings"]))
            self.assertTrue(any("403" in str(item.value) for item in app.text))
            self.assertFalse(any("Aucune règle violée détectée" in str(item.value) for item in app.success))
            self.assertTrue(self.widget(app, "button", "Contrôler les corrections du mois précédent").disabled)

    def test_current_analysis_and_previous_corrections_have_independent_buttons_and_requests(self):
        app = self.new_app(client=PATHLESS_VALIDATION_CLIENT)
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            control_label = "Contrôler les corrections du mois précédent"
            self.assertTrue(self.widget(app, "button", control_label).disabled)

            details = self.analyse(app).copy(deep=True)
            current = app.session_state["_test_export_frame"].copy(deep=True)
            saved = app.session_state["categorisation_result"]
            timestamp = saved["analysed_at"]
            self.assertEqual(len(app.session_state["_fixture_detail_queries"]), 1)
            self.assertFalse(saved["corrections_checked"])
            self.assertFalse(self.widget(app, "button", control_label).disabled)
            self.assertTrue(current["Règles violées (M-1)"].isna().all())
            self.assertTrue(current["Règles corrigées (M-1 -> M)"].isna().all())
            self.assertTrue(current[TOTAL_COL].isna().all())
            self.assertEqual(current["Règles violées (M)"].tolist(), [1, 1, 1])

            self.widget(app, "button", control_label).click()
            self.run_app(app)
            checked = app.session_state["categorisation_result"]
            summary = app.session_state["_test_export_frame"]
            self.assertEqual(len(app.session_state["_fixture_detail_queries"]), 2)
            self.assertTrue(checked["corrections_checked"])
            self.assertTrue(checked.get("corrections_checked_at"))
            self.assertEqual(checked["analysed_at"], timestamp)
            pd.testing.assert_frame_equal(self.detail_view(app), details)
            for column in [
                "Complétude globale (%)", "Promptitude globale (%)", "Reports_Actual",
                "Règles violées (M)", "Ratio / 100 rapports", "Score de qualité",
                "Score de qualité /20", "Score violations /20",
            ]:
                pd.testing.assert_series_equal(summary[column], current[column])
            self.assertEqual(summary["Règles violées (M-1)"].tolist(), [0, 0, 0])
            self.assertTrue(summary[TOTAL_COL].notna().all())
            self.assertEqual([request["Période"] for request in checked["validation_requests"]], ["M", "M-1"])

            self.widget(app, "text_input", "Rechercher dans les violations").input("etoile")
            self.run_app(app)
            self.assertEqual(len(app.session_state["_fixture_detail_queries"]), 2)

            self.analyse(app)
            renewed = app.session_state["categorisation_result"]
            self.assertEqual(len(app.session_state["_fixture_detail_queries"]), 3)
            self.assertFalse(renewed["corrections_checked"])
            self.assertFalse(renewed.get("corrections_checked_at"))
            self.assertTrue(app.session_state["_test_export_frame"]["Règles violées (M-1)"].isna().all())
            self.assertEqual([request["Période"] for request in renewed["validation_requests"]], ["M"])

            self.widget(app, "selectbox", "Sélectionner le Groupe de Règles de Validation :").select("Groupe test")
            self.run_app(app)
            self.assertNotIn("categorisation_result", app.session_state)
            self.assertTrue(self.widget(app, "button", control_label).disabled)

    def test_previous_month_failure_preserves_the_current_dhis2_results_and_notes(self):
        app = self.new_app(client=PREVIOUS_FAILURE_CLIENT)
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            details = self.analyse(app).copy(deep=True)
            current = app.session_state["_test_export_frame"].copy(deep=True)
            timestamp = app.session_state["categorisation_result"]["analysed_at"]
            self.assertEqual(len(app.session_state["_fixture_detail_queries"]), 1)

            self.widget(app, "button", "Contrôler les corrections du mois précédent").click()
            self.run_app(app)
            saved = app.session_state["categorisation_result"]
            summary = app.session_state["_test_export_frame"]
            self.assertEqual(len(app.session_state["_fixture_detail_queries"]), 2)
            self.assertEqual(saved["analysed_at"], timestamp)
            self.assertEqual(saved["validation_warnings"], [])
            self.assertTrue(any("403" in str(message) for message in saved["correction_warnings"]))
            pd.testing.assert_frame_equal(self.detail_view(app), details)
            for column in [
                "Complétude globale (%)", "Promptitude globale (%)", "Reports_Actual",
                "Règles violées (M)", "Ratio / 100 rapports", "Score de qualité",
                "Score de qualité /20", "Score violations /20",
            ]:
                pd.testing.assert_series_equal(summary[column], current[column])
            self.assertTrue(summary["Règles violées (M-1)"].isna().all())
            self.assertTrue(summary["Règles corrigées (M-1 -> M)"].isna().all())
            self.assertTrue(summary[TOTAL_COL].isna().all())
            self.assertFalse(self.widget(app, "button", "Contrôler les corrections du mois précédent").disabled)

    def test_missing_metadata_preserves_identifiers_and_actual_values(self):
        app = self.new_app(metadata=MISSING_METADATA)
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            frame = self.analyse(app)
            self.assertEqual(len(frame), 3)
            self.assertEqual(set(frame["Unité d’organisation"]), {"CentreAAAAA", "CentreBBBBB", "CentreCCCCC"})
            self.assertEqual(set(frame["Règle de validation"]), {"RuleAAAAAAA", "RuleBBBBBBB", "RuleCCCCCCC"})
            self.assertTrue(frame["Importance"].str.contains("renseign", case=False).all())
            self.assertEqual(frame.loc[frame["Règle de validation"].eq("RuleAAAAAAA"), "Valeur gauche"].iloc[0], 0)
            self.assertGreater(len(app.warning) + len(app.info), 0)


class ValidationDetailMetadataApiTests(unittest.TestCase):
    def test_metadata_is_batched_and_partial_failure_remains_nonfatal(self):
        calls = []

        class Client:
            def get(self, endpoint, params=None):
                calls.append((endpoint, params))
                ids = params["filter"].removeprefix("id:in:[").removesuffix("]").split(",")
                if endpoint == "organisationUnits":
                    raise RuntimeError("fixture métadonnée indisponible")
                return {endpoint: [{"id": uid, "displayName": "Règle " + uid} for uid in ids]}

        get_metadata = dashboard_test.dashboard_functions(
            "get_validation_detail_metadata", clean_ou_id=clean_ou_id,
            get_dhis2_client=lambda: Client(),
        )["get_validation_detail_metadata"]
        ids = [f"Rule{number:07d}" for number in range(121)]
        rules, orgs, errors = get_metadata(ids + [ids[0]], ["CentreAAAAA"])
        self.assertEqual(set(rules), set(ids))
        self.assertEqual(orgs, {})
        self.assertTrue(errors)
        rule_calls = [params for endpoint, params in calls if endpoint == "validationRules"]
        self.assertEqual(len(rule_calls), 3)
        self.assertTrue(all(params["paging"] == "false" for _, params in calls))
        self.assertTrue(all("importance" in params["fields"] and "leftSide" in params["fields"] for params in rule_calls))

    def test_no_detail_does_not_require_a_client_or_metadata_queries(self):
        client_factory = mock.Mock(side_effect=AssertionError("Aucun client nécessaire"))
        get_metadata = dashboard_test.dashboard_functions(
            "get_validation_detail_metadata", clean_ou_id=clean_ou_id,
            get_dhis2_client=client_factory,
        )["get_validation_detail_metadata"]
        self.assertEqual(get_metadata([], []), ({}, {}, []))
        client_factory.assert_not_called()


if __name__ == "__main__":
    unittest.main()

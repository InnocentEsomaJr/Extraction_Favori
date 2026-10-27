"""Périmètre propre à la validation et dates exactes, sans réseau réel."""

from datetime import date
import unittest
from unittest import mock

import pandas as pd
import streamlit as st
from streamlit.testing.v1 import AppTest

import test_dashboard_categorisation as dashboard_test
import test_dashboard_filters as filter_test
import test_dashboard_validation_details as detail_test


SCOPED_DATA = '''
@st.cache_data(ttl=3600)
def get_data(favori_id, period=None, cache_user=None, ou_dimension=None):
    requests = list(st.session_state.get("_date_data_requests", []))
    requests.append({"period": period, "ou_dimension": ou_dimension, "cache_user": cache_user})
    st.session_state["_date_data_requests"] = requests
    units = [
        ("ZoneAAAAAAA", "ZS Commune", "ProvinceAAA", "CountryAAAA", 10),
        ("ZoneBBBBBBB", "ZS B", "ProvinceAAA", "CountryAAAA", 20),
        ("ZoneCCCCCCC", "ZS Commune", "ProvinceBBB", "CountryAAAA", 30),
        ("ZoneDDDDDDD", "ZS D", "ProvinceCCC", "CountryBBBB", 40),
    ]
    if ou_dimension and ou_dimension.startswith("LEVEL-3;"):
        boundary = ou_dimension.split(";", 1)[1]
        selected = [unit for unit in units if boundary in (unit[2], unit[3])]
    else:
        selected = [unit for unit in units if unit[0] == ou_dimension]
    multiplier = len(str(period).split(";"))
    rows = []
    for uid, name, province, country, base in selected:
        row = {"Organisation unit": name, "Organisation unit ID": uid}
        for number in range(10):
            prefix = f"Dataset {number} "
            row[prefix + "Actual reports"] = base * multiplier
            row[prefix + "Expected reports"] = base * multiplier
            row[prefix + "Actual reports on time"] = base * multiplier
            row[prefix + "Reporting rate"] = 100
            row[prefix + "Reporting rate on time"] = 100
        rows.append(row)
    return pd.DataFrame(rows)
'''


MULTI_COUNTRY_CATALOG = filter_test.CATALOG_FIXTURE.replace(
    "    return build_geography_catalog(units)",
    '''    second = {"id": "CountryBBBB", "displayName": "Pays secondaire", "level": 1}
    province = {"id": "ProvinceCCC", "displayName": "Province C", "level": 2}
    units.extend([second, {**province, "ancestors": [second]}, {
        "id": "ZoneDDDDDDD", "displayName": "ZS D", "level": 3, "ancestors": [second, province]
    }])
    return build_geography_catalog(units)''',
)


MULTI_COUNTRY_CLIENT = detail_test.VALIDATION_CLIENT.replace(
    '"ZoneCCCCCCC": ("CentreCCCCC", "RuleCCCCCCC", 25, 0),',
    '"ZoneCCCCCCC": ("CentreCCCCC", "RuleCCCCCCC", 25, 0),\n                "ZoneDDDDDDD": ("CentreDDDDD", "RuleDDDDDDD", 1, 0),',
).replace(
    'province = "ProvinceBBB" if params["ou"] == "ZoneCCCCCCC" else "ProvinceAAA"',
    'province = {"ZoneCCCCCCC": "ProvinceBBB", "ZoneDDDDDDD": "ProvinceCCC"}.get(params["ou"], "ProvinceAAA")\n            country = "CountryBBBB" if params["ou"] == "ZoneDDDDDDD" else "CountryAAAA"',
).replace('/CountryAAAA/{province}/', '/{country}/{province}/')


class ValidationDateFilterTests(unittest.TestCase):
    def widget(self, app, kind, label):
        return next(widget for widget in getattr(app, kind) if widget.label == label)

    def run_app(self, app):
        app.run(timeout=30)
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])

    def new_app(self, multi_country=False):
        source = dashboard_test.app_fixture_source({
            "get_organisation_catalog": MULTI_COUNTRY_CATALOG if multi_country else filter_test.CATALOG_FIXTURE,
            "get_data": SCOPED_DATA,
            "get_dhis2_client": MULTI_COUNTRY_CLIENT if multi_country else detail_test.VALIDATION_CLIENT,
            "get_validation_detail_metadata": detail_test.DETAIL_METADATA,
        })
        app = AppTest.from_string(source, default_timeout=30)
        app.session_state["connected"] = True
        app.session_state["dhis2_user"] = "fixture-date-" + self._testMethodName
        app.session_state["dhis2_pass"] = "fixture-password"
        if multi_country:
            app.session_state["filter_country_main"] = "CountryAAAA"
        return app

    def analyse(self, app):
        self.widget(app, "button", "Lancer l'analyse des violations").click()
        self.run_app(app)
        return app.session_state["_test_export_frame"].copy(deep=True)

    def choose_dates(self, app, start, end):
        self.widget(app, "date_input", "Date de début").set_value(start)
        self.widget(app, "date_input", "Date de fin").set_value(end)
        self.run_app(app)

    def test_exact_days_clip_current_analysis_and_keep_full_previous_month(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            year = int(self.widget(app, "selectbox", "Choisir l'année :").value)
            self.choose_dates(app, date(year, 1, 5), date(year, 1, 10))
            summary = self.analyse(app)
            self.assertEqual(sorted(summary["Reports_Actual"]), [100, 200, 300])
            queries = list(app.session_state["_fixture_detail_queries"])
            current = [query for query in queries if query["startDate"].startswith(f"{year}-01")]
            previous = [query for query in queries if query["startDate"].startswith(f"{year - 1}-12")]
            self.assertEqual(len(current), 3)
            self.assertTrue(all(query["startDate"] == f"{year}-01-05" and query["endDate"] == f"{year}-01-10" for query in current))
            self.assertEqual(len(previous), 3)
            self.assertTrue(all(query["startDate"] == f"{year - 1}-12-01" and query["endDate"] == f"{year - 1}-12-31" for query in previous))
            self.assertEqual(len(app.session_state["_date_data_requests"]), 1)

            self.widget(app, "text_input", "Rechercher dans les violations").input("etoile")
            self.run_app(app)
            self.assertEqual(app.session_state["_fixture_detail_queries"], queries)
            self.assertEqual(len(app.session_state["_date_data_requests"]), 1)

            self.widget(app, "date_input", "Date de fin").set_value(date(year, 1, 11))
            self.run_app(app)
            self.assertNotIn("categorisation_result", app.session_state)
            self.assertTrue(app.session_state["_test_export_frame"].empty)
            self.assertEqual(app.session_state["_fixture_detail_queries"], queries)

    def test_cross_year_dates_load_matching_months_for_reporting_denominators(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            year = int(self.widget(app, "selectbox", "Choisir l'année :").value)
            self.choose_dates(app, date(year, 12, 15), date(year + 1, 1, 10))
            summary = self.analyse(app)
            self.assertEqual(sorted(summary["Reports_Actual"]), [200, 400, 600])
            data_requests = app.session_state["_date_data_requests"]
            self.assertTrue(any(
                request["period"] == f"{year}12;{year + 1}01" and request["ou_dimension"] == "LEVEL-3;CountryAAAA"
                for request in data_requests
            ))
            date_pairs = {(query["startDate"], query["endDate"]) for query in app.session_state["_fixture_detail_queries"]}
            self.assertEqual(date_pairs, {
                (f"{year}-12-15", f"{year}-12-31"),
                (f"{year + 1}-01-01", f"{year + 1}-01-10"),
                (f"{year}-11-01", f"{year}-11-30"),
            })

    def test_tab_province_scope_can_differ_from_global_zone_and_uses_its_own_metrics(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            self.widget(app, "selectbox", "Niveau d'analyse :").select("Zone de santé")
            self.run_app(app)
            self.widget(app, "selectbox", "Zone de santé :").select("ZoneAAAAAAA")
            self.run_app(app)
            self.widget(app, "selectbox", "Province de l’analyse").select("ProvinceBBB")
            self.run_app(app)
            self.widget(app, "selectbox", "Zone de santé de l’analyse").select("ZoneCCCCCCC")
            self.run_app(app)
            summary = self.analyse(app)
            self.assertEqual(self.widget(app, "selectbox", "Zone de santé :").value, "ZoneAAAAAAA")
            self.assertEqual(summary["Reports_Actual"].tolist(), [300])
            self.assertEqual({query["ou"] for query in app.session_state["_fixture_detail_queries"]}, {"ZoneCCCCCCC"})
            self.assertTrue(any(
                request["ou_dimension"] in {"LEVEL-3;ProvinceBBB", "ZoneCCCCCCC"}
                for request in app.session_state["_date_data_requests"]
            ))

    def test_validation_country_is_independent_and_cascades_to_its_own_zones(self):
        app = self.new_app(multi_country=True)
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            self.widget(app, "selectbox", "Pays de l’analyse").select("CountryBBBB")
            self.run_app(app)
            summary = self.analyse(app)
            self.assertEqual(self.widget(app, "selectbox", "Pays :").value, "CountryAAAA")
            self.assertEqual(summary["Reports_Actual"].tolist(), [400])
            self.assertEqual({query["ou"] for query in app.session_state["_fixture_detail_queries"]}, {"ZoneDDDDDDD"})
            self.assertTrue(any(request["ou_dimension"] == "LEVEL-3;CountryBBBB" for request in app.session_state["_date_data_requests"]))

    def test_returning_to_all_provinces_resets_zone_and_restores_the_country_scope(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            self.widget(app, "selectbox", "Province de l’analyse").select("ProvinceAAA")
            self.run_app(app)
            self.widget(app, "selectbox", "Zone de santé de l’analyse").select("ZoneAAAAAAA")
            self.run_app(app)
            self.assertEqual(self.analyse(app)["Reports_Actual"].tolist(), [100])

            self.widget(app, "selectbox", "Province de l’analyse").select("")
            self.run_app(app)
            self.assertEqual(self.widget(app, "selectbox", "Zone de santé de l’analyse").value, "")
            self.assertNotIn("categorisation_result", app.session_state)
            summary = self.analyse(app)
            self.assertEqual(sorted(summary["Reports_Actual"]), [100, 200, 300])
            self.assertEqual({query["ou"] for query in app.session_state["_fixture_detail_queries"]}, {"ZoneAAAAAAA", "ZoneBBBBBBB", "ZoneCCCCCCC"})

    def test_invalid_tab_dates_discard_analysis_without_removing_other_dashboard_tabs(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            self.analyse(app)
            year = int(self.widget(app, "selectbox", "Choisir l'année :").value)
            queries = list(app.session_state["_fixture_detail_queries"])
            self.choose_dates(app, date(year, 1, 20), date(year, 1, 10))
            self.assertNotIn("categorisation_result", app.session_state)
            self.assertTrue(app.session_state["_test_export_frame"].empty)
            self.assertGreater(len(app.error), 0)
            launchers = [button for button in app.button if button.label == "Lancer l'analyse des violations"]
            self.assertTrue(not launchers or all(button.disabled for button in launchers))
            self.assertTrue(any(tab.label.endswith("Complétude") for tab in app.tabs))
            self.assertTrue(any(tab.label.endswith("Promptitude") for tab in app.tabs))
            self.assertEqual(app.session_state["_fixture_detail_queries"], queries)

    def test_excel_context_records_tab_dates_selected_parent_and_rule_group(self):
        app = self.new_app()
        downloads = {}
        real_download = st.download_button

        def capture(label, *args, **kwargs):
            downloads[label] = kwargs.get("data", args[0] if args else None)
            return real_download(label, *args, **kwargs)

        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")), \
                mock.patch("streamlit.download_button", side_effect=capture):
            self.run_app(app)
            year = int(self.widget(app, "selectbox", "Choisir l'année :").value)
            self.widget(app, "selectbox", "Province de l’analyse").select("ProvinceAAA")
            self.run_app(app)
            self.widget(app, "selectbox", "Zone de santé de l’analyse").select("ZoneAAAAAAA")
            self.widget(app, "selectbox", "Sélectionner le Groupe de Règles de Validation :").select("Groupe test")
            start, end = date(year, 1, 5), date(year, 1, 10)
            self.choose_dates(app, start, end)
            self.analyse(app)
            headers, rows = detail_test.excel_sheet(downloads["Télécharger Excel"], "Contexte_Validation")
            values = {str(value) for row in rows for value in row.values()}
            for expected in ["ZoneAAAAAAA", "ZS Commune", "Groupe test"]:
                self.assertIn(expected, values)
            for day in [start, end]:
                representations = {day.isoformat(), day.strftime("%d/%m/%Y"), str(float((day - date(1899, 12, 30)).days))}
                self.assertTrue(values.intersection(representations), f"Date non retrouvée dans le contexte : {day}")
            self.assertTrue(headers)


if __name__ == "__main__":
    unittest.main()

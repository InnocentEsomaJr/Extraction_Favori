"""Les filtres pays/province/zone portent sur les véritables IDs DHIS2."""

import unittest
from unittest import mock

import pandas as pd
import requests
from streamlit.testing.v1 import AppTest

from geography import build_geography_catalog
import test_dashboard_categorisation as dashboard_test


CATALOG_FIXTURE = '''
def get_organisation_catalog(cache_user=None):
    from geography import build_geography_catalog
    country = {"id": "CountryAAAA", "displayName": "Pays test", "level": 1}
    pa = {"id": "ProvinceAAA", "displayName": "Province A", "level": 2}
    pb = {"id": "ProvinceBBB", "displayName": "Province B", "level": 2}
    units = [country, {**pa, "ancestors": [country]}, {**pb, "ancestors": [country]}]
    for uid, name, parent in [
        ("ZoneAAAAAAA", "ZS Commune", pa),
        ("ZoneBBBBBBB", "ZS B", pa),
        ("ZoneCCCCCCC", "ZS Commune", pb),
    ]:
        units.append({"id": uid, "displayName": name, "level": 3, "ancestors": [country, parent]})
    return build_geography_catalog(units)
'''


DATA_FIXTURE = '''
def get_data(favori_id, period=None, cache_user=None, ou_dimension=None):
    units = [
        ("ZoneAAAAAAA", "ZS Commune", "ProvinceAAA", 10),
        ("ZoneBBBBBBB", "ZS B", "ProvinceAAA", 20),
        ("ZoneCCCCCCC", "ZS Commune", "ProvinceBBB", 30),
    ]
    if ou_dimension == "LEVEL-3;CountryAAAA":
        selected = units
    elif ou_dimension and ou_dimension.startswith("LEVEL-3;"):
        parent_id = ou_dimension.split(";", 1)[1]
        selected = [unit for unit in units if unit[2] == parent_id]
    else:
        selected = [unit for unit in units if unit[0] == ou_dimension]
    st.session_state["_fixture_data_scope"] = ou_dimension
    st.session_state["_fixture_loaded_ids"] = [unit[0] for unit in selected]
    rows = []
    for uid, name, province, count in selected:
        row = {"Organisation unit": name, "Organisation unit ID": uid, "Organisation unit is parent": True}
        for number in range(10):
            prefix = f"Dataset {number} "
            row[prefix + "Actual reports"] = count
            row[prefix + "Expected reports"] = count
            row[prefix + "Actual reports on time"] = count
            row[prefix + "Reporting rate"] = 100
            row[prefix + "Reporting rate on time"] = 100
        rows.append(row)
    return pd.DataFrame(rows)
'''


class ScopedLoadingApiTests(unittest.TestCase):
    def test_catalog_queries_country_province_zone_levels_independently_of_favorite(self):
        country = {"id": "CountryAAAA", "displayName": "Pays test", "level": 1}
        province = {"id": "ProvinceAAA", "displayName": "Province A", "level": 2}
        units = [country, {**province, "ancestors": [country]}, {
            "id": "ZoneAAAAAAA", "displayName": "ZS A", "level": 3, "ancestors": [country, province],
        }]
        client = mock.Mock()
        client.get.return_value = {"organisationUnits": units}
        get_catalog = dashboard_test.dashboard_functions(
            "get_organisation_catalog", get_dhis2_client=lambda: client,
            build_geography_catalog=build_geography_catalog,
        )["get_organisation_catalog"]
        catalog = get_catalog(cache_user="fixture")
        self.assertEqual(catalog.loc[catalog["level"].eq(3), "Zone_ID"].tolist(), ["ZoneAAAAAAA"])
        endpoint, = client.get.call_args.args
        self.assertEqual(endpoint, "organisationUnits")
        params = client.get.call_args.kwargs["params"]
        self.assertEqual(params["filter"], "level:le:3")
        self.assertEqual(params["paging"], "false")
        self.assertIn("ancestors[id,displayName,level]", params["fields"])

    def test_failed_or_malformed_catalog_is_not_a_successful_empty_catalog(self):
        for response in [{}, {"organisationUnits": "invalid"}, RuntimeError("indisponible")]:
            with self.subTest(response=response):
                client = mock.Mock()
                if isinstance(response, Exception):
                    client.get.side_effect = response
                else:
                    client.get.return_value = response
                get_catalog = dashboard_test.dashboard_functions(
                    "get_organisation_catalog", get_dhis2_client=lambda: client,
                    build_geography_catalog=build_geography_catalog,
                )["get_organisation_catalog"]
                self.assertIsNone(get_catalog())
        client = mock.Mock()
        client.get.return_value = {"organisationUnits": []}
        get_catalog = dashboard_test.dashboard_functions(
            "get_organisation_catalog", get_dhis2_client=lambda: client,
            build_geography_catalog=build_geography_catalog,
        )["get_organisation_catalog"]
        self.assertTrue(get_catalog().empty)

    def test_scoped_favorite_load_uses_analytics_and_preserves_requested_boundary(self):
        visualization = {"id": "FavoriteAAA", "dataDimensionItems": []}
        response = {"headers": [], "rows": []}
        client = mock.Mock()
        client.get.side_effect = [visualization, response]
        params = [("dimension", "dx:DatasetAAAA.REPORTING_RATE"), ("dimension", "ou:LEVEL-3;ProvinceAAA")]
        labels = {"DatasetAAAA.REPORTING_RATE": "Dataset A Reporting rate"}
        build_query = mock.Mock(return_value=(params, labels))
        expected = pd.DataFrame({"Organisation unit": ["ZS A"], "Organisation unit ID": ["ZoneAAAAAAA"]})
        convert = mock.Mock(return_value=expected)
        get_data = dashboard_test.dashboard_functions(
            "get_data", get_dhis2_client=lambda: client, requests=requests, st=mock.Mock(),
            build_analytics_query=build_query, analytics_to_reporting_frame=convert,
        )["get_data"]
        actual = get_data("FavoriteAAA", period="202501", cache_user="fixture", ou_dimension="LEVEL-3;ProvinceAAA")
        self.assertIs(actual, expected)
        build_query.assert_called_once_with(visualization, "202501", "LEVEL-3;ProvinceAAA")
        self.assertEqual(client.get.call_count, 2)
        self.assertEqual(client.get.call_args_list[0].args, ("visualizations/FavoriteAAA",))
        fields = client.get.call_args_list[0].kwargs["params"]["fields"]
        self.assertIn("dataDimensionItems", fields)
        self.assertIn("dataSet[id,displayName,name]", fields)
        self.assertEqual(client.get.call_args_list[1], mock.call("analytics.json", params=params))
        convert.assert_called_once_with(response, labels)

    def test_scoped_load_failure_does_not_fall_back_to_provincial_favorite_rows(self):
        client = mock.Mock()
        client.get.return_value = {"id": "FavoriteAAA"}
        ui = mock.Mock()
        get_data = dashboard_test.dashboard_functions(
            "get_data", get_dhis2_client=lambda: client, requests=requests, st=ui,
            build_analytics_query=mock.Mock(side_effect=ValueError("dimension indisponible")),
            analytics_to_reporting_frame=mock.Mock(),
        )["get_data"]
        self.assertIsNone(get_data("FavoriteAAA", period="202501", ou_dimension="LEVEL-3;CountryAAAA"))
        self.assertEqual(client.get.call_count, 1)
        self.assertEqual(client.get.call_args.args, ("visualizations/FavoriteAAA",))
        ui.error.assert_called_once()


class DashboardFilterTests(unittest.TestCase):
    def widget(self, app, kind, label):
        return next(widget for widget in getattr(app, kind) if widget.label == label)

    def run_app(self, app):
        app.run(timeout=30)
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])

    def new_app(self):
        source = dashboard_test.app_fixture_source({
            "get_organisation_catalog": CATALOG_FIXTURE,
            "get_data": DATA_FIXTURE,
        })
        app = AppTest.from_string(source, default_timeout=30)
        app.session_state["connected"] = True
        # Le cache DHIS2 reste actif, mais chaque scénario possède son utilisateur.
        app.session_state["dhis2_user"] = "fixture-" + self._testMethodName
        app.session_state["dhis2_pass"] = "fixture-password"
        return app

    def analyse(self, app, expected_ids):
        self.widget(app, "button", "Lancer l'analyse des violations").click()
        self.run_app(app)
        frame = app.session_state["_test_export_frame"]
        self.assertEqual(len(frame), len(expected_ids))
        self.assertEqual(frame["Barème datasets"].tolist(), ["ZS"] * len(expected_ids))
        self.assertEqual(set(app.session_state["_fixture_validation_targets"]), set(expected_ids))
        self.assertEqual(len(app.session_state["_fixture_validation_targets"]), 2 * len(expected_ids))
        return frame

    def test_country_loads_and_validates_all_health_zones_including_homonyms(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            self.assertEqual(app.session_state["_fixture_data_scope"], "LEVEL-3;CountryAAAA")
            expected = ["ZoneAAAAAAA", "ZoneBBBBBBB", "ZoneCCCCCCC"]
            self.assertEqual(set(app.session_state["_fixture_loaded_ids"]), set(expected))
            frame = self.analyse(app, expected)
            self.assertEqual(sorted(frame["Reports_Actual"].tolist()), [100, 200, 300])

    def test_province_loads_and_validates_only_its_health_zones(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            self.widget(app, "selectbox", "Niveau d'analyse :").select("Province")
            self.run_app(app)
            self.widget(app, "selectbox", "Province :").select("ProvinceAAA")
            self.run_app(app)
            self.assertEqual(app.session_state["_fixture_data_scope"], "LEVEL-3;ProvinceAAA")
            self.assertEqual(set(app.session_state["_fixture_loaded_ids"]), {"ZoneAAAAAAA", "ZoneBBBBBBB"})
            zone_options = self.widget(app, "selectbox", "Zone de santé :").options
            self.assertFalse(any("Province B" in option for option in zone_options))
            self.analyse(app, ["ZoneAAAAAAA", "ZoneBBBBBBB"])

    def test_selected_zone_is_resolved_by_id_under_its_province(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            self.widget(app, "selectbox", "Niveau d'analyse :").select("Zone de santé")
            self.run_app(app)
            self.widget(app, "selectbox", "Province :").select("ProvinceBBB")
            self.run_app(app)
            self.widget(app, "selectbox", "Zone de santé :").select("ZoneCCCCCCC")
            self.run_app(app)
            self.assertEqual(app.session_state["_fixture_data_scope"], "ZoneCCCCCCC")
            self.assertEqual(app.session_state["_fixture_loaded_ids"], ["ZoneCCCCCCC"])
            frame = self.analyse(app, ["ZoneCCCCCCC"])
            self.assertEqual(frame["Reports_Actual"].tolist(), [300])

    def test_province_can_drill_down_to_one_of_its_zones(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            self.widget(app, "selectbox", "Niveau d'analyse :").select("Province")
            self.run_app(app)
            self.widget(app, "selectbox", "Province :").select("ProvinceAAA")
            self.run_app(app)
            self.widget(app, "selectbox", "Zone de santé :").select("ZoneBBBBBBB")
            self.run_app(app)
            self.assertEqual(app.session_state["_fixture_data_scope"], "ZoneBBBBBBB")
            self.analyse(app, ["ZoneBBBBBBB"])

    def test_changing_province_resets_selected_zone_and_invalidates_previous_analysis(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            self.widget(app, "selectbox", "Niveau d'analyse :").select("Province")
            self.run_app(app)
            self.widget(app, "selectbox", "Province :").select("ProvinceAAA")
            self.run_app(app)
            self.widget(app, "selectbox", "Zone de santé :").select("ZoneBBBBBBB")
            self.run_app(app)
            self.analyse(app, ["ZoneBBBBBBB"])

            self.widget(app, "selectbox", "Province :").select("ProvinceBBB")
            self.run_app(app)
            self.assertNotEqual(app.session_state["filter_zone_main"], "ZoneBBBBBBB")
            self.assertNotIn("ZS B", self.widget(app, "selectbox", "Zone de santé :").options)
            self.assertEqual(app.session_state["_fixture_loaded_ids"], ["ZoneCCCCCCC"])
            self.assertTrue(app.session_state["_test_export_frame"].empty)

            app.session_state["_fixture_validation_targets"] = []
            frame = self.analyse(app, ["ZoneCCCCCCC"])
            self.assertEqual(frame["Reports_Actual"].tolist(), [300])


if __name__ == "__main__":
    unittest.main()

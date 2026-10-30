"""Actualisation, période et lecture des résultats sans réseau ni secrets."""

from datetime import datetime, timedelta
import unittest
from unittest import mock

import pandas as pd
from streamlit.testing.v1 import AppTest

from categorisation import TOTAL_COL
import test_dashboard_categorisation as dashboard_test
import test_dashboard_filters as filter_test


def cached_fixture(source, counter):
    """Suit les vraies exécutions du cache, plutôt que les simples reruns UI."""
    lines = source.strip().splitlines()
    lines.insert(0, "@st.cache_data(ttl=3600)")
    lines.insert(2, f'    st.session_state["{counter}"] = st.session_state.get("{counter}", 0) + 1')
    lines.insert(3, f'    st.session_state["{counter}_key"] = cache_user')
    return "\n".join(lines) + "\n"


CACHED_CATALOG = cached_fixture(filter_test.CATALOG_FIXTURE, "_fixture_catalog_loads")
CACHED_DATA = cached_fixture(filter_test.DATA_FIXTURE, "_fixture_data_loads")
CACHED_GROUPS = cached_fixture('''
def get_validation_groups(cache_user=None):
    return [{"id": "groupAAAAAA", "displayName": "Groupe test"}]
''', "_fixture_group_loads")


class ValidationGroupApiTests(unittest.TestCase):
    def test_empty_success_is_distinct_from_a_failed_request(self):
        client = mock.Mock()
        get_groups = dashboard_test.dashboard_functions(
            "get_validation_groups", get_dhis2_client=lambda: client,
        )["get_validation_groups"]
        client.get.return_value = {"validationRuleGroups": []}
        self.assertEqual(get_groups(), [])
        client.get.side_effect = RuntimeError("fixture indisponible")
        self.assertIsNone(get_groups())


class DashboardImprovementTests(unittest.TestCase):
    def widget(self, app, kind, label):
        return next(widget for widget in getattr(app, kind) if widget.label == label)

    def run_app(self, app):
        app.run(timeout=30)
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])

    def new_app(self, replacements=None, suffix=""):
        fixtures = {
            "get_organisation_catalog": CACHED_CATALOG,
            "get_data": CACHED_DATA,
            "get_validation_groups": CACHED_GROUPS,
        }
        fixtures.update(replacements or {})
        app = AppTest.from_string(dashboard_test.app_fixture_source(fixtures), default_timeout=30)
        app.session_state["connected"] = True
        app.session_state["dhis2_user"] = "fixture-improvement-" + self._testMethodName + suffix
        app.session_state["dhis2_pass"] = "fixture-password"
        return app

    def analyse(self, app):
        self.widget(app, "button", "Lancer l'analyse des violations").click()
        self.run_app(app)
        frame = app.session_state["_test_export_frame"]
        self.assertEqual(len(frame), 3)
        self.assertEqual(frame[TOTAL_COL].tolist(), [100, 100, 100])
        return frame.copy(deep=True)

    def categorisation_view(self, app):
        frames = [table.value for table in app.dataframe if TOTAL_COL in table.value.columns]
        return frames[0] if frames else pd.DataFrame(columns=[TOTAL_COL])

    def test_refresh_changes_session_cache_key_and_discards_old_analysis(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")), \
                mock.patch("streamlit.cache_data.clear", side_effect=AssertionError("Purge globale interdite")):
            self.run_app(app)
            self.analyse(app)
            first_revision = app.session_state["data_revision"]
            first_key = app.session_state["_fixture_data_loads_key"]
            self.assertIsInstance(first_key, tuple)
            self.assertIn(first_revision, first_key)
            for name in ["_fixture_catalog_loads", "_fixture_data_loads", "_fixture_group_loads"]:
                self.assertEqual(app.session_state[name], 1)
                self.assertEqual(app.session_state[name + "_key"], first_key)

            self.run_app(app)
            self.assertEqual(app.session_state["_fixture_data_loads"], 1)
            self.widget(app, "button", "Actualiser les données").click()
            self.run_app(app)
            self.assertNotEqual(app.session_state["data_revision"], first_revision)
            self.assertNotEqual(app.session_state["_fixture_data_loads_key"], first_key)
            for name in ["_fixture_catalog_loads", "_fixture_data_loads", "_fixture_group_loads"]:
                self.assertEqual(app.session_state[name], 2)
            self.assertTrue(app.session_state["_test_export_frame"].empty)
            self.assertNotIn("categorisation_result", app.session_state)

            self.analyse(app)
            self.assertEqual(app.session_state["_fixture_validation_calls"], 4)

    def test_analysis_time_and_full_export_survive_view_search_and_sort(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            full_export = self.analyse(app)
            analysed_at = app.session_state["categorisation_result"]["analysed_at"]
            parsed = datetime.fromisoformat(analysed_at)
            self.assertEqual(parsed.utcoffset(), timedelta(hours=1))
            self.assertTrue(any("Analyse" in item.value and "Kinshasa" in item.value for item in app.caption))

            self.widget(app, "text_input", "Rechercher une zone").input("zs b")
            self.run_app(app)
            view = self.categorisation_view(app)
            self.assertEqual(view["Zone de santé"].tolist(), ["ZS B"])
            self.assertEqual(app.session_state["categorisation_result"]["analysed_at"], analysed_at)
            pd.testing.assert_frame_equal(app.session_state["_test_export_frame"], full_export)

            self.widget(app, "text_input", "Rechercher une zone").input("")
            self.widget(app, "selectbox", "Trier les résultats").select("Nom de zone")
            self.run_app(app)
            self.assertEqual(len(self.categorisation_view(app)), 3)
            pd.testing.assert_frame_equal(app.session_state["_test_export_frame"], full_export)

            self.widget(app, "selectbox", "Afficher les résultats").select("Scores incomplets")
            self.run_app(app)
            self.assertTrue(self.categorisation_view(app).empty)
            pd.testing.assert_frame_equal(app.session_state["_test_export_frame"], full_export)
            self.assertEqual(app.session_state["_fixture_validation_calls"], 2)
            self.assertEqual(app.session_state["_fixture_data_loads"], 1)
            self.assertTrue(any("export" in item.value.casefold() and "compl" in item.value.casefold() for item in app.caption))

    def test_reversed_period_blocks_loading_and_invalidates_previous_analysis(self):
        app = self.new_app()
        with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
            self.run_app(app)
            self.analyse(app)
            self.widget(app, "selectbox", "Mois début :").select("Mars")
            self.run_app(app)
            self.assertGreater(len(app.error), 0)
            self.assertNotIn("categorisation_result", app.session_state)
            self.assertEqual(app.session_state["_fixture_data_loads"], 1)
            self.assertEqual(app.session_state["_fixture_catalog_loads"], 1)
            self.assertFalse(any(button.label == "Lancer l'analyse des violations" for button in app.button))

    def test_global_analysis_remains_available_without_validation_groups(self):
        for value in ["[]", "None"]:
            with self.subTest(groups=value):
                app = self.new_app({"get_validation_groups": f'''
def get_validation_groups(cache_user=None):
    return {value}
'''}, suffix=value)
                with mock.patch("requests.sessions.Session.request", side_effect=AssertionError("Réseau interdit")):
                    self.run_app(app)
                    selector = self.widget(app, "selectbox", "Sélectionner le Groupe de Règles de Validation :")
                    self.assertEqual(selector.options, ["Toutes les règles (Global)"])
                    self.analyse(app)


if __name__ == "__main__":
    unittest.main()

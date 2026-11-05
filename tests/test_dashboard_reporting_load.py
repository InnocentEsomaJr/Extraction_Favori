"""Le dashboard utilise réellement les lots Analytics et garde les échecs absents."""

import ast
from collections import Counter
import unittest
from unittest import mock

import pandas as pd
from streamlit.testing.v1 import AppTest

from categorisation import TOTAL_COL
import test_dashboard_categorisation as dashboard_test
import test_dashboard_filters as filter_test


REPORTING_CLIENT = '''
def get_dhis2_client():
    class Client:
        def get(self, endpoint, params=None):
            if endpoint.startswith("visualizations/"):
                st.session_state["_reporting_visualization_calls"] = st.session_state.get("_reporting_visualization_calls", 0) + 1
                metrics = ["ACTUAL_REPORTS", "EXPECTED_REPORTS", "REPORTING_RATE", "ACTUAL_REPORTS_ON_TIME", "REPORTING_RATE_ON_TIME"]
                return {"dataDimensionItems": [
                    {"reportingRate": {"id": f"Data{number:07d}.{metric}", "dataSet": {
                        "id": f"Data{number:07d}", "displayName": f"Dataset {number}",
                    }}} for number in range(9) for metric in metrics
                ]}
            if endpoint != "analytics.json":
                raise AssertionError("GET fixture inattendu : " + endpoint)
            dimensions = [value for key, value in params if key == "dimension"]
            units = next(value[3:] for value in dimensions if value.startswith("ou:"))
            items = next(value[3:].split(";") for value in dimensions if value.startswith("dx:"))
            filters = [value for key, value in params if key == "filter"]
            queries = list(st.session_state.get("_reporting_queries", []))
            queries.append({"units": units, "items": items, "filters": filters})
            st.session_state["_reporting_queries"] = queries
            if units.startswith("LEVEL-") or ";" in units:
                raise AssertionError("Le dashboard ne doit pas lancer une requête Analytics globale.")
            if units in st.session_state.get("_reporting_failed_units", []):
                import requests
                raise requests.ReadTimeout("fixture-secret https://fixture.invalid/api/analytics.json?dimension=private")
            if units not in ["ZoneAAAAAAA", "ZoneBBBBBBB"]:
                raise AssertionError("Unité fixture hors périmètre : " + units)
            actual, timely = (1, 0.68) if units == "ZoneAAAAAAA" else (2, 1.6)
            values = {
                "ACTUAL_REPORTS": actual, "EXPECTED_REPORTS": actual,
                "REPORTING_RATE": 100, "ACTUAL_REPORTS_ON_TIME": timely,
                "REPORTING_RATE_ON_TIME": timely / actual * 100,
            }
            return {
                "headers": [{"name": name} for name in ["ou", "dx", "value"]],
                "rows": [[units, item, values[item.rpartition(".")[2]]] for item in items],
                "metaData": {"items": {units: {"displayName": "ZS Commune" if units == "ZoneAAAAAAA" else "ZS B"}}},
            }

        def post(self, endpoint, data=None):
            if endpoint != "dataAnalysis/validationRules":
                raise AssertionError("POST fixture inattendu : " + endpoint)
            queries = list(st.session_state.get("_reporting_validation_queries", []))
            queries.append(dict(data))
            st.session_state["_reporting_validation_queries"] = queries
            return []
    return Client()
'''


def actual_loader_functions():
    tree = ast.parse(dashboard_test.DASHBOARD.read_text(encoding='utf-8-sig'))
    names = {'get_data', 'get_quality_reporting_data'}
    return {node.name: ast.unparse(node) for node in tree.body if isinstance(node, ast.FunctionDef) and node.name in names}


class DashboardReportingLoadTests(unittest.TestCase):
    def widget(self, app, kind, label):
        return next(item for item in getattr(app, kind) if item.label == label)

    def run_app(self, app):
        app.run(timeout=45)
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])

    def new_app(self, failed=()):
        source = dashboard_test.app_fixture_source({
            **actual_loader_functions(),
            'get_organisation_catalog': filter_test.CATALOG_FIXTURE,
            'get_dhis2_client': REPORTING_CLIENT,
        })
        app = AppTest.from_string(source, default_timeout=45)
        app.session_state['connected'] = True
        app.session_state['dhis2_user'] = 'fixture-reporting-' + self._testMethodName
        app.session_state['dhis2_pass'] = 'fixture-password'
        app.session_state['filter_level_main'] = 'Province'
        app.session_state['filter_province_main'] = 'ProvinceAAA'
        app.session_state['_reporting_failed_units'] = list(failed)
        return app

    def scores(self, app):
        return app.session_state['_test_export_frame'].copy()

    def displayed_text(self, app):
        return '\n'.join(str(item.value) for kind in ['error', 'warning', 'info', 'text'] for item in getattr(app, kind))

    def assert_no_technical_exception(self, app):
        text = self.displayed_text(app)
        for forbidden in ['HTTPSConnectionPool', 'Max retries', 'dimension=private', 'fixture-secret', '/api/analytics.json?']:
            self.assertNotIn(forbidden, text)

    def test_actual_loaders_request_one_unit_and_follow_the_selected_zone(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            frame = self.scores(app).set_index('Zone de santé')
            self.assertEqual(frame.loc['ZS Commune', 'Complétude globale (%)'], 100)
            self.assertEqual(frame.loc['ZS Commune', 'Promptitude globale (%)'], 68)
            self.assertEqual(frame.loc['ZS Commune', 'Reports_Actual'], 9)
            self.assertEqual(frame.loc['ZS Commune', 'Nombre de datasets complétude ≥95 %'], 9)
            self.assertEqual(frame.loc['ZS Commune', 'Score de qualité'], 84)
            self.assertTrue(frame[TOTAL_COL].isna().all())
            queries = list(app.session_state['_reporting_queries'])
            self.assertEqual(Counter(query['units'] for query in queries), {'ZoneAAAAAAA': 1, 'ZoneBBBBBBB': 1})
            self.assertTrue(all(len(query['items']) == 45 for query in queries))
            self.assertTrue(all(len(query['filters']) == 1 and query['filters'][0].startswith('pe:') for query in queries))
            self.assertNotIn('_reporting_validation_queries', app.session_state)

            self.widget(app, 'selectbox', 'Zone de santé :').select('ZoneBBBBBBB')
            self.run_app(app)
            selected = self.scores(app)
            self.assertEqual(selected['Zone de santé'].tolist(), ['ZS B'])
            self.assertEqual(selected['Reports_Actual'].tolist(), [18])
            self.assertEqual(selected['Promptitude globale (%)'].tolist(), [80])
            self.assertEqual(selected['Score de qualité'].tolist(), [90])
            self.assertEqual(list(app.session_state['_reporting_queries']), queries)

    def test_failed_unit_is_missing_and_successful_unit_cache_survives_scope_changes(self):
        app = self.new_app(failed=['ZoneBBBBBBB'])
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            frame = self.scores(app).set_index('Zone de santé')
            self.assertEqual(frame.loc['ZS Commune', 'Reports_Actual'], 9)
            missing_columns = [
                'Complétude globale (%)', 'Promptitude globale (%)', 'Reports_Actual',
                'Nombre de datasets complétude ≥95 %', 'Score de qualité', TOTAL_COL,
            ]
            self.assertTrue(frame.loc['ZS B', missing_columns].isna().all())
            self.assertTrue(any(button.label == "Lancer l'analyse des violations" for button in app.button))
            self.assert_no_technical_exception(app)
            first_counts = Counter(query['units'] for query in app.session_state['_reporting_queries'])
            self.assertEqual(first_counts, {'ZoneAAAAAAA': 1, 'ZoneBBBBBBB': 1})

            self.widget(app, 'text_input', 'Rechercher une zone').input('commune')
            self.run_app(app)
            self.assertEqual(Counter(query['units'] for query in app.session_state['_reporting_queries']), first_counts)
            self.widget(app, 'selectbox', 'Zone de santé :').select('ZoneAAAAAAA')
            self.run_app(app)
            self.assertEqual(Counter(query['units'] for query in app.session_state['_reporting_queries']), first_counts)

            self.widget(app, 'selectbox', 'Zone de santé :').select('ZoneBBBBBBB')
            self.run_app(app)
            # Une nouvelle sélection retente la seule unité échouée ; sa réponse
            # d'erreur n'a pas été stockée dans le cache Analytics d'une heure.
            self.assertEqual(Counter(query['units'] for query in app.session_state['_reporting_queries']), {
                'ZoneAAAAAAA': 1, 'ZoneBBBBBBB': 2,
            })
            self.assertTrue(self.scores(app)[missing_columns].isna().all().all())
            self.assert_no_technical_exception(app)

    def test_total_analytics_failure_keeps_tabs_and_validation_accessible(self):
        app = self.new_app(failed=['ZoneAAAAAAA', 'ZoneBBBBBBB'])
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.assertEqual(len(app.tabs), 5)
            self.assertTrue(self.scores(app)[['Complétude globale (%)', 'Reports_Actual', 'Score de qualité']].isna().all().all())
            self.assertFalse(any('Données indisponibles.' in str(item.value) for item in app.error))
            self.assert_no_technical_exception(app)
            self.widget(app, 'button', "Lancer l'analyse des violations").click()
            self.run_app(app)
            self.assertEqual({query['ou'] for query in app.session_state['_reporting_validation_queries']}, {'ZoneAAAAAAA', 'ZoneBBBBBBB'})
            self.assertEqual(self.scores(app)['Règles violées (M)'].tolist(), [0, 0])
            self.assertTrue(self.scores(app)[['Complétude globale (%)', 'Reports_Actual', 'Score de qualité']].isna().all().all())
            self.assert_no_technical_exception(app)


if __name__ == '__main__':
    unittest.main()

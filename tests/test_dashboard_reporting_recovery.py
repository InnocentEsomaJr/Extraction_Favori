"""De nombreux datasets et un refus isolé ne masquent pas les cellules reçues."""

import ast
import unittest
from unittest import mock

import pandas as pd
from streamlit.testing.v1 import AppTest

from categorisation import TOTAL_COL
import test_dashboard_categorisation as dashboard_test


REPORTING_CLIENT = '''
def get_dhis2_client():
    class Client:
        def get(self, endpoint, params=None):
            if endpoint.startswith("visualizations/"):
                return {"dataDimensionItems": [{"reportingRate": {
                    "id": f"Data{number:07d}.{metric}", "metric": metric,
                    "dataSet": {"id": f"Data{number:07d}", "displayName": f"Dataset {number}"},
                }} for number in range(96) for metric in [
                    "ACTUAL_REPORTS", "EXPECTED_REPORTS", "REPORTING_RATE",
                    "ACTUAL_REPORTS_ON_TIME", "REPORTING_RATE_ON_TIME"]]}
            if endpoint != "analytics.json":
                raise AssertionError("GET inattendu " + endpoint)
            dimensions = dict(value.split(":", 1) for key, value in params if key == "dimension")
            items = dimensions["dx"].split(";")
            queries = list(st.session_state.get("_recovery_queries", []))
            query = {"unit": dimensions["ou"], "items": items,
                     "filters": [value for key, value in params if key == "filter"]}
            queries.append(query)
            st.session_state["_recovery_queries"] = queries
            if dimensions["ou"] != "ZoneAAAAAAA":
                raise AssertionError("Le périmètre doit rester limité à la zone sélectionnée.")
            if len(items) > 15 or any(item.startswith(st.session_state.get("_recovery_bad_dataset", "NO_BAD") + ".") for item in items):
                import requests
                response = requests.Response()
                response.status_code = 400
                raise requests.HTTPError("fixture-secret https://fixture.invalid/private", response=response)
            values = {"ACTUAL_REPORTS": 1, "EXPECTED_REPORTS": 1, "REPORTING_RATE": 100,
                      "ACTUAL_REPORTS_ON_TIME": 1, "REPORTING_RATE_ON_TIME": 100}
            return {"headers": [{"name": name} for name in ["ou", "dx", "value"]],
                    "rows": [[dimensions["ou"], item, values[item.rpartition(".")[2]]] for item in items],
                    "metaData": {"items": {"ZoneAAAAAAA": {"displayName": "ZS A"}}}}

        def post(self, endpoint, data=None):
            if endpoint != "dataAnalysis/validationRules":
                raise AssertionError("POST inattendu " + endpoint)
            calls = list(st.session_state.get("_recovery_validation_posts", []))
            calls.append(dict(data))
            st.session_state["_recovery_validation_posts"] = calls
            return []
    return Client()
'''


class ReportingRecoveryAppTests(unittest.TestCase):
    def widget(self, app, kind, label):
        return next(item for item in getattr(app, kind) if item.label == label)

    def run_app(self, app):
        app.run(timeout=75)
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])

    def new_app(self, bad_dataset=None):
        tree = ast.parse(dashboard_test.DASHBOARD.read_text(encoding='utf-8-sig'))
        boundaries = {
            node.name: ast.unparse(node) for node in tree.body
            if isinstance(node, ast.FunctionDef) and node.name in {'get_data', 'get_quality_reporting_data'}
        }
        source = dashboard_test.app_fixture_source({**boundaries, 'get_dhis2_client': REPORTING_CLIENT})
        source += '\nst.session_state["_recovery_raw"] = target_df.copy()\n'
        source += 'st.session_state["_recovery_completeness"] = df_final_c.copy()\n'
        source += 'st.session_state["_recovery_timeliness"] = df_final_p.copy()\n'
        app = AppTest.from_string(source, default_timeout=75)
        app.session_state['connected'] = True
        app.session_state['dhis2_user'] = 'fixture-recovery-' + self._testMethodName
        app.session_state['dhis2_pass'] = 'fixture-password'
        if bad_dataset:
            app.session_state['_recovery_bad_dataset'] = bad_dataset
        return app

    def assert_no_catalogue_regression(self, app):
        self.assertFalse(any(widget.label == 'Datasets évalués pour le score qualité' for widget in app.multiselect))
        displayed = '\n'.join(str(item.value) for kind in ['warning', 'error', 'info', 'text'] for item in getattr(app, kind))
        self.assertNotIn('1y-Surveillance', displayed)
        self.assertNotIn('fixture-secret', displayed)
        self.assertNotIn('fixture.invalid/private', displayed)

    def test_ninety_six_datasets_load_values_with_bounded_queries_and_cached_validation_actions(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            result = app.session_state['_test_export_frame']
            self.assertEqual(result['Reports_Actual'].tolist(), [96])
            self.assertEqual(result['Complétude globale (%)'].tolist(), [100])
            self.assertEqual(result['Promptitude globale (%)'].tolist(), [100])
            self.assertEqual(result['Nombre de datasets complétude ≥95 %'].tolist(), [96])
            raw = app.session_state['_recovery_raw']
            self.assertEqual(raw['Dataset 0 Actual reports'].tolist(), [1])
            self.assertEqual(raw['Dataset 95 Actual reports'].tolist(), [1])
            queries = list(app.session_state['_recovery_queries'])
            self.assertEqual(len(queries), 32)
            self.assertTrue(all(len(query['items']) <= 15 for query in queries))
            self.assertEqual(len({item for query in queries for item in query['items']}), 480)
            self.assert_no_catalogue_regression(app)
            self.widget(app, 'button', 'Calculer tous les scores (M et M-1)').click()
            self.run_app(app)
            self.assertEqual(app.session_state['_test_export_frame'][TOTAL_COL].tolist(), [100])
            self.assertEqual(len(app.session_state['_recovery_validation_posts']), 2)
            self.assertEqual(list(app.session_state['_recovery_queries']), queries)
            self.widget(app, 'button', "Lancer l'analyse des violations").click()
            self.run_app(app)
            self.assertEqual(len(app.session_state['_recovery_validation_posts']), 3)
            self.assertEqual(list(app.session_state['_recovery_queries']), queries)

    def test_one_rejected_dataset_keeps_other_cells_visible_and_validation_independent(self):
        app = self.new_app('Data0000003')
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            raw = app.session_state['_recovery_raw']
            self.assertEqual(raw['Dataset 0 Actual reports'].tolist(), [1])
            self.assertEqual(raw['Dataset 95 Actual reports'].tolist(), [1])
            self.assertTrue(raw['Dataset 3 Actual reports'].isna().all())
            completeness = app.session_state['_recovery_completeness']
            self.assertEqual(completeness['Dataset 0 Reporting rate'].tolist(), [100])
            self.assertTrue(completeness['Dataset 3 Reporting rate'].isna().all())
            timeliness = app.session_state['_recovery_timeliness']
            self.assertEqual(timeliness['Dataset 95 Reporting rate on time'].tolist(), [100])
            self.assertTrue(timeliness['Dataset 3 Reporting rate on time'].isna().all())
            result = app.session_state['_test_export_frame']
            self.assertTrue(result[['Reports_Actual', 'Complétude globale (%)', TOTAL_COL]].isna().all().all())
            self.assert_no_catalogue_regression(app)
            top_warnings = '\n'.join(str(item.value) for item in app.warning)
            self.assertNotIn('HTTP 400', top_warnings)
            queries = list(app.session_state['_recovery_queries'])
            self.widget(app, 'button', "Lancer l'analyse des violations").click()
            self.run_app(app)
            self.assertEqual(len(app.session_state['_recovery_validation_posts']), 1)
            self.assertEqual(app.session_state['_test_export_frame']['Règles violées (M)'].tolist(), [0])
            self.assertTrue(app.session_state['_test_export_frame'][TOTAL_COL].isna().all())
            # Les erreurs ne sont pas mises en cache une heure. Si le rendu des
            # 480 colonnes dépasse les 60 s du cache de synthèse, seul un lot
            # contenant le dataset refusé peut être retenté ; les lots sains
            # restent en cache, indépendamment de l'action de validation.
            retries = list(app.session_state['_recovery_queries'])[len(queries):]
            self.assertTrue(all(
                any(item.startswith('Data0000003.') for item in query['items'])
                for query in retries
            ))


if __name__ == '__main__':
    unittest.main()

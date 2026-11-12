"""La sélection qualité pilote réellement les requêtes et le score exporté."""

import ast
import unittest
from unittest import mock

from streamlit.testing.v1 import AppTest

from categorisation import TOTAL_COL
import test_dashboard_categorisation as dashboard_test


DATASET_CLIENT = '''
def get_dhis2_client():
    class Client:
        def get(self, endpoint, params=None):
            catalog = [{"id": f"Data{n:07d}", "displayName": f"Service {n}"} for n in range(9)]
            catalog += [{"id": "SurveillAAA", "displayName": "1y-Surveillance épidémiologique"},
                        {"id": "TertiaryAAA", "displayName": "Service tertiaire"}]
            if endpoint == "dataSets":
                return {"dataSets": catalog}
            if endpoint.startswith("visualizations/"):
                favorite = catalog[:9] + [catalog[-1]]
                return {"dataDimensionItems": [{"reportingRate": {
                    "id": item["id"] + "." + metric, "dataSet": item,
                }} for item in favorite for metric in ["ACTUAL_REPORTS", "EXPECTED_REPORTS",
                    "ACTUAL_REPORTS_ON_TIME", "REPORTING_RATE", "REPORTING_RATE_ON_TIME"]]}
            if endpoint != "analytics.json":
                raise AssertionError("Métadonnées inattendues " + endpoint)
            dimensions = dict((value.split(":", 1)) for key, value in params if key == "dimension")
            items = dimensions["dx"].split(";")
            queries = list(st.session_state.get("_dataset_queries", []))
            queries.append({"unit": dimensions["ou"], "items": items,
                            "filters": [v for k, v in params if k == "filter"]})
            st.session_state["_dataset_queries"] = queries
            values = []
            for item in items:
                dataset, metric = item.rsplit(".", 1)
                actual, expected = (999, 999) if dataset == "TertiaryAAA" else (
                    (16, 16) if dataset == "SurveillAAA" else (
                        (12, 10) if dataset == "Data0000000" else (
                            (8, 10) if dataset == "Data0000001" else (14, 14))))
                number = {"ACTUAL_REPORTS": actual, "EXPECTED_REPORTS": expected,
                          "ACTUAL_REPORTS_ON_TIME": actual, "REPORTING_RATE": actual / expected * 100,
                          "REPORTING_RATE_ON_TIME": actual / expected * 100}[metric]
                values.append([dimensions["ou"], item, number])
            return {"headers": [{"name": name} for name in ["ou", "dx", "value"]], "rows": values,
                    "metaData": {"items": {"ZoneAAAAAAA": {"displayName": "ZS A"}}}}

        def post(self, endpoint, data=None):
            assert endpoint == "dataAnalysis/validationRules"
            current = data["startDate"] == st.session_state["validation_start_date"].isoformat()
            st.session_state["_dataset_posts"] = st.session_state.get("_dataset_posts", 0) + 1
            return [{"organisationUnitId": "ZoneAAAAAAA", "organisationUnitDisplayName": "ZS A",
                     "validationRuleId": f"Rule{'Now' if current else 'Old'}{n:04d}",
                     "periodId": data["startDate"][:7].replace("-", ""),
                     "leftSideValue": 2, "rightSideValue": 1} for n in range(5 if current else 2)]
    return Client()
'''


class DashboardQualityDatasetsTests(unittest.TestCase):
    def widget(self, app, kind, label):
        return next(item for item in getattr(app, kind) if item.label == label)

    def run_app(self, app):
        app.run(timeout=60)
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])

    def test_surveillance_replaces_tertiary_in_quality_only_and_recalculates_on_selection(self):
        tree = ast.parse(dashboard_test.DASHBOARD.read_text(encoding='utf-8-sig'))
        names = {'get_data', 'get_quality_reporting_data', 'get_quality_dataset_options'}
        boundaries = {
            node.name: ast.unparse(node) for node in tree.body
            if isinstance(node, ast.FunctionDef) and node.name in names
        }
        source = dashboard_test.app_fixture_source({**boundaries,
            'get_dhis2_client': DATASET_CLIENT,
            'get_validation_detail_metadata': '''
def get_validation_detail_metadata(rule_ids, org_ids, cache_user=None):
    return {}, {}, []
''',
        })
        app = AppTest.from_string(source, default_timeout=60)
        app.session_state['connected'] = True
        app.session_state['dhis2_user'] = 'fixture-datasets-' + self._testMethodName
        app.session_state['dhis2_pass'] = 'fixture-password'
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            selector = self.widget(app, 'multiselect', 'Datasets évalués pour le score qualité')
            selected = app.session_state['quality_evaluation_datasets']
            self.assertEqual(len(selected), 10)
            self.assertIn('SurveillAAA', selected)
            self.assertNotIn('TertiaryAAA', selected)
            self.assertNotIn('Service tertiaire', selector.options)
            quality = app.session_state['_test_export_frame']
            self.assertEqual(quality['Reports_Actual'].tolist(), [134])
            self.assertEqual(quality['Nombre de datasets complétude ≥95 %'].tolist(), [9])
            self.assertEqual(quality['Complétude globale (%)'].tolist(), [100])
            self.assertEqual(quality['Promptitude globale (%)'].tolist(), [100])
            queries = app.session_state['_dataset_queries']
            self.assertEqual(len(queries), 2)
            self.assertTrue(any('TertiaryAAA.' in item for item in queries[0]['items']))
            self.assertEqual(len(queries[1]['items']), 50)
            self.assertTrue(any('SurveillAAA.' in item for item in queries[1]['items']))
            self.assertFalse(any('TertiaryAAA.' in item for item in queries[1]['items']))
            self.widget(app, 'button', 'Calculer tous les scores (M et M-1)').click()
            self.run_app(app)
            quality = app.session_state['_test_export_frame']
            self.assertEqual(quality[TOTAL_COL].tolist(), [75])
            self.assertEqual(quality['Score corrections /20'].tolist(), [0])
            self.assertEqual(app.session_state['_dataset_posts'], 2)
            self.widget(app, 'multiselect', 'Datasets évalués pour le score qualité').unselect('SurveillAAA')
            self.run_app(app)
            quality = app.session_state['_test_export_frame']
            self.assertEqual(quality['Reports_Actual'].tolist(), [118])
            self.assertEqual(quality['Nombre de datasets complétude ≥95 %'].tolist(), [8])
            self.assertEqual(quality['Règles violées (M)'].tolist(), [5])
            self.assertEqual(app.session_state['_dataset_posts'], 2)
            last = app.session_state['_dataset_queries'][-1]
            self.assertFalse(any('SurveillAAA.' in item or 'TertiaryAAA.' in item for item in last['items']))
            self.assertIn('pe:', last['filters'][0])


if __name__ == '__main__':
    unittest.main()

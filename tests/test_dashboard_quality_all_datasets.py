"""Le score qualité suit les datasets du favori applicables à chaque zone."""

import unittest
from unittest import mock

import pandas as pd
from streamlit.testing.v1 import AppTest

from categorisation import TOTAL_COL
from scoped_analytics import METRIC_LABELS
import test_dashboard_categorisation as dashboard_test
import test_dashboard_filters as filter_test
from test_dashboard_reporting_load import actual_loader_functions


REPORTING_CLIENT = '''
def get_dhis2_client():
    class Client:
        def get(self, endpoint, params=None):
            if endpoint.startswith("visualizations/"):
                metrics = ["ACTUAL_REPORTS", "EXPECTED_REPORTS", "REPORTING_RATE",
                           "ACTUAL_REPORTS_ON_TIME", "REPORTING_RATE_ON_TIME"]
                return {"dataDimensionItems": [
                    {"reportingRate": {
                        "id": f"Data{number:07d}.{metric}", "metric": metric,
                        "dataSet": {"id": f"Data{number:07d}", "displayName": f"Dataset {number}"},
                    }} for number in range(12) for metric in metrics
                ] + [
                    {"indicator": {"id": "IndicatorAA", "displayName": "Autre Actual reports"}},
                    {"indicator": {"id": "IndicatorBB", "displayName": "Indicateur refuse"}},
                ]}
            if endpoint != "analytics.json":
                raise AssertionError("GET fixture inattendu : " + endpoint)
            dimensions = dict(value.split(":", 1) for key, value in params if key == "dimension")
            unit = dimensions["ou"]
            items = dimensions["dx"].split(";")
            if unit not in ["ZoneAAAAAAA", "ZoneBBBBBBB"]:
                raise AssertionError("Le chargement doit rester limite aux zones selectionnees.")
            queries = list(st.session_state.get("_all_datasets_queries", []))
            queries.append({"unit": unit, "items": items,
                            "filters": [value for key, value in params if key == "filter"]})
            st.session_state["_all_datasets_queries"] = queries
            if "IndicatorBB" in items and st.session_state.get("_all_datasets_reject_indicator", True):
                import requests
                response = requests.Response()
                response.status_code = 400
                raise requests.HTTPError("Indicateur generique refuse.", response=response)
            active_count, actual, timely = (9, 10, 8) if unit == "ZoneAAAAAAA" else (6, 5, 4)
            values = {"ACTUAL_REPORTS": actual, "EXPECTED_REPORTS": 10,
                      "REPORTING_RATE": actual * 10, "ACTUAL_REPORTS_ON_TIME": timely,
                      "REPORTING_RATE_ON_TIME": timely * 10}
            rows = []
            for item in items:
                if item == "IndicatorAA":
                    rows.append([unit, item, 999])
                    continue
                if item == "IndicatorBB":
                    rows.append([unit, item, 1])
                    continue
                dataset, metric = item.rsplit(".", 1)
                if int(dataset[4:]) < active_count:
                    rows.append([unit, item, values[metric]])
            return {"headers": [{"name": name} for name in ["ou", "dx", "value"]],
                    "rows": rows,
                    "metaData": {"items": {unit: {
                        "displayName": "ZS Commune" if unit == "ZoneAAAAAAA" else "ZS B",
                    }}}}

        def post(self, endpoint, data=None):
            calls = list(st.session_state.get("_all_datasets_posts", []))
            calls.append({"endpoint": endpoint, "data": data})
            st.session_state["_all_datasets_posts"] = calls
            raise AssertionError("Aucune analyse des violations ne doit etre lancee automatiquement.")
    return Client()
'''


class AllFavoriteDatasetsQualityTests(unittest.TestCase):
    quality_columns = [
        'Reports_Actual', 'Complétude globale (%)', 'Promptitude globale (%)',
        'Nombre de datasets complétude ≥95 %', 'Score de qualité',
    ]

    def new_app(self, reject_indicator=True):
        source = dashboard_test.app_fixture_source({
            **actual_loader_functions(),
            'get_organisation_catalog': filter_test.CATALOG_FIXTURE,
            'get_dhis2_client': REPORTING_CLIENT,
        })
        source += '\nst.session_state["_all_datasets_raw"] = target_df.copy()\n'
        app = AppTest.from_string(source, default_timeout=45)
        app.session_state['connected'] = True
        app.session_state['dhis2_user'] = 'fixture-all-datasets-' + self._testMethodName
        app.session_state['dhis2_pass'] = 'fixture-password'
        app.session_state['_all_datasets_reject_indicator'] = reject_indicator
        app.session_state['filter_level_main'] = 'Province'
        app.session_state['filter_province_main'] = 'ProvinceAAA'
        return app

    def run_app(self, app):
        app.run(timeout=45)
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])

    def assert_reporting_export(self, app, expected):
        # Ce tableau est celui transmis à l'export Excel par le dashboard.
        frame = app.session_state['_test_export_frame'].set_index('Zone de santé')
        self.assertCountEqual(frame.index.tolist(), expected)
        for name, values in expected.items():
            self.assertEqual(frame.loc[name, self.quality_columns].tolist(), values)
        for column in self.quality_columns:
            self.assertTrue(pd.api.types.is_float_dtype(frame[column]), (column, frame[column].dtype))
        self.assertTrue(frame[['Règles violées (M)', TOTAL_COL]].isna().all().all())
        self.assertNotIn('_all_datasets_posts', app.session_state)
        self.assertFalse(any(
            item.label == 'Datasets évalués pour le score qualité' for item in app.multiselect
        ))
        warnings = '\n'.join(str(item.value) for item in app.warning)
        self.assertNotIn('Certains indicateurs de déclaration sont indisponibles.', warnings)

    def test_all_datasets_use_per_zone_reporting_and_ignore_generic_indicator_failures(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.assert_reporting_export(app, {
                'ZS Commune': [90.0, 100.0, 80.0, 9.0, 90.0],
                'ZS B': [30.0, 50.0, 40.0, 0.0, 45.0],
            })
            raw = app.session_state['_all_datasets_raw'].set_index('Organisation unit ID')
            self.assertEqual(raw['Autre Actual reports'].tolist(), [999, 999])
            self.assertTrue(raw['Indicateur refuse'].isna().all())
            self.assertTrue(raw['Dataset 11 Actual reports'].isna().all())
            self.assertEqual(raw.loc['ZoneAAAAAAA', 'Dataset 8 Actual reports'], 10)
            self.assertTrue(pd.isna(raw.loc['ZoneBBBBBBB', 'Dataset 8 Actual reports']))
            self.assertEqual(raw.attrs['failed_metrics_by_unit'], {
                'ZoneAAAAAAA': ['IndicatorBB'], 'ZoneBBBBBBB': ['IndicatorBB'],
            })

            queries = list(app.session_state['_all_datasets_queries'])
            reporting_items = {f'Data{number:07d}.{metric}' for number in range(12) for metric in METRIC_LABELS}
            self.assertTrue(all(len(query['items']) <= 15 for query in queries))
            for unit in ['ZoneAAAAAAA', 'ZoneBBBBBBB']:
                requested = {item for query in queries if query['unit'] == unit for item in query['items']}
                self.assertTrue(reporting_items.issubset(requested))
            self.assertTrue(all(query['filters'] == ['pe:202501'] for query in queries))

            next(item for item in app.selectbox if item.label == 'Zone de santé :').select('ZoneBBBBBBB')
            self.run_app(app)
            self.assert_reporting_export(app, {'ZS B': [30.0, 50.0, 40.0, 0.0, 45.0]})
            # Seul l'indicateur générique refusé peut être retenté si le cache
            # court expire ; les réponses de reporting réussies sont réutilisées.
            retries = list(app.session_state['_all_datasets_queries'])[len(queries):]
            self.assertTrue(all('IndicatorBB' in query['items'] for query in retries))

    def test_successful_generic_count_cannot_seed_quality_from_the_displayed_summary(self):
        app = self.new_app(reject_indicator=False)
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.assert_reporting_export(app, {
                'ZS Commune': [90.0, 100.0, 80.0, 9.0, 90.0],
                'ZS B': [30.0, 50.0, 40.0, 0.0, 45.0],
            })
            raw = app.session_state['_all_datasets_raw']
            self.assertEqual(raw['Autre Actual reports'].tolist(), [999, 999])
            self.assertEqual(raw['Indicateur refuse'].tolist(), [1, 1])
            self.assertEqual(raw.attrs['failed_units'], [])


if __name__ == '__main__':
    unittest.main()

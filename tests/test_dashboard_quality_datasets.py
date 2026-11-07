"""La sélection qualité pilote réellement les requêtes et le score exporté."""

import ast
from datetime import date
import unittest
from unittest import mock

import streamlit as st
from streamlit.testing.v1 import AppTest

from categorisation import TOTAL_COL
import test_dashboard_categorisation as dashboard_test
from test_dashboard_validation_details import excel_sheet


DATASET_CLIENT = '''
def get_dhis2_client():
    class Client:
        def get(self, endpoint, params=None):
            catalog = [{"id": f"Data{n:07d}", "displayName": f"Service {n}", "periodType": "Monthly"} for n in range(9)]
            catalog += [{"id": "SurveillAAA", "displayName": "eSurveillance Hebdomadaire-Edition 2025", "periodType": "Weekly"},
                        {"id": "TertiaryAAA", "displayName": "Service tertiaire", "periodType": "Monthly"}]
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
            periods = next(value[3:].split(";") for key, value in params if key == "filter" and value.startswith("pe:"))
            queries = list(st.session_state.get("_dataset_queries", []))
            queries.append({"unit": dimensions["ou"], "items": items,
                            "filters": [v for k, v in params if k == "filter"]})
            st.session_state["_dataset_queries"] = queries
            if st.session_state.get("_dataset_fail_weekly") and any(item.startswith("SurveillAAA.") for item in items):
                import requests
                raise requests.ReadTimeout("Échec de surveillance hebdomadaire simulé")
            values = []
            for item in items:
                dataset, metric = item.rsplit(".", 1)
                actual, expected = (999, 999) if dataset == "TertiaryAAA" else (
                    (17 * len(periods), 17 * len(periods)) if dataset == "SurveillAAA" else (
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
        downloads = {}
        real_download = st.download_button

        def capture(label, *args, **kwargs):
            downloads[label] = kwargs.get('data', args[0] if args else None)
            return real_download(label, *args, **kwargs)

        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')), \
                mock.patch('streamlit.download_button', side_effect=capture):
            self.run_app(app)
            self.widget(app, 'selectbox', "Choisir l'année :").select('2026')
            self.widget(app, 'selectbox', 'Mois début :').select('Septembre')
            self.widget(app, 'selectbox', 'Mois fin :').select('Septembre')
            app.session_state['_dataset_queries'] = []
            self.run_app(app)
            selector = self.widget(app, 'multiselect', 'Datasets évalués pour le score qualité')
            selected = app.session_state['quality_evaluation_datasets']
            self.assertEqual(len(selected), 10)
            self.assertIn('SurveillAAA', selected)
            self.assertNotIn('TertiaryAAA', selected)
            self.assertNotIn('Service tertiaire', selector.options)
            quality = app.session_state['_test_export_frame']
            self.assertEqual(quality['Reports_Actual'].tolist(), [203])
            self.assertEqual(quality['Nombre de datasets complétude ≥95 %'].tolist(), [9])
            self.assertEqual(quality['Complétude globale (%)'].tolist(), [100])
            self.assertEqual(quality['Promptitude globale (%)'].tolist(), [100])
            queries = app.session_state['_dataset_queries']
            self.assertEqual(len(queries), 3)
            self.assertTrue(any('TertiaryAAA.' in item for item in queries[0]['items']))
            self.assertEqual(len(queries[1]['items']), 45)
            self.assertEqual(queries[1]['filters'], ['pe:202609'])
            self.assertFalse(any('SurveillAAA.' in item for item in queries[1]['items']))
            self.assertFalse(any('TertiaryAAA.' in item for item in queries[1]['items']))
            self.assertEqual(len(queries[2]['items']), 5)
            self.assertTrue(all('SurveillAAA.' in item for item in queries[2]['items']))
            self.assertEqual(queries[2]['filters'], ['pe:2026W35;2026W36;2026W37;2026W38;2026W39'])
            self.assertTrue(all(query['unit'] == 'ZoneAAAAAAA' for query in queries))
            _, periods = excel_sheet(downloads['Télécharger Excel'], 'Periodes_Surveillance')
            self.assertEqual([row['Semaine'] for row in periods], [f'2026W{week}' for week in range(35, 40)])
            self.assertEqual(periods[0], {
                'Mois': '202609', 'Semaine': '2026W35', 'Début': '2026-08-24', 'Fin': '2026-08-30', 'Reportée': 1,
            })
            self.assertEqual([row['Reportée'] for row in periods], [1, 0, 0, 0, 0])
            headers, weekly = excel_sheet(downloads['Télécharger Excel'], 'Reporting_Surveillance')
            weekly_actual = next(column for column in headers if column.endswith('Actual reports'))
            weekly_expected = next(column for column in headers if column.endswith('Expected reports'))
            self.assertEqual(weekly[0][weekly_actual], 85)
            self.assertEqual(weekly[0][weekly_expected], 85)
            self.assertEqual(weekly[0]['Organisation unit ID'], 'ZoneAAAAAAA')
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
            self.assertTrue(any(
                len(query['items']) == 45 and not any('SurveillAAA.' in item or 'TertiaryAAA.' in item
                                                     for item in query['items'])
                for query in app.session_state['_dataset_queries']
            ))
            self.widget(app, 'multiselect', 'Datasets évalués pour le score qualité').select('SurveillAAA')
            self.widget(app, 'selectbox', 'Mois début :').select('Novembre')
            self.widget(app, 'selectbox', 'Mois fin :').select('Novembre')
            self.run_app(app)
            self.assertEqual(app.session_state['_test_export_frame']['Reports_Actual'].tolist(), [186])
            _, november = excel_sheet(downloads['Télécharger Excel'], 'Periodes_Surveillance')
            self.assertEqual([row['Semaine'] for row in november], [f'2026W{week}' for week in range(44, 48)])
            self.assertEqual([row['Reportée'] for row in november], [0, 0, 0, 0])
            self.widget(app, 'selectbox', 'Mois début :').select('Décembre')
            self.widget(app, 'selectbox', 'Mois fin :').select('Décembre')
            self.run_app(app)
            self.assertEqual(app.session_state['_test_export_frame']['Reports_Actual'].tolist(), [203])
            _, december = excel_sheet(downloads['Télécharger Excel'], 'Periodes_Surveillance')
            self.assertEqual([row['Semaine'] for row in december], [f'2026W{week}' for week in range(48, 53)])
            self.assertEqual(december[0], {
                'Mois': '202612', 'Semaine': '2026W48', 'Début': '2026-11-23', 'Fin': '2026-11-29', 'Reportée': 1,
            })
            self.assertEqual([row['Reportée'] for row in december], [1, 0, 0, 0, 0])
            self.assertEqual(app.session_state['_dataset_posts'], 2)
            self.widget(app, 'button', "Lancer l'analyse des violations").click()
            self.run_app(app)
            self.assertIn('categorisation_result', app.session_state)
            self.assertEqual(app.session_state['_dataset_posts'], 3)
            app.session_state['_dataset_fail_weekly'] = True
            self.widget(app, 'date_input', 'Date de début').set_value(date(2026, 10, 1))
            self.widget(app, 'date_input', 'Date de fin').set_value(date(2026, 10, 31))
            self.run_app(app)
            failed = app.session_state['_test_export_frame']
            self.assertTrue(failed[[
                'Complétude globale (%)', 'Promptitude globale (%)', 'Reports_Actual',
                'Nombre de datasets complétude ≥95 %', TOTAL_COL,
            ]].isna().all().all())
            self.assertTrue(failed['Règles violées (M)'].isna().all())
            self.assertNotIn('categorisation_result', app.session_state)
            self.assertEqual(app.session_state['_dataset_posts'], 3)
            warnings = '\n'.join(str(item.value) for kind in ['warning', 'text'] for item in getattr(app, kind))
            self.assertIn('weekly', warnings)
            self.assertIn('Analytics', warnings)


if __name__ == '__main__':
    unittest.main()

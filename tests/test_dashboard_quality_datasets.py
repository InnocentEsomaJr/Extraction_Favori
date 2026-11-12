"""Périodes et semaines explicites partagées par tous les tableaux du dashboard."""

import ast
from datetime import date
import unittest
from unittest import mock

import pandas as pd
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
                if dataset != "SurveillAAA":
                    actual, expected = actual * len(periods), expected * len(periods)
                if dataset == "Data0000000" and st.session_state.get("_dataset_zero_monthly"):
                    actual = 0
                number = {"ACTUAL_REPORTS": actual, "EXPECTED_REPORTS": expected,
                          "ACTUAL_REPORTS_ON_TIME": actual, "REPORTING_RATE": actual / expected * 100,
                          "REPORTING_RATE_ON_TIME": actual / expected * 100}[metric]
                values.append([dimensions["ou"], item, number])
            return {"headers": [{"name": name} for name in ["ou", "dx", "value"]], "rows": values,
                    "metaData": {"items": {"ZoneAAAAAAA": {"displayName": "ZS A"}}}}

        def post(self, endpoint, data=None):
            assert endpoint == "dataAnalysis/validationRules"
            requests = list(st.session_state.get("_dataset_post_queries", []))
            requests.append(dict(data))
            st.session_state["_dataset_post_queries"] = requests
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

    def choice(self, app, label):
        return next(item for kind in ('selectbox', 'radio') for item in getattr(app, kind) if item.label == label)

    def run_app(self, app):
        app.run(timeout=60)
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])

    def new_app(self, replacements=None):
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
            **(replacements or {}),
        })
        app = AppTest.from_string(source, default_timeout=60)
        app.session_state['connected'] = True
        app.session_state['dhis2_user'] = 'fixture-datasets-' + self._testMethodName
        app.session_state['dhis2_pass'] = 'fixture-password'
        return app

    def available_view(self, app):
        return next(item.value for item in app.dataframe if 'Rapports mensuels disponibles' in item.value.columns)

    def assert_partial_monthly_visible(self, app, monthly=118):
        raw = app.session_state['_test_main_raw']
        monthly_columns = raw.attrs['monthly_reporting_columns']
        actual_columns = [column for column in monthly_columns if column.endswith(' Actual reports')]
        self.assertEqual(raw[actual_columns].sum(axis=1, min_count=1).tolist(), [monthly])
        for key, suffix in [('_test_main_comp', ' Reporting rate'), ('_test_main_prompt', ' Reporting rate on time')]:
            frame = app.session_state[key]
            columns = [column for column in monthly_columns if column.endswith(suffix)]
            self.assertTrue(columns)
            self.assertFalse(frame[columns].isna().any().any())
            self.assertIn(80, frame[columns].iloc[0].tolist())
        weekly_columns = raw.attrs.get('weekly_reporting_columns', [])
        if weekly_columns:
            self.assertTrue(raw[weekly_columns].isna().all().all())
        for key, columns in [
            ('_test_main_summary', ['Reports_Actual', 'Complétude_Globale (%)', 'Promptitude_Globale (%)']),
            ('_test_main_comp', ['Reports_Actual', 'Complétude_Globale (%)']),
            ('_test_main_prompt', ['Reports_Actual_On_Time', 'Promptitude_Globale (%)']),
            ('_test_export_frame', ['Reports_Actual', 'Complétude globale (%)', 'Promptitude globale (%)', TOTAL_COL]),
        ]:
            self.assertTrue(app.session_state[key][columns].isna().all().all())
        available = self.available_view(app)
        self.assertEqual(pd.to_numeric(available['Rapports mensuels disponibles']).tolist(), [monthly])
        self.assertTrue(pd.to_numeric(available['Rapports hebdomadaires disponibles'], errors='coerce').isna().all())
        self.assertEqual(available['État des sources'].tolist(), ['Incomplet'])
        text = '\n'.join(str(item.value) for kind in ['markdown', 'warning', 'text'] for item in getattr(app, kind))
        self.assertIn('Rapports disponibles par source', text)
        self.assertFalse(any('tertiaire' in str(column).casefold() for column in raw.columns))

    def select_month_range(self, app, first, last, first_year='2026', last_year='2026'):
        self.widget(app, 'selectbox', "Choisir l'année :").select(first_year)
        self.widget(app, 'selectbox', 'Année de fin').select(last_year)
        self.widget(app, 'selectbox', 'Mois début :').select(first)
        self.widget(app, 'selectbox', 'Mois fin :').select(last)
        self.run_app(app)

    def assert_shared_reporting(self, app, actual, datasets):
        for key in ['_test_main_summary', '_test_main_comp']:
            frame = app.session_state[key]
            self.assertEqual(frame['Reports_Actual'].tolist(), [actual])
            self.assertEqual(frame['Reports_Attendu'].tolist(), [actual])
            self.assertEqual(frame['Complétude_Globale (%)'].tolist(), [100])
        prompt = app.session_state['_test_main_prompt']
        self.assertEqual(prompt['Reports_Actual_On_Time'].tolist(), [actual])
        self.assertEqual(prompt['Promptitude_Globale (%)'].tolist(), [100])
        self.assertEqual(app.session_state['_test_main_comp']['Nombre des data set complétude >/=95%'].tolist(), [datasets])
        quality = app.session_state['_test_export_frame']
        self.assertEqual(quality['Reports_Actual'].tolist(), [actual])
        self.assertEqual(quality['Nombre de datasets complétude ≥95 %'].tolist(), [datasets])
        self.assertEqual(quality['Complétude globale (%)'].tolist(), [100])
        self.assertEqual(quality['Promptitude globale (%)'].tolist(), [100])

    def test_selected_weeks_are_aggregated_in_every_table_and_export(self):
        app = self.new_app()
        downloads = {}
        real_download = st.download_button

        def capture(label, *args, **kwargs):
            downloads[label] = kwargs.get('data', args[0] if args else None)
            return real_download(label, *args, **kwargs)

        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')), \
                mock.patch('streamlit.download_button', side_effect=capture):
            self.run_app(app)
            app.session_state['_dataset_queries'] = []
            self.select_month_range(app, 'Septembre', 'Septembre')
            selected = app.session_state['quality_evaluation_datasets']
            self.assertEqual(len(selected), 10)
            self.assertIn('SurveillAAA', selected)
            self.assertNotIn('TertiaryAAA', selected)
            self.assertEqual(app.session_state['reporting_weeks_selected'], [f'2026W{week}' for week in range(36, 40)])
            self.assert_shared_reporting(app, 186, 9)
            queries = app.session_state['_dataset_queries']
            self.assertTrue(any(query['filters'] == ['pe:202609'] and len(query['items']) == 45 for query in queries))
            self.assertTrue(any(query['filters'] == ['pe:2026W36;2026W37;2026W38;2026W39']
                                and len(query['items']) == 5 for query in queries))
            self.assertFalse(any('tertiaire' in str(column).casefold() for column in app.session_state['_test_main_raw'].columns))
            self.assertTrue(all(query['unit'] == 'ZoneAAAAAAA' for query in queries))
            for sheet in ['Perf_Final_Completude', 'Synthese_Comparative', 'Resultats_Regles']:
                _, rows = excel_sheet(downloads['Télécharger Excel'], sheet)
                self.assertEqual(rows[0]['Reports_Actual'], 186)
            _, prompt = excel_sheet(downloads['Télécharger Excel'], 'Analyse_Promptitude')
            self.assertEqual(prompt[0]['Reports_Actual_On_Time'], 186)
            _, periods = excel_sheet(downloads['Télécharger Excel'], 'Periodes_Surveillance')
            self.assertEqual([row['Semaine'] for row in periods], [f'2026W{week}' for week in range(36, 40)])
            headers, weekly = excel_sheet(downloads['Télécharger Excel'], 'Reporting_Surveillance')
            actual = next(column for column in headers if column.endswith('Actual reports'))
            self.assertEqual(weekly[0][actual], 68)

            self.widget(app, 'button', 'Calculer tous les scores (M et M-1)').click()
            self.run_app(app)
            self.assertEqual(app.session_state['_test_export_frame'][TOTAL_COL].tolist(), [75])
            self.assertEqual(app.session_state['_dataset_posts'], 2)
            self.widget(app, 'multiselect', 'Semaines de surveillance').unselect('2026W39')
            self.run_app(app)
            self.assert_shared_reporting(app, 169, 9)
            self.assertEqual(app.session_state['_dataset_posts'], 2)
            self.assertEqual(app.session_state['_test_export_frame']['Règles violées (M)'].tolist(), [5])
            self.assertEqual(app.session_state['_test_export_frame'][TOTAL_COL].tolist(), [75])

            self.widget(app, 'multiselect', 'Datasets évalués pour le score qualité').unselect('SurveillAAA')
            self.run_app(app)
            self.assert_shared_reporting(app, 118, 8)
            self.assertEqual(app.session_state['_dataset_posts'], 2)
            self.widget(app, 'multiselect', 'Datasets évalués pour le score qualité').select('SurveillAAA')
            self.run_app(app)
            self.choice(app, 'Choix des semaines').set_value('Intervalle de semaines')
            self.run_app(app)
            self.widget(app, 'selectbox', 'Semaine de début').select('2026W36')
            self.widget(app, 'selectbox', 'Semaine de fin').select('2026W39')
            self.run_app(app)
            self.assert_shared_reporting(app, 186, 9)
            self.assertEqual(app.session_state['_dataset_posts'], 2)

            app.session_state['_dataset_fail_weekly'] = True
            self.select_month_range(app, 'Octobre', 'Octobre')
            self.assert_partial_monthly_visible(app)
            _, available = excel_sheet(downloads['Télécharger Excel'], 'Reporting_Disponible')
            self.assertEqual(available[0]['Rapports mensuels disponibles'], 118)
            self.assertEqual(available[0]['Attendus mensuels disponibles'], 118)
            self.assertIsNone(available[0]['Rapports hebdomadaires disponibles'])
            self.assertEqual(available[0]['État des sources'], 'Incomplet')
            self.assertNotIn('categorisation_result', app.session_state)
            self.assertEqual(app.session_state['_dataset_posts'], 2)

    def test_no_selected_week_preserves_monthly_data_without_a_weekly_request_or_complete_score(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.select_month_range(app, 'Septembre', 'Septembre')
            app.session_state['_dataset_queries'] = []
            for period in tuple(app.session_state['reporting_weeks_selected']):
                self.widget(app, 'multiselect', 'Semaines de surveillance').unselect(period)
            self.run_app(app)
            self.assertEqual(app.session_state['reporting_weeks_selected'], [])
            self.assert_partial_monthly_visible(app)
            self.assertFalse(any(
                any(item.startswith('SurveillAAA.') for item in query['items'])
                for query in app.session_state['_dataset_queries']
            ))
            self.assertNotIn('_dataset_posts', app.session_state)

    def test_outer_combined_failure_recovers_only_selected_monthly_datasets(self):
        app = self.new_app({'get_quality_reporting_data': '''
def get_quality_reporting_data(favori_id, period=None, cache_user=None, ou_dimension=None,
                               unit_ids=None, dataset_ids=None, weekly_periods=None):
    result = pd.DataFrame()
    result.attrs['quality_warnings'] = ['Préparation combinée indisponible pour cette analyse.']
    return result
'''})
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.select_month_range(app, 'Septembre', 'Septembre')
            self.assert_partial_monthly_visible(app)
            self.assertNotEqual(self.available_view(app)['Rapports mensuels disponibles'].iloc[0], 1117)
            self.assertTrue(app.session_state['_test_main_raw'].attrs['monthly_reporting_fallback'])
            self.assertNotIn('_dataset_posts', app.session_state)

    def test_observed_monthly_zero_stays_visible_during_a_weekly_failure(self):
        app = self.new_app()
        app.session_state['_dataset_zero_monthly'] = True
        app.session_state['_dataset_fail_weekly'] = True
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.select_month_range(app, 'Septembre', 'Septembre')
            self.assert_partial_monthly_visible(app, monthly=106)
            raw = app.session_state['_test_main_raw']
            for column in ['Service 0 Actual reports', 'Service 0 Reporting rate', 'Service 0 Reporting rate on time']:
                self.assertEqual(raw[column].tolist(), [0])
            self.assertEqual(app.session_state['_test_main_comp']['Service 0 Reporting rate'].tolist(), [0])
            self.assertEqual(app.session_state['_test_main_prompt']['Service 0 Reporting rate on time'].tolist(), [0])

    def test_fixed_semester_quarter_year_and_cross_year_interval_use_all_selected_periods(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.widget(app, 'selectbox', "Choisir l'année :").select('2026')
            self.choice(app, 'Mode de sélection de la période').set_value('Période fixe')
            self.choice(app, 'Périodicité').set_value('Semestriel')
            self.run_app(app)
            self.widget(app, 'selectbox', 'Semestre').select('2026S1')
            self.run_app(app)
            self.assert_shared_reporting(app, 118 * 6 + 17 * 26, 9)
            self.assertEqual(app.session_state['reporting_weeks_selected'], [f'2026W{week}' for week in range(1, 27)])
            self.assertTrue(any(query['filters'] == ['pe:202601;202602;202603;202604;202605;202606']
                                for query in app.session_state['_dataset_queries']))
            self.choice(app, 'Périodicité').set_value('Trimestriel')
            self.run_app(app)
            self.widget(app, 'selectbox', 'Trimestre').select('2026Q1')
            self.run_app(app)
            self.assert_shared_reporting(app, 118 * 3 + 17 * 13, 9)
            self.choice(app, 'Périodicité').set_value('Annuel')
            self.run_app(app)
            self.assert_shared_reporting(app, 118 * 12 + 17 * 52, 9)

            self.choice(app, 'Mode de sélection de la période').set_value('Intervalle')
            self.choice(app, 'Périodicité').set_value('Mensuel')
            self.run_app(app)
            self.select_month_range(app, 'Novembre', 'Février', last_year='2027')
            self.assert_shared_reporting(app, 118 * 4 + 17 * 18, 9)
            self.assertTrue(any(query['filters'] == ['pe:202611;202612;202701;202702']
                                for query in app.session_state['_dataset_queries']))
            self.assertIn('2026W53', app.session_state['reporting_weeks_selected'])
            self.assertIn('2027W8', app.session_state['reporting_weeks_selected'])
            self.assertEqual(len(app.session_state['reporting_weeks_selected']), 18)
            self.assertNotIn('_dataset_posts', app.session_state)

    def test_weekly_and_quarterly_intervals_keep_explicit_weeks_and_year_bounds(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.widget(app, 'selectbox', "Choisir l'année :").select('2026')
            self.widget(app, 'selectbox', 'Année de fin').select('2026')
            self.choice(app, 'Périodicité').set_value('Hebdomadaire')
            self.run_app(app)
            self.widget(app, 'selectbox', 'Semaine de début de période').select('2026W36')
            self.widget(app, 'selectbox', 'Semaine de fin de période').select('2026W39')
            self.run_app(app)
            weeks = [f'2026W{week}' for week in range(36, 40)]
            self.assertEqual(app.session_state['reporting_weeks_selected'], weeks)
            self.assert_shared_reporting(app, 118 + 17 * 4, 9)
            self.assertEqual(self.widget(app, 'date_input', 'Date de début').value, date(2026, 8, 31))
            self.assertEqual(self.widget(app, 'date_input', 'Date de fin').value, date(2026, 9, 27))
            queries = app.session_state['_dataset_queries']
            self.assertTrue(any(query['filters'] == ['pe:202609'] and len(query['items']) == 45 for query in queries))
            self.assertFalse(any(query['filters'] == ['pe:202608;202609'] for query in queries))
            self.assertTrue(any(query['filters'] == ['pe:' + ';'.join(weeks)] for query in queries))
            captions = '\n'.join(str(item.value) for item in app.caption)
            self.assertIn('202609', captions)
            self.assertIn('31/08/2026', captions)
            self.assertIn('27/09/2026', captions)
            self.widget(app, 'button', "Lancer l'analyse des violations").click()
            self.run_app(app)
            self.assertEqual(app.session_state['_dataset_post_queries'][-1]['startDate'], '2026-08-31')
            self.assertEqual(app.session_state['_dataset_post_queries'][-1]['endDate'], '2026-09-27')

            self.choice(app, 'Périodicité').set_value('Annuel')
            self.run_app(app)
            self.assertEqual(self.widget(app, 'selectbox', 'Année de fin').value, '2026')
            self.assert_shared_reporting(app, 118 * 12 + 17 * 52, 9)
            self.choice(app, 'Mode de sélection de la période').set_value('Période fixe')
            self.run_app(app)
            self.choice(app, 'Mode de sélection de la période').set_value('Intervalle')
            self.choice(app, 'Périodicité').set_value('Mensuel')
            self.run_app(app)
            self.assertEqual(self.widget(app, 'selectbox', 'Année de fin').value, '2026')

            self.choice(app, 'Périodicité').set_value('Trimestriel')
            self.widget(app, 'selectbox', 'Année de fin').select('2027')
            self.run_app(app)
            self.widget(app, 'selectbox', 'Trimestre de début').select('2026Q3')
            self.widget(app, 'selectbox', 'Trimestre de fin').select('2027Q1')
            self.run_app(app)
            self.assert_shared_reporting(app, 118 * 9 + 17 * 39, 9)
            self.assertEqual(len(app.session_state['reporting_weeks_selected']), 39)
            self.assertEqual(app.session_state['reporting_weeks_selected'][0], '2026W27')
            self.assertEqual(app.session_state['reporting_weeks_selected'][-1], '2027W12')
            months = ['202607', '202608', '202609', '202610', '202611', '202612', '202701', '202702', '202703']
            self.assertTrue(any(query['filters'] == ['pe:' + ';'.join(months)]
                                for query in app.session_state['_dataset_queries']))
            self.assertEqual(app.session_state['_dataset_posts'], 1)

    def test_daily_inclusive_bounds_and_week_attachment_change_only_the_requested_scope(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.choice(app, 'Périodicité').set_value('Journalier')
            self.run_app(app)
            self.widget(app, 'date_input', 'Date de début (incluse)').set_value(date(2026, 9, 1))
            self.widget(app, 'date_input', 'Date de fin (incluse)').set_value(date(2026, 9, 30))
            self.run_app(app)
            self.assertEqual(app.session_state['reporting_weeks_selected'], [f'2026W{week}' for week in range(36, 40)])
            self.assert_shared_reporting(app, 186, 9)
            self.assertEqual(self.widget(app, 'date_input', 'Date de début').value, date(2026, 9, 1))
            self.assertEqual(self.widget(app, 'date_input', 'Date de fin').value, date(2026, 9, 30))

            self.widget(app, 'date_input', 'Date de fin (incluse)').set_value(date(2026, 9, 26))
            self.run_app(app)
            self.assertEqual(app.session_state['reporting_weeks_selected'], [f'2026W{week}' for week in range(36, 39)])
            self.assert_shared_reporting(app, 169, 9)
            self.assertEqual(self.widget(app, 'date_input', 'Date de fin').value, date(2026, 9, 26))

            self.widget(app, 'date_input', 'Date de fin (incluse)').set_value(date(2026, 9, 30))
            self.choice(app, 'Rattachement des semaines').set_value('Semaines entièrement comprises')
            self.run_app(app)
            self.assertEqual(app.session_state['reporting_weeks_selected'], [f'2026W{week}' for week in range(37, 40)])
            self.assert_shared_reporting(app, 169, 9)
            before = len(app.session_state['_dataset_queries'])
            self.widget(app, 'date_input', 'Date de début (incluse)').set_value(date(2026, 10, 1))
            self.run_app(app)
            self.assertGreater(len(app.error), 0)
            self.assertEqual(len(app.session_state['_dataset_queries']), before)
            self.assertNotIn('_dataset_posts', app.session_state)

    def test_sunday_quarter_end_is_included_once_and_next_quarter_opens_the_next_week(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.widget(app, 'selectbox', "Choisir l'année :").select('2024')
            self.choice(app, 'Mode de sélection de la période').set_value('Période fixe')
            self.choice(app, 'Périodicité').set_value('Trimestriel')
            self.run_app(app)
            self.widget(app, 'selectbox', 'Trimestre').select('2024Q1')
            self.run_app(app)
            first = list(app.session_state['reporting_weeks_selected'])
            self.assertEqual(first, [f'2024W{week}' for week in range(1, 14)])
            self.assertEqual(self.widget(app, 'date_input', 'Date de début').value, date(2024, 1, 1))
            self.assertEqual(self.widget(app, 'date_input', 'Date de fin').value, date(2024, 3, 31))
            self.assert_shared_reporting(app, 118 * 3 + 17 * 13, 9)

            self.widget(app, 'selectbox', 'Trimestre').select('2024Q2')
            self.run_app(app)
            second = list(app.session_state['reporting_weeks_selected'])
            self.assertEqual(second, [f'2024W{week}' for week in range(14, 27)])
            self.assertFalse(set(first) & set(second))
            self.assertEqual(self.widget(app, 'date_input', 'Date de début').value, date(2024, 4, 1))
            self.assertEqual(self.widget(app, 'date_input', 'Date de fin').value, date(2024, 6, 30))
            self.assert_shared_reporting(app, 118 * 3 + 17 * 13, 9)
            self.assertNotIn('_dataset_posts', app.session_state)

    def test_independent_dates_within_same_month_use_their_own_weeks_and_weekly_m1_uses_reporting_month(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.select_month_range(app, 'Septembre', 'Septembre')
            weeks = [f'2026W{week}' for week in range(36, 40)]
            self.assertEqual(app.session_state['reporting_weeks_selected'], weeks)
            self.assert_shared_reporting(app, 186, 9)

            self.widget(app, 'date_input', 'Date de début').set_value(date(2026, 9, 7))
            self.widget(app, 'date_input', 'Date de fin').set_value(date(2026, 9, 13))
            self.run_app(app)
            self.assertEqual(app.session_state['validation_weeks_selected'], ['2026W37'])
            self.assertEqual(app.session_state['reporting_weeks_selected'], weeks)
            self.assertEqual(app.session_state['_test_main_summary']['Reports_Actual'].tolist(), [186])
            self.assertEqual(app.session_state['_test_export_frame']['Reports_Actual'].tolist(), [135])
            self.assertTrue(any(query['filters'] == ['pe:2026W37']
                                for query in app.session_state['_dataset_queries']))
            self.widget(app, 'button', "Lancer l'analyse des violations").click()
            self.run_app(app)
            self.assertEqual(app.session_state['_dataset_post_queries'][-1]['startDate'], '2026-09-07')
            self.assertEqual(app.session_state['_dataset_post_queries'][-1]['endDate'], '2026-09-13')

            self.choice(app, 'Périodicité').set_value('Hebdomadaire')
            self.run_app(app)
            self.widget(app, 'selectbox', 'Semaine de début de période').select('2026W36')
            self.widget(app, 'selectbox', 'Semaine de fin de période').select('2026W39')
            self.run_app(app)
            self.assert_shared_reporting(app, 186, 9)
            self.widget(app, 'button', 'Calculer tous les scores (M et M-1)').click()
            self.run_app(app)
            self.assertEqual(app.session_state['_dataset_posts'], 3)
            requests = app.session_state['_dataset_post_queries']
            self.assertEqual((requests[-2]['startDate'], requests[-2]['endDate']), ('2026-08-31', '2026-09-27'))
            self.assertEqual((requests[-1]['startDate'], requests[-1]['endDate']), ('2026-08-01', '2026-08-31'))
            self.assert_shared_reporting(app, 186, 9)
            self.assertEqual(app.session_state['_test_export_frame'][TOTAL_COL].tolist(), [75])


if __name__ == '__main__':
    unittest.main()

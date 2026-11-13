"""Le tableau qualité peut couvrir les datasets accessibles hors du favori."""

from datetime import date
import unittest
from unittest import mock

import pandas as pd
from streamlit.testing.v1 import AppTest

from categorisation import TOTAL_COL
from scoped_analytics import METRIC_LABELS
import test_dashboard_categorisation as dashboard_test
from test_dashboard_reporting_load import actual_loader_functions


GEOGRAPHY_CATALOG = '''
def get_organisation_catalog(cache_user=None):
    from geography import build_geography_catalog
    country = {"id": "CountryAAAA", "displayName": "Pays test", "level": 1}
    province = {"id": "ProvinceAAA", "displayName": "Province A", "level": 2}
    return build_geography_catalog([
        country, {**province, "ancestors": [country]},
        {"id": "ZoneAAAAAAA", "displayName": "ZS A", "level": 3, "ancestors": [country, province]},
        {"id": "ZoneBBBBBBB", "displayName": "ZS B", "level": 3, "ancestors": [country, province]},
    ])
'''


DATASET_CATALOG = '''
def get_quality_dataset_catalog(cache_user=None):
    st.session_state["_catalog_dataset_calls"] = st.session_state.get("_catalog_dataset_calls", 0) + 1
    return [
        {"id": "Data0000000", "displayName": "Dataset du favori", "periodType": "Monthly"},
        {"id": "Data0000001", "displayName": "Hors favori complet", "periodType": "Monthly"},
        {"id": "Data0000002", "displayName": "Hors favori incomplet", "periodType": "Monthly"},
    ]
'''


REPORTING_AND_VALIDATION_CLIENT = '''
def get_dhis2_client():
    class Client:
        def get(self, endpoint, params=None):
            if endpoint.startswith("visualizations/"):
                return {"dataDimensionItems": [
                    {"reportingRate": {
                        "id": "Data0000000." + metric, "metric": metric,
                        "dataSet": {"id": "Data0000000", "displayName": "Dataset du favori"},
                    }} for metric in ["ACTUAL_REPORTS", "EXPECTED_REPORTS", "REPORTING_RATE",
                                      "ACTUAL_REPORTS_ON_TIME", "REPORTING_RATE_ON_TIME"]
                ]}
            if endpoint != "analytics.json":
                raise AssertionError("GET fixture inattendu : " + endpoint)
            dimensions = dict(value.split(":", 1) for key, value in params if key == "dimension")
            unit = dimensions["ou"]
            if unit not in ["ZoneAAAAAAA", "ZoneBBBBBBB"]:
                raise AssertionError("Chaque requete Analytics doit cibler une seule zone accessible.")
            items = dimensions["dx"].split(";")
            filters = [value for key, value in params if key == "filter"]
            queries = list(st.session_state.get("_catalog_analytics_queries", []))
            queries.append({"unit": unit, "items": items, "filters": filters})
            st.session_state["_catalog_analytics_queries"] = queries
            period = next(value[3:] for value in filters if value.startswith("pe:"))
            factor = 0.5 if period.endswith("02") else 1
            # expected, actual, timely : les deux datasets hors favori modifient
            # les sommes, les taux ponderes et le nombre de datasets >=95 %.
            values = [(10, 10, 8), (20, 20, 18), (10, 5, 2)] if unit == "ZoneAAAAAAA" else [
                (10, 5, 4), (10, 10, 10), (20, 20, 18)
            ]
            rows = []
            for item in items:
                dataset, metric = item.rsplit(".", 1)
                expected, actual, timely = values[int(dataset[4:])]
                actual, timely = actual * factor, timely * factor
                metrics = {
                    "ACTUAL_REPORTS": actual, "EXPECTED_REPORTS": expected,
                    "REPORTING_RATE": actual / expected * 100,
                    "ACTUAL_REPORTS_ON_TIME": timely,
                    "REPORTING_RATE_ON_TIME": timely / expected * 100,
                }
                rows.append([unit, item, metrics[metric]])
            return {"headers": [{"name": name} for name in ["ou", "dx", "value"]],
                    "rows": rows, "metaData": {"items": {unit: {
                        "displayName": "ZS A" if unit == "ZoneAAAAAAA" else "ZS B",
                    }}}}

        def post(self, endpoint, data=None):
            if endpoint != "dataAnalysis/validationRules":
                raise AssertionError("POST fixture inattendu : " + endpoint)
            calls = list(st.session_state.get("_catalog_validation_posts", []))
            calls.append(dict(data))
            st.session_state["_catalog_validation_posts"] = calls
            previous = data["startDate"] < st.session_state["validation_start_date"].isoformat()
            if data["ou"] == "ZoneAAAAAAA":
                rules = ["RuleAAAAAAA", "RuleCCCCCCC" if previous else "RuleBBBBBBB"]
                centre = "CentreAAAAA"
            elif data["ou"] == "ZoneBBBBBBB":
                rules = ["RuleDDDDDDD"] if previous else []
                centre = "CentreBBBBB"
            else:
                raise AssertionError("L'analyse doit cibler chaque zone separement.")
            return [{
                "organisationUnitId": centre,
                "organisationUnitPath": "/CountryAAAA/ProvinceAAA/" + data["ou"] + "/" + centre,
                "validationRuleId": rule, "periodId": data["startDate"][:7].replace("-", ""),
                "leftSideValue": 5, "rightSideValue": 1,
            } for rule in rules]
    return Client()
'''


VALIDATION_METADATA = '''
def get_validation_detail_metadata(rule_ids, org_ids, cache_user=None):
    rules = {uid: {"id": uid, "displayName": "Regle " + uid, "operator": "equal_to", "importance": "HIGH"}
             for uid in rule_ids}
    orgs = {centre: {"id": centre, "displayName": "Centre " + label,
            "path": "/CountryAAAA/ProvinceAAA/" + zone + "/" + centre}
            for centre, zone, label in [("CentreAAAAA", "ZoneAAAAAAA", "A"), ("CentreBBBBB", "ZoneBBBBBBB", "B")]
            if centre in org_ids}
    return rules, orgs, []
'''


class QualityDatasetCatalogAppTests(unittest.TestCase):
    metric_columns = [
        'Reports_Actual', 'Complétude globale (%)', 'Promptitude globale (%)',
        'Nombre de datasets complétude ≥95 %', 'Score de qualité',
    ]
    comparison_columns = [
        'Règles violées (M)', 'Règles violées (M-1)',
        'Règles corrigées (M-1 -> M)', 'Règles restantes (M-1 -> M)',
    ]

    def new_app(self):
        source = dashboard_test.app_fixture_source({
            **actual_loader_functions(),
            'get_quality_dataset_catalog': DATASET_CATALOG,
            'get_organisation_catalog': GEOGRAPHY_CATALOG,
            'get_dhis2_client': REPORTING_AND_VALIDATION_CLIENT,
            'get_validation_detail_metadata': VALIDATION_METADATA,
        })
        source += '\nst.session_state["_catalog_main_source"] = target_df.copy()\n'
        source += 'st.session_state["_catalog_quality_source"] = quality_reporting_source.copy()\n'
        source += 'st.session_state["_catalog_export_context"] = df_validation_context_export.copy()\n'
        app = AppTest.from_string(source, default_timeout=45)
        app.session_state['connected'] = True
        app.session_state['dhis2_user'] = 'fixture-quality-catalog-' + self._testMethodName
        app.session_state['dhis2_pass'] = 'fixture-password'
        app.session_state['filter_level_main'] = 'Province'
        app.session_state['filter_province_main'] = 'ProvinceAAA'
        # Le prefixe des fixtures preserve les anciens tests du favori ; cette
        # selection explicite verifie le mode par defaut de la production.
        app.session_state['quality_dataset_scope'] = 'Tous les datasets DHIS2'
        return app

    def widget(self, app, kind, label):
        return next(item for item in getattr(app, kind) if item.label == label)

    def run_app(self, app):
        # AppTest conserve parfois plusieurs versions du widget de detail apres
        # les reruns du batch ; une selection explicite les resynchronise.
        for widget in app.selectbox:
            if widget.key == 'validation_selected_row':
                widget.select(0)
        app.run(timeout=45)
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])

    def frame(self, app):
        return app.session_state['_test_export_frame'].set_index('Zone de santé').copy()

    def assert_metrics(self, app, expected):
        frame = self.frame(app)
        self.assertCountEqual(frame.index.tolist(), expected)
        for name, values in expected.items():
            for column, expected_value in zip(self.metric_columns, values):
                self.assertAlmostEqual(frame.loc[name, column], expected_value, places=2, msg=(name, column))
        for column in self.metric_columns:
            self.assertTrue(pd.api.types.is_numeric_dtype(frame[column]), (column, frame[column].dtype))
        return frame

    def choose_external_dataset(self, app):
        self.widget(app, 'selectbox', 'Datasets du tableau qualité').select('Choisir les datasets')
        self.run_app(app)
        self.widget(app, 'multiselect', 'Choisir les datasets à évaluer').set_value(['Data0000002'])
        self.run_app(app)

    def test_all_accessible_datasets_contribute_even_when_absent_from_the_favorite(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            scope = self.widget(app, 'selectbox', 'Datasets du tableau qualité')
            self.assertEqual(scope.value, 'Tous les datasets DHIS2')
            self.assertEqual(scope.options, [
                'Tous les datasets DHIS2', 'Datasets du rapport sélectionné', 'Choisir les datasets',
            ])
            frame = self.assert_metrics(app, {
                'ZS A': [35, 87.5, 70, 2, 78.8],
                'ZS B': [35, 87.5, 80, 2, 83.8],
            })
            self.assertTrue(frame[['Règles violées (M)', TOTAL_COL]].isna().all().all())
            self.assertNotIn('_catalog_validation_posts', app.session_state)
            main_source = app.session_state['_catalog_main_source']
            self.assertIn('Dataset du favori Actual reports', main_source.columns)
            self.assertFalse(any('Hors favori' in column for column in main_source.columns))
            quality_source = app.session_state['_catalog_quality_source'].set_index('Organisation unit ID')
            self.assertEqual(quality_source.loc['ZoneAAAAAAA', 'Hors favori complet Actual reports'], 20)
            self.assertEqual(quality_source.loc['ZoneBBBBBBB', 'Hors favori incomplet Actual reports'], 20)
            queries = list(app.session_state['_catalog_analytics_queries'])
            all_items = {f'Data{number:07d}.{metric}' for number in range(3) for metric in METRIC_LABELS}
            self.assertTrue(all(len(query['items']) <= 15 for query in queries))
            self.assertTrue(all(query['filters'] == ['pe:202501'] for query in queries))
            for unit in ['ZoneAAAAAAA', 'ZoneBBBBBBB']:
                requested = {item for query in queries if query['unit'] == unit for item in query['items']}
                self.assertTrue(all_items.issubset(requested))

    def test_custom_dataset_recomputes_reporting_and_preserves_validation_comparisons(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.widget(app, 'button', 'Calculer tous les scores (M et M-1)').click()
            self.run_app(app)
            before = self.frame(app)
            self.assertEqual(before.loc['ZS A', self.comparison_columns].tolist(), [2, 2, 1, 1])
            self.assertEqual(before.loc['ZS B', self.comparison_columns].tolist(), [0, 1, 1, 0])
            self.assertAlmostEqual(before.loc['ZS A', 'Ratio / 100 rapports'], 2 / 35 * 100)
            saved = app.session_state['categorisation_result']
            current = saved['violations'].copy(deep=True)
            previous = saved['previous_violations'].copy(deep=True)
            posts = list(app.session_state['_catalog_validation_posts'])
            self.assertEqual(len(posts), 4)

            self.choose_external_dataset(app)
            selected = self.assert_metrics(app, {
                'ZS A': [5, 50, 20, 0, 35],
                'ZS B': [20, 100, 90, 1, 95],
            })
            pd.testing.assert_frame_equal(selected[self.comparison_columns], before[self.comparison_columns])
            pd.testing.assert_frame_equal(app.session_state['categorisation_result']['violations'], current)
            pd.testing.assert_frame_equal(app.session_state['categorisation_result']['previous_violations'], previous)
            self.assertAlmostEqual(selected.loc['ZS A', 'Ratio / 100 rapports'], 40)
            self.assertLess(selected.loc['ZS A', 'Score complétude /20'], before.loc['ZS A', 'Score complétude /20'])
            self.assertEqual(list(app.session_state['_catalog_validation_posts']), posts)
            context = app.session_state['_catalog_export_context'].astype(str).to_string(index=False)
            self.assertIn('Data0000002', context)
            self.assertIn('Hors favori incomplet', context)

            self.widget(app, 'selectbox', 'Datasets du tableau qualité').select('Datasets du rapport sélectionné')
            self.run_app(app)
            favorite = self.assert_metrics(app, {
                'ZS A': [10, 100, 80, 1, 90],
                'ZS B': [5, 50, 40, 0, 45],
            })
            pd.testing.assert_frame_equal(favorite[self.comparison_columns], before[self.comparison_columns])
            self.assertAlmostEqual(favorite.loc['ZS A', 'Ratio / 100 rapports'], 20)
            self.assertEqual(list(app.session_state['_catalog_validation_posts']), posts)

    def test_province_zone_and_date_filters_recompute_numeric_export_values_without_post(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.choose_external_dataset(app)
            self.widget(app, 'selectbox', 'Zone de santé de l’analyse').select('ZoneAAAAAAA')
            self.run_app(app)
            self.assert_metrics(app, {'ZS A': [5, 50, 20, 0, 35]})
            year = int(self.widget(app, 'selectbox', "Choisir l'année :").value)
            self.widget(app, 'date_input', 'Date de début').set_value(date(year, 2, 1))
            self.widget(app, 'date_input', 'Date de fin').set_value(date(year, 2, 28))
            self.run_app(app)
            self.assert_metrics(app, {'ZS A': [2.5, 25, 10, 0, 17.5]})
            self.assertTrue(any(query['filters'] == [f'pe:{year}02']
                                for query in app.session_state['_catalog_analytics_queries']))

            self.widget(app, 'selectbox', 'Zone de santé de l’analyse').select('')
            self.run_app(app)
            self.assert_metrics(app, {
                'ZS A': [2.5, 25, 10, 0, 17.5],
                'ZS B': [10, 50, 45, 0, 47.5],
            })
            self.assertEqual(app.session_state['quality_selected_dataset_ids'], ['Data0000002'])
            self.assertNotIn('_catalog_validation_posts', app.session_state)
            self.assertTrue(self.frame(app)[[TOTAL_COL, 'Règles violées (M)']].isna().all().all())

            queries = list(app.session_state['_catalog_analytics_queries'])
            self.widget(app, 'multiselect', 'Choisir les datasets à évaluer').set_value([])
            self.run_app(app)
            empty_selection = self.frame(app)
            self.assertTrue(empty_selection[self.metric_columns].isna().all().all())
            self.assertTrue(empty_selection[[TOTAL_COL, 'Règles violées (M)']].isna().all().all())
            self.assertEqual(list(app.session_state['_catalog_analytics_queries']), queries)
            self.assertNotIn('_catalog_validation_posts', app.session_state)
            self.assertTrue(any('Sélectionnez au moins un dataset' in str(item.value) for item in app.info))


if __name__ == '__main__':
    unittest.main()

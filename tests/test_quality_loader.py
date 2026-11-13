import copy
import unittest
from unittest import mock

import numpy as np
import pandas as pd
import requests

from quality_datasets import select_quality_datasets
from reporting_loader import fetch_reporting_chunks
from reporting_metrics import clean_ou_id
from scoped_analytics import (
    METRIC_LABELS, analytics_to_reporting_frame, build_analytics_query,
    build_quality_reporting_query,
)
from test_dashboard_categorisation import dashboard_functions


class QualityLoaderTests(unittest.TestCase):
    def loader(self, client, **overrides):
        dependencies = {
            'requests': requests,
            'get_dhis2_client': lambda: client,
            'build_quality_reporting_query': build_quality_reporting_query,
            'analytics_to_reporting_frame': analytics_to_reporting_frame,
            'select_quality_datasets': select_quality_datasets,
            **overrides,
        }
        return dashboard_functions('get_quality_reporting_data', **dependencies)['get_quality_reporting_data']

    def favorite(self):
        return {
            'id': 'FavoriteAAA',
            'dataDimensionItems': [
                {'reportingRate': {
                    'id': 'DatasetAAAA.REPORTING_RATE',
                    'dataSet': {'id': 'DatasetAAAA', 'displayName': 'Consultations'},
                }},
                {'indicator': {'id': 'IndicatorAA', 'displayName': 'Autre indicateur'}},
            ],
            'columns': [{'dimension': 'dx', 'items': [{'id': 'DatasetAAAA.EXPECTED_REPORTS'}]}],
            'rows': [{'dimension': 'ou', 'items': [{'id': 'OldZoneAAAA'}]}],
            'filters': [
                {'dimension': 'pe', 'items': [{'id': '202501'}]},
                {'dimension': 'CategoryAAA', 'items': [{'id': 'OptionAAAAA'}]},
            ],
            'aggregationType': 'SUM', 'skipRounding': True, 'completedOnly': False,
        }

    def catalog(self):
        return [
            {'id': 'DatasetAAAA', 'displayName': 'Consultations', 'periodType': 'Monthly'},
            {'id': 'DatasetBBBB', 'displayName': 'Vaccination', 'periodType': 'Monthly'},
            {'id': 'DatasetCCCC', 'displayName': 'Surveillance', 'periodType': 'Monthly'},
        ]

    def test_rate_only_favorite_loads_counts_for_the_selected_zone_and_period(self):
        calls = []

        class Client:
            def get(self, endpoint, params=None):
                calls.append((endpoint, params))
                if endpoint.startswith('visualizations/'):
                    return {'dataDimensionItems': [{'reportingRate': {
                        'id': 'DatasetAAAA.REPORTING_RATE',
                        'dataSet': {'id': 'DatasetAAAA', 'displayName': 'Consultations'},
                    }}]}
                return {
                    'headers': [{'name': 'ou'}, {'name': 'dx'}, {'name': 'value'}],
                    'rows': [['ZoneAAAAAAA', 'DatasetAAAA.' + metric, value] for metric, value in [
                        ('REPORTING_RATE', '90'), ('REPORTING_RATE_ON_TIME', '80'),
                        ('ACTUAL_REPORTS', '9'), ('ACTUAL_REPORTS_ON_TIME', '8'), ('EXPECTED_REPORTS', '10'),
                    ]],
                }

            def post(self, *args, **kwargs):
                raise AssertionError('Le tableau qualité ne doit pas analyser les violations.')

        catalog = mock.Mock(side_effect=AssertionError('Le mode favori ne charge pas le catalogue global.'))
        frame = self.loader(Client(), get_quality_dataset_catalog=catalog)(
            'FavoriteAAA', period='202609', ou_dimension='ZoneAAAAAAA',
        )
        catalog.assert_not_called()
        self.assertEqual(len(calls), 2)
        self.assertEqual(calls[1][0], 'analytics.json')
        self.assertIn(('dimension', 'ou:ZoneAAAAAAA'), calls[1][1])
        self.assertIn(('filter', 'pe:202609'), calls[1][1])
        self.assertEqual(frame['Organisation unit ID'].tolist(), ['ZoneAAAAAAA'])
        self.assertEqual(frame['Consultations Actual reports'].tolist(), [9])
        self.assertEqual(frame['Consultations Expected reports'].tolist(), [10])
        self.assertEqual(frame['Consultations Actual reports on time'].tolist(), [8])

    def test_explicit_catalog_selection_loads_only_its_five_metrics_even_outside_the_favorite(self):
        favorite, catalog = self.favorite(), self.catalog()
        original_favorite, original_catalog = copy.deepcopy(favorite), copy.deepcopy(catalog)
        original_query = build_analytics_query(favorite, '202609;202610', 'LEVEL-3;ProvinceAAA')
        selected = ('DatasetCCCC', 'DatasetBBBB')
        get_catalog = mock.Mock(return_value=catalog)
        selector = mock.Mock(wraps=select_quality_datasets)
        client = mock.Mock()

        def get(endpoint, params=None):
            if endpoint == 'visualizations/FavoriteAAA':
                return favorite
            self.assertEqual(endpoint, 'analytics.json')
            identifiers = next(value[3:].split(';') for key, value in params
                               if key == 'dimension' and value.startswith('dx:'))
            return {
                'headers': [{'name': name} for name in ('ou', 'dx', 'value')],
                'rows': [['ZoneAAAAAAA', identifier, '10'] for identifier in identifiers],
            }

        client.get.side_effect = get
        frame = self.loader(
            client, get_quality_dataset_catalog=get_catalog, select_quality_datasets=selector,
        )('FavoriteAAA', period='202609;202610', cache_user='fixture-account',
          ou_dimension='LEVEL-3;ProvinceAAA', dataset_ids=selected)

        get_catalog.assert_called_once_with(cache_user='fixture-account')
        selector.assert_called_once_with(favorite, selected, catalog)
        self.assertEqual(client.get.call_count, 2)
        query = client.get.call_args_list[-1].kwargs['params']
        dimensions = [value[3:].split(';') for key, value in query
                      if key == 'dimension' and value.startswith('dx:')]
        expected_items = {f'{identifier}.{metric}' for identifier in selected for metric in METRIC_LABELS}
        self.assertEqual(len(dimensions), 1)
        self.assertEqual(set(dimensions[0]), expected_items)
        self.assertEqual(len(dimensions[0]), len(expected_items))
        for parameter in [
            ('dimension', 'ou:LEVEL-3;ProvinceAAA'), ('filter', 'pe:202609;202610'),
            ('filter', 'CategoryAAA:OptionAAAAA'), ('aggregationType', 'SUM'),
            ('skipRounding', 'true'), ('completedOnly', 'false'),
        ]:
            self.assertIn(parameter, query)
        self.assertEqual(set(frame.columns[2:]), {
            f'{name} {suffix}' for name in ('Surveillance', 'Vaccination')
            for suffix in METRIC_LABELS.values()
        })
        self.assertEqual(frame['Vaccination Actual reports'].tolist(), [10])
        self.assertEqual(frame['Surveillance Expected reports'].tolist(), [10])
        client.post.assert_not_called()
        self.assertEqual(favorite, original_favorite)
        self.assertEqual(catalog, original_catalog)
        self.assertEqual(build_analytics_query(favorite, '202609;202610', 'LEVEL-3;ProvinceAAA'), original_query)

    def test_explicit_empty_invalid_or_unknown_datasets_never_fall_back_to_the_favorite(self):
        invalid_selections = [
            (), ('DatasetBBBB', 'invalid'), ('DatasetBBBB', 'MissingAAAA'),
            ('DatasetBBBB;DatasetCCCC',), ('LEVEL-3',), (None,),
        ]
        for selected in invalid_selections:
            with self.subTest(selected=selected):
                favorite, catalog = self.favorite(), self.catalog()
                original_favorite, original_catalog = copy.deepcopy(favorite), copy.deepcopy(catalog)
                client = mock.Mock()

                def get(endpoint, params=None):
                    self.assertEqual(endpoint, 'visualizations/FavoriteAAA')
                    return favorite

                client.get.side_effect = get
                frame = self.loader(client, get_quality_dataset_catalog=mock.Mock(return_value=catalog))(
                    'FavoriteAAA', period='202609', cache_user='fixture-account',
                    ou_dimension='ZoneAAAAAAA', dataset_ids=selected,
                )
                self.assertTrue(frame.empty)
                self.assertTrue(frame.attrs['quality_warnings'])
                self.assertFalse(any(call.args[0] == 'analytics.json' for call in client.get.call_args_list))
                client.post.assert_not_called()
                self.assertEqual(favorite, original_favorite)
                self.assertEqual(catalog, original_catalog)

    def test_unavailable_catalog_never_loads_the_favorite_in_place_of_selected_datasets(self):
        for outcome in (None, requests.ConnectionError('password=fixture')):
            with self.subTest(outcome=type(outcome).__name__):
                get_catalog = mock.Mock()
                if isinstance(outcome, Exception):
                    get_catalog.side_effect = outcome
                else:
                    get_catalog.return_value = outcome
                client = mock.Mock()
                client.get.return_value = self.favorite()
                frame = self.loader(client, get_quality_dataset_catalog=get_catalog)(
                    'FavoriteAAA', period='202609', ou_dimension='ZoneAAAAAAA',
                    dataset_ids=('DatasetBBBB',),
                )
                self.assertTrue(frame.empty)
                self.assertTrue(frame.attrs['quality_warnings'])
                self.assertNotIn('password', str(frame.attrs))
                get_catalog.assert_called_once_with(cache_user=None)
                self.assertFalse(any(call.args[0] == 'analytics.json' for call in client.get.call_args_list))
                client.post.assert_not_called()

    def test_explicit_unit_loader_receives_the_same_datasets_account_period_and_scope(self):
        result = pd.DataFrame({'Organisation unit ID': ['ZoneAAAAAAA']})
        result.attrs['reporting_warnings'] = ['Lot incomplet']
        load_for_units = mock.Mock(return_value=result)
        client = mock.Mock()
        get_catalog = mock.Mock()
        frame = self.loader(client, load_reporting_for_units=load_for_units,
                            get_quality_dataset_catalog=get_catalog)(
            'FavoriteAAA', period='202609;202610', cache_user='fixture-account',
            ou_dimension='LEVEL-3;ProvinceAAA', unit_ids=('ZoneAAAAAAA',),
            dataset_ids=('DatasetBBBB', 'DatasetCCCC'),
        )
        load_for_units.assert_called_once_with(
            'FavoriteAAA', '202609;202610', 'LEVEL-3;ProvinceAAA', ('ZoneAAAAAAA',),
            cache_user='fixture-account', quality=True, dataset_ids=('DatasetBBBB', 'DatasetCCCC'),
        )
        self.assertIs(frame, result)
        self.assertEqual(frame.attrs['quality_warnings'], ['Lot incomplet'])
        client.get.assert_not_called()
        client.post.assert_not_called()
        get_catalog.assert_not_called()

    def test_unit_chunks_keep_catalog_selection_and_account_in_the_analytics_cache(self):
        favorite, catalog = self.favorite(), self.catalog()
        original_favorite = copy.deepcopy(favorite)
        get_visualization = mock.Mock(return_value=favorite)
        get_catalog = mock.Mock(return_value=catalog)
        cached_analytics = mock.Mock()
        st = mock.Mock()

        def analytics(params, cache_user=None):
            self.assertEqual(cache_user, 'fixture-account')
            self.assertIsInstance(params, tuple)
            dimensions = {value.split(':', 1)[0]: value.split(':', 1)[1].split(';')
                          for key, value in params if key == 'dimension'}
            return {
                'headers': [{'name': name} for name in ('ou', 'dx', 'value')],
                'rows': [[unit, identifier, '10'] for unit in dimensions['ou']
                         for identifier in dimensions['dx']],
            }

        cached_analytics.side_effect = analytics
        load_for_units = dashboard_functions(
            'load_reporting_for_units', st=st,
            get_reporting_visualization=get_visualization,
            get_quality_dataset_catalog=get_catalog,
            get_cached_analytics=cached_analytics,
            select_quality_datasets=select_quality_datasets,
            build_quality_reporting_query=build_quality_reporting_query,
            build_analytics_query=build_analytics_query,
            fetch_reporting_chunks=fetch_reporting_chunks,
            clean_ou_id=clean_ou_id,
        )['load_reporting_for_units']
        frame = load_for_units(
            'FavoriteAAA', '202609;202610', 'LEVEL-3;ProvinceAAA',
            ('ZoneAAAAAAA', 'ZoneBBBBBBB'), cache_user='fixture-account',
            quality=True, dataset_ids=('DatasetBBBB',),
        )

        get_visualization.assert_called_once_with('FavoriteAAA', cache_user='fixture-account')
        get_catalog.assert_called_once_with(cache_user='fixture-account')
        self.assertEqual(cached_analytics.call_count, 2)
        requested_units = []
        for call in cached_analytics.call_args_list:
            query = call.args[0]
            requested_units.extend(value[3:] for key, value in query
                                   if key == 'dimension' and value.startswith('ou:'))
            requested_items = next(value[3:].split(';') for key, value in query
                                   if key == 'dimension' and value.startswith('dx:'))
            self.assertEqual(set(requested_items), {f'DatasetBBBB.{metric}' for metric in METRIC_LABELS})
            self.assertEqual(len(requested_items), len(METRIC_LABELS))
            self.assertIn(('filter', 'pe:202609;202610'), query)
            self.assertIn(('filter', 'CategoryAAA:OptionAAAAA'), query)
            self.assertIn(('skipRounding', 'true'), query)
            self.assertIn(('completedOnly', 'false'), query)
            self.assertEqual(call.kwargs, {'cache_user': 'fixture-account'})
        self.assertEqual(requested_units, ['ZoneAAAAAAA', 'ZoneBBBBBBB'])
        self.assertEqual(frame['Organisation unit ID'].tolist(), requested_units)
        self.assertEqual(frame['Vaccination Actual reports'].tolist(), [10, 10])
        self.assertEqual(favorite, original_favorite)
        st.progress.return_value.empty.assert_called_once_with()

    def test_reporting_failure_is_explicit_without_copying_server_exception_text(self):
        class Client:
            def get(self, *args, **kwargs):
                response = requests.Response()
                response.status_code = 504
                raise requests.HTTPError('password=fixture', response=response)

        frame = self.loader(Client())('FavoriteAAA', period='202609', ou_dimension='ZoneAAAAAAA')
        self.assertTrue(frame.empty)
        message = ' '.join(frame.attrs['quality_warnings'])
        self.assertIn('504', message)
        self.assertNotIn('password', message)

    def test_missing_numbers_render_as_dashes_while_source_values_remain_numeric(self):
        styler = dashboard_functions('style_taux', 'style_score', '_build_styler')['_build_styler']
        frame = pd.DataFrame({'Score de qualité': [89.0, np.nan]})
        html = styler(frame, score_cols=['Score de qualité'], missing_text='—').to_html()
        self.assertIn('—</td>', html)
        self.assertNotIn('>None</td>', html)
        self.assertNotIn('>nan</td>', html)
        self.assertTrue(pd.api.types.is_numeric_dtype(frame['Score de qualité']))
        self.assertTrue(pd.isna(frame.iloc[1, 0]))


if __name__ == '__main__':
    unittest.main()

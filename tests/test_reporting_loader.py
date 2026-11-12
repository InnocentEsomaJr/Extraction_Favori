"""Les lots Analytics restent exacts et les erreurs ne deviennent pas des zéros."""

import unittest
from unittest import mock

import pandas as pd
import requests
from urllib3.exceptions import MaxRetryError, ReadTimeoutError

from reporting_loader import fetch_reporting_chunks, format_reporting_error


METRICS = ['ACTUAL_REPORTS', 'EXPECTED_REPORTS', 'ACTUAL_REPORTS_ON_TIME',
           'REPORTING_RATE', 'REPORTING_RATE_ON_TIME']
SUFFIXES = ['Actual reports', 'Expected reports', 'Actual reports on time',
            'Reporting rate', 'Reporting rate on time']


class ReportingChunkTests(unittest.TestCase):
    def http_error(self, status):
        response = requests.Response()
        response.status_code = status
        return requests.HTTPError('private password URL', response=response)

    def query(self, count=4):
        labels = {f'Dataset{number:04d}.{metric}': f'Dataset {number} {suffix}'
                  for number in range(count) for metric, suffix in zip(METRICS, SUFFIXES)}
        params = [('dimension', 'dx:' + ';'.join(labels)), ('dimension', 'ou:LEVEL-3;ProvinceAAA'),
                  ('filter', 'pe:202609;202610'), ('filter', 'CategoryAAA:OptionAAAAA'),
                  ('includeMetadataDetails', 'true'), ('aggregationType', 'DEFAULT'), ('skipRounding', 'false')]
        return params, labels

    def client(self, failure=None):
        client = mock.Mock()

        def get(endpoint, params):
            self.assertEqual(endpoint, 'analytics.json')
            dimensions = {value.split(':', 1)[0]: value.split(':', 1)[1].split(';')
                          for key, value in params if key == 'dimension'}
            if failure:
                failure(dimensions)
            units, items = dimensions['ou'], dimensions['dx']
            return {'headers': [{'name': name} for name in ['ou', 'dx', 'value']],
                    'rows': [[unit, item, {'ACTUAL_REPORTS': 134, 'EXPECTED_REPORTS': 134,
                                           'ACTUAL_REPORTS_ON_TIME': 100, 'REPORTING_RATE': 100,
                                           'REPORTING_RATE_ON_TIME': 100 / 134 * 100}[item.split('.')[-1]]]
                             for unit in units for item in items],
                    'metaData': {'items': {unit: {'displayName': 'Zone ' + unit} for unit in units}}}
        client.get.side_effect = get
        return client

    def test_partitions_units_and_datasets_without_splitting_a_dataset_or_period(self):
        params, labels = self.query(4)
        units = [f'Zone{number:07d}' for number in range(7)]
        client = self.client()
        progress = mock.Mock()
        result = fetch_reporting_chunks(client, params, labels, units, on_progress=progress)
        self.assertEqual(client.get.call_count, 6)
        self.assertEqual(progress.call_count, 6)
        self.assertEqual(result['Organisation unit ID'].tolist(), units)
        self.assertEqual(result['Dataset 0 Actual reports'].tolist(), [134] * 7)
        self.assertEqual(result['Dataset 3 Actual reports'].tolist(), [134] * 7)
        self.assertFalse(result.iloc[:, 2:].isna().any().any())
        self.assertEqual(result.attrs['reporting_status'], 'complete')
        for call in client.get.call_args_list:
            chunk = call.kwargs['params']
            for parameter in params[2:]:
                self.assertIn(parameter, chunk)
            dx = next(value[3:].split(';') for key, value in chunk if key == 'dimension' and value.startswith('dx:'))
            self.assertLessEqual(len(dx), 15)
            for dataset in {item.split('.')[0] for item in dx}:
                self.assertEqual({item.split('.')[1] for item in dx if item.startswith(dataset + '.')}, set(METRICS))
            ou = next(value[3:].split(';') for key, value in chunk if key == 'dimension' and value.startswith('ou:'))
            self.assertLessEqual(len(ou), 3)
            self.assertTrue(set(ou).issubset(units))

    def test_failed_dataset_chunk_leaves_nan_and_preserves_other_units_and_datasets(self):
        params, labels = self.query(4)

        def failure(dimensions):
            if dimensions['ou'] == ['ZoneAAAAAAA'] and dimensions['dx'][0].startswith('Dataset0003.'):
                raise requests.ReadTimeout('private password and long URL')
        client = self.client(failure)
        result = fetch_reporting_chunks(client, params, labels, ['ZoneAAAAAAA', 'ZoneBBBBBBB'], units_per_request=1)
        frame = result.set_index('Organisation unit ID')
        self.assertEqual(frame.loc['ZoneAAAAAAA', 'Dataset 0 Actual reports'], 134)
        self.assertTrue(pd.isna(frame.loc['ZoneAAAAAAA', 'Dataset 3 Actual reports']))
        self.assertEqual(frame.loc['ZoneBBBBBBB', 'Dataset 3 Actual reports'], 134)
        self.assertEqual(client.get.call_count, 4)
        self.assertEqual(result.attrs['reporting_status'], 'partial')
        self.assertEqual(result.attrs['failed_units'], ['ZoneAAAAAAA'])
        self.assertEqual(set(result.attrs['failed_metrics_by_unit']['ZoneAAAAAAA']),
                         {item for item in labels if item.startswith('Dataset0003.')})
        self.assertNotIn('private', str(result.attrs))

    def test_all_failed_is_empty_unavailable_with_numeric_schema_not_false_zeros(self):
        params, labels = self.query()
        client = mock.Mock()
        client.get.side_effect = requests.ConnectionError(
            MaxRetryError(None, 'https://secret.invalid/private', ReadTimeoutError(None, '/private', 'timed out'))
        )
        result = fetch_reporting_chunks(client, params, labels, ['ZoneAAAAAAA'])
        self.assertTrue(result.empty)
        self.assertEqual(result.attrs['reporting_status'], 'unavailable')
        self.assertEqual(client.get.call_count, 2)
        self.assertEqual(set(result.attrs['failed_metrics_by_unit']['ZoneAAAAAAA']), set(labels))
        self.assertTrue(all('Délai' in message for message in result.attrs['reporting_warnings']))
        self.assertNotIn('secret.invalid', str(result.attrs))

    def test_successful_empty_is_distinct_from_an_unavailable_response(self):
        params, labels = self.query(1)
        client = mock.Mock()
        client.get.return_value = {'rows': []}
        result = fetch_reporting_chunks(client, params, labels, ['ZoneAAAAAAA'])
        self.assertTrue(result.empty)
        self.assertEqual(result.attrs['reporting_status'], 'complete')
        self.assertEqual(result.attrs['failed_units'], [])

    def test_authorization_failure_stops_further_network_calls_without_widening_scope(self):
        params, labels = self.query()
        client = mock.Mock()
        response = requests.Response()
        response.status_code = 403
        client.get.side_effect = requests.HTTPError('password private URL', response=response)
        result = fetch_reporting_chunks(client, params, labels, ['ZoneAAAAAAA', 'ZoneBBBBBBB'], units_per_request=1)
        self.assertEqual(client.get.call_count, 1)
        self.assertEqual(result.attrs['reporting_status'], 'unavailable')
        self.assertEqual(set(result.attrs['failed_units']), {'ZoneAAAAAAA', 'ZoneBBBBBBB'})
        self.assertEqual(len(result.attrs['reporting_warnings']), 1)
        self.assertIn('403', result.attrs['reporting_warnings'][0])

    def test_out_of_scope_data_and_unrequested_metrics_are_rejected_per_chunk(self):
        params, labels = self.query(1)
        for row in [['OtherZoneAA', next(iter(labels)), 1], ['ZoneAAAAAAA', 'UnexpectedA', 1]]:
            with self.subTest(row=row):
                client = mock.Mock()
                client.get.return_value = {'headers': [{'name': name} for name in ['ou', 'dx', 'value']], 'rows': [row]}
                result = fetch_reporting_chunks(client, params, labels, ['ZoneAAAAAAA'])
                self.assertTrue(result.empty)
                self.assertEqual(result.attrs['reporting_status'], 'unavailable')

    def test_legacy_unsplit_request_and_empty_explicit_selection_are_distinct(self):
        params, labels = self.query(1)
        direct = [(key, 'ou:ZoneAAAAAAA') if value.startswith('ou:') else (key, value) for key, value in params]
        client = self.client()
        frame = fetch_reporting_chunks(client, direct, labels, None)
        self.assertEqual(len(frame), 1)
        client.get.assert_called_once_with('analytics.json', params=direct)
        client.reset_mock()
        empty = fetch_reporting_chunks(client, params, labels, [])
        self.assertTrue(empty.empty)
        self.assertEqual(empty.attrs['reporting_status'], 'unavailable')
        client.get.assert_not_called()

    def test_duplicate_unit_ids_do_not_duplicate_rows_and_invalid_ids_never_widen_selection(self):
        params, labels = self.query(1)
        frame = fetch_reporting_chunks(self.client(), params, labels, ['ZoneAAAAAAA', 'ZoneAAAAAAA'])
        self.assertEqual(frame['Organisation unit ID'].tolist(), ['ZoneAAAAAAA'])
        for value in ['', None, 'LEVEL-3;CountryAAAA', 'LEVEL-3', 'USER_ORGUNIT', 'ZoneAAAAAAA;OtherZoneAA']:
            with self.subTest(value=value), self.assertRaises(ValueError):
                fetch_reporting_chunks(self.client(), params, labels, [value])
        with self.assertRaises(ValueError):
            fetch_reporting_chunks(self.client(), params, labels, 'ZoneAAAAAAA')

    def test_error_messages_do_not_expose_exception_text_and_detect_nested_timeouts(self):
        timed_out = requests.ConnectionError(MaxRetryError(None, '/password', ReadTimeoutError(None, '/', 'private')))
        self.assertIn('Délai', format_reporting_error(timed_out))
        self.assertNotIn('password', format_reporting_error(timed_out))
        self.assertNotIn('private', format_reporting_error(ValueError('private URL')))

    def test_ninety_six_datasets_use_small_requests_and_keep_every_value(self):
        params, labels = self.query(96)

        def reject_large(dimensions):
            if len(dimensions['dx']) > 15:
                raise self.http_error(400)

        client = self.client(reject_large)
        units = ['ZoneAAAAAAA', 'ZoneBBBBBBB']
        result = fetch_reporting_chunks(client, params, labels, units, units_per_request=1, items_per_request=15)
        self.assertEqual(client.get.call_count, 64)
        self.assertEqual(result['Organisation unit ID'].tolist(), units)
        self.assertEqual(result['Dataset 95 Actual reports'].tolist(), [134, 134])
        self.assertFalse(result.iloc[:, 2:].isna().any().any())
        self.assertEqual(result.attrs['reporting_status'], 'complete')
        self.assertEqual(result.attrs['reporting_warnings'], [])
        self.assertEqual(result.attrs['failed_units'], [])

    def test_bad_request_or_long_uri_recovers_by_disjoint_complete_datasets(self):
        params, labels = self.query(4)
        for status in (400, 414):
            with self.subTest(status=status):
                def reject_large(dimensions):
                    if len(dimensions['dx']) > 5:
                        raise self.http_error(status)
                client = self.client(reject_large)
                frame = fetch_reporting_chunks(
                    client, params, labels, ['ZoneAAAAAAA'], units_per_request=1, items_per_request=20,
                )
                self.assertEqual(client.get.call_count, 7)  # 2 * 4 familles - 1.
                self.assertEqual(frame.attrs['reporting_status'], 'complete')
                self.assertEqual(frame.attrs['failed_units'], [])
                self.assertEqual(frame.attrs['reporting_warnings'], [])
                self.assertFalse(frame.iloc[:, 2:].isna().any().any())
                leaf_items = []
                for call in client.get.call_args_list:
                    query = call.kwargs['params']
                    self.assertEqual(call.args[0], 'analytics.json')
                    for parameter in params[2:]:
                        self.assertIn(parameter, query)
                    self.assertIn(('dimension', 'ou:ZoneAAAAAAA'), query)
                    items = next(value[3:].split(';') for key, value in query
                                 if key == 'dimension' and value.startswith('dx:'))
                    for dataset in {item.split('.')[0] for item in items}:
                        self.assertEqual({item.split('.')[1] for item in items if item.startswith(dataset + '.')}, set(METRICS))
                    if len(items) == 5:
                        leaf_items.extend(items)
                self.assertEqual(len(leaf_items), len(set(leaf_items)))
                self.assertEqual(set(leaf_items), set(labels))
                self.assertTrue(any(entry['Statut'] == 'Fractionné' for entry in frame.attrs['reporting_requests']))

    def test_one_rejected_dataset_keeps_all_other_cells_and_only_marks_its_metrics(self):
        params, labels = self.query(4)

        def reject_bad_dataset(dimensions):
            if dimensions['ou'] == ['ZoneAAAAAAA'] and any(item.startswith('Dataset0002.') for item in dimensions['dx']):
                raise self.http_error(400)

        client = self.client(reject_bad_dataset)
        result = fetch_reporting_chunks(client, params, labels, ['ZoneAAAAAAA', 'ZoneBBBBBBB'],
                                        units_per_request=1, items_per_request=20)
        frame = result.set_index('Organisation unit ID')
        self.assertEqual(client.get.call_count, 6)
        for number in (0, 1, 3):
            self.assertEqual(frame.loc['ZoneAAAAAAA', f'Dataset {number} Actual reports'], 134)
        self.assertTrue(pd.isna(frame.loc['ZoneAAAAAAA', 'Dataset 2 Actual reports']))
        self.assertEqual(frame.loc['ZoneBBBBBBB', 'Dataset 2 Actual reports'], 134)
        self.assertEqual(result.attrs['failed_units'], ['ZoneAAAAAAA'])
        self.assertEqual(set(result.attrs['failed_metrics_by_unit']['ZoneAAAAAAA']),
                         {item for item in labels if item.startswith('Dataset0002.')})
        self.assertEqual(result.attrs['reporting_status'], 'partial')
        self.assertEqual(len(result.attrs['reporting_warnings']), 1)
        self.assertNotIn('private', str(result.attrs))
        self.assertNotIn('password', str(result.attrs))

    def test_authorization_failure_inside_a_subdivision_stops_remaining_gets(self):
        params, labels = self.query(4)
        client = self.client()
        client.get.side_effect = [self.http_error(400), self.http_error(403)]
        result = fetch_reporting_chunks(client, params, labels, ['ZoneAAAAAAA', 'ZoneBBBBBBB'],
                                        units_per_request=1, items_per_request=20)
        self.assertEqual(client.get.call_count, 2)
        self.assertEqual(result.attrs['reporting_status'], 'unavailable')
        self.assertEqual(set(result.attrs['failed_units']), {'ZoneAAAAAAA', 'ZoneBBBBBBB'})
        self.assertEqual(len(result.attrs['reporting_warnings']), 1)
        self.assertIn('403', result.attrs['reporting_warnings'][0])
        for failed in result.attrs['failed_metrics_by_unit'].values():
            self.assertEqual(set(failed), set(labels))

    def test_timeout_or_server_failure_never_triggers_subdivision(self):
        params, labels = self.query(4)
        for error in (requests.ReadTimeout('private'), self.http_error(500), self.http_error(504)):
            with self.subTest(error=type(error).__name__):
                client = mock.Mock()
                client.get.side_effect = error
                result = fetch_reporting_chunks(client, params, labels, ['ZoneAAAAAAA'], items_per_request=20)
                self.assertEqual(client.get.call_count, 1)
                self.assertEqual(result.attrs['reporting_status'], 'unavailable')

    def test_recovered_request_keeps_legacy_initial_scope_and_rejects_out_of_scope_leaf(self):
        params, labels = self.query(2)
        params[1] = ('dimension', 'ou:ZoneAAAAAAA')
        client = self.client()
        normal = client.get.side_effect

        def get(endpoint, params):
            items = next(value[3:].split(';') for key, value in params
                         if key == 'dimension' and value.startswith('dx:'))
            if len(items) > 5:
                raise self.http_error(400)
            response = normal(endpoint, params)
            if items[0].startswith('Dataset0001.'):
                response['rows'][0][0] = 'OtherZoneAA'
            return response

        client.get.side_effect = get
        result = fetch_reporting_chunks(client, params, labels, None)
        self.assertEqual(client.get.call_count, 3)
        self.assertEqual(client.get.call_args_list[0], mock.call('analytics.json', params=params))
        self.assertEqual(result['Organisation unit ID'].tolist(), ['ZoneAAAAAAA'])
        self.assertEqual(result['Dataset 0 Actual reports'].tolist(), [134])
        self.assertTrue(result['Dataset 1 Actual reports'].isna().all())
        self.assertEqual(result.attrs['failed_units'], ['ZoneAAAAAAA'])


if __name__ == '__main__':
    unittest.main()

"""Analytics possède son délai et ne relance pas une requête volumineuse."""

from types import SimpleNamespace
import unittest
from unittest import mock

import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

from test_dashboard_categorisation import dashboard_functions


class ReportingClientTests(unittest.TestCase):
    def client(self, **kwargs):
        factory = dashboard_functions(
            '_build_dhis2_client', requests=requests, Retry=Retry, HTTPAdapter=HTTPAdapter,
        )['_build_dhis2_client']
        return factory('https://fixture.invalid/', 'fixture-user', 'fixture-password', **kwargs)

    def response(self):
        response = mock.Mock()
        response.text = '{}'
        response.json.return_value = {}
        return response

    def configuration(self, config):
        factory = mock.Mock()
        namespace = dashboard_functions(
            '_read_int_config_value', 'get_dhis2_client',
            st=SimpleNamespace(session_state={
                'dhis2_user': 'fixture-user', 'dhis2_pass': 'fixture-password',
            }),
            _read_config_value=lambda key: config.get(key), _build_dhis2_client=factory,
        )
        namespace['get_dhis2_client']()
        return factory.call_args.kwargs

    def test_default_analytics_wait_is_separate_from_metadata_and_validation(self):
        client = self.client()
        with mock.patch.object(client.session, 'request', return_value=self.response()) as request:
            client.get('analytics.json', params=[('dimension', 'ou:ZoneAAAAAAA')])
            client.get('organisationUnits')
            client.post('dataAnalysis/validationRules', data={'ou': 'ZoneAAAAAAA'})
        self.assertEqual([call.kwargs['timeout'] for call in request.call_args_list], [
            (10, 300), (10, 90), (10, 600),
        ])
        self.assertEqual(client.session.auth, ('fixture-user', 'fixture-password'))
        self.assertEqual(client.session.headers['Accept'], 'application/json')

    def test_normalized_analytics_paths_share_the_budget_and_preserve_query_parameters(self):
        client = self.client(timeout_connect=7, timeout_read=23, analytics_timeout_read=480)
        params = [('dimension', 'dx:DatasetAAAA.ACTUAL_REPORTS'), ('filter', 'pe:202609')]
        with mock.patch.object(client.session, 'request', return_value=self.response()) as request:
            for endpoint in ['analytics', 'analytics.json', '/api/analytics.json/']:
                client.get(endpoint, params=params)
        for call in request.call_args_list:
            self.assertEqual(call.kwargs['timeout'], (7, 480))
            self.assertEqual(call.kwargs['params'], params)

    def test_expensive_analytics_calls_have_no_transport_retry_but_metadata_keeps_its_policy(self):
        client = self.client(retries=2)
        for endpoint in ['analytics', 'analytics.json', 'analytics.json?dimension=ou:ZoneAAAAAAA']:
            adapter = client.session.get_adapter(f'https://fixture.invalid/api/{endpoint}')
            self.assertEqual(adapter.max_retries.total, 0)
        metadata_adapter = client.session.get_adapter('https://fixture.invalid/api/organisationUnits')
        self.assertEqual(metadata_adapter.max_retries.total, 2)
        self.assertEqual(metadata_adapter.max_retries.read, 2)

    def test_analytics_timeout_propagates_once_without_a_second_application_request(self):
        client = self.client()
        with mock.patch.object(client.session, 'request', side_effect=requests.ReadTimeout('fixture')) as request:
            with self.assertRaises(requests.ReadTimeout):
                client.get('analytics.json', params={'dimension': 'ou:ZoneAAAAAAA'})
        self.assertEqual(request.call_count, 1)

    def test_analytics_setting_is_separate_from_the_general_and_validation_settings(self):
        config = self.configuration({
            'DHIS2_URL': 'https://fixture.invalid', 'DHIS2_TIMEOUT_CONNECT': '7',
            'DHIS2_TIMEOUT_READ': '120', 'DHIS2_ANALYTICS_TIMEOUT_READ': '480',
            'DHIS2_VALIDATION_TIMEOUT_READ': '900',
        })
        self.assertEqual(config['timeout_connect'], 7)
        self.assertEqual(config['timeout_read'], 120)
        self.assertEqual(config['analytics_timeout_read'], 480)
        self.assertEqual(config['validation_timeout_read'], 900)

    def test_default_analytics_budget_preserves_an_existing_larger_general_budget(self):
        for general, expected in [(None, 300), ('90', 300), ('600', 600)]:
            with self.subTest(general=general):
                config = self.configuration({
                    'DHIS2_URL': 'https://fixture.invalid', 'DHIS2_TIMEOUT_READ': general,
                })
                self.assertEqual(config['analytics_timeout_read'], expected)

    def test_invalid_analytics_budget_uses_a_positive_default(self):
        for invalid in [None, '', 'bad', '-1', '0']:
            with self.subTest(invalid=invalid):
                config = self.configuration({
                    'DHIS2_URL': 'https://fixture.invalid', 'DHIS2_ANALYTICS_TIMEOUT_READ': invalid,
                })
                self.assertEqual(config['analytics_timeout_read'], 300)


if __name__ == '__main__':
    unittest.main()

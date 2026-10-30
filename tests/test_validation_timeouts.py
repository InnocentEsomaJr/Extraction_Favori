"""Budgets d'attente isolés pour l'analyse DHIS2, sans connexion réelle."""

from types import SimpleNamespace
import unittest
from unittest import mock

import requests

from test_dashboard_categorisation import dashboard_functions


class ValidationTimeoutTests(unittest.TestCase):
    def client(self, **kwargs):
        session = mock.Mock()
        session.request.return_value.text = '[]'
        session.request.return_value.json.return_value = []
        requests_module = mock.Mock()
        requests_module.Session.return_value = session
        retry = mock.Mock()
        factory = dashboard_functions(
            '_build_dhis2_client', requests=requests_module, Retry=retry, HTTPAdapter=mock.Mock(),
        )['_build_dhis2_client']
        return factory('https://fixture.invalid/', 'fixture-user', 'fixture-password', **kwargs), session, retry

    def test_long_analysis_wait_does_not_extend_metadata_or_analytics_requests(self):
        client, session, _ = self.client()
        client.post('dataAnalysis/validationRules', data={'ou': 'ZoneAAAAAAA'})
        client.get('organisationUnits', params={'paging': 'false'})
        client.get('analytics', params={'dimension': 'ou:ZoneAAAAAAA'})
        client.get('validationRules')
        self.assertEqual([call.kwargs['timeout'] for call in session.request.call_args_list], [
            (10, 600), (10, 90), (10, 90), (10, 90),
        ])
        self.assertEqual(client.timeout, (10, 90))
        self.assertEqual(client.validation_timeout, (10, 600))

    def test_legacy_get_analysis_uses_the_same_extended_budget(self):
        client, session, _ = self.client(timeout_connect=7, timeout_read=23, validation_timeout_read=1200)
        client.get('/api/dataAnalysis/validationRules/', params={'ou': 'ZoneAAAAAAA'})
        client.get('/api/validationRules')
        self.assertEqual([call.kwargs['timeout'] for call in session.request.call_args_list], [
            (7, 1200), (7, 23),
        ])

    def test_timed_out_post_is_not_automatically_relaunched(self):
        client, session, retry = self.client()
        session.request.side_effect = requests.ReadTimeout('fixture')
        with self.assertRaises(requests.ReadTimeout):
            client.post('dataAnalysis/validationRules', data={'ou': 'ZoneAAAAAAA'})
        self.assertEqual(session.request.call_count, 1)
        self.assertNotIn('POST', retry.call_args.kwargs['allowed_methods'])

    def configured_client(self, config):
        builder = mock.Mock()
        namespace = dashboard_functions(
            '_read_int_config_value', 'get_dhis2_client',
            st=SimpleNamespace(session_state={'dhis2_user': 'fixture-user', 'dhis2_pass': 'fixture-password'}),
            _read_config_value=lambda key: config.get(key), _build_dhis2_client=builder,
        )
        namespace['get_dhis2_client']()
        return builder.call_args.kwargs

    def test_validation_setting_is_separate_from_existing_general_timeout(self):
        kwargs = self.configured_client({
            'DHIS2_URL': 'https://fixture.invalid', 'DHIS2_TIMEOUT_READ': '120',
            'DHIS2_VALIDATION_TIMEOUT_READ': '1200',
        })
        self.assertEqual(kwargs['timeout_read'], 120)
        self.assertEqual(kwargs['validation_timeout_read'], 1200)
        self.assertEqual(kwargs['timeout_connect'], 10)

    def test_default_wait_is_ten_minutes_and_preserves_a_larger_existing_budget(self):
        for generic, expected in [(None, 600), ('120', 600), ('900', 900)]:
            with self.subTest(generic=generic):
                kwargs = self.configured_client({
                    'DHIS2_URL': 'https://fixture.invalid', 'DHIS2_TIMEOUT_READ': generic,
                })
                self.assertEqual(kwargs['validation_timeout_read'], expected)

    def test_invalid_validation_budget_uses_a_positive_default(self):
        for value in [None, '', 'bad', '-1', '0']:
            with self.subTest(value=value):
                kwargs = self.configured_client({
                    'DHIS2_URL': 'https://fixture.invalid', 'DHIS2_VALIDATION_TIMEOUT_READ': value,
                })
                self.assertEqual(kwargs['validation_timeout_read'], 600)


if __name__ == '__main__':
    unittest.main()

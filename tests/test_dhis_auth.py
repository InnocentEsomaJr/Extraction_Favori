"""La connexion DHIS2 tolère un incident bref sans divulguer les identifiants."""

import unittest
from unittest import mock

import requests

from dhis_auth import check_dhis2_credentials


class DhisAuthenticationTests(unittest.TestCase):
    def response(self, status):
        response = mock.Mock()
        response.status_code = status
        response.text = 'fixture-password fixture-user secret server body'
        return response

    def check(self, get, **kwargs):
        return check_dhis2_credentials(
            'https://fixture.invalid/', 'fixture-user', 'fixture-password', request_get=get, **kwargs
        )

    def assert_safe(self, result):
        for secret in ['fixture-password', 'fixture-user', 'secret server body', 'https://fixture.invalid']:
            self.assertNotIn(secret, str(result))

    def test_success_uses_separate_connect_and_read_timeouts(self):
        get = mock.Mock(return_value=self.response(200))
        self.assertEqual(self.check(get), (True, None))
        get.assert_called_once_with(
            'https://fixture.invalid/api/me', auth=('fixture-user', 'fixture-password'), timeout=(10, 60)
        )

    def test_read_timeout_is_retried_once_and_a_later_success_logs_in(self):
        get = mock.Mock(side_effect=[requests.ReadTimeout('fixture-password'), self.response(200)])
        self.assertEqual(self.check(get), (True, None))
        self.assertEqual(get.call_count, 2)

    def test_connection_error_is_retried_once_then_reported_without_exception_text(self):
        for error in [requests.ConnectionError, requests.ConnectTimeout]:
            with self.subTest(error=error):
                get = mock.Mock(side_effect=error('fixture-user fixture-password https://fixture.invalid'))
                result = self.check(get)
                self.assertFalse(result[0])
                self.assertIn('Connexion au serveur', result[1])
                self.assertEqual(get.call_count, 2)
                self.assert_safe(result)

    def test_repeated_read_timeout_explains_the_wait_and_respects_custom_timeouts(self):
        get = mock.Mock(side_effect=requests.ReadTimeout('fixture-password'))
        result = self.check(get, timeout_connect=15, timeout_read=90)
        self.assertFalse(result[0])
        self.assertIn('90 secondes', result[1])
        self.assertEqual(get.call_count, 2)
        self.assertEqual(get.call_args.kwargs['timeout'], (15, 90))
        self.assert_safe(result)

    def test_gateway_errors_allow_one_retry_and_preserve_the_final_http_code(self):
        for status in [502, 503, 504]:
            with self.subTest(status=status):
                get = mock.Mock(return_value=self.response(status))
                result = self.check(get)
                self.assertFalse(result[0])
                self.assertIn(f'HTTP {status}', result[1])
                self.assertEqual(get.call_count, 2)
                self.assert_safe(result)
                get = mock.Mock(side_effect=[self.response(status), self.response(200)])
                self.assertEqual(self.check(get), (True, None))
                self.assertEqual(get.call_count, 2)

    def test_unauthorized_and_forbidden_are_not_retried(self):
        for status in [401, 403]:
            with self.subTest(status=status):
                get = mock.Mock(return_value=self.response(status))
                result = self.check(get)
                self.assertFalse(result[0])
                self.assertIn('identifiants', result[1])
                self.assertEqual(get.call_count, 1)
                self.assert_safe(result)

    def test_tls_error_is_not_retried_or_exposed_as_bad_credentials(self):
        get = mock.Mock(side_effect=requests.exceptions.SSLError('fixture-password private-certificate'))
        result = self.check(get)
        self.assertFalse(result[0])
        self.assertIn('certificat TLS', result[1])
        self.assertEqual(get.call_count, 1)
        self.assert_safe(result)
        self.assertNotIn('private-certificate', result[1])

    def test_other_http_errors_are_not_retried(self):
        get = mock.Mock(return_value=self.response(500))
        result = self.check(get)
        self.assertFalse(result[0])
        self.assertIn('HTTP 500', result[1])
        self.assertEqual(get.call_count, 1)
        self.assert_safe(result)

    def test_non_transient_request_errors_never_expose_exception_text(self):
        for error in [requests.RequestException, requests.exceptions.InvalidURL, RuntimeError]:
            with self.subTest(error=error):
                get = mock.Mock(side_effect=error('fixture-password https://fixture.invalid'))
                result = self.check(get)
                self.assertFalse(result[0])
                self.assertEqual(get.call_count, 1)
                self.assert_safe(result)

    def test_retry_configuration_is_bounded_and_can_disable_retries(self):
        for retries, count in [(0, 1), (-1, 1), (1, 2), (10, 2), ('invalid', 2)]:
            with self.subTest(retries=retries):
                get = mock.Mock(side_effect=requests.ReadTimeout('fixture-password'))
                self.assertFalse(self.check(get, retries=retries)[0])
                self.assertEqual(get.call_count, count)

    def test_invalid_timeout_configuration_uses_defaults(self):
        for invalid in [None, 0, -1, 'invalid', float('inf'), float('nan')]:
            with self.subTest(invalid=invalid):
                get = mock.Mock(return_value=self.response(200))
                self.assertEqual(self.check(get, timeout_connect=invalid, timeout_read=invalid), (True, None))
                self.assertEqual(get.call_args.kwargs['timeout'], (10, 60))

    def test_missing_configuration_or_credentials_do_not_make_a_request(self):
        get = mock.Mock()
        for url, username, password in [('', 'user', 'pass'), ('https://fixture.invalid', '', 'pass'), ('https://fixture.invalid', 'user', '')]:
            with self.subTest(url=url, username=username):
                result = check_dhis2_credentials(url, username, password, request_get=get)
                self.assertFalse(result[0])
        get.assert_not_called()


if __name__ == '__main__':
    unittest.main()

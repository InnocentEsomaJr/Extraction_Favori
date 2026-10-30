import unittest
from datetime import date, datetime
from unittest.mock import MagicMock, Mock, patch

import requests

from validation_api import fetch_validation_analysis


class FakeClient:
    def __init__(self, responses):
        self.responses = responses
        self.calls = []
        self.methods = []

    def _call(self, method, endpoint, params):
        self.methods.append(method)
        self.calls.append((endpoint, params.copy()))
        response = self.responses[len(self.calls) - 1]
        if isinstance(response, Exception):
            raise response
        return response

    def post(self, endpoint, data=None):
        return self._call('POST', endpoint, data)

    def get(self, endpoint, params=None):
        return self._call('GET', endpoint, params)


def http_error(status):
    response = requests.Response()
    response.status_code = status
    return requests.HTTPError('contenu sensible à ne jamais afficher', response=response)


def violation(ou="LeafAAAAAAA", period="202402", rule="RuleAAAAAAA", **changes):
    row = {
        "organisationUnitId": ou, "organisationUnitDisplayName": "Aire test",
        "organisationUnitPath": f"/CountryAAAA/ProvinceAAA/ZoneAAAAAAA/{ou}",
        "validationRuleId": rule, "validationRuleDescription": "Règle test",
        "periodId": period, "periodDisplayName": "Février 2024",
        "attributeOptionComboId": "AocAAAAAAAA", "importance": "MEDIUM",
        "leftSideValue": 0, "operator": "<=", "rightSideValue": 5,
    }
    row.update(changes)
    return row


class ValidationAnalysisTests(unittest.TestCase):
    def test_exact_dates_use_one_post_for_the_complete_interval(self):
        client = FakeClient([[]])
        result = fetch_validation_analysis(
            client, "ZoneAAAAAAA", ["202606", "202607", "202608", "202609", "202610", "202611"],
            start_date=date(2026, 7, 27), end_date="2026-10-27",
        )
        self.assertEqual([(params["startDate"], params["endDate"]) for _, params in client.calls], [
            ("2026-07-27", "2026-10-27"),
        ])
        self.assertEqual(client.methods, ['POST'])
        self.assertNotIn('maxResults', client.calls[0][1])
        self.assertNotIn('vrg', client.calls[0][1])
        self.assertEqual(result.attrs["start_date"], "2026-07-27")
        self.assertEqual(result.attrs["end_date"], "2026-10-27")
        self.assertEqual(result.attrs["failed_periods"], [])

    def test_exact_dates_can_derive_months_cross_year(self):
        client = FakeClient([[]])
        fetch_validation_analysis(client, "ZoneAAAAAAA", None,
                                  start_date="2025-12-31", end_date=datetime(2026, 1, 1, 12, 30))
        self.assertEqual([(params["startDate"], params["endDate"]) for _, params in client.calls],
                         [("2025-12-31", "2026-01-01")])

    def test_same_day_leap_date_is_sent_unchanged_without_refiltering_server_periods(self):
        client = FakeClient([[violation(period="20240228"), violation(period="20240229"), violation(period="20240301")]])
        result = fetch_validation_analysis(client, "ZoneAAAAAAA", ["202402"],
                                           start_date="2024-02-29", end_date=date(2024, 2, 29))
        self.assertEqual(client.calls[0][1]["startDate"], "2024-02-29")
        self.assertEqual(client.calls[0][1]["endDate"], "2024-02-29")
        self.assertEqual(result["periodId"].tolist(), ["20240228", "20240229", "20240301"])
        self.assertEqual(result.attrs["failed_periods"], [])

    def test_server_periods_are_preserved_as_in_the_official_application(self):
        client = FakeClient([[violation(period="202607"), violation(period="20260726"),
                              violation(period="20260727"), violation(period="2026W30"),
                              violation(period="2026W31")]])
        result = fetch_validation_analysis(client, "ZoneAAAAAAA", ["202607"],
                                           start_date="2026-07-27", end_date="2026-07-31")
        self.assertEqual(result["periodId"].tolist(), ["202607", "20260726", "20260727", "2026W30", "2026W31"])

    def test_invalid_explicit_dates_fail_before_any_network_request(self):
        for start, end in [("2026-10-27", None), (None, "2026-10-27"),
                           ("2026-10-28", "2026-10-27"), ("bad", "2026-10-27"),
                           ("2026-10-27", "2026-02-30")]:
            with self.subTest(start=start, end=end):
                client = FakeClient([])
                with self.assertRaises(ValueError):
                    fetch_validation_analysis(client, "ZoneAAAAAAA", ["202610"], start_date=start, end_date=end)
                self.assertEqual(client.calls, [])

    def test_post_parameters_use_one_calendar_interval_and_disable_persistence_and_notifications(self):
        client = FakeClient([[]])
        result = fetch_validation_analysis(client, "ZoneAAAAAAA", ["202402", "202501"], "GroupAAAAAA", limit=123)
        self.assertTrue(result.empty)
        self.assertEqual(result.attrs["failed_periods"], [])
        self.assertEqual([call[0] for call in client.calls], ["dataAnalysis/validationRules"])
        self.assertEqual(client.methods, ['POST'])
        self.assertEqual(client.calls[0][1], {
            "ou": "ZoneAAAAAAA", "startDate": "2024-02-01", "endDate": "2025-01-31",
            "persist": False, "notification": False, "maxResults": 123, "vrg": "GroupAAAAAA",
        })
        self.assertEqual(result.attrs['request_method'], 'POST')
        self.assertEqual(result.attrs['request_params'], client.calls[0][1])

    def test_flat_fields_remain_available_with_legacy_aliases_and_zero_values(self):
        row = violation()
        result = fetch_validation_analysis(FakeClient([[row]]), "ZoneAAAAAAA", ["202402"])
        actual = result.iloc[0]
        self.assertEqual(actual["organisationUnit"]["id"], "LeafAAAAAAA")
        self.assertEqual(actual["validationRule"]["id"], "RuleAAAAAAA")
        self.assertEqual(actual["period"]["id"], "202402")
        self.assertEqual(actual["leftSideValue"], 0)
        self.assertEqual(actual["leftsideValue"], 0)
        self.assertEqual(actual["rightsideValue"], 5)
        self.assertEqual(actual["organisationUnitPath"], row["organisationUnitPath"])

    def test_explicit_out_of_scope_path_is_rejected_and_marks_the_interval_incomplete(self):
        other_scope = violation(ou="OtherAAAAAA", organisationUnitPath="/CountryAAAA/OtherZoneAA/OtherAAAAAA")
        response = [violation(), other_scope, violation(period="202401"), violation(period="2024Q1", rule="RuleBBBBBBB")]
        result = fetch_validation_analysis(FakeClient([response]), "ZoneAAAAAAA", ["202402"])
        self.assertEqual(result["periodId"].tolist(), ["202402", "202401", "2024Q1"])
        self.assertEqual(result.attrs["failed_periods"], ['202402'])
        self.assertIn('chemin', ' '.join(result.attrs['errors']['202402']))

    def test_scope_matching_uses_complete_path_segments(self):
        response = [violation(organisationUnitPath="/CountryAAAA/ZoneAAAAAAA2/LeafAAAAAAA")]
        result = fetch_validation_analysis(FakeClient([response]), "ZoneAAAAAAA", ["202402"])
        self.assertTrue(result.empty)
        self.assertEqual(result.attrs["failed_periods"], ['202402'])

    def test_direct_unit_can_be_verified_without_descendant_path(self):
        result = fetch_validation_analysis(FakeClient([[violation(ou="ZoneAAAAAAA", organisationUnitPath="")]]), "ZoneAAAAAAA", "202402")
        self.assertEqual(len(result), 1)
        self.assertEqual(result.attrs["failed_periods"], [])

    def test_missing_required_identity_is_incomplete_instead_of_a_healthy_zero(self):
        for changes in [{"periodId": None}, {"validationRuleId": None}, {"organisationUnitId": None}]:
            with self.subTest(changes=changes):
                result = fetch_validation_analysis(FakeClient([[violation(**changes)]]), "ZoneAAAAAAA", ["202402"])
                self.assertTrue(result.empty)
                self.assertEqual(result.attrs["failed_periods"], ["202402"])

    def test_cap_is_checked_before_scope_filtering_and_retains_available_lines(self):
        result = fetch_validation_analysis(FakeClient([[violation(), violation(rule="RuleBBBBBBB")]]), "ZoneAAAAAAA", ["202402"], limit=2)
        self.assertEqual(len(result), 2)
        self.assertEqual(result.attrs["truncated_periods"], ["202402"])
        self.assertEqual(result.attrs["failed_periods"], ["202402"])

    def test_network_and_malformed_response_failures_mark_the_whole_interval(self):
        for response in [RuntimeError('fixture réseau'), {'unexpected': []}, None, '<html>login</html>']:
            with self.subTest(response=type(response).__name__):
                client = FakeClient([response])
                result = fetch_validation_analysis(client, 'ZoneAAAAAAA', ['202402', '202403', '202404'])
                self.assertTrue(result.empty)
                self.assertEqual(result.attrs['failed_periods'], ['202402', '202403', '202404'])
                self.assertEqual(result.attrs['source'], 'calculated')
                self.assertEqual(len(client.calls), 1)
                self.assertIn('POST /api/dataAnalysis/validationRules', result.attrs['errors']['202402'][0])

    def test_supported_envelopes_and_nested_result_aliases(self):
        row = {"organisationUnit": {"id": "ZoneAAAAAAA"}, "period": {"id": "202402"},
               "validationRule": {"id": "RuleAAAAAAA"}, "leftsideValue": 0, "rightsideValue": 2}
        for envelope in ["validationResults", "validationRuleViolations"]:
            with self.subTest(envelope=envelope):
                result = fetch_validation_analysis(FakeClient([{envelope: [row]}]), "ZoneAAAAAAA", ["202402"])
                self.assertEqual(len(result), 1)
                self.assertEqual(result.attrs["failed_periods"], [])

    def test_legacy_period_iso_precedes_numeric_database_identifier(self):
        for iso_field in ["iso", "isoDate"]:
            with self.subTest(iso_field=iso_field):
                row = {"organisationUnit": {"id": "ZoneAAAAAAA"},
                       "period": {"id": 12345, iso_field: "202609"},
                       "validationRule": {"id": "RuleAAAAAAA"}}
                result = fetch_validation_analysis(FakeClient([[row]]), "ZoneAAAAAAA", ["202609"])
                self.assertEqual(len(result), 1)
                self.assertEqual(result.iloc[0]["period"]["id"], "202609")
                self.assertEqual(result.attrs["failed_periods"], [])

    def test_flat_period_id_precedes_legacy_iso_period(self):
        row = {"organisationUnit": {"id": "ZoneAAAAAAA"}, "periodId": "202609",
               "period": {"id": 12345, "iso": "202608", "isoDate": "202607"},
               "validationRule": {"id": "RuleAAAAAAA"}}
        result = fetch_validation_analysis(FakeClient([[row]]), "ZoneAAAAAAA", ["202609"])
        self.assertEqual(len(result), 1)
        self.assertEqual(result.iloc[0]["periodId"], "202609")
        self.assertEqual(result.iloc[0]["period"]["id"], "202609")
        self.assertEqual(result.attrs["failed_periods"], [])

    def test_quarterly_result_is_requested_once_over_the_complete_interval(self):
        row = violation(period="2024Q1")
        client = FakeClient([[row]])
        result = fetch_validation_analysis(client, "ZoneAAAAAAA", ["202401", "202402"])
        self.assertEqual(len(result), 1)
        self.assertEqual(len(client.calls), 1)

    def test_server_order_and_duplicate_observations_are_preserved(self):
        first = violation(period="2024Q1", leftSideValue=0, rightSideValue=5)
        duplicate = violation(period="2024Q1", leftSideValue="0.0", rightSideValue="5.00")
        changed = violation(period="2024Q1", leftSideValue=1, rightSideValue=5)
        missing = violation(period="2024Q1", leftSideValue=None, rightSideValue=5)
        result = fetch_validation_analysis(FakeClient([[first, duplicate, changed, missing]]), "ZoneAAAAAAA", ["202401", "202402"])
        self.assertEqual(len(result), 4)
        self.assertEqual(result.iloc[0]["leftSideValue"], 0)
        self.assertEqual(result.iloc[1]["leftSideValue"], '0.0')
        self.assertEqual(result.iloc[2]["leftSideValue"], 1)
        self.assertTrue(result["leftSideValue"].isna().iloc[3])

    def test_flat_alias_identifiers_override_conflicting_nested_ids_and_keep_metadata(self):
        row = violation()
        row.update(
            organisationUnit={"id": "WrongUnitAA", "path": "/old/path", "code": "retained"},
            validationRule={"id": "WrongRuleAA", "displayName": "Ancien libellé", "description": "Métadonnée conservée"},
            period={"id": "202401", "startDate": "2024-02-01"},
            attributeOptionCombo={"id": "WrongAocAAA", "code": "combo-code"},
        )
        actual = fetch_validation_analysis(FakeClient([[row]]), "ZoneAAAAAAA", ["202402"]).iloc[0]
        self.assertEqual(actual["organisationUnit"]["id"], row["organisationUnitId"])
        self.assertEqual(actual["organisationUnit"]["displayName"], row["organisationUnitDisplayName"])
        self.assertEqual(actual["organisationUnit"]["code"], "retained")
        self.assertEqual(actual["validationRule"]["id"], row["validationRuleId"])
        self.assertEqual(actual["validationRule"]["displayName"], row["validationRuleDescription"])
        self.assertEqual(actual["validationRule"]["description"], "Métadonnée conservée")
        self.assertEqual(actual["period"]["id"], "202402")
        self.assertEqual(actual["period"]["startDate"], "2024-02-01")
        self.assertEqual(actual["attributeOptionCombo"]["id"], row["attributeOptionComboId"])
        self.assertEqual(actual["attributeOptionCombo"]["code"], "combo-code")

    def test_invalid_requests_never_call_client(self):
        client = FakeClient([])
        result = fetch_validation_analysis(client, "ZoneAAAAAAA", ["202413", "2024Q1", "bad"])
        self.assertEqual(result.attrs["failed_periods"], ["202413", "2024Q1", "bad"])
        result = fetch_validation_analysis(client, None, ["202402"])
        self.assertEqual(result.attrs["failed_periods"], ["202402"])
        self.assertEqual(client.calls, [])
        with self.assertRaises(ValueError):
            fetch_validation_analysis(client, "ZoneAAAAAAA", ["202402"], limit=0)

    def test_descendant_without_path_is_accepted_like_the_official_ui(self):
        row = violation()
        row.pop('organisationUnitPath')
        result = fetch_validation_analysis(FakeClient([[row]]), 'ZoneAAAAAAA', ['202402'])
        self.assertEqual(len(result), 1)
        self.assertEqual(result.attrs['failed_periods'], [])

    def test_malformed_record_keeps_available_rows_without_claiming_completeness(self):
        result = fetch_validation_analysis(FakeClient([[violation(), 'invalid record']]), 'ZoneAAAAAAA', ['202402'])
        self.assertEqual(len(result), 1)
        self.assertEqual(result.attrs['failed_periods'], ['202402'])

    def test_default_limit_is_not_sent_but_warns_at_the_ui_threshold(self):
        client = FakeClient([[violation(rule=f'Rule{number:07d}') for number in range(500)]])
        result = fetch_validation_analysis(client, 'ZoneAAAAAAA', ['202402', '202403'])
        self.assertNotIn('maxResults', client.calls[0][1])
        self.assertEqual(result.attrs['max_results'], 500)
        self.assertEqual(len(result), 500)
        self.assertEqual(result.attrs['failed_periods'], ['202402', '202403'])
        self.assertEqual(result.attrs['truncated_periods'], ['202402', '202403'])

    def test_get_fallback_is_only_used_for_explicitly_unsupported_post(self):
        for status in [404, 405, 501]:
            with self.subTest(status=status):
                client = FakeClient([http_error(status), [violation()]])
                result = fetch_validation_analysis(client, 'ZoneAAAAAAA', ['202402'], 'GroupAAAAAA')
                self.assertEqual(client.methods, ['POST', 'GET'])
                self.assertIs(client.calls[0][1]['persist'], False)
                self.assertIs(client.calls[0][1]['notification'], False)
                self.assertEqual(client.calls[1][1]['persist'], 'false')
                self.assertEqual(client.calls[1][1]['notification'], 'false')
                self.assertEqual(client.calls[1][1]['vrg'], 'GroupAAAAAA')
                self.assertNotIn('maxResults', client.calls[1][1])
                self.assertEqual(result.attrs['request_method'], 'GET')
                self.assertEqual(result.attrs['request_params'], client.calls[1][1])
                self.assertEqual(result.attrs['failed_periods'], [])
                self.assertEqual(len(result), 1)

    def test_auth_permission_and_server_errors_do_not_trigger_get_fallback(self):
        for status in [400, 401, 403, 422, 429, 500, 502, 503]:
            with self.subTest(status=status):
                client = FakeClient([http_error(status)])
                result = fetch_validation_analysis(client, 'ZoneAAAAAAA', ['202402'])
                self.assertEqual(client.methods, ['POST'])
                self.assertEqual(result.attrs['failed_periods'], ['202402'])
                message = result.attrs['errors']['202402'][0]
                self.assertIn(f'HTTP {status}', message)
                self.assertIn('POST /api/dataAnalysis/validationRules', message)
                self.assertNotIn('contenu sensible', message)

    def test_failed_get_fallback_reports_its_actual_method_and_status(self):
        client = FakeClient([http_error(405), http_error(403)])
        result = fetch_validation_analysis(client, 'ZoneAAAAAAA', ['202402', '202403'])
        self.assertEqual(client.methods, ['POST', 'GET'])
        self.assertEqual(result.attrs['request_method'], 'GET')
        self.assertEqual(result.attrs['failed_periods'], ['202402', '202403'])
        self.assertIn('GET /api/dataAnalysis/validationRules — HTTP 403', result.attrs['errors']['202402'][0])

    def test_timeout_diagnostic_never_exposes_exception_content(self):
        client = FakeClient([requests.ReadTimeout('https://user:password@example.invalid?secret=fixture')])
        result = fetch_validation_analysis(client, 'ZoneAAAAAAA', ['202402'])
        message = result.attrs['errors']['202402'][0]
        self.assertIn('délai d’attente dépassé', message)
        self.assertNotIn('password', message)
        self.assertNotIn('secret', message)
        self.assertEqual(client.methods, ['POST'])

    def test_elapsed_time_and_validation_read_limit_are_recorded_without_changing_the_request(self):
        client = FakeClient([[violation()]])
        client.timeout = (10, 90)
        client.validation_timeout = (10, 600)
        with patch('validation_api.perf_counter', side_effect=[100.0, 101.25]):
            result = fetch_validation_analysis(client, 'ZoneAAAAAAA', ['202402'])
        self.assertEqual(result.attrs['elapsed_seconds'], 1.25)
        self.assertEqual(result.attrs['validation_read_timeout_seconds'], 600)
        self.assertEqual(client.methods, ['POST'])
        self.assertEqual(client.calls[0][1], {
            'ou': 'ZoneAAAAAAA', 'startDate': '2024-02-01', 'endDate': '2024-02-29',
            'persist': False, 'notification': False,
        })

    def test_read_timeout_reports_configured_limit_and_elapsed_time_without_retrying(self):
        client = FakeClient([requests.ReadTimeout('url password=fixture secret=fixture')])
        client.validation_timeout = (10, 600)
        with patch('validation_api.perf_counter', side_effect=[100.0, 700.5]):
            result = fetch_validation_analysis(client, 'ZoneAAAAAAA', ['202402'])
        message = result.attrs['errors']['202402'][0]
        self.assertIn('lecture de la réponse DHIS2', message)
        self.assertIn('limite 600 s', message)
        self.assertNotIn('password', message)
        self.assertNotIn('secret', message)
        self.assertEqual(result.attrs['elapsed_seconds'], 600.5)
        self.assertEqual(result.attrs['validation_read_timeout_seconds'], 600)
        self.assertEqual(client.methods, ['POST'])
        self.assertEqual(result.attrs['failed_periods'], ['202402'])

    def test_connect_timeout_is_distinguished_from_waiting_for_the_analysis_response(self):
        client = FakeClient([requests.ConnectTimeout('private connection info')])
        client.validation_timeout = (10, 600)
        result = fetch_validation_analysis(client, 'ZoneAAAAAAA', ['202402'])
        message = result.attrs['errors']['202402'][0]
        self.assertIn('délai d’attente dépassé à la connexion', message)
        self.assertNotIn('lecture', message)
        self.assertNotIn('600', message)
        self.assertNotIn('private', message)
        self.assertEqual(client.methods, ['POST'])

    def test_generic_timeouts_keep_a_safe_waiting_diagnostic_without_method_fallback(self):
        for error in [requests.Timeout('private timeout info'), TimeoutError('private timeout info')]:
            with self.subTest(error_type=type(error).__name__):
                client = FakeClient([error])
                result = fetch_validation_analysis(client, 'ZoneAAAAAAA', ['202402'])
                message = result.attrs['errors']['202402'][0]
                self.assertIn('délai d’attente dépassé', message)
                self.assertNotIn('private', message)
                self.assertEqual(client.methods, ['POST'])

    def test_gateway_failures_identify_the_http_relay_without_retrying(self):
        for status in [502, 504, 524]:
            with self.subTest(status=status):
                client = FakeClient([http_error(status)])
                result = fetch_validation_analysis(client, 'ZoneAAAAAAA', ['202402'])
                message = result.attrs['errors']['202402'][0]
                self.assertIn(f'HTTP {status}', message)
                self.assertIn('relais HTTP de SNIS', message)
                self.assertNotIn('contenu sensible', message)
                self.assertEqual(client.methods, ['POST'])

    def test_read_limit_falls_back_to_the_generic_tuple_when_validation_config_is_invalid(self):
        invalid_configs = [
            None, Mock(), {'read': 600}, (), (10,), (10, 600, 700),
            (10, Mock()), (10, MagicMock(spec=float)), (10, {'read': 600}),
            (10, True), (10, 0), (10, -1), (10, float('nan')), (10, float('inf')), (10, '600'),
        ]
        for configured in invalid_configs:
            with self.subTest(config_type=type(configured).__name__):
                client = FakeClient([[]])
                client.validation_timeout = configured
                client.timeout = (10, 125.5)
                result = fetch_validation_analysis(client, 'ZoneAAAAAAA', ['202402'])
                self.assertEqual(result.attrs['validation_read_timeout_seconds'], 125.5)

    def test_missing_or_mocked_read_limit_is_unavailable_and_never_printed(self):
        client = FakeClient([requests.ReadTimeout('private')])
        client.validation_timeout = Mock()
        client.timeout = {'read': 600}
        result = fetch_validation_analysis(client, 'ZoneAAAAAAA', ['202402'])
        self.assertIsNone(result.attrs['validation_read_timeout_seconds'])
        message = result.attrs['errors']['202402'][0]
        self.assertIn('lecture de la réponse DHIS2', message)
        self.assertNotIn('limite', message)
        self.assertNotIn('Mock', message)
        self.assertNotIn('private', message)

    def test_fiscal_period_returned_by_dhis2_is_preserved_without_local_calendar_parsing(self):
        result = fetch_validation_analysis(FakeClient([[violation(period='2026April')]]), 'ZoneAAAAAAA', ['202609'])
        self.assertEqual(result['periodId'].tolist(), ['2026April'])
        self.assertEqual(result.attrs['failed_periods'], [])

    def test_non_json_response_explains_format_failure_without_response_content(self):
        client = FakeClient([ValueError('<html>connexion password=fixture</html>')])
        result = fetch_validation_analysis(client, 'ZoneAAAAAAA', ['202609'])
        self.assertTrue(result.empty)
        self.assertEqual(result.attrs['failed_periods'], ['202609'])
        message = result.attrs['errors']['202609'][0]
        self.assertIn('réponse non JSON', message)
        self.assertNotIn('password', message)
        self.assertEqual(client.methods, ['POST'])


if __name__ == "__main__":
    unittest.main()

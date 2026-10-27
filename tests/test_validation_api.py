import unittest
from datetime import date, datetime

from validation_api import fetch_validation_analysis


class FakeClient:
    def __init__(self, responses):
        self.responses = responses
        self.calls = []

    def get(self, endpoint, params=None):
        self.calls.append((endpoint, params.copy()))
        response = self.responses[len(self.calls) - 1]
        if isinstance(response, Exception):
            raise response
        return response


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
    def test_exact_dates_clip_each_month_and_skip_explicit_months_outside_range(self):
        client = FakeClient([[], [], [], []])
        result = fetch_validation_analysis(
            client, "ZoneAAAAAAA", ["202606", "202607", "202608", "202609", "202610", "202611"],
            start_date=date(2026, 7, 27), end_date="2026-10-27",
        )
        self.assertEqual([(params["startDate"], params["endDate"]) for _, params in client.calls], [
            ("2026-07-27", "2026-07-31"), ("2026-08-01", "2026-08-31"),
            ("2026-09-01", "2026-09-30"), ("2026-10-01", "2026-10-27"),
        ])
        self.assertEqual(result.attrs["start_date"], "2026-07-27")
        self.assertEqual(result.attrs["end_date"], "2026-10-27")
        self.assertEqual(result.attrs["failed_periods"], [])

    def test_exact_dates_can_derive_months_cross_year(self):
        client = FakeClient([[], []])
        fetch_validation_analysis(client, "ZoneAAAAAAA", None,
                                  start_date="2025-12-31", end_date=datetime(2026, 1, 1, 12, 30))
        self.assertEqual([(params["startDate"], params["endDate"]) for _, params in client.calls],
                         [("2025-12-31", "2025-12-31"), ("2026-01-01", "2026-01-01")])

    def test_same_day_leap_date_uses_exact_day_and_filters_daily_results(self):
        client = FakeClient([[violation(period="20240228"), violation(period="20240229"), violation(period="20240301")]])
        result = fetch_validation_analysis(client, "ZoneAAAAAAA", ["202402"],
                                           start_date="2024-02-29", end_date=date(2024, 2, 29))
        self.assertEqual(client.calls[0][1]["startDate"], "2024-02-29")
        self.assertEqual(client.calls[0][1]["endDate"], "2024-02-29")
        self.assertEqual(result["periodId"].tolist(), ["20240229"])
        self.assertEqual(result.attrs["failed_periods"], [])

    def test_monthly_and_weekly_period_overlap_uses_clipped_dates(self):
        client = FakeClient([[violation(period="202607"), violation(period="20260726"),
                              violation(period="20260727"), violation(period="2026W30"),
                              violation(period="2026W31")]])
        result = fetch_validation_analysis(client, "ZoneAAAAAAA", ["202607"],
                                           start_date="2026-07-27", end_date="2026-07-31")
        self.assertEqual(result["periodId"].tolist(), ["202607", "20260727", "2026W31"])

    def test_invalid_explicit_dates_fail_before_any_network_request(self):
        for start, end in [("2026-10-27", None), (None, "2026-10-27"),
                           ("2026-10-28", "2026-10-27"), ("bad", "2026-10-27"),
                           ("2026-10-27", "2026-02-30")]:
            with self.subTest(start=start, end=end):
                client = FakeClient([])
                with self.assertRaises(ValueError):
                    fetch_validation_analysis(client, "ZoneAAAAAAA", ["202610"], start_date=start, end_date=end)
                self.assertEqual(client.calls, [])

    def test_get_parameters_use_calendar_dates_and_disable_all_writes_and_notifications(self):
        client = FakeClient([[], []])
        result = fetch_validation_analysis(client, "ZoneAAAAAAA", ["202402", "202501"], "GroupAAAAAA", limit=123)
        self.assertTrue(result.empty)
        self.assertEqual(result.attrs["failed_periods"], [])
        self.assertEqual([call[0] for call in client.calls], ["dataAnalysis/validationRules"] * 2)
        self.assertEqual(client.calls[0][1], {
            "ou": "ZoneAAAAAAA", "startDate": "2024-02-01", "endDate": "2024-02-29",
            "persist": "false", "notification": "false", "maxResults": 123, "vrg": "GroupAAAAAA",
        })
        self.assertEqual(client.calls[1][1]["endDate"], "2025-01-31")

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

    def test_only_selected_scope_and_overlapping_periods_are_kept(self):
        other_scope = violation(ou="OtherAAAAAA", organisationUnitPath="/CountryAAAA/OtherZoneAA/OtherAAAAAA")
        response = [violation(), other_scope, violation(period="202401"), violation(period="2024Q1", rule="RuleBBBBBBB")]
        result = fetch_validation_analysis(FakeClient([response]), "ZoneAAAAAAA", ["202402"])
        self.assertEqual(result["periodId"].tolist(), ["202402", "2024Q1"])
        self.assertEqual(result.attrs["failed_periods"], [])

    def test_scope_matching_uses_complete_path_segments(self):
        response = [violation(organisationUnitPath="/CountryAAAA/ZoneAAAAAAA2/LeafAAAAAAA")]
        result = fetch_validation_analysis(FakeClient([response]), "ZoneAAAAAAA", ["202402"])
        self.assertTrue(result.empty)
        self.assertEqual(result.attrs["failed_periods"], [])

    def test_direct_unit_can_be_verified_without_descendant_path(self):
        result = fetch_validation_analysis(FakeClient([[violation(ou="ZoneAAAAAAA", organisationUnitPath="")]]), "ZoneAAAAAAA", "202402")
        self.assertEqual(len(result), 1)
        self.assertEqual(result.attrs["failed_periods"], [])

    def test_missing_identity_period_or_path_is_incomplete_instead_of_a_healthy_zero(self):
        for changes in [{"periodId": None}, {"organisationUnitPath": ""}, {"validationRuleId": None}, {"organisationUnitId": None}]:
            with self.subTest(changes=changes):
                result = fetch_validation_analysis(FakeClient([[violation(**changes)]]), "ZoneAAAAAAA", ["202402"])
                self.assertTrue(result.empty)
                self.assertEqual(result.attrs["failed_periods"], ["202402"])

    def test_cap_is_checked_before_scope_filtering_and_retains_available_lines(self):
        result = fetch_validation_analysis(FakeClient([[violation(), violation(rule="RuleBBBBBBB")]]), "ZoneAAAAAAA", ["202402"], limit=2)
        self.assertEqual(len(result), 2)
        self.assertEqual(result.attrs["truncated_periods"], ["202402"])
        self.assertEqual(result.attrs["failed_periods"], ["202402"])

    def test_network_and_malformed_response_failures_preserve_other_months(self):
        client = FakeClient([[violation()], RuntimeError("fixture réseau"), {"unexpected": []}])
        result = fetch_validation_analysis(client, "ZoneAAAAAAA", ["202402", "202403", "202404"])
        self.assertEqual(len(result), 1)
        self.assertEqual(result.attrs["failed_periods"], ["202403", "202404"])
        self.assertEqual(result.attrs["source"], "calculated")

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

    def test_repeated_quarterly_result_is_not_duplicated_across_requested_months(self):
        row = violation(period="2024Q1")
        result = fetch_validation_analysis(FakeClient([[row], [row]]), "ZoneAAAAAAA", ["202401", "202402"])
        self.assertEqual(len(result), 1)

    def test_same_identity_keeps_distinct_observations_but_normalizes_numeric_duplicates(self):
        first = violation(period="2024Q1", leftSideValue=0, rightSideValue=5)
        duplicate = violation(period="2024Q1", leftSideValue="0.0", rightSideValue="5.00")
        changed = violation(period="2024Q1", leftSideValue=1, rightSideValue=5)
        missing = violation(period="2024Q1", leftSideValue=None, rightSideValue=5)
        result = fetch_validation_analysis(FakeClient([[first], [duplicate, changed, missing]]), "ZoneAAAAAAA", ["202401", "202402"])
        self.assertEqual(len(result), 3)
        self.assertEqual(result.iloc[0]["leftSideValue"], 0)
        self.assertEqual(result.iloc[1]["leftSideValue"], 1)
        self.assertTrue(result["leftSideValue"].isna().iloc[2])

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


if __name__ == "__main__":
    unittest.main()

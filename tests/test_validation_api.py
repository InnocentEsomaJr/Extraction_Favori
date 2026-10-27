import unittest

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

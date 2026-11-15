import copy
import unittest

import requests

from validation_history import rule_fingerprint, violation_key
from validation_history_checks import verify_history_absences


TARGET = "ZoneAAAAAAA"
UNIT = "CentreAAAAA"
RULE = "RuleAAAAAAA"
GROUP = "GroupAAAAAA"
METADATA = {
    "id": RULE, "operator": "less_than_or_equal_to",
    "leftSide": {"expression": "#{LeftAAAAAAA}"},
    "rightSide": {"expression": "#{RightAAAAAA}"},
}


def violation(**changes):
    row = {
        "targetID": TARGET, "organisationUnitID": UNIT, "validationRuleID": RULE,
        "periodID": "202601", "dayInPeriod": None, "attributeOptionComboID": None,
        "categoryOptionComboID": None, "dataSetID": None,
        "ruleFingerprint": rule_fingerprint(METADATA),
    }
    row.update(changes)
    return row


def snapshot(rows=None, **changes):
    result = {
        "sourceID": "a" * 64, "accountID": "b" * 64,
        "filters": {
            "startDate": "2026-01-01", "endDate": "2026-01-31",
            "groupID": None, "parentID": "CountryAAAA",
        },
        "targets": [{"ID": TARGET, "status": "complete"}],
        "violations": [] if rows is None else rows,
    }
    result.update(changes)
    return result


def http_error(status):
    response = requests.Response()
    response.status_code = status
    return requests.HTTPError("mot-de-passe-secret https://compte:secret@hote.invalid", response=response)


class MetadataClient:
    def __init__(self):
        self.units = {UNIT: {"id": UNIT, "path": f"/CountryAAAA/{TARGET}/{UNIT}"}}
        self.rules = {RULE: copy.deepcopy(METADATA)}
        self.group = {"id": GROUP, "validationRules": [{"id": RULE}]}
        self.responses = {}
        self.calls = []

    def get(self, endpoint, params=None):
        self.calls.append((endpoint, copy.deepcopy(params)))
        if endpoint in self.responses:
            result = self.responses[endpoint]
            if isinstance(result, Exception):
                raise result
            return copy.deepcopy(result)
        if endpoint == f"validationRuleGroups/{GROUP}":
            return copy.deepcopy(self.group)
        if endpoint not in {"organisationUnits", "validationRules"}:
            raise AssertionError(f"Endpoint inattendu : {endpoint}")
        identifiers = params["filter"].removeprefix("id:in:[").removesuffix("]").split(",")
        source = self.units if endpoint == "organisationUnits" else self.rules
        return {endpoint: [copy.deepcopy(source[identifier]) for identifier in identifiers if identifier in source]}

    def post(self, *args, **kwargs):
        raise AssertionError("Aucune analyse POST n'est autorisée")

    def put(self, *args, **kwargs):
        raise AssertionError("Aucune écriture n'est autorisée")


class HistoryAbsenceVerificationTests(unittest.TestCase):
    def setUp(self):
        self.client = MetadataClient()
        self.old = snapshot([violation()])
        self.new = snapshot()

    def verify(self, allowed=None):
        return verify_history_absences(
            self.client, self.old, self.new, {TARGET} if allowed is None else allowed,
        )

    def test_complete_absence_checks_only_read_only_metadata_and_returns_original_key(self):
        verified, warnings = self.verify()
        self.assertEqual(verified, {violation_key(self.old["violations"][0])})
        self.assertEqual(warnings, [])
        self.assertEqual([endpoint for endpoint, _ in self.client.calls], ["organisationUnits", "validationRules"])
        self.assertEqual(self.client.calls[0][1], {
            "fields": "id,path", "filter": f"id:in:[{UNIT}]", "paging": "false",
        })
        self.assertEqual(self.client.calls[1][1]["fields"], "id,operator,leftSide[expression],rightSide[expression]")

    def test_existing_violation_and_its_duplicates_are_never_an_absence(self):
        self.new["violations"] = [violation(), violation()]
        self.assertEqual(self.verify(), (set(), []))
        self.assertEqual(self.client.calls, [])

    def test_current_dates_must_cover_every_old_date(self):
        self.new["filters"].update(startDate="2025-12-01", endDate="2026-02-28")
        self.assertTrue(self.verify()[0])
        for changes in [{"startDate": "2026-01-02"}, {"endDate": "2026-01-30"},
                        {"startDate": "20260101"}, {"endDate": "2026-02-30"}]:
            with self.subTest(changes=changes):
                self.client.calls.clear()
                self.new = snapshot()
                self.new["filters"].update(changes)
                verified, warnings = self.verify()
                self.assertFalse(verified)
                self.assertTrue(warnings)
                self.assertEqual(self.client.calls, [])

    def test_sources_accounts_and_rule_groups_are_isolated(self):
        for field, value in [("sourceID", "c" * 64), ("accountID", "c" * 64), ("sourceID", "")]:
            with self.subTest(field=field):
                self.new = snapshot(**{field: value})
                self.assertFalse(self.verify()[0])
        self.new = snapshot()
        self.new["filters"]["groupID"] = GROUP
        self.assertFalse(self.verify()[0])
        self.assertEqual(self.client.calls, [])

    def test_parent_can_change_when_the_same_individual_target_is_allowed(self):
        self.new["filters"]["parentID"] = TARGET
        self.assertTrue(self.verify()[0])

    def test_a_narrower_current_selection_certifies_only_its_complete_authorized_targets(self):
        other_target, other_unit = "ZoneBBBBBBB", "CentreBBBBB"
        other = violation(targetID=other_target, organisationUnitID=other_unit)
        self.old["targets"].append({"ID": other_target, "status": "complete"})
        self.old["violations"].append(other)
        self.new["filters"]["parentID"] = TARGET
        verified, warnings = self.verify()
        self.assertEqual(verified, {violation_key(violation())})
        self.assertTrue(warnings)
        self.assertTrue(all(other_unit not in str(params) for _, params in self.client.calls))

    def test_incomplete_missing_and_unauthorized_targets_are_never_checked(self):
        for side in ("old", "new"):
            for state in ("failed", "pending", "truncated"):
                with self.subTest(side=side, state=state):
                    self.old, self.new = snapshot([violation()]), snapshot()
                    getattr(self, side)["targets"][0]["status"] = state
                    self.assertFalse(self.verify()[0])
        self.old, self.new = snapshot([violation()]), snapshot()
        self.new["targets"] = []
        self.assertFalse(self.verify()[0])
        self.new = snapshot()
        self.assertFalse(self.verify(set())[0])
        self.assertEqual(self.client.calls, [])

    def test_duplicate_target_declarations_do_not_prove_completeness(self):
        self.new["targets"].append(copy.deepcopy(self.new["targets"][0]))
        self.assertFalse(self.verify()[0])
        self.assertEqual(self.client.calls, [])

    def test_rule_group_membership_is_read_and_required(self):
        self.old["filters"]["groupID"] = self.new["filters"]["groupID"] = GROUP
        verified, warnings = self.verify()
        self.assertTrue(verified)
        self.assertEqual(warnings, [])
        self.assertEqual(self.client.calls[0], (
            f"validationRuleGroups/{GROUP}", {"fields": "id,validationRules[id]"},
        ))
        self.client.calls.clear()
        self.client.group["validationRules"] = []
        verified, warnings = self.verify()
        self.assertFalse(verified)
        self.assertIn("groupe", " ".join(warnings))
        self.assertEqual(len(self.client.calls), 1)

    def test_inaccessible_or_incomplete_rule_group_never_certifies_absence(self):
        self.old["filters"]["groupID"] = self.new["filters"]["groupID"] = GROUP
        for response in [http_error(403), {}, {"id": GROUP}, {"id": GROUP, "validationRules": None},
                         {"id": GROUP, "validationRules": [{}]}]:
            with self.subTest(response=type(response).__name__):
                self.client.responses[f"validationRuleGroups/{GROUP}"] = response
                verified, warnings = self.verify()
                self.assertFalse(verified)
                self.assertTrue(warnings)

    def test_unit_access_loss_api_errors_and_bad_responses_are_not_corrections(self):
        for endpoint in ("organisationUnits", "validationRules"):
            for response in [http_error(403), http_error(500), TimeoutError("secret"),
                             ValueError("secret"), None, [], {}, {endpoint: None}, {endpoint: [None]}, {endpoint: []}]:
                with self.subTest(endpoint=endpoint, response=type(response).__name__):
                    self.client.responses = {endpoint: response}
                    verified, warnings = self.verify()
                    self.assertFalse(verified)
                    self.assertTrue(warnings)
                    self.assertNotIn("secret", " ".join(warnings))
                    self.assertNotIn("https:", " ".join(warnings))
                    self.assertNotIn("403", " ".join(warnings))

    def test_unit_must_still_belong_to_the_exact_target_and_path_end(self):
        for path in [None, "", f"/CountryAAAA/ZoneBBBBBBB/{UNIT}",
                     f"/CountryAAAA/{TARGET}2/{UNIT}", f"/CountryAAAA/{TARGET}/CentreBBBBB",
                     f"/CountryAAAA/{TARGET}/{UNIT}/", f"/CountryAAAA/{TARGET}/{UNIT}/OtherAAAAAA"]:
            with self.subTest(path=path):
                self.client.units[UNIT]["path"] = path
                verified, warnings = self.verify()
                self.assertFalse(verified)
                self.assertIn("rattachement", " ".join(warnings))

    def test_missing_changed_or_incomplete_rule_fingerprints_are_not_corrections(self):
        for metadata in [{}, {"id": RULE}, {**METADATA, "operator": "greater_than"},
                         {**METADATA, "leftSide": {}}, {**METADATA, "rightSide": {"expression": ""}},
                         {**METADATA, "leftSide": {"expression": "#{ChangedAAAA}"}}]:
            with self.subTest(metadata=metadata):
                self.client.rules[RULE] = metadata
                verified, warnings = self.verify()
                self.assertFalse(verified)
                self.assertIn("règles", " ".join(warnings))
        self.client.rules[RULE] = copy.deepcopy(METADATA)
        self.old["violations"][0]["ruleFingerprint"] = None
        self.assertFalse(self.verify()[0])

    def test_duplicate_old_rows_share_one_verification_without_inflating_count(self):
        self.old["violations"].append(copy.deepcopy(self.old["violations"][0]))
        verified, warnings = self.verify()
        self.assertEqual(verified, {violation_key(violation())})
        self.assertEqual(warnings, [])
        self.assertEqual(len(self.client.calls), 2)

    def test_old_duplicates_with_conflicting_rule_fingerprints_or_targets_remain_unverified(self):
        for changes in [{"ruleFingerprint": "c" * 64}, {"targetID": "ZoneBBBBBBB"}]:
            with self.subTest(changes=changes):
                self.old["violations"] = [violation(), violation(**changes)]
                self.assertFalse(self.verify()[0])
        self.assertEqual(self.client.calls, [])

    def test_partial_metadata_can_only_certify_units_explicitly_returned(self):
        second = violation(organisationUnitID="CentreBBBBB")
        self.old["violations"].append(second)
        verified, warnings = self.verify()
        self.assertEqual(verified, {violation_key(violation())})
        self.assertNotIn(violation_key(second), verified)
        self.assertTrue(warnings)

    def test_duplicate_metadata_objects_are_ambiguous(self):
        for endpoint, row in [("organisationUnits", self.client.units[UNIT]), ("validationRules", METADATA)]:
            with self.subTest(endpoint=endpoint):
                self.client.responses = {
                    endpoint: {endpoint: [copy.deepcopy(row), copy.deepcopy(row)]},
                }
                self.assertFalse(self.verify()[0])

    def test_paged_metadata_is_incomplete_even_when_it_contains_the_requested_id(self):
        self.client.responses["organisationUnits"] = {
            "organisationUnits": [self.client.units[UNIT]], "pager": {"pageCount": 2, "total": 2},
        }
        self.assertFalse(self.verify()[0])

    def test_invalid_current_rows_prevent_certification(self):
        for rows in [None, {}, [None], [{}], [violation(organisationUnitID=None)]]:
            with self.subTest(rows=rows):
                self.new["violations"] = rows
                self.assertFalse(self.verify()[0])
        self.assertEqual(self.client.calls, [])

    def test_unsafe_group_identifier_never_reaches_the_client(self):
        self.old["filters"]["groupID"] = self.new["filters"]["groupID"] = "../../users"
        self.assertFalse(self.verify()[0])
        self.assertEqual(self.client.calls, [])

    def test_metadata_requests_are_deduplicated_and_bounded_to_sixty_ids(self):
        rows = []
        self.client.units, self.client.rules = {}, {}
        for index in range(121):
            unit_id, rule_id = f"U{index:010d}", f"R{index:010d}"
            rule = {**copy.deepcopy(METADATA), "id": rule_id}
            rows.append(violation(organisationUnitID=unit_id, validationRuleID=rule_id,
                                  ruleFingerprint=rule_fingerprint(rule)))
            self.client.units[unit_id] = {"id": unit_id, "path": f"/CountryAAAA/{TARGET}/{unit_id}"}
            self.client.rules[rule_id] = rule
        self.old["violations"] = rows + [copy.deepcopy(rows[0])]
        verified, warnings = self.verify()
        self.assertEqual(len(verified), 121)
        self.assertEqual(warnings, [])
        for endpoint in ("organisationUnits", "validationRules"):
            sizes = [len(params["filter"].removeprefix("id:in:[").removesuffix("]").split(","))
                     for called, params in self.client.calls if called == endpoint]
            self.assertEqual(sizes, [60, 60, 1])


if __name__ == "__main__":
    unittest.main()

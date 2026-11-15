"""Contrat public de l'historique et preuves nécessaires aux corrections."""

from copy import deepcopy
import json
import re
import unittest

import pandas as pd

from validation_history import (
    build_snapshot, compare_snapshots, snapshot_filename, validate_snapshot,
    violation_key, source_fingerprint, account_fingerprint, rule_fingerprint,
    snapshot_comparison_filters,
)


COUNTRY = 'CountryAAAA'
PROVINCE = 'ProvinceAAA'
ZONE = 'ZoneAAAAAAA'
OTHER_ZONE = 'ZoneBBBBBBB'
CENTRE = 'CentreAAAAA'
RULE = 'RuleAAAAAAA'
OTHER_RULE = 'RuleBBBBBBB'
GROUP = 'GroupAAAAAA'
STAMP = '2026-09-21T12:34:56+01:00'
RULE_METADATA = {
    'operator': 'less_than_or_equal_to',
    'leftSide': {'expression': '#{PRIVATE_EXPRESSION_LEFT}'},
    'rightSide': {'expression': '#{PRIVATE_EXPRESSION_RIGHT}'},
}


class HistoryTests(unittest.TestCase):
    def row(self, **changes):
        result = {
            'organisationUnitId': CENTRE, 'organisationUnitDisplayName': 'Centre A',
            'validationRuleId': RULE, 'validationRuleDescription': 'Règle A',
            'periodId': '202609', 'periodDisplayName': 'Septembre 2026',
            'attributeOptionComboId': 'AocAAAAAAAA', 'categoryOptionCombo': {'id': 'CocAAAAAAAA'},
            'dataSet': {'id': 'DatasetAAAA'}, '_request_unit_id': ZONE,
            'leftSideValue': 8, 'rightSideValue': 4,
        }
        result.update(changes)
        return result

    def snapshot(self, rows=None, *, status='complete', start='2026-09-01', end='2026-09-20',
                 source='https://fixture.invalid/dhis', account='PRIVATE_ACTOR',
                 metadata=None, group=GROUP, extra_zone=False, snapshot_id='a' * 32):
        rows = [self.row()] if rows is None else rows
        target_rows = [{'Zone_ID': ZONE, 'Zone': 'ZS A', 'Pays_ID': COUNTRY, 'Pays': 'RDC',
                        'Province_ID': PROVINCE, 'Province': 'Province A', 'level': 3}]
        if extra_zone:
            target_rows.append({**target_rows[0], 'Zone_ID': OTHER_ZONE, 'Zone': 'ZS B'})
        frames = {row['Zone_ID']: pd.DataFrame() for row in target_rows}
        batch = {
            'frames_by_id': frames,
            'processed_ids': [row['Zone_ID'] for row in target_rows] if status != 'pending' else [],
            'pending_ids': [ZONE] if status == 'pending' else [],
            'succeeded_by_id': {row['Zone_ID']: status == 'complete' for row in target_rows},
            'errors_by_id': {ZONE: ['PRIVATE_ERROR']} if status == 'failed' else {},
            'context': {'token': 'PRIVATE_CONTEXT_TOKEN'},
        }
        if status == 'truncated':
            frames[ZONE].attrs['truncated_periods'] = ['202609']
        frames[ZONE].attrs['request_params'] = {'token': 'PRIVATE_FRAME_TOKEN'}
        return build_snapshot(
            source_url=source, account=account, start_date=start, end_date=end,
            parent_id=PROVINCE, parent_name='Province A', group_id=group, group_name='Groupe A',
            targets=pd.DataFrame(target_rows), batch=batch, raw=pd.DataFrame(rows),
            details=pd.DataFrame(), rule_metadata=metadata if metadata is not None else {RULE: RULE_METADATA},
            analysed_at=STAMP, captured_at=STAMP, snapshot_id=snapshot_id,
        )

    def compare(self, old, new, proofs=(), allowed=None):
        return compare_snapshots(old, new, {ZONE} if allowed is None else allowed, proofs)

    def count(self, comparison, column, zone=ZONE):
        return int(comparison['summary'].set_index('Zone_ID').loc[zone, column])

    def test_snapshot_whitelist_never_serializes_context_errors_credentials_or_expressions(self):
        row = self.row(password='PRIVATE_ROW_PASSWORD', access_token='PRIVATE_ROW_TOKEN')
        snapshot = self.snapshot([row], source='https://PRIVATE_USER:PRIVATE_PASSWORD@fixture.invalid/dhis?token=PRIVATE_QUERY')
        encoded = json.dumps(snapshot, allow_nan=False)
        for secret in ('PRIVATE_USER', 'PRIVATE_PASSWORD', 'PRIVATE_ACTOR', 'PRIVATE_QUERY',
                       'PRIVATE_ROW_PASSWORD', 'PRIVATE_ROW_TOKEN', 'PRIVATE_FRAME_TOKEN',
                       'PRIVATE_CONTEXT_TOKEN', 'PRIVATE_EXPRESSION_LEFT', 'fixture.invalid'):
            with self.subTest(secret=secret):
                self.assertNotIn(secret, encoded)
        self.assertEqual(snapshot['violations'][0]['ruleFingerprint'], rule_fingerprint(RULE_METADATA))
        self.assertEqual(snapshot['targets'][0]['status'], 'complete')
        self.assertEqual(validate_snapshot(json.loads(encoded)), snapshot)

    def test_flat_and_nested_uids_period_iso_and_duplicates_survive(self):
        flat = self.row(dayInPeriod=0)
        nested = self.row(organisationUnitId=None, organisationUnit={'id': CENTRE},
                          validationRuleId=None, validationRule={'id': RULE}, periodId=None,
                          period={'id': 1234, 'iso': '202609'}, dayInPeriod=0)
        snapshot = self.snapshot([flat, nested, deepcopy(flat)])
        self.assertEqual(len(snapshot['violations']), 3)
        self.assertEqual(len({violation_key(row) for row in snapshot['violations']}), 1)
        self.assertEqual(snapshot['violations'][0]['dayInPeriod'], 0)

    def test_labels_are_never_used_as_identifiers_and_numbers_are_json_finite(self):
        snapshot = self.snapshot([self.row(organisationUnitId=None,
                                          organisationUnit={'displayName': 'Centre A'},
                                          leftSideValue=float('nan'), rightSideValue=float('inf'))])
        row = snapshot['violations'][0]
        self.assertIsNone(row['organisationUnitID'])
        self.assertIsNone(row['leftValue'])
        self.assertIsNone(row['rightValue'])
        self.assertIsNone(violation_key(row))
        json.dumps(snapshot, allow_nan=False)

    def test_period_day_and_combinations_all_belong_to_identity_but_values_do_not(self):
        baseline = self.snapshot()['violations'][0]
        key = violation_key(baseline)
        same = self.snapshot([self.row(leftSideValue=100, rightSideValue=2)])['violations'][0]
        self.assertEqual(key, violation_key(same))
        for changes in ({'periodId': '202610'}, {'dayInPeriod': 0},
                        {'attributeOptionComboId': 'AocBBBBBBBB'},
                        {'categoryOptionCombo': {'id': 'CocBBBBBBBB'}},
                        {'dataSet': {'id': 'DatasetBBBB'}}):
            with self.subTest(changes=changes):
                changed = self.snapshot([self.row(**changes)])['violations'][0]
                self.assertNotEqual(key, violation_key(changed))

    def test_status_preserves_success_failures_pending_and_truncation_separately(self):
        for status in ('complete', 'failed', 'pending', 'truncated'):
            with self.subTest(status=status):
                self.assertEqual(self.snapshot([], status=status)['targets'][0]['status'], status)

    def test_filename_is_ascii_unique_and_uses_real_capture_in_utc(self):
        first = self.snapshot()
        second = self.snapshot(snapshot_id='b' * 32)
        name = snapshot_filename(first)
        name.encode('ascii')
        self.assertTrue(name.startswith('septembre_2026-09-01_2026-09-20_' + PROVINCE))
        self.assertIn('20260921T113456000000Z', name)
        self.assertNotEqual(name, snapshot_filename(second))
        self.assertRegex(name, re.compile(r'^[A-Za-z0-9_.-]+\.json$'))

    def test_source_and_account_fingerprints_are_canonical_and_do_not_include_credentials(self):
        self.assertEqual(source_fingerprint('https://fixture.invalid/dhis/'),
                         source_fingerprint('https://name:password@FIXTURE.invalid:443/dhis?token=no#fragment'))
        self.assertNotEqual(source_fingerprint('http://fixture.invalid/dhis'),
                            source_fingerprint('https://fixture.invalid/dhis'))
        self.assertNotEqual(account_fingerprint('one'), account_fingerprint('two'))
        self.assertIsNone(rule_fingerprint({'displayName': 'Nom seul'}))
        self.assertIsNone(rule_fingerprint({'operator': 'equal', 'leftSide': {}, 'rightSide': {}}))
        changed = deepcopy(RULE_METADATA)
        changed['rightSide']['expression'] += '+1'
        self.assertNotEqual(rule_fingerprint(changed), rule_fingerprint(RULE_METADATA))

    def test_strict_validation_rejects_unknown_fields_invalid_refs_dates_and_nan(self):
        good = self.snapshot()
        mutations = [
            lambda value: value.update(token='SECRET'),
            lambda value: value.update(schemaVersion=True),
            lambda value: value.update(snapshotID='../escape'),
            lambda value: value.update(capturedAt='2026-09-21T12:34:56'),
            lambda value: value.update(capturedAt='2026-09-20T12:34:56+01:00'),
            lambda value: value['filters'].update(startDate='2026-09-31'),
            lambda value: value['filters'].update(parentID='../escape'),
            lambda value: value['targets'][0].update(status={'bad': True}),
            lambda value: value['targets'].append(deepcopy(value['targets'][0])),
            lambda value: value['violations'][0].update(targetID=OTHER_ZONE),
            lambda value: value['violations'][0].update(leftValue=float('nan')),
            lambda value: value['violations'][0].update(dayInPeriod=True),
            lambda value: value['violations'][0].update(periodID='202613'),
            lambda value: value['violations'][0].update(unexpected='SECRET'),
        ]
        for mutation in mutations:
            value = deepcopy(good)
            mutation(value)
            with self.subTest(value=value):
                with self.assertRaises(ValueError):
                    validate_snapshot(value)

    def test_presence_is_persistent_even_if_values_changed(self):
        old = self.snapshot()
        new = self.snapshot([self.row(leftSideValue=9)], end='2026-09-30')
        result = self.compare(old, new)
        self.assertEqual(self.count(result, 'Persistantes'), 1)
        self.assertEqual(self.count(result, 'Corrigées'), 0)
        self.assertEqual(result['details']['État historique'].tolist(), ['Toujours violée'])

    def test_absence_requires_explicit_metadata_proof_even_for_successful_empty_result(self):
        old, new = self.snapshot(), self.snapshot([])
        unproven = self.compare(old, new)
        self.assertEqual(self.count(unproven, 'Corrigées'), 0)
        self.assertEqual(self.count(unproven, 'Non vérifiables'), 1)
        key = violation_key(old['violations'][0])
        proven = self.compare(old, new, {key})
        self.assertEqual(self.count(proven, 'Corrigées'), 1)

    def test_disjoint_partial_horizons_never_become_corrections_even_with_proof(self):
        old = self.snapshot([self.row(periodId='20260907', dayInPeriod=0)])
        key = violation_key(old['violations'][0])
        for start, end in (('2026-09-20', '2026-09-30'), ('2026-09-01', '2026-09-10')):
            with self.subTest(start=start, end=end):
                result = self.compare(old, self.snapshot([], start=start, end=end), {key})
                self.assertEqual(self.count(result, 'Corrigées'), 0)
                self.assertEqual(self.count(result, 'Non vérifiables'), 1)

    def test_wider_interval_can_confirm_absence_only_with_proof(self):
        old = self.snapshot()
        key = violation_key(old['violations'][0])
        result = self.compare(old, self.snapshot([], end='2026-09-30'), {key})
        self.assertEqual(self.count(result, 'Corrigées'), 1)

    def test_source_account_group_mismatch_and_incomplete_requests_block_proofs(self):
        old = self.snapshot()
        key = violation_key(old['violations'][0])
        variants = [self.snapshot([], source='https://other.invalid'),
                    self.snapshot([], account='other'), self.snapshot([], group='GroupBBBBBB')]
        variants.extend(self.snapshot([], status=status) for status in ('failed', 'pending', 'truncated'))
        for new in variants:
            with self.subTest(new=new):
                result = self.compare(old, new, {key})
                self.assertEqual(self.count(result, 'Corrigées'), 0)
                self.assertEqual(self.count(result, 'Non vérifiables'), 1)
        failed_old = self.snapshot(status='failed')
        self.assertEqual(self.count(self.compare(failed_old, self.snapshot([]), {key}), 'Corrigées'), 0)

    def test_missing_or_changed_rule_definition_is_unverifiable(self):
        old = self.snapshot(metadata={})
        key = violation_key(old['violations'][0])
        self.assertEqual(self.count(self.compare(old, self.snapshot([]), {key}), 'Corrigées'), 0)
        changed = deepcopy(RULE_METADATA)
        changed['leftSide']['expression'] += '+2'
        result = self.compare(self.snapshot(), self.snapshot(metadata={RULE: changed}))
        self.assertEqual(self.count(result, 'Persistantes'), 0)
        self.assertEqual(self.count(result, 'Non vérifiables'), 1)

    def test_older_current_analysis_and_conflicting_duplicate_definitions_block_correction(self):
        old, new = self.snapshot(), self.snapshot([])
        key = violation_key(old['violations'][0])
        new['analysedAt'] = '2026-09-20T12:34:56+01:00'
        result = self.compare(old, new, {key})
        self.assertEqual(self.count(result, 'Corrigées'), 0)
        self.assertEqual(self.count(result, 'Non vérifiables'), 1)
        duplicate = deepcopy(old['violations'][0])
        duplicate['ruleFingerprint'] = 'b' * 64
        old['violations'].append(duplicate)
        result = self.compare(old, self.snapshot([]), {key})
        self.assertEqual(self.count(result, 'Corrigées'), 0)
        self.assertEqual(self.count(result, 'Non vérifiables'), 1)

    def test_new_observations_distinguish_historical_coverage_from_extended_filter(self):
        new_row = self.row(validationRuleId=OTHER_RULE)
        old = self.snapshot([])
        equal = self.compare(old, self.snapshot([new_row]))
        extended = self.compare(old, self.snapshot([new_row], end='2026-09-30'))
        failed = self.compare(self.snapshot([], status='failed'), self.snapshot([new_row]))
        self.assertEqual(self.count(equal, 'Nouvelles'), 1)
        self.assertEqual(self.count(extended, 'Nouvelles dans le filtre actuel'), 1)
        self.assertEqual(self.count(extended, 'Nouvelles'), 0)
        self.assertEqual(self.count(failed, 'Non vérifiables'), 1)

    def test_duplicates_are_preserved_but_counts_use_distinct_identity(self):
        old = self.snapshot([self.row(), self.row()])
        result = self.compare(old, self.snapshot())
        self.assertEqual(self.count(result, 'Anciennes'), 1)
        self.assertEqual(self.count(result, 'Persistantes'), 1)
        self.assertEqual(self.count(result, 'Corrigées'), 0)
        self.assertEqual(len(result['details']), 2)
        self.assertEqual(result['details']['Occurrences anciennes'].tolist(), [2, 2])
        proven = self.compare(old, self.snapshot([]), {violation_key(old['violations'][0])})
        self.assertEqual(self.count(proven, 'Corrigées'), 1)

    def test_authorized_zone_filter_excludes_other_zones_and_never_mutates_snapshots(self):
        other = self.row(_request_unit_id=OTHER_ZONE, organisationUnitId='CentreBBBBB')
        old, new = self.snapshot([self.row(), other], extra_zone=True), self.snapshot([], extra_zone=True)
        originals = deepcopy((old, new))
        result = self.compare(old, new)
        self.assertEqual(result['summary']['Zone_ID'].tolist(), [ZONE])
        self.assertEqual(result['details']['Zone_ID'].tolist(), [ZONE])
        self.assertEqual((old, new), originals)
        self.assertTrue(self.compare(old, new, allowed=set())['summary'].empty)
        with self.assertRaises(ValueError):
            self.compare(old, new, allowed={'../outside'})
        filters = snapshot_comparison_filters(old)
        filters['startDate'] = 'changed'
        self.assertEqual(old['filters']['startDate'], '2026-09-01')


if __name__ == '__main__':
    unittest.main()

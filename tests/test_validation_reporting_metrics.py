import unittest

import numpy as np
import pandas as pd

from reporting_metrics import (
    aggregate_validation_metrics, has_complete_reporting_metrics, reporting_quality_warnings,
)


class ValidationReportingMetricTests(unittest.TestCase):
    def targets(self, ids=("ZoneAAAAAAA", "ZoneBBBBBBB")):
        return pd.DataFrame({
            "Organisation unit ID": list(ids), "Zone_ID": list(ids),
            "Zone": ["Zone homonyme"] * len(ids),
        })

    def row(self, identifier="ZoneAAAAAAA", actual=95, expected=100, on_time=80, **extra):
        return {
            "Organisation unit": "Zone homonyme", "Organisation unit ID": identifier,
            "Dataset Actual reports": actual, "Dataset Expected reports": expected,
            "Dataset Actual reports on time": on_time, "Dataset Reporting rate": 95,
            **extra,
        }

    def test_same_names_use_ids_and_exclude_other_zones_and_parents(self):
        source = pd.DataFrame([
            self.row(actual=95, on_time=80),
            self.row("ZoneBBBBBBB", actual=50, on_time=40),
            self.row("OtherZoneAA", actual=9999, on_time=9999),
            self.row("ProvinceAAA", actual=9999, on_time=9999),
        ])
        actual = aggregate_validation_metrics(source, self.targets()).set_index("Zone_ID")
        self.assertEqual(actual.loc["ZoneAAAAAAA"].tolist(), [95, 80, 95, 1])
        self.assertEqual(actual.loc["ZoneBBBBBBB"].tolist(), [50, 40, 50, 0])

    def test_missing_ids_are_never_associated_by_a_matching_name(self):
        source = pd.DataFrame([self.row()]).drop(columns="Organisation unit ID")
        actual = aggregate_validation_metrics(source, self.targets())
        self.assertTrue(actual.iloc[:, 1:].isna().all().all())

    def test_target_without_source_row_stays_missing_not_zero(self):
        actual = aggregate_validation_metrics(pd.DataFrame([self.row()]), self.targets())
        self.assertTrue(actual.iloc[1, 1:].isna().all())
        self.assertEqual(actual.iloc[0]["Actual_Reports"], 95)

    def test_parent_data_is_not_added_to_children_or_duplicate_sources(self):
        parent = self.row()
        child = self.row("CentreAAAAA", actual=500, on_time=500)
        source = pd.DataFrame([parent, dict(parent), child])
        actual = aggregate_validation_metrics(source, self.targets(("ZoneAAAAAAA",))).iloc[0]
        self.assertEqual(actual["Actual_Reports"], 95)
        self.assertEqual(actual["Complétude_Globale"], 95)
        self.assertEqual(actual["Nombre_Datasets_95"], 1)

    def test_distinct_months_sum_reports_and_weight_rates(self):
        source = pd.DataFrame([
            self.row(actual=1, expected=1, on_time=1, **{"Reporting month": "202612"}),
            self.row(actual=94, expected=100, on_time=50, **{"Reporting month": "202701"}),
        ])
        actual = aggregate_validation_metrics(source, self.targets(("ZoneAAAAAAA",))).iloc[0]
        self.assertEqual(actual["Actual_Reports"], 95)
        self.assertEqual(actual["Complétude_Globale"], 94.06)
        self.assertEqual(actual["Promptitude_Globale"], 50.5)
        self.assertEqual(actual["Nombre_Datasets_95"], 0)

    def test_multiple_datasets_sum_reports_and_count_each_dataset_once(self):
        row = self.row(actual=10, expected=10, on_time=9)
        for number in range(1, 10):
            row.update({
                f"Dataset {number} Actual reports": 10,
                f"Dataset {number} Expected reports": 10,
                f"Dataset {number} Actual reports on time": 9,
                f"Dataset {number} Reporting rate": 100,
            })
        actual = aggregate_validation_metrics(pd.DataFrame([row]), self.targets(("ZoneAAAAAAA",))).iloc[0]
        self.assertEqual(actual["Actual_Reports"], 100)
        self.assertEqual(actual["Complétude_Globale"], 100)
        self.assertEqual(actual["Promptitude_Globale"], 90)
        self.assertEqual(actual["Nombre_Datasets_95"], 10)

    def test_zero_reports_remain_zero_but_zero_expected_makes_rates_unavailable(self):
        source = pd.DataFrame([self.row(actual=0, expected=0, on_time=0)])
        actual = aggregate_validation_metrics(source, self.targets(("ZoneAAAAAAA",))).iloc[0]
        self.assertEqual(actual["Actual_Reports"], 0)
        self.assertTrue(pd.isna(actual["Complétude_Globale"]))
        self.assertTrue(pd.isna(actual["Promptitude_Globale"]))
        self.assertTrue(pd.isna(actual["Nombre_Datasets_95"]))
        source["Dataset Expected reports"] = 10
        actual = aggregate_validation_metrics(source, self.targets(("ZoneAAAAAAA",))).iloc[0]
        self.assertEqual(actual["Complétude_Globale"], 0)
        self.assertEqual(actual["Promptitude_Globale"], 0)
        self.assertEqual(actual["Nombre_Datasets_95"], 0)

    def test_absent_metric_columns_are_unavailable(self):
        source = pd.DataFrame([self.row()]).drop(columns="Dataset Actual reports on time")
        actual = aggregate_validation_metrics(source, self.targets(("ZoneAAAAAAA",))).iloc[0]
        self.assertEqual(actual["Complétude_Globale"], 95)
        self.assertTrue(pd.isna(actual["Promptitude_Globale"]))
        source = source.drop(columns=["Dataset Actual reports", "Dataset Expected reports"])
        actual = aggregate_validation_metrics(source, self.targets(("ZoneAAAAAAA",))).iloc[0]
        self.assertTrue(pd.isna(actual["Actual_Reports"]))
        self.assertTrue(pd.isna(actual["Complétude_Globale"]))

    def test_invalid_counts_do_not_produce_partial_healthy_totals(self):
        for invalid in [None, np.nan, np.inf, -1, "inconnu"]:
            with self.subTest(invalid=invalid):
                source = pd.DataFrame([self.row(actual=invalid)])
                actual = aggregate_validation_metrics(source, self.targets(("ZoneAAAAAAA",))).iloc[0]
                self.assertTrue(pd.isna(actual["Actual_Reports"]))
                self.assertTrue(pd.isna(actual["Complétude_Globale"]))
                self.assertTrue(pd.isna(actual["Nombre_Datasets_95"]))

    def test_targets_are_unique_and_keep_their_order(self):
        targets = self.targets(("ZoneBBBBBBB", "ZoneAAAAAAA", "ZoneBBBBBBB", ""))
        actual = aggregate_validation_metrics(pd.DataFrame([self.row()]), targets)
        self.assertEqual(actual["Zone_ID"].tolist(), ["ZoneBBBBBBB", "ZoneAAAAAAA"])

    def test_empty_inputs_and_no_mutation(self):
        source = pd.DataFrame([self.row()])
        targets = self.targets()
        original_source, original_targets = source.copy(deep=True), targets.copy(deep=True)
        aggregate_validation_metrics(source, targets)
        pd.testing.assert_frame_equal(source, original_source)
        pd.testing.assert_frame_equal(targets, original_targets)
        self.assertTrue(aggregate_validation_metrics(source, pd.DataFrame()).empty)
        missing = aggregate_validation_metrics(pd.DataFrame(), targets)
        self.assertEqual(len(missing), 2)
        self.assertTrue(missing.iloc[:, 1:].isna().all().all())


class ReportingQualityWarningTests(unittest.TestCase):
    targets = ValidationReportingMetricTests.targets
    row = ValidationReportingMetricTests.row

    def warnings(self, source, targets=None):
        selected = self.targets(("ZoneAAAAAAA",)) if targets is None else targets
        metrics = aggregate_validation_metrics(source, selected)
        return reporting_quality_warnings(source, metrics, selected)

    def test_complete_reporting_families_are_structural_and_ignore_rates_on_time(self):
        source = pd.DataFrame([self.row()])
        self.assertTrue(has_complete_reporting_metrics(source))
        self.assertFalse(has_complete_reporting_metrics(source.drop(columns='Dataset Actual reports on time')))
        rate_only = pd.DataFrame({
            'Dataset Reporting rate': [95], 'Dataset Reporting rate on time': [80],
        })
        self.assertFalse(has_complete_reporting_metrics(rate_only))
        self.assertFalse(has_complete_reporting_metrics(None))
        self.assertFalse(has_complete_reporting_metrics(pd.DataFrame()))
        missing_values = source.copy()
        missing_values['Dataset Actual reports'] = np.nan
        self.assertTrue(has_complete_reporting_metrics(missing_values))
        self.assertTrue(has_complete_reporting_metrics(source.rename(columns={
            'Dataset Actual reports': ' Dataset  ACTUAL REPORTS ',
        })))

    def test_valid_reporting_metrics_have_no_warning(self):
        self.assertEqual(self.warnings(pd.DataFrame([self.row()])), [])

    def test_rate_only_dataset_is_not_hidden_by_another_complete_dataset(self):
        source = pd.DataFrame([self.row()])
        for suffix in ['Reporting rate', 'Reporting rate on time']:
            with self.subTest(suffix=suffix):
                mixed = source.assign(**{f'Other {suffix}': 100})
                self.assertFalse(has_complete_reporting_metrics(mixed))

    def test_multiple_complete_dataset_count_pairs_do_not_require_rate_columns(self):
        row = self.row()
        row.update({
            'Other Actual reports': 8, 'Other Expected reports': 10,
            'Other Actual reports on time': 7,
        })
        source = pd.DataFrame([row]).drop(columns='Dataset Reporting rate')
        self.assertTrue(has_complete_reporting_metrics(source))
        self.assertFalse(has_complete_reporting_metrics(source.drop(columns='Other Actual reports on time')))

    def test_duplicate_dataset_count_family_is_not_a_complete_source(self):
        source = pd.DataFrame([self.row()])
        source[' DATASET  actual REPORTS '] = 95
        self.assertFalse(has_complete_reporting_metrics(source))

    def test_missing_metric_family_is_explained_without_fabricated_counts(self):
        source = pd.DataFrame([self.row()]).drop(columns='Dataset Actual reports on time')
        warnings = self.warnings(source)
        self.assertTrue(any('rapports à temps' in message for message in warnings))
        self.assertTrue(any('promptitude' in message and 'ZoneAAAAAAA' in message for message in warnings))
        metrics = aggregate_validation_metrics(source, self.targets(("ZoneAAAAAAA",)))
        self.assertTrue(pd.isna(metrics.iloc[0]['Promptitude_Globale']))

    def test_target_without_data_is_identified_by_id_even_when_names_match(self):
        source = pd.DataFrame([self.row('ZoneBBBBBBB')])
        warnings = self.warnings(source)
        self.assertTrue(any('ZoneAAAAAAA' in message and 'aucune donnée' in message for message in warnings))
        self.assertFalse(any('ZoneBBBBBBB' in message for message in warnings))

    def test_partial_dataset_values_are_explained_for_the_affected_zone(self):
        row = self.row()
        row.update({
            'Other Actual reports': np.nan, 'Other Expected reports': 10,
            'Other Actual reports on time': 8, 'Other Reporting rate': np.nan,
        })
        warnings = self.warnings(pd.DataFrame([row]))
        self.assertTrue(any(
            'ZoneAAAAAAA' in message and 'valeurs manquantes' in message and 'rapports réels' in message
            for message in warnings
        ))

    def test_incomplete_dataset_columns_are_not_hidden_by_another_complete_dataset(self):
        row = self.row()
        row['Other Actual reports'] = 10
        warnings = self.warnings(pd.DataFrame([row]))
        self.assertTrue(any('datasets' in message and 'other' in message for message in warnings))

    def test_invalid_counts_and_zero_denominator_have_explicit_reasons(self):
        for value in [-1, np.inf, 'inconnu']:
            with self.subTest(value=value):
                warnings = self.warnings(pd.DataFrame([self.row(actual=value)]))
                self.assertTrue(any('comptes invalides' in message and 'rapports réels' in message for message in warnings))
        warnings = self.warnings(pd.DataFrame([self.row(actual=0, expected=0, on_time=0)]))
        self.assertTrue(any('aucun rapport attendu' in message for message in warnings))

    def test_missing_ids_empty_targets_and_source_mutation_are_safe(self):
        source = pd.DataFrame([self.row()])
        selected = self.targets()
        metrics = aggregate_validation_metrics(source, selected)
        original_source, original_metrics, original_targets = source.copy(deep=True), metrics.copy(deep=True), selected.copy(deep=True)
        messages = reporting_quality_warnings(source, metrics, selected)
        self.assertTrue(any('ZoneBBBBBBB' in message for message in messages))
        pd.testing.assert_frame_equal(source, original_source)
        pd.testing.assert_frame_equal(metrics, original_metrics)
        pd.testing.assert_frame_equal(selected, original_targets)
        missing_ids = self.warnings(source.drop(columns='Organisation unit ID'))
        self.assertTrue(any('identifiants' in message for message in missing_ids))
        self.assertEqual(reporting_quality_warnings(source, metrics, pd.DataFrame()), [])


if __name__ == "__main__":
    unittest.main()

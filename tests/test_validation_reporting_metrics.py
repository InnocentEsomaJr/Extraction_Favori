import unittest

import numpy as np
import pandas as pd

from reporting_metrics import aggregate_validation_metrics


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


if __name__ == "__main__":
    unittest.main()

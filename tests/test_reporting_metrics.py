import unittest

import numpy as np
import pandas as pd

from reporting_metrics import attach_reporting_scope, clean_ou_id, count_datasets_95


class ReportingScopeTests(unittest.TestCase):
    def metadata(self):
        return pd.DataFrame({
            "Organisation unit": ["Homonyme", "Homonyme", "Zone A"],
            "Organisation unit ID": ["ou-a", "ou-b", "zone-a"],
            "Zone": ["Zone A", "Zone B", "Zone A"],
            "Zone_ID": ["zone-a", "zone-b", "zone-a"],
        })

    def test_ids_distinguish_homonymous_units_without_cartesian_join(self):
        source = pd.DataFrame({"Organisation unit": ["Homonyme", "Homonyme"],
                               "Organisation unit ID": ["ou-a", "ou-b"], "Actual": [10, 20]})
        result = attach_reporting_scope(source, self.metadata())
        self.assertEqual(len(result), 2)
        self.assertEqual(result["ou_id"].tolist(), ["ou-a", "ou-b"])
        self.assertEqual(result["Zone_ID"].tolist(), ["zone-a", "zone-b"])
        self.assertEqual(result["Actual"].tolist(), [10, 20])

    def test_ambiguous_name_without_id_stays_unmatched(self):
        source = pd.DataFrame({"Organisation unit": [" Homonyme "]})
        result = attach_reporting_scope(source, self.metadata())
        self.assertEqual(result.iloc[0]["ou_id"], "")
        self.assertEqual(result.iloc[0]["Zone_ID"], "")

    def test_unique_name_without_id_matches_normalized_metadata(self):
        source = pd.DataFrame({"Organisation unit": [" ZONE   A "]})
        result = attach_reporting_scope(source, self.metadata())
        self.assertEqual(result.iloc[0]["ou_id"], "zone-a")
        self.assertEqual(result.iloc[0]["Zone_ID"], "zone-a")

    def test_parent_rows_replace_descendants_in_same_group(self):
        source = pd.DataFrame({"Organisation unit": ["Homonyme", "Zone A", "Homonyme"],
                               "Organisation unit ID": ["ou-a", "zone-a", "ou-b"],
                               "Actual": [50, 40, 30]}, index=[5, 6, 7])
        result = attach_reporting_scope(source, self.metadata())
        self.assertEqual(result.index.tolist(), [6, 7])
        self.assertEqual(result["ou_id"].tolist(), ["zone-a", "ou-b"])
        self.assertEqual(result["Actual"].sum(), 70)

    def test_missing_hierarchy_uses_source_identity(self):
        source = pd.DataFrame({"Organisation unit": ["Unité connue"],
                               "Organisation unit ID": [" known-id "]})
        result = attach_reporting_scope(source, pd.DataFrame())
        self.assertEqual(result.iloc[0]["ou_id"], "known-id")
        self.assertEqual(result.iloc[0]["Zone_ID"], "known-id")
        self.assertEqual(result.iloc[0]["Zone"], "Unité connue")

    def test_unique_name_does_not_override_conflicting_source_id(self):
        source = pd.DataFrame({"Organisation unit": ["Zone A"], "Organisation unit ID": ["new-id"]})
        result = attach_reporting_scope(source, self.metadata())
        self.assertEqual(result.iloc[0]["ou_id"], "new-id")
        self.assertEqual(result.iloc[0]["Zone_ID"], "new-id")

    def test_repeated_metadata_does_not_duplicate_rows_or_make_name_ambiguous(self):
        metadata = pd.concat([self.metadata(), self.metadata()], ignore_index=True)
        source = pd.DataFrame({"Organisation unit": ["Zone A"]})
        result = attach_reporting_scope(source, metadata)
        self.assertEqual(len(result), 1)
        self.assertEqual(result.iloc[0]["ou_id"], "zone-a")

    def test_missing_id_tokens_are_cleaned(self):
        for value in [None, pd.NA, np.nan, "", " nan ", "NONE", "null", "<NA>"]:
            with self.subTest(value=value):
                self.assertEqual(clean_ou_id(value), "")
        self.assertEqual(clean_ou_id("  AbCd1234567 "), "AbCd1234567")

    def test_empty_source_is_supported(self):
        self.assertTrue(attach_reporting_scope(pd.DataFrame(), self.metadata()).empty)


class DatasetAggregationTests(unittest.TestCase):
    def frame(self, actual, expected, rates=None, ids=None):
        return pd.DataFrame({
            "ou_id": ids or [f"child-{i}" for i in range(len(actual))],
            "Dataset Reporting rate": rates if rates is not None else [0] * len(actual),
            "Dataset Actual reports": actual,
            "Dataset Expected reports": expected,
        })

    def test_children_count_dataset_once(self):
        data = self.frame([95] * 10, [100] * 10)
        self.assertEqual(count_datasets_95(data, ["Dataset Reporting rate"], "zone"), 1)

    def test_parent_excludes_children(self):
        data = self.frame([0, 100, 100], [100, 100, 100], ids=["zone", "a", "b"])
        self.assertEqual(count_datasets_95(data, ["Dataset Reporting rate"], "zone"), 0)

    def test_uses_weighted_counts_instead_of_mean_rates(self):
        data = self.frame([1, 940], [1, 1000], rates=[100, 94])
        self.assertEqual(count_datasets_95(data, ["Dataset Reporting rate"], "zone"), 0)

    def test_zero_actual_is_valid(self):
        data = self.frame([0], [10])
        self.assertEqual(count_datasets_95(data, ["Dataset Reporting rate"]), 0)

    def test_nonapplicable_zero_rows_are_excluded(self):
        data = self.frame([0, 95], [0, 100])
        self.assertEqual(count_datasets_95(data, ["Dataset Reporting rate"]), 1)
        self.assertTrue(np.isnan(count_datasets_95(data.iloc[:1], ["Dataset Reporting rate"])))

    def test_missing_and_invalid_counts_do_not_become_partial_success(self):
        for actual, expected in [([95, np.nan], [100, 10]), ([95, np.inf], [100, 10]),
                                 ([95, -1], [100, 10]), ([95, 1], [100, np.nan]),
                                 ([95, 1], [100, -1]), ([95, 1], [100, 0])]:
            with self.subTest(actual=actual, expected=expected):
                data = self.frame(actual, expected)
                self.assertTrue(np.isnan(count_datasets_95(data, ["Dataset Reporting rate"])))

    def test_counts_multiple_datasets_uniquely(self):
        data = self.frame([95, 95], [100, 100])
        data["Other Reporting rate"] = [100, 100]
        data["Other Actual reports"] = [100, 100]
        data["Other Expected reports"] = [100, 100]
        self.assertEqual(count_datasets_95(data, ["Dataset Reporting rate", "Other Reporting rate", "Other Reporting rate"]), 2)

    def test_direct_single_rate_without_counts(self):
        data = pd.DataFrame({"ou_id": ["zone", "child"], "Dataset Reporting rate": [95, 0]})
        self.assertEqual(count_datasets_95(data, ["Dataset Reporting rate"], "zone"), 1)

    def test_direct_multiple_rates_without_counts_are_not_averaged(self):
        data = pd.DataFrame({"ou_id": ["zone", "zone"], "Dataset Reporting rate": [100, 90]})
        self.assertTrue(np.isnan(count_datasets_95(data, ["Dataset Reporting rate"], "zone")))

    def test_child_rates_without_counts_are_not_averaged(self):
        data = pd.DataFrame({"Dataset Reporting rate": [100, 100]})
        self.assertTrue(np.isnan(count_datasets_95(data, ["Dataset Reporting rate"], "zone")))

    def test_normalises_case_spacing_and_ignores_on_time(self):
        data = pd.DataFrame({"ou_id": ["zone"], " Dataset  Reporting RATE ": [0],
                             "dataset Actual reports": [95], "DATASET expected REPORTS": [100],
                             "Other Reporting rate on time": [100]})
        self.assertEqual(count_datasets_95(data, [" Dataset  Reporting RATE ", "Other Reporting rate on time"], "zone"), 1)

    def test_empty_or_missing_rate_information_is_unavailable(self):
        self.assertTrue(np.isnan(count_datasets_95(pd.DataFrame(), [])))
        self.assertTrue(np.isnan(count_datasets_95(pd.DataFrame({"a": [1]}), ["missing"])))
        self.assertTrue(np.isnan(count_datasets_95(pd.DataFrame({"a": [1]}), [])))


if __name__ == "__main__":
    unittest.main()

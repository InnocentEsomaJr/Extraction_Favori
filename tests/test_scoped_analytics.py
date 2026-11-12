import unittest

import pandas as pd

from scoped_analytics import (
    METRIC_LABELS, analytics_to_reporting_frame, build_analytics_query, build_quality_reporting_query,
)


class AnalyticsQueryTests(unittest.TestCase):
    def favorite(self):
        return {
            "dataDimensionItems": [
                {"dataDimensionItemType": "REPORTING_RATE", "reportingRate": {
                    "dataSet": {"id": "DatasetAAAA", "displayName": "Consultations"}, "metric": metric,
                }}
                for metric in ["REPORTING_RATE", "REPORTING_RATE_ON_TIME", "ACTUAL_REPORTS", "ACTUAL_REPORTS_ON_TIME", "EXPECTED_REPORTS"]
            ],
            "columns": [{"id": "dx"}], "rows": [{"id": "ou"}],
            "filters": [{"dimension": "pe", "items": [{"id": "202401"}]}],
            "organisationUnits": [{"id": "ProvinceAAA"}],
        }

    def test_selected_scope_and_period_override_favorite_and_metrics_are_canonical(self):
        params, labels = build_analytics_query(self.favorite(), "202501;202502", "LEVEL-3;CountryAAAA")
        self.assertIn(("dimension", "ou:LEVEL-3;CountryAAAA"), params)
        self.assertIn(("filter", "pe:202501;202502"), params)
        self.assertNotIn(("dimension", "pe:202501;202502"), params)
        self.assertNotIn("ProvinceAAA", str(params))
        self.assertNotIn("202401", str(params))
        self.assertEqual(len([key for key, _ in params if key == "dimension"]), 2)
        self.assertEqual(list(labels.values()), [
            "Consultations Reporting rate", "Consultations Reporting rate on time", "Consultations Actual reports",
            "Consultations Actual reports on time", "Consultations Expected reports",
        ])

    def test_other_explicit_filters_and_aggregation_options_are_preserved(self):
        favorite = self.favorite()
        favorite["filters"].append({"dimension": "CategoryAAA", "items": [{"id": "OptionAAAAA"}]})
        favorite.update(aggregationType="SUM", skipRounding=True, completedOnly=False)
        params, _ = build_analytics_query(favorite, ["202501", "202502"], "LEVEL-3;ProvinceAAA")
        self.assertIn(("filter", "CategoryAAA:OptionAAAAA"), params)
        self.assertIn(("aggregationType", "SUM"), params)
        self.assertIn(("skipRounding", "true"), params)
        self.assertIn(("completedOnly", "false"), params)

    def test_fallback_dx_items_and_detail_category_selections(self):
        favorite = {"columns": [{"dimension": "dx", "items": [{"id": "SetAAAAAAA.ACTUAL_REPORTS", "name": "Consultations Actual reports"}]}],
                    "filters": [{"id": "CategoryAAA"}],
                    "categoryDimensions": [{"category": {"id": "CategoryAAA"}, "categoryOptions": [{"id": "OptionAAAAA"}]}]}
        params, labels = build_analytics_query(favorite, "202501", "ZoneAAAAAAA")
        self.assertEqual(labels, {"SetAAAAAAA.ACTUAL_REPORTS": "Consultations Actual reports"})
        self.assertIn(("filter", "CategoryAAA:OptionAAAAA"), params)

    def test_report_metric_id_only_and_generic_indicator(self):
        favorite = {"dataDimensionItems": [
            {"reportingRate": {"id": "SetAAAAAAA.EXPECTED_REPORTS", "displayName": "Consultations Expected reports"}},
            {"indicator": {"id": "IndicatorAA", "displayName": "Indicateur"}},
        ]}
        _, labels = build_analytics_query(favorite, "202501", "ZoneAAAAAAA")
        self.assertEqual(labels["SetAAAAAAA.EXPECTED_REPORTS"], "Consultations Expected reports")
        self.assertEqual(labels["IndicatorAA"], "Indicateur")

    def test_unknown_or_unselected_extra_dimensions_are_not_silently_aggregated(self):
        favorite = self.favorite()
        favorite["rows"].append({"id": "CategoryAAA"})
        with self.assertRaisesRegex(ValueError, "sélection explicite"):
            build_analytics_query(favorite, "202501", "LEVEL-3;ProvinceAAA")
        with self.assertRaises(ValueError):
            build_analytics_query({}, "202501", "ZoneAAAAAAA")
        with self.assertRaises(ValueError):
            build_analytics_query(self.favorite(), "", "ZoneAAAAAAA")

    def test_tracker_attribute_dimensions_are_not_mistaken_for_category_filters(self):
        favorite = self.favorite()
        favorite["attributeDimensions"] = [{"attribute": {"id": "AttributeAA"}, "filter": "EQ:valeur"}]
        with self.assertRaisesRegex(ValueError, "Tracker"):
            build_analytics_query(favorite, "202501", "ZoneAAAAAAA")

    def test_homonymous_datasets_receive_distinct_consistent_metric_prefixes(self):
        favorite = {"dataDimensionItems": [
            {"reportingRate": {"dataSet": {"id": dataset, "name": "Même nom"}, "metric": metric}}
            for dataset in ["DatasetAAAA", "DatasetBBBB"] for metric in ["REPORTING_RATE", "ACTUAL_REPORTS", "EXPECTED_REPORTS"]
        ]}
        _, labels = build_analytics_query(favorite, "202501", "ZoneAAAAAAA")
        self.assertEqual(len(set(labels.values())), 6)
        self.assertEqual(labels["DatasetAAAA.REPORTING_RATE"], "Même nom (DatasetAAAA) Reporting rate")
        self.assertEqual(labels["DatasetAAAA.ACTUAL_REPORTS"], "Même nom (DatasetAAAA) Actual reports")


class QualityReportingQueryTests(unittest.TestCase):
    def favorite(self, metric='REPORTING_RATE'):
        return {
            'dataDimensionItems': [{
                'reportingRate': {
                    'dataSet': {'id': 'DatasetAAAA', 'displayName': 'Consultations'},
                    'metric': metric,
                },
            }],
            'filters': [{'dimension': 'pe', 'items': [{'id': '202401'}]}],
            'organisationUnits': [{'id': 'OldZoneAAAA'}],
        }

    def test_default_does_not_change_the_original_favorite_selection(self):
        params, labels = build_analytics_query(self.favorite(), '202609', 'ZoneAAAAAAA')
        self.assertEqual(list(labels), ['DatasetAAAA.REPORTING_RATE'])
        self.assertIn(('dimension', 'dx:DatasetAAAA.REPORTING_RATE'), params)

    def test_rate_only_favorite_requests_the_counts_needed_for_the_quality_score(self):
        params, labels = build_analytics_query(
            self.favorite(), '202609', 'ZoneAAAAAAA', complete_reporting_metrics=True,
        )
        expected_ids = {f'DatasetAAAA.{metric}' for metric in METRIC_LABELS}
        self.assertEqual(set(labels), expected_ids)
        self.assertEqual(set(labels.values()), {f'Consultations {suffix}' for suffix in METRIC_LABELS.values()})
        dimension = next(value for key, value in params if key == 'dimension' and value.startswith('dx:'))
        self.assertEqual(set(dimension[3:].split(';')), expected_ids)
        self.assertIn(('dimension', 'ou:ZoneAAAAAAA'), params)
        self.assertIn(('filter', 'pe:202609'), params)
        self.assertNotIn('OldZoneAAAA', str(params))
        self.assertNotIn('202401', str(params))

    def test_counts_only_favorite_can_also_complete_its_reporting_family(self):
        _, labels = build_analytics_query(
            self.favorite('ACTUAL_REPORTS'), '202609', 'ZoneAAAAAAA', complete_reporting_metrics=True,
        )
        self.assertEqual(len(labels), 5)
        self.assertEqual(labels['DatasetAAAA.REPORTING_RATE_ON_TIME'], 'Consultations Reporting rate on time')
        self.assertEqual(labels['DatasetAAAA.EXPECTED_REPORTS'], 'Consultations Expected reports')

    def test_extra_filters_and_aggregation_options_are_preserved_for_each_selected_scope(self):
        favorite = self.favorite()
        favorite['filters'].append({'dimension': 'CategoryAAA', 'items': [{'id': 'OptionAAAAA'}]})
        favorite.update(aggregationType='SUM', completedOnly=False)
        for scope in ['LEVEL-3;CountryAAAA', 'LEVEL-3;ProvinceAAA', 'ZoneAAAAAAA']:
            with self.subTest(scope=scope):
                params, _ = build_analytics_query(
                    favorite, ['202609', '202610'], scope, complete_reporting_metrics=True,
                )
                self.assertIn(('dimension', 'ou:' + scope), params)
                self.assertIn(('filter', 'pe:202609;202610'), params)
                self.assertIn(('filter', 'CategoryAAA:OptionAAAAA'), params)
                self.assertIn(('aggregationType', 'SUM'), params)
                self.assertIn(('completedOnly', 'false'), params)

    def test_existing_metrics_are_never_requested_twice(self):
        favorite = AnalyticsQueryTests().favorite()
        favorite['columns'] = [{
            'dimension': 'dx', 'items': [{'id': 'DatasetAAAA.ACTUAL_REPORTS'}],
        }]
        params, labels = build_analytics_query(
            favorite, '202609', 'ZoneAAAAAAA', complete_reporting_metrics=True,
        )
        self.assertEqual(len(labels), 5)
        dimension = next(value for key, value in params if key == 'dimension' and value.startswith('dx:'))
        self.assertEqual(len(dimension[3:].split(';')), 5)

    def test_homonymous_datasets_keep_distinct_and_matching_metric_prefixes(self):
        favorite = {'dataDimensionItems': [{
            'reportingRate': {'dataSet': {'id': uid, 'displayName': 'Consultations'}, 'metric': 'REPORTING_RATE'},
        } for uid in ['DatasetAAAA', 'DatasetBBBB']]}
        _, labels = build_analytics_query(
            favorite, '202609', 'ZoneAAAAAAA', complete_reporting_metrics=True,
        )
        self.assertEqual(len(set(labels.values())), 10)
        for uid in ['DatasetAAAA', 'DatasetBBBB']:
            for metric, suffix in METRIC_LABELS.items():
                self.assertEqual(labels[f'{uid}.{metric}'], f'Consultations ({uid}) {suffix}')

    def test_generic_indicators_are_not_mistaken_for_dataset_ids(self):
        favorite = {'dataDimensionItems': [{
            'indicator': {'id': 'IndicatorAA', 'displayName': 'Complétude générale'},
        }]}
        with self.assertRaisesRegex(ValueError, 'aucun dataset de reporting'):
            build_analytics_query(favorite, '202609', 'ZoneAAAAAAA', complete_reporting_metrics=True)
        _, labels = build_analytics_query(favorite, '202609', 'ZoneAAAAAAA')
        self.assertEqual(labels, {'IndicatorAA': 'Complétude générale'})

    def test_quality_source_excludes_generic_counts_with_a_similar_label(self):
        favorite = self.favorite()
        favorite['dataDimensionItems'].extend([
            {'indicator': {'id': 'IndicatorAA', 'displayName': 'Consultations Actual reports'}},
            {'dataElement': {'id': 'ElementAAAA', 'displayName': 'Autre donnée'}},
        ])
        params, labels = build_quality_reporting_query(favorite, '202609', 'ZoneAAAAAAA')
        self.assertEqual(set(labels), {f'DatasetAAAA.{metric}' for metric in METRIC_LABELS})
        self.assertNotIn('IndicatorAA', str(params))
        self.assertNotIn('ElementAAAA', str(params))
        self.assertEqual(labels['DatasetAAAA.ACTUAL_REPORTS'], 'Consultations Actual reports')
        self.assertEqual(labels['DatasetAAAA.EXPECTED_REPORTS'], 'Consultations Expected reports')

    def test_quality_source_keeps_reporting_prefixes_aligned_after_label_collisions(self):
        favorite = self.favorite()
        favorite['dataDimensionItems'].append({
            'indicator': {'id': 'IndicatorAA', 'displayName': 'Consultations Reporting rate'},
        })
        _, labels = build_quality_reporting_query(favorite, '202609', 'ZoneAAAAAAA')
        prefixes = {
            labels[f'DatasetAAAA.{metric}'][:-len(suffix)].strip()
            for metric, suffix in METRIC_LABELS.items()
        }
        self.assertEqual(len(prefixes), 1)

    def test_quality_source_preserves_the_selected_scope_period_and_other_filters(self):
        favorite = self.favorite()
        favorite['filters'].append({'dimension': 'CategoryAAA', 'items': [{'id': 'OptionAAAAA'}]})
        favorite.update(skipRounding=True, completedOnly=False)
        params, _ = build_quality_reporting_query(favorite, '202609;202610', 'LEVEL-3;ProvinceAAA')
        self.assertIn(('dimension', 'ou:LEVEL-3;ProvinceAAA'), params)
        self.assertIn(('filter', 'pe:202609;202610'), params)
        self.assertIn(('filter', 'CategoryAAA:OptionAAAAA'), params)
        self.assertIn(('skipRounding', 'true'), params)
        self.assertIn(('completedOnly', 'false'), params)


class AnalyticsFrameTests(unittest.TestCase):
    def response(self):
        return {
            "headers": [{"name": "dx"}, {"name": "ou"}, {"name": "value"}],
            "rows": [["DatasetAAAA.ACTUAL_REPORTS", "ZoneAAAAAAA", "10"],
                     ["DatasetAAAA.ACTUAL_REPORTS", "ZoneBBBBBBB", "20"],
                     ["DatasetAAAA.EXPECTED_REPORTS", "ZoneAAAAAAA", "10"]],
            "metaData": {"items": {"ZoneAAAAAAA": {"name": "Homonyme"}, "ZoneBBBBBBB": {"name": "Homonyme"}, "DatasetAAAA": {"name": "Consultations"}}},
        }

    def test_homonymous_org_units_stay_separate_by_identifier(self):
        frame = analytics_to_reporting_frame(self.response())
        self.assertEqual(frame["Organisation unit"].tolist(), ["Homonyme", "Homonyme"])
        self.assertEqual(frame["Organisation unit ID"].tolist(), ["ZoneAAAAAAA", "ZoneBBBBBBB"])
        self.assertEqual(frame["Consultations Actual reports"].tolist(), [10, 20])
        self.assertTrue(pd.isna(frame.loc[1, "Consultations Expected reports"]))

    def test_requested_columns_order_and_absent_metrics_are_preserved(self):
        labels = {"DatasetAAAA.EXPECTED_REPORTS": "Consultations Expected reports", "DatasetAAAA.ACTUAL_REPORTS": "Consultations Actual reports", "DatasetAAAA.REPORTING_RATE": "Consultations Reporting rate"}
        frame = analytics_to_reporting_frame(self.response(), labels)
        self.assertEqual(frame.columns.tolist(), ["Organisation unit", "Organisation unit ID", *labels.values()])
        self.assertTrue(frame["Consultations Reporting rate"].isna().all())

    def test_metadata_enriches_only_identifier_fallbacks_and_keeps_explicit_names(self):
        response = self.response()
        labels = {"DatasetAAAA.ACTUAL_REPORTS": "DatasetAAAA Actual reports",
                  "DatasetAAAA.EXPECTED_REPORTS": "Libellé choisi Expected reports"}
        frame = analytics_to_reporting_frame(response, labels)
        self.assertEqual(frame.columns.tolist(), ["Organisation unit", "Organisation unit ID",
                                                "Consultations Actual reports", "Libellé choisi Expected reports"])
        response["metaData"]["items"].update({"IndicatorAA": {"name": "Indicateur lisible"}})
        frame = analytics_to_reporting_frame({"rows": [], "metaData": response["metaData"]}, {"IndicatorAA": "IndicatorAA"})
        self.assertIn("Indicateur lisible", frame.columns)

    def test_item_column_identity_survives_metadata_enrichment_and_missing_metrics(self):
        labels = {f'DatasetAAAA.{metric}': f'DatasetAAAA {suffix}' for metric, suffix in METRIC_LABELS.items()}
        frame = analytics_to_reporting_frame(self.response(), labels)
        self.assertEqual(frame.attrs['reporting_item_labels'], {
            f'DatasetAAAA.{metric}': f'Consultations {suffix}' for metric, suffix in METRIC_LABELS.items()
        })
        self.assertTrue(frame[frame.attrs['reporting_item_labels']['DatasetAAAA.REPORTING_RATE']].isna().all())

    def test_empty_item_column_identity_uses_final_names_and_excludes_generic_indicators(self):
        frame = analytics_to_reporting_frame({'rows': [], 'metaData': {'items': {
            'DatasetAAAA': {'displayName': 'Consultations'},
        }}}, {'DatasetAAAA.ACTUAL_REPORTS': 'DatasetAAAA Actual reports', 'IndicatorAA': 'Faux Actual reports'})
        self.assertEqual(frame.attrs['reporting_item_labels'], {'DatasetAAAA.ACTUAL_REPORTS': 'Consultations Actual reports'})
        self.assertIn('Faux Actual reports', frame.columns)

    def test_metadata_enrichment_resolves_homonymous_dataset_names_without_collision(self):
        response = self.response()
        response["rows"].append(["DatasetBBBB.ACTUAL_REPORTS", "ZoneAAAAAAA", "7"])
        response["metaData"]["items"]["DatasetBBBB"] = {"name": "Consultations"}
        labels = {"DatasetAAAA.ACTUAL_REPORTS": "DatasetAAAA Actual reports",
                  "DatasetBBBB.ACTUAL_REPORTS": "DatasetBBBB Actual reports"}
        frame = analytics_to_reporting_frame(response, labels)
        self.assertIn("Consultations (DatasetAAAA) Actual reports", frame.columns)
        self.assertIn("Consultations (DatasetBBBB) Actual reports", frame.columns)
        self.assertEqual(frame.loc[0, "Consultations (DatasetBBBB) Actual reports"], 7)

    def test_duplicate_values_raise_instead_of_averaging(self):
        response = self.response()
        response["rows"].append(response["rows"][0])
        with self.assertRaisesRegex(ValueError, "plusieurs valeurs"):
            analytics_to_reporting_frame(response)

    def test_extra_dimensions_and_invalid_values_are_explicit(self):
        response = self.response()
        response["headers"].append({"name": "pe"})
        with self.assertRaisesRegex(ValueError, "uniquement"):
            analytics_to_reporting_frame(response)
        response = self.response()
        response["rows"][0][2] = "indisponible"
        with self.assertRaisesRegex(ValueError, "numérique"):
            analytics_to_reporting_frame(response)

    def test_empty_response_keeps_identifier_and_requested_columns(self):
        frame = analytics_to_reporting_frame({"rows": []}, {"set.ACTUAL_REPORTS": "DS Actual reports"})
        self.assertTrue(frame.empty)
        self.assertEqual(frame.columns.tolist(), ["Organisation unit", "Organisation unit ID", "DS Actual reports"])

    def test_reporting_column_metadata_excludes_generic_indicators_with_matching_names(self):
        response = self.response()
        response['rows'].append(['IndicatorAA', 'ZoneAAAAAAA', '999'])
        labels = {
            'DatasetAAAA.ACTUAL_REPORTS': 'DS Actual reports',
            'DatasetAAAA.EXPECTED_REPORTS': 'DS Expected reports',
            'IndicatorAA': 'Faux Actual reports',
        }
        frame = analytics_to_reporting_frame(response, labels)
        self.assertEqual(frame.attrs['reporting_columns'], ['DS Actual reports', 'DS Expected reports'])
        self.assertIn('Faux Actual reports', frame.columns)
        self.assertNotIn('Faux Actual reports', frame.attrs['reporting_columns'])

    def test_empty_response_also_certifies_only_dataset_reporting_columns(self):
        frame = analytics_to_reporting_frame({'rows': []}, {
            'DatasetAAAA.ACTUAL_REPORTS': 'DS Actual reports', 'IndicatorAA': 'Faux Actual reports',
        })
        self.assertTrue(frame.empty)
        self.assertEqual(frame.attrs['reporting_columns'], ['DS Actual reports'])


if __name__ == "__main__":
    unittest.main()

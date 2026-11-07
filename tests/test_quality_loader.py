import unittest

import numpy as np
import pandas as pd
import requests

from scoped_analytics import analytics_to_reporting_frame, build_quality_reporting_query
from test_dashboard_categorisation import dashboard_functions


class QualityLoaderTests(unittest.TestCase):
    def loader(self, client):
        return dashboard_functions(
            'get_quality_reporting_data', requests=requests,
            get_dhis2_client=lambda: client,
            build_quality_reporting_query=build_quality_reporting_query,
            analytics_to_reporting_frame=analytics_to_reporting_frame,
        )['get_quality_reporting_data']

    def test_rate_only_favorite_loads_counts_for_the_selected_zone_and_period(self):
        calls = []

        class Client:
            def get(self, endpoint, params=None):
                calls.append((endpoint, params))
                if endpoint.startswith('visualizations/'):
                    return {'dataDimensionItems': [{'reportingRate': {
                        'id': 'DatasetAAAA.REPORTING_RATE',
                        'dataSet': {'id': 'DatasetAAAA', 'displayName': 'Consultations'},
                    }}]}
                return {
                    'headers': [{'name': 'ou'}, {'name': 'dx'}, {'name': 'value'}],
                    'rows': [['ZoneAAAAAAA', 'DatasetAAAA.' + metric, value] for metric, value in [
                        ('REPORTING_RATE', '90'), ('REPORTING_RATE_ON_TIME', '80'),
                        ('ACTUAL_REPORTS', '9'), ('ACTUAL_REPORTS_ON_TIME', '8'), ('EXPECTED_REPORTS', '10'),
                    ]],
                }

            def post(self, *args, **kwargs):
                raise AssertionError('Le tableau qualité ne doit pas analyser les violations.')

        frame = self.loader(Client())('FavoriteAAA', period='202609', ou_dimension='ZoneAAAAAAA')
        self.assertEqual(len(calls), 2)
        self.assertEqual(calls[1][0], 'analytics.json')
        self.assertIn(('dimension', 'ou:ZoneAAAAAAA'), calls[1][1])
        self.assertIn(('filter', 'pe:202609'), calls[1][1])
        self.assertEqual(frame['Organisation unit ID'].tolist(), ['ZoneAAAAAAA'])
        self.assertEqual(frame['Consultations Actual reports'].tolist(), [9])
        self.assertEqual(frame['Consultations Expected reports'].tolist(), [10])
        self.assertEqual(frame['Consultations Actual reports on time'].tolist(), [8])

    def test_reporting_failure_is_explicit_without_copying_server_exception_text(self):
        class Client:
            def get(self, *args, **kwargs):
                response = requests.Response()
                response.status_code = 504
                raise requests.HTTPError('password=fixture', response=response)

        frame = self.loader(Client())('FavoriteAAA', period='202609', ou_dimension='ZoneAAAAAAA')
        self.assertTrue(frame.empty)
        message = ' '.join(frame.attrs['quality_warnings'])
        self.assertIn('504', message)
        self.assertNotIn('password', message)

    def test_missing_numbers_render_as_dashes_while_source_values_remain_numeric(self):
        styler = dashboard_functions('style_taux', 'style_score', '_build_styler')['_build_styler']
        frame = pd.DataFrame({'Moyenne complétude / promptitude (%)': [89.0, np.nan]})
        html = styler(frame, score_cols=['Moyenne complétude / promptitude (%)'], missing_text='—').to_html()
        self.assertIn('—</td>', html)
        self.assertNotIn('>None</td>', html)
        self.assertNotIn('>nan</td>', html)
        self.assertTrue(pd.api.types.is_numeric_dtype(frame['Moyenne complétude / promptitude (%)']))
        self.assertTrue(pd.isna(frame.iloc[1, 0]))


if __name__ == '__main__':
    unittest.main()

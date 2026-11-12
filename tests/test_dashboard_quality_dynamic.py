"""Le score qualité suit les données et filtres DHIS2 sans lancer les violations."""

from datetime import date
import unittest
from unittest import mock

import pandas as pd
import pyarrow as pa
import streamlit as st
from streamlit.testing.v1 import AppTest

from categorisation import TOTAL_COL
import test_dashboard_categorisation as dashboard_test
import test_dashboard_filters as filter_test
import test_dashboard_validation_details as detail_test


QUALITY_DATA = '''
def get_data(favori_id, period=None, cache_user=None, ou_dimension=None):
    units = [
        ("ZoneAAAAAAA", "ZS Commune", "ProvinceAAA", 100, 90, 75),
        ("ZoneBBBBBBB", "ZS B", "ProvinceAAA", 200, 100, 80),
        ("ZoneCCCCCCC", "ZS Commune", "ProvinceBBB", 100, 98, 97),
    ]
    if ou_dimension == "LEVEL-3;CountryAAAA":
        selected = units
    elif ou_dimension and ou_dimension.startswith("LEVEL-3;"):
        parent = ou_dimension.split(";", 1)[1]
        selected = [unit for unit in units if unit[2] == parent]
    else:
        selected = [unit for unit in units if unit[0] == ou_dimension]
    periods = str(period).split(";")
    rows = []
    for uid, name, province, expected, january_actual, january_on_time in selected:
        actual = on_time = expected_total = 0
        for month in periods:
            expected_total += expected
            if month.endswith("02"):
                actual += {"ZoneAAAAAAA": 40, "ZoneBBBBBBB": 180, "ZoneCCCCCCC": 60}[uid]
                on_time += {"ZoneAAAAAAA": 30, "ZoneBBBBBBB": 100, "ZoneCCCCCCC": 30}[uid]
            else:
                actual += january_actual
                on_time += january_on_time
        row = {"Organisation unit": name, "Organisation unit ID": uid}
        for number in range(10):
            prefix = f"Dataset {number} "
            row[prefix + "Actual reports"] = actual
            row[prefix + "Expected reports"] = expected_total
            row[prefix + "Actual reports on time"] = on_time
            row[prefix + "Reporting rate"] = actual / expected_total * 100
            row[prefix + "Reporting rate on time"] = on_time / expected_total * 100
        rows.append(row)
    return pd.DataFrame(rows)
'''


QUALITY_REPORTING = '''
def get_quality_reporting_data(favori_id, period=None, cache_user=None, ou_dimension=None):
    requests = list(st.session_state.get("_quality_requests", []))
    requests.append({"period": period, "ou_dimension": ou_dimension})
    st.session_state["_quality_requests"] = requests
    return get_data(favori_id, period=period, cache_user=cache_user, ou_dimension=ou_dimension)
'''


GATEWAY_FAILURE = '''
def get_dhis2_client():
    class Client:
        def post(self, endpoint, data=None):
            import requests
            calls = list(st.session_state.get("_quality_validation_calls", []))
            calls.append(dict(data))
            st.session_state["_quality_validation_calls"] = calls
            response = requests.Response()
            response.status_code = 504
            raise requests.HTTPError("Relais indisponible", response=response)
    return Client()
'''


SCREENSHOT_CATALOG = filter_test.CATALOG_FIXTURE.replace('"ZS Commune", pa', '"Ango", pa')
SCREENSHOT_DATA = QUALITY_DATA.replace(
    '"ProvinceAAA", 100, 90, 75', '"ProvinceAAA", 100, 100, 68',
).replace(
    '"ProvinceAAA", 200, 100, 80', '"ProvinceAAA", 200, 80, 60',
).replace('range(10)', 'range(9)').replace(
    '    return pd.DataFrame(rows)',
    '''    frame = pd.DataFrame(rows)
    for metric in ["Actual reports", "Expected reports", "Actual reports on time", "Reporting rate", "Reporting rate on time"]:
        frame["Dataset inutilisé " + metric] = np.nan
    return frame''',
)

FAILED_EXTRA_REPORTING = '''
def get_quality_reporting_data(favori_id, period=None, cache_user=None, ou_dimension=None):
    attempts = list(st.session_state.get("_extra_quality_attempts", []))
    attempts.append({"period": period, "ou_dimension": ou_dimension})
    st.session_state["_extra_quality_attempts"] = attempts
    return None
'''

FOUR_CURRENT_VIOLATIONS = '''
def get_dhis2_client():
    class Client:
        def post(self, endpoint, data=None):
            queries = list(st.session_state.get("_screenshot_validation_queries", []))
            queries.append(dict(data))
            st.session_state["_screenshot_validation_queries"] = queries
            rows = []
            for number, (left, right) in enumerate([(84, 59), (20, 18), (567, 566), (6, 5)]):
                centre = f"Centre{number:05d}"
                rows.append({
                    "organisationUnitId": centre,
                    "organisationUnitDisplayName": f"Centre {number}",
                    "organisationUnitPath": f"/CountryAAAA/ProvinceAAA/{data['ou']}/{centre}",
                    "validationRuleId": f"Rule{number:07d}",
                    "validationRuleDescription": f"Règle {number}",
                    "periodId": data["startDate"][:7].replace("-", ""),
                    "leftSideValue": left, "rightSideValue": right, "operator": "<=",
                    "importance": "MEDIUM",
                })
            return rows
    return Client()
'''


class DynamicQualityDashboardTests(unittest.TestCase):
    def widget(self, app, kind, label):
        return next(widget for widget in getattr(app, kind) if widget.label == label)

    def run_app(self, app):
        app.run(timeout=30)
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])

    def new_app(self, client=detail_test.VALIDATION_CLIENT, reporting=QUALITY_REPORTING, data=QUALITY_DATA, catalog=filter_test.CATALOG_FIXTURE):
        source = dashboard_test.app_fixture_source({
            "get_organisation_catalog": catalog,
            "get_data": data,
            "get_quality_reporting_data": reporting,
            "get_dhis2_client": client,
            "get_validation_detail_metadata": detail_test.DETAIL_METADATA,
        })
        app = AppTest.from_string(source, default_timeout=30)
        app.session_state['connected'] = True
        app.session_state['dhis2_user'] = 'fixture-quality-' + self._testMethodName
        app.session_state['dhis2_pass'] = 'fixture-password'
        return app

    def summary(self, app):
        return app.session_state['_test_export_frame'].copy(deep=True)

    def select_zone(self, app, province, zone):
        self.widget(app, 'selectbox', 'Province de l’analyse').select(province)
        self.run_app(app)
        self.widget(app, 'selectbox', 'Zone de santé de l’analyse').select(zone)
        self.run_app(app)

    def test_quality_is_available_before_any_validation_request(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            frame = self.summary(app)
            self.assertEqual(len(frame), 3)
            self.assertEqual(sorted(frame['Reports_Actual'].tolist()), [900, 980, 1000])
            self.assertEqual(sorted(frame['Score de qualité'].tolist()), [45, 82.5, 97.5])
            self.assertEqual(sorted(frame['Score de qualité /20'].tolist()), [0, 15, 20])
            self.assertTrue(frame['Règles violées (M)'].isna().all())
            self.assertTrue(frame['Règles violées (M-1)'].isna().all())
            self.assertTrue(frame[TOTAL_COL].isna().all())
            self.assertNotIn('categorisation_result', app.session_state)
            self.assertNotIn('_fixture_detail_queries', app.session_state)
            quality_grid = next(grid for grid in app.dataframe if TOTAL_COL in grid.value.columns)
            displayed = pa.ipc.open_stream(quality_grid.proto.styler.display_values).read_all().to_pandas()
            self.assertEqual(displayed['Règles violées (M)'].tolist(), ['—', '—', '—'])
            self.assertEqual(sorted(displayed['Score de qualité'].tolist()), ['45.00', '82.50', '97.50'])

    def test_zone_homonyms_and_dates_use_the_selected_ids_and_months(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.select_zone(app, 'ProvinceAAA', 'ZoneAAAAAAA')
            january = self.summary(app)
            self.assertEqual(january['Reports_Actual'].tolist(), [900])
            self.assertEqual(january['Complétude globale (%)'].tolist(), [90])
            self.assertEqual(january['Promptitude globale (%)'].tolist(), [75])
            self.assertEqual(january['Score de qualité'].tolist(), [82.5])

            year = int(self.widget(app, 'selectbox', "Choisir l'année :").value)
            self.widget(app, 'date_input', 'Date de début').set_value(date(year, 2, 1))
            self.widget(app, 'date_input', 'Date de fin').set_value(date(year, 2, 28))
            self.run_app(app)
            february = self.summary(app)
            self.assertEqual(february['Reports_Actual'].tolist(), [400])
            self.assertEqual(february['Score de qualité'].tolist(), [35])
            self.assertEqual(february['Score de qualité /20'].tolist(), [0])

            self.select_zone(app, 'ProvinceBBB', 'ZoneCCCCCCC')
            other = self.summary(app)
            self.assertEqual(len(other), 1)
            self.assertEqual(other['Reports_Actual'].tolist(), [600])
            self.assertEqual(other['Score de qualité'].tolist(), [45])
            self.assertEqual(app.session_state['_quality_requests'][-1], {
                'period': f'{year}02', 'ou_dimension': 'ZoneCCCCCCC',
            })
            self.assertNotIn('_fixture_detail_queries', app.session_state)

    def test_gateway_error_does_not_erase_quality_metrics(self):
        app = self.new_app(client=GATEWAY_FAILURE)
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            before = self.summary(app)
            self.widget(app, 'button', "Lancer l'analyse des violations").click()
            self.run_app(app)
            after = self.summary(app)
            for column in [
                'Zone de santé', 'Complétude globale (%)', 'Promptitude globale (%)',
                'Reports_Actual', 'Score de qualité', 'Score de qualité /20',
            ]:
                pd.testing.assert_series_equal(after[column], before[column])
            self.assertTrue(after['Règles violées (M)'].isna().all())
            self.assertTrue(any('504' in message for message in app.session_state['categorisation_result']['validation_warnings']))
            self.assertEqual(len(app.session_state['_quality_validation_calls']), 3)

    def test_province_all_zones_and_country_are_rebuilt_without_running_validation(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.assertEqual(len(self.summary(app)), 3)
            self.widget(app, 'selectbox', 'Province de l’analyse').select('ProvinceAAA')
            self.run_app(app)
            self.assertEqual(sorted(self.summary(app)['Reports_Actual'].tolist()), [900, 1000])
            self.widget(app, 'selectbox', 'Zone de santé de l’analyse').select('ZoneBBBBBBB')
            self.run_app(app)
            self.assertEqual(self.summary(app)['Reports_Actual'].tolist(), [1000])
            self.widget(app, 'selectbox', 'Zone de santé de l’analyse').select('')
            self.run_app(app)
            self.assertEqual(sorted(self.summary(app)['Reports_Actual'].tolist()), [900, 1000])
            self.widget(app, 'selectbox', 'Province de l’analyse').select('')
            self.run_app(app)
            self.assertEqual(sorted(self.summary(app)['Reports_Actual'].tolist()), [900, 980, 1000])
            self.assertNotIn('_fixture_detail_queries', app.session_state)

    def test_excel_contains_dynamic_quality_before_validation_and_keeps_unknown_counts_empty(self):
        app = self.new_app()
        downloads = {}
        real_download = st.download_button

        def capture(label, *args, **kwargs):
            downloads[label] = kwargs.get('data', args[0] if args else None)
            return real_download(label, *args, **kwargs)

        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')), \
                mock.patch('streamlit.download_button', side_effect=capture):
            self.run_app(app)
            self.select_zone(app, 'ProvinceBBB', 'ZoneCCCCCCC')
            year = int(self.widget(app, 'selectbox', "Choisir l'année :").value)
            start, end = date(year, 1, 5), date(year, 1, 10)
            self.widget(app, 'date_input', 'Date de début').set_value(start)
            self.widget(app, 'date_input', 'Date de fin').set_value(end)
            self.run_app(app)
            _, rows = detail_test.excel_sheet(downloads['Télécharger Excel'], 'Resultats_Regles')
            self.assertEqual(len(rows), 1)
            self.assertEqual(rows[0]['Reports_Actual'], 980)
            self.assertEqual(rows[0]['Score de qualité'], 97.5)
            self.assertEqual(rows[0]['Score de qualité /20'], 20)
            self.assertIsNone(rows[0]['Règles violées (M)'])
            self.assertIsNone(rows[0][TOTAL_COL])
            _, context_rows = detail_test.excel_sheet(downloads['Télécharger Excel'], 'Contexte_Validation')
            context = {row['Paramètre']: row['Valeur'] for row in context_rows}
            self.assertEqual(context['ID unité parent'], 'ZoneCCCCCCC')
            self.assertEqual(context['Date de début'], start.isoformat())
            self.assertEqual(context['Date de fin'], end.isoformat())
            self.assertEqual(context['Mois des indicateurs de déclaration'], f'{year}01')
            self.assertTrue(any(label in context['Source des indicateurs'] for label in ['Analytics', 'Synthèse']))
            self.assertEqual(context['Analyse des violations'], 'Non lancée')
            self.assertNotIn('_fixture_detail_queries', app.session_state)

    def test_scope_switch_replaces_quality_and_discards_old_validation_results(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.select_zone(app, 'ProvinceAAA', 'ZoneAAAAAAA')
            self.widget(app, 'button', "Lancer l'analyse des violations").click()
            self.run_app(app)
            self.assertEqual(self.summary(app)['Règles violées (M)'].tolist(), [1])
            self.assertEqual(len(app.session_state['_fixture_detail_queries']), 1)

            self.widget(app, 'text_input', 'Rechercher une zone').input('commune')
            self.widget(app, 'selectbox', 'Afficher les résultats').select('Scores complets')
            self.run_app(app)
            current_grid = next(grid for grid in app.dataframe if TOTAL_COL in grid.value.columns)
            self.assertTrue(current_grid.value.empty)

            self.select_zone(app, 'ProvinceAAA', 'ZoneBBBBBBB')
            frame = self.summary(app)
            self.assertEqual(frame['Zone de santé'].tolist(), ['ZS B'])
            self.assertEqual(frame['Reports_Actual'].tolist(), [1000])
            self.assertEqual(frame['Score de qualité'].tolist(), [45])
            self.assertTrue(frame['Règles violées (M)'].isna().all())
            self.assertNotIn('categorisation_result', app.session_state)
            self.assertEqual(len(app.session_state['_fixture_detail_queries']), 1)
            self.assertEqual(self.widget(app, 'text_input', 'Rechercher une zone').value, '')
            self.assertEqual(self.widget(app, 'selectbox', 'Afficher les résultats').value, 'Toutes')
            updated_grid = next(grid for grid in app.dataframe if TOTAL_COL in grid.value.columns)
            self.assertEqual(updated_grid.value['Zone de santé'].tolist(), ['ZS B'])

    def test_reporting_failure_is_missing_and_does_not_invent_scores(self):
        missing_source = QUALITY_DATA.replace(
            'return pd.DataFrame(rows)',
            "return pd.DataFrame(rows)[['Organisation unit', 'Organisation unit ID']]",
        )
        app = self.new_app(data=missing_source, reporting='''
def get_quality_reporting_data(favori_id, period=None, cache_user=None, ou_dimension=None):
    return None
''')
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            frame = self.summary(app)
            self.assertEqual(len(frame), 3)
            self.assertTrue(frame['Complétude globale (%)'].isna().all())
            self.assertTrue(frame['Promptitude globale (%)'].isna().all())
            self.assertTrue(frame['Reports_Actual'].isna().all())
            self.assertTrue(frame['Score de qualité'].isna().all())
            self.assertTrue(frame[TOTAL_COL].isna().all())
            self.assertGreater(len(app.warning) + len(app.info), 0)

    def test_visible_comparative_values_feed_ango_quality_without_an_extra_failed_query(self):
        app = self.new_app(
            data=SCREENSHOT_DATA, catalog=SCREENSHOT_CATALOG,
            reporting=FAILED_EXTRA_REPORTING, client=FOUR_CURRENT_VIOLATIONS,
        )
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.select_zone(app, 'ProvinceAAA', 'ZoneAAAAAAA')
            current = self.summary(app)
            self.assertEqual(current['Zone de santé'].tolist(), ['Ango'])
            self.assertEqual(current['Complétude globale (%)'].tolist(), [100])
            self.assertEqual(current['Promptitude globale (%)'].tolist(), [68])
            self.assertEqual(current['Reports_Actual'].tolist(), [900])
            self.assertEqual(current['Nombre de datasets complétude ≥95 %'].tolist(), [9])
            self.assertEqual(current['Score de qualité'].tolist(), [84])
            self.assertEqual(current['Score de qualité /20'].tolist(), [15])
            self.assertNotIn('_extra_quality_attempts', app.session_state)

            self.widget(app, 'button', "Lancer l'analyse des violations").click()
            self.run_app(app)
            analysed = self.summary(app)
            self.assertEqual(analysed['Règles violées (M)'].tolist(), [4])
            self.assertEqual(analysed['Score de qualité'].tolist(), [84])
            self.assertEqual(analysed['Nombre de datasets complétude ≥95 %'].tolist(), [9])
            self.assertEqual(len(app.session_state['_screenshot_validation_queries']), 1)
            self.assertNotIn('_extra_quality_attempts', app.session_state)

            self.select_zone(app, 'ProvinceAAA', 'ZoneBBBBBBB')
            other = self.summary(app)
            self.assertEqual(other['Reports_Actual'].tolist(), [720])
            self.assertEqual(other['Complétude globale (%)'].tolist(), [40])
            self.assertEqual(other['Promptitude globale (%)'].tolist(), [30])
            self.assertEqual(other['Score de qualité'].tolist(), [35])
            self.assertTrue(other['Règles violées (M)'].isna().all())
            self.assertEqual(len(app.session_state['_screenshot_validation_queries']), 1)
            self.assertNotIn('_extra_quality_attempts', app.session_state)

    def test_month_outside_the_visible_dashboard_loads_its_own_reporting_values(self):
        app = self.new_app()
        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')):
            self.run_app(app)
            self.select_zone(app, 'ProvinceAAA', 'ZoneAAAAAAA')
            self.assertEqual(self.summary(app)['Reports_Actual'].tolist(), [900])
            self.assertNotIn('_quality_requests', app.session_state)
            year = int(self.widget(app, 'selectbox', "Choisir l'année :").value)
            self.widget(app, 'date_input', 'Date de début').set_value(date(year, 2, 1))
            self.widget(app, 'date_input', 'Date de fin').set_value(date(year, 2, 28))
            self.run_app(app)
            frame = self.summary(app)
            self.assertEqual(frame['Reports_Actual'].tolist(), [400])
            self.assertEqual(frame['Complétude globale (%)'].tolist(), [40])
            self.assertEqual(frame['Promptitude globale (%)'].tolist(), [30])
            self.assertEqual(frame['Score de qualité'].tolist(), [35])
            self.assertEqual(app.session_state['_quality_requests'][-1], {
                'period': f'{year}02', 'ou_dimension': 'ZoneAAAAAAA',
            })
            self.assertNotIn('_fixture_detail_queries', app.session_state)


if __name__ == '__main__':
    unittest.main()

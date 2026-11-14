"""Périodes fixes, sommes exportées et dates libres des règles, sans réseau."""

from contextlib import contextmanager
from datetime import date
import unittest
from unittest import mock

import streamlit as st
from streamlit.testing.v1 import AppTest

import test_dashboard_categorisation as dashboard_test
import test_dashboard_filters as filter_test
import test_dashboard_validation_details as detail_test


PERIOD_DATA = '''
def get_data(favori_id, period=None, cache_user=None, ou_dimension=None, unit_ids=None):
    calls = list(st.session_state.get("_fixed_data_requests", []))
    calls.append({"period": period, "ou_dimension": ou_dimension, "unit_ids": unit_ids})
    st.session_state["_fixed_data_requests"] = calls
    units = [
        ("ZoneAAAAAAA", "ZS Commune", "ProvinceAAA", 1),
        ("ZoneBBBBBBB", "ZS B", "ProvinceAAA", 2),
        ("ZoneCCCCCCC", "ZS Commune", "ProvinceBBB", 3),
    ]
    if ou_dimension == "LEVEL-3;CountryAAAA":
        selected = units
    elif ou_dimension and ou_dimension.startswith("LEVEL-3;"):
        parent = ou_dimension.split(";", 1)[1]
        selected = [unit for unit in units if unit[2] == parent]
    else:
        selected = [unit for unit in units if unit[0] == ou_dimension]
    if unit_ids is not None:
        selected = [unit for unit in selected if unit[0] in unit_ids]
    periods = str(period).split(";")
    values = []
    for item in periods:
        if "W" in item:
            values.append(int(item.rsplit("W", 1)[1]))
        elif "Q" in item:
            first = (int(item.rsplit("Q", 1)[1]) - 1) * 3 + 1
            values.extend(range(first, first + 3))
        elif "S" in item:
            first = (int(item.rsplit("S", 1)[1]) - 1) * 6 + 1
            values.extend(range(first, first + 6))
        elif len(item) == 4:
            values.extend(range(1, 13))
        else:
            values.append(int(item[-2:]))
    expected_per_period = 60 if any("W" in item for item in periods) else 12
    rows = []
    for uid, name, province, base in selected:
        actual = base * sum(values)
        expected = base * expected_per_period * len(values)
        row = {"Organisation unit": name, "Organisation unit ID": uid}
        for number in range(10):
            prefix = f"Dataset {number} "
            row[prefix + "Actual reports"] = actual
            row[prefix + "Expected reports"] = expected
            row[prefix + "Actual reports on time"] = actual / 2
            row[prefix + "Reporting rate"] = actual / expected * 100
            row[prefix + "Reporting rate on time"] = actual / expected * 50
        rows.append(row)
    return pd.DataFrame(rows)
'''


PERIOD_QUALITY_DATA = PERIOD_DATA.replace(
    'def get_data(favori_id, period=None, cache_user=None, ou_dimension=None, unit_ids=None):',
    'def get_quality_reporting_data(favori_id, period=None, cache_user=None, ou_dimension=None, unit_ids=None, dataset_ids=None):',
).replace('_fixed_data_requests', '_fixed_quality_requests')


PERIOD_VALIDATION_CLIENT = '''
def get_dhis2_client():
    class Client:
        def post(self, endpoint, data=None):
            if endpoint != "dataAnalysis/validationRules":
                raise AssertionError("Appel fixture non prévu : " + endpoint)
            calls = list(st.session_state.get("_fixed_validation_requests", []))
            calls.append(dict(data))
            st.session_state["_fixed_validation_requests"] = calls
            return []
    return Client()
'''


class FixedPeriodDashboardTests(unittest.TestCase):
    def widget(self, app, kind, label):
        return next(item for item in getattr(app, kind) if item.label == label)

    def run_app(self, app):
        app.run(timeout=30)
        self.assertEqual(len(app.exception), 0, [item.message for item in app.exception])

    def new_app(self):
        source = dashboard_test.app_fixture_source({
            'get_organisation_catalog': filter_test.CATALOG_FIXTURE,
            'get_data': PERIOD_DATA,
            'get_quality_reporting_data': PERIOD_QUALITY_DATA,
            'get_dhis2_client': PERIOD_VALIDATION_CLIENT,
        })
        app = AppTest.from_string(source, default_timeout=30)
        app.session_state['connected'] = True
        app.session_state['dhis2_user'] = 'fixture-fixed-period-' + self._testMethodName
        app.session_state['dhis2_pass'] = 'fixture-password'
        return app

    @contextmanager
    def offline_downloads(self):
        downloads = {}
        real_download = st.download_button

        def capture(label, *args, **kwargs):
            downloads[label] = kwargs.get('data', args[0] if args else None)
            return real_download(label, *args, **kwargs)

        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')), \
                mock.patch('streamlit.download_button', side_effect=capture):
            yield downloads

    def period(self, app, periodicity):
        self.widget(app, 'selectbox', 'Périodicité').select(periodicity)
        self.run_app(app)

    def choose_range(self, app, start_label, end_label, start, end):
        year = self.widget(app, 'selectbox', "Choisir l'année :").value

        def identifier(value):
            if isinstance(value, int):
                return f'{year}W{value}'
            if value.startswith('T'):
                return f'{year}Q{value[1:]}'
            return f'{year}{value}'

        self.widget(app, 'selectbox', start_label).select(identifier(start))
        self.widget(app, 'selectbox', end_label).select(identifier(end))
        self.run_app(app)

    def requests(self, app, key='_fixed_data_requests'):
        return list(app.session_state[key]) if key in app.session_state else []

    def assert_no_validation(self, app):
        self.assertEqual(self.requests(app, '_fixed_validation_requests'), [])
        self.assertNotIn('categorisation_result', app.session_state)

    def assert_request(self, app, period, scope, ids, key='_fixed_data_requests'):
        self.assertTrue(any(
            call['period'] == period and call['ou_dimension'] == scope
            and set(call['unit_ids'] or []) == set(ids)
            for call in self.requests(app, key)
        ), self.requests(app, key))

    def assert_numeric_exports(self, downloads, actuals, expected, on_time):
        report = downloads['Télécharger Excel']
        for sheet in ['Detail_Donnees_Reelles', 'Perf_Final_Completude', 'Synthese_Comparative']:
            with self.subTest(sheet=sheet):
                _, rows = detail_test.excel_sheet(report, sheet)
                values = [row['Reports_Actual'] for row in rows]
                self.assertTrue(all(isinstance(value, float) for value in values))
                self.assertEqual(sorted(values), sorted(actuals))
        _, rows = detail_test.excel_sheet(report, 'Synthese_Comparative')
        self.assertEqual(sorted(row['Reports_Attendu'] for row in rows), sorted(expected))
        self.assertEqual(sorted(row['Reports_Actual_On_Time'] for row in rows), sorted(on_time))

    def assert_export_context(self, downloads, periodicity, period, start, end):
        headers, rows = detail_test.excel_sheet(downloads['Télécharger Excel'], 'Contexte_Periode')
        values = {str(value) for row in rows for value in row.values()}
        self.assertIn(periodicity, values)
        self.assertTrue(any(period in value for value in values))
        for label in ['Périodicité', 'Périodes DHIS2', 'Date de début', 'Date de fin']:
            self.assertIn(label, set(headers) | values)
        for day in [start, end]:
            representations = {
                day.isoformat(), day.strftime('%d/%m/%Y'),
                str(float((day - date(1899, 12, 30)).days)),
            }
            self.assertTrue(values.intersection(representations), f'Date absente du contexte : {day}')

    def test_monthly_default_keeps_existing_widgets_and_numeric_exports(self):
        app = self.new_app()
        with self.offline_downloads() as downloads:
            self.run_app(app)
            self.assertEqual(self.widget(app, 'selectbox', 'Périodicité').value, 'Mensuel')
            self.assertEqual(self.widget(app, 'selectbox', 'Périodicité').options,
                             ['Mensuel', 'Trimestriel', 'Semestriel', 'Annuel', 'Hebdomadaire'])
            year = self.widget(app, 'selectbox', "Choisir l'année :").value
            self.assertEqual(self.widget(app, 'selectbox', 'Mois début :').value, 'Janvier')
            self.assertEqual(self.widget(app, 'selectbox', 'Mois fin :').value, 'Janvier')
            self.assert_request(app, f'{year}01', 'LEVEL-3;CountryAAAA',
                                ['ZoneAAAAAAA', 'ZoneBBBBBBB', 'ZoneCCCCCCC'])
            self.assert_numeric_exports(downloads, [10, 20, 30], [120, 240, 360], [5, 10, 15])
            self.assert_no_validation(app)

    def test_second_quarter_loads_april_may_june_and_sums_each_zone(self):
        app = self.new_app()
        with self.offline_downloads() as downloads:
            self.run_app(app)
            self.period(app, 'Trimestriel')
            self.choose_range(app, 'Trimestre début :', 'Trimestre fin :', 'T2', 'T2')
            year = int(self.widget(app, 'selectbox', "Choisir l'année :").value)
            self.assert_request(app, f'{year}Q2', 'LEVEL-3;CountryAAAA',
                                ['ZoneAAAAAAA', 'ZoneBBBBBBB', 'ZoneCCCCCCC'])
            self.assert_request(app, f'{year}04;{year}05;{year}06', 'LEVEL-3;CountryAAAA',
                                ['ZoneAAAAAAA', 'ZoneBBBBBBB', 'ZoneCCCCCCC'], key='_fixed_quality_requests')
            self.assert_numeric_exports(downloads, [150, 300, 450], [360, 720, 1080], [75, 150, 225])
            self.assert_export_context(downloads, 'Trimestriel', f'{year}Q2', date(year, 4, 1), date(year, 6, 30))
            self.assertEqual(self.widget(app, 'date_input', 'Date de début').value, date(year, 4, 1))
            self.assertEqual(self.widget(app, 'date_input', 'Date de fin').value, date(year, 6, 30))
            self.assertEqual(sorted(app.session_state['_test_export_frame']['Reports_Actual']), [150, 300, 450])
            self.assert_no_validation(app)

    def test_second_semester_keeps_the_selected_province_scope(self):
        app = self.new_app()
        with self.offline_downloads() as downloads:
            self.run_app(app)
            self.period(app, 'Semestriel')
            self.choose_range(app, 'Semestre début :', 'Semestre fin :', 'S2', 'S2')
            self.widget(app, 'selectbox', "Niveau d'analyse :").select('Province')
            self.run_app(app)
            self.widget(app, 'selectbox', 'Province :').select('ProvinceAAA')
            self.run_app(app)
            year = int(self.widget(app, 'selectbox', "Choisir l'année :").value)
            self.assert_request(app, f'{year}S2', 'LEVEL-3;ProvinceAAA', ['ZoneAAAAAAA', 'ZoneBBBBBBB'])
            self.assert_request(app, ';'.join(f'{year}{month:02d}' for month in range(7, 13)),
                                'LEVEL-3;ProvinceAAA', ['ZoneAAAAAAA', 'ZoneBBBBBBB'],
                                key='_fixed_quality_requests')
            self.assert_numeric_exports(downloads, [570, 1140], [720, 1440], [285, 570])
            self.assertEqual(self.widget(app, 'date_input', 'Date de début').value, date(year, 7, 1))
            self.assertEqual(self.widget(app, 'date_input', 'Date de fin').value, date(year, 12, 31))
            self.assert_no_validation(app)

    def test_annual_loads_twelve_months_and_preserves_the_selected_zone_id(self):
        app = self.new_app()
        with self.offline_downloads() as downloads:
            self.run_app(app)
            self.period(app, 'Annuel')
            self.widget(app, 'selectbox', "Niveau d'analyse :").select('Zone de santé')
            self.run_app(app)
            self.widget(app, 'selectbox', 'Province :').select('ProvinceBBB')
            self.run_app(app)
            self.widget(app, 'selectbox', 'Zone de santé :').select('ZoneCCCCCCC')
            self.run_app(app)
            year = int(self.widget(app, 'selectbox', "Choisir l'année :").value)
            self.assert_request(app, str(year), 'ZoneCCCCCCC', ['ZoneCCCCCCC'])
            self.assert_request(app, ';'.join(f'{year}{month:02d}' for month in range(1, 13)),
                                'ZoneCCCCCCC', ['ZoneCCCCCCC'], key='_fixed_quality_requests')
            self.assert_numeric_exports(downloads, [2340], [4320], [1170])
            self.assertEqual(self.widget(app, 'date_input', 'Date de début').value, date(year, 1, 1))
            self.assertEqual(self.widget(app, 'date_input', 'Date de fin').value, date(year, 12, 31))
            self.assert_no_validation(app)

    def test_weekly_loads_native_iso_weeks_and_exports_their_sum(self):
        app = self.new_app()
        with self.offline_downloads() as downloads:
            self.run_app(app)
            self.widget(app, 'selectbox', "Choisir l'année :").select('2026')
            self.run_app(app)
            self.period(app, 'Hebdomadaire')
            options = self.widget(app, 'selectbox', 'Semaine début :').options
            self.assertTrue(any(
                '36' in option and any(day in option for day in ['31/08', '2026-08-31'])
                and any(day in option for day in ['06/09', '2026-09-06'])
                for option in options
            ), options)
            self.choose_range(app, 'Semaine début :', 'Semaine fin :', 36, 39)
            year = int(self.widget(app, 'selectbox', "Choisir l'année :").value)
            period = ';'.join(f'{year}W{week}' for week in range(36, 40))
            self.assert_request(app, period, 'LEVEL-3;CountryAAAA',
                                ['ZoneAAAAAAA', 'ZoneBBBBBBB', 'ZoneCCCCCCC'])
            self.assert_numeric_exports(downloads, [1500, 3000, 4500], [2400, 4800, 7200], [750, 1500, 2250])
            self.assert_export_context(downloads, 'Hebdomadaire', period,
                                       date.fromisocalendar(year, 36, 1), date.fromisocalendar(year, 39, 7))
            self.assert_request(app, f'{year}08;{year}09', 'LEVEL-3;CountryAAAA',
                                ['ZoneAAAAAAA', 'ZoneBBBBBBB', 'ZoneCCCCCCC'], key='_fixed_quality_requests')
            self.assertEqual(self.widget(app, 'date_input', 'Date de début').value, date(year, 8, 1))
            self.assertEqual(self.widget(app, 'date_input', 'Date de fin').value, date(year, 9, 30))
            self.assertEqual(sorted(app.session_state['_test_export_frame']['Reports_Actual']), [170, 340, 510])
            self.assert_no_validation(app)

    def test_weekly_range_can_include_more_than_four_iso_weeks(self):
        app = self.new_app()
        with self.offline_downloads() as downloads:
            self.run_app(app)
            self.period(app, 'Hebdomadaire')
            self.choose_range(app, 'Semaine début :', 'Semaine fin :', 36, 41)
            year = self.widget(app, 'selectbox', "Choisir l'année :").value
            period = ';'.join(f'{year}W{week}' for week in range(36, 42))
            self.assert_request(app, period, 'LEVEL-3;CountryAAAA',
                                ['ZoneAAAAAAA', 'ZoneBBBBBBB', 'ZoneCCCCCCC'])
            self.assert_numeric_exports(downloads, [2310, 4620, 6930], [3600, 7200, 10800], [1155, 2310, 3465])
            self.assert_no_validation(app)

    def test_rule_dates_remain_exact_and_previous_month_stays_calendar_based(self):
        app = self.new_app()
        with self.offline_downloads():
            self.run_app(app)
            self.period(app, 'Trimestriel')
            self.choose_range(app, 'Trimestre début :', 'Trimestre fin :', 'T2', 'T2')
            year = int(self.widget(app, 'selectbox', "Choisir l'année :").value)
            self.widget(app, 'selectbox', 'Province de l’analyse').select('ProvinceBBB')
            self.run_app(app)
            self.widget(app, 'selectbox', 'Zone de santé de l’analyse').select('ZoneCCCCCCC')
            self.widget(app, 'selectbox', 'Sélectionner le Groupe de Règles de Validation :').select('Groupe test')
            self.widget(app, 'date_input', 'Date de début').set_value(date(year, 5, 9))
            self.widget(app, 'date_input', 'Date de fin').set_value(date(year, 6, 4))
            self.run_app(app)
            self.assert_no_validation(app)
            self.assert_request(app, f'{year}05;{year}06', 'ZoneCCCCCCC', ['ZoneCCCCCCC'],
                                key='_fixed_quality_requests')
            self.assertEqual(app.session_state['_test_export_frame']['Reports_Actual'].tolist(), [330])
            self.widget(app, 'button', "Lancer l'analyse des violations").click()
            self.run_app(app)
            calls = self.requests(app, '_fixed_validation_requests')
            self.assertEqual(len(calls), 1)
            self.assertEqual((calls[0]['startDate'], calls[0]['endDate']), (f'{year}-05-09', f'{year}-06-04'))
            self.assertEqual(calls[0]['ou'], 'ZoneCCCCCCC')
            self.assertEqual(calls[0]['vrg'], 'groupAAAAAA')
            self.widget(app, 'button', 'Contrôler les corrections du mois précédent').click()
            self.run_app(app)
            calls = self.requests(app, '_fixed_validation_requests')
            self.assertEqual(len(calls), 2)
            self.assertEqual((calls[1]['startDate'], calls[1]['endDate']), (f'{year}-04-01', f'{year}-04-30'))
            self.assertTrue(all(call['ou'] == 'ZoneCCCCCCC' and call['vrg'] == 'groupAAAAAA' for call in calls))
            self.assertTrue(all(call['persist'] is False and call['notification'] is False for call in calls))

    def test_reversed_fixed_period_blocks_before_loading_any_new_data(self):
        for periodicity, start_label, end_label, start, end in [
            ('Trimestriel', 'Trimestre début :', 'Trimestre fin :', 'T3', 'T2'),
            ('Semestriel', 'Semestre début :', 'Semestre fin :', 'S2', 'S1'),
            ('Hebdomadaire', 'Semaine début :', 'Semaine fin :', 39, 36),
        ]:
            with self.subTest(periodicity=periodicity), self.offline_downloads():
                app = self.new_app()
                self.run_app(app)
                self.period(app, periodicity)
                data = self.requests(app)
                quality = self.requests(app, '_fixed_quality_requests')
                app.session_state['categorisation_result'] = {'stale': True}
                app.session_state['validation_job'] = {'stale': True}
                self.choose_range(app, start_label, end_label, start, end)
                self.assertGreater(len(app.error), 0)
                self.assertEqual(self.requests(app), data)
                self.assertEqual(self.requests(app, '_fixed_quality_requests'), quality)
                self.assert_no_validation(app)
                self.assertNotIn('validation_job', app.session_state)


if __name__ == '__main__':
    unittest.main()

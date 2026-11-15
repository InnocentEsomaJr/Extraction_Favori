"""Le vrai dashboard conserve l'archive M et ses tableaux après le contrôle M-1."""

from copy import deepcopy
import io
import json
from types import SimpleNamespace
import unittest
from unittest import mock

import pandas as pd
import streamlit as st
from streamlit.testing.v1 import AppTest

from validation_history import validate_snapshot
import test_dashboard_categorisation as dashboard_test
import test_dashboard_validation_details as detail_test
from test_validation_history_ui import MemoryHistoryStore


CONFIG = '''
def _read_config_value(key):
    return {
        'DHIS2_URL': 'https://history-dashboard-fixture.invalid',
        'GITHUB_HISTORY_TOKEN': 'PRIVATE_DASHBOARD_HISTORY_TOKEN',
        'GITHUB_HISTORY_REPOSITORY': 'fixture/history',
        'GITHUB_HISTORY_BRANCH': 'fixture-history',
        'GITHUB_HISTORY_DIRECTORY': 'fixture-archive',
    }.get(key)
'''


CLIENT = '''
def get_dhis2_client():
    class Client:
        def post(self, endpoint, data=None):
            if endpoint != 'dataAnalysis/validationRules':
                raise AssertionError('POST fixture inattendu : ' + endpoint)
            calls = list(st.session_state.get('_history_integration_posts', []))
            calls.append(dict(data))
            st.session_state['_history_integration_posts'] = calls
            previous = data['startDate'] < st.session_state['validation_start_date'].isoformat()
            rules = ['RuleAAAAAAA', 'RuleBBBBBBB'] if previous else ['RuleAAAAAAA']
            return [{
                'organisationUnitId': 'CentreAAAAA',
                'organisationUnitDisplayName': 'Centre test',
                'organisationUnitPath': '/CountryAAAA/ProvinceAAA/ZoneAAAAAAA/CentreAAAAA',
                'validationRuleId': rule,
                'periodId': data['startDate'][:7].replace('-', ''),
                'leftSideValue': 8, 'rightSideValue': 4,
                'attributeOptionComboId': 'AttributeAA',
            } for rule in rules]

        def get(self, endpoint, params=None):
            raise AssertionError('GET fixture inattendu : ' + endpoint)
    return Client()
'''


class DashboardValidationHistoryTests(unittest.TestCase):
    def widget(self, app, kind, label):
        return next(item for item in getattr(app, kind) if item.label == label)

    def run_app(self, app):
        app.run(timeout=45)
        self.assertEqual(len(app.exception), 0, [entry.message for entry in app.exception])

    def click(self, app, label):
        # AppTest garde parfois un ancien widget après les autoreruns des
        # tableaux dynamiques. Le rendu stable supprime ce doublon ; sinon
        # deux états du même ID (True, False) annulent artificiellement le clic.
        matches = [item for item in app.button if item.label == label]
        if len({item.id for item in matches}) < len(matches):
            self.run_app(app)
        self.widget(app, 'button', label).click()
        self.run_app(app)

    def test_m_archive_is_identical_after_m_minus_one_and_both_detail_views_survive_saving(self):
        source = dashboard_test.app_fixture_source({
            '_read_config_value': CONFIG,
            'get_dhis2_client': CLIENT,
            'get_validation_detail_metadata': detail_test.DETAIL_METADATA,
        })
        app = AppTest.from_string(source, default_timeout=45)
        app.session_state['connected'] = True
        app.session_state['dhis2_user'] = 'PRIVATE_DASHBOARD_HISTORY_USER'
        app.session_state['dhis2_pass'] = 'PRIVATE_DASHBOARD_HISTORY_PASSWORD'
        memory = SimpleNamespace(archive={}, configurations=[], save_calls=[], read_calls=[], list_calls=0)
        downloads = {}
        real_download = st.download_button

        def capture(label, *args, **kwargs):
            if label in {'Télécharger cette extraction JSON', 'Télécharger les violations affichées en CSV'}:
                downloads[label] = kwargs.get('data', args[0] if args else None)
            return real_download(label, *args, **kwargs)

        def store(repository, token, **kwargs):
            return MemoryHistoryStore(memory, repository, token, **kwargs)

        def detail_csv():
            return pd.read_csv(io.BytesIO(downloads['Télécharger les violations affichées en CSV']),
                               sep=None, engine='python', dtype=str, keep_default_na=False)

        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')), \
                mock.patch('validation_history_ui.GitHubHistoryStore', side_effect=store), \
                mock.patch('streamlit.download_button', side_effect=capture):
            self.run_app(app)
            self.assertNotIn('categorisation_result', app.session_state)
            self.click(app, "Lancer l'analyse des violations")
            current = deepcopy(app.session_state['categorisation_result'])
            self.assertEqual(current['current_batch']['processed_ids'], ['ZoneAAAAAAA'])
            self.assertEqual(len(current['violations']), 1)
            self.assertEqual(len(detail_csv()), 1)
            first_json = downloads['Télécharger cette extraction JSON']
            first_snapshot = validate_snapshot(json.loads(first_json))
            self.assertEqual(len(first_snapshot['violations']), 1)
            self.assertFalse(self.widget(app, 'button', 'Enregistrer cette extraction dans GitHub').disabled,
                             str(memory.configurations))
            self.click(app, 'Enregistrer cette extraction dans GitHub')
            first_archive = dict(memory.archive)
            self.assertEqual(len(first_archive), 1)
            self.assertEqual(len(app.session_state['_history_integration_posts']), 1)

            self.click(app, 'Contrôler les corrections du mois précédent')
            saved = app.session_state['categorisation_result']
            self.assertEqual(saved['current_extraction_id'], current['current_extraction_id'])
            self.assertEqual(saved['analysed_at'], current['analysed_at'])
            for key in ('current_raw', 'current_assigned', 'violations'):
                pd.testing.assert_frame_equal(saved[key], current[key])
            self.assertTrue(saved['corrections_checked'])
            self.assertEqual(len(saved['previous_violations']), 2)
            self.assertEqual(self.widget(app, 'selectbox', 'Période des violations affichées').value,
                             'Mois précédent (M-1)')
            self.assertEqual(len(detail_csv()), 2)
            current_details = deepcopy(saved['violations'])
            previous_details = deepcopy(saved['previous_violations'])
            score_table = deepcopy(app.session_state['_test_export_frame'])
            requests_before = deepcopy(app.session_state['_history_integration_posts'])
            self.assertEqual(len(requests_before), 2)
            self.assertEqual({request['ou'] for request in requests_before}, {'ZoneAAAAAAA'})
            self.assertNotEqual(requests_before[0]['startDate'][:7], requests_before[1]['startDate'][:7])
            self.assertEqual(downloads['Télécharger cette extraction JSON'], first_json)

            self.click(app, 'Enregistrer cette extraction dans GitHub')
            self.click(app, 'Enregistrer cette extraction dans GitHub')
            self.assertEqual(memory.archive, first_archive)
            self.assertEqual(len(memory.save_calls), 3)
            self.assertEqual(len(set(memory.save_calls)), 1)
            self.assertTrue(any('déjà enregistrée' in entry.value for entry in app.success))
            self.assertEqual(app.session_state['_history_integration_posts'], requests_before)
            self.assertEqual(downloads['Télécharger cette extraction JSON'], first_json)
            pd.testing.assert_frame_equal(app.session_state['categorisation_result']['violations'], current_details)
            pd.testing.assert_frame_equal(app.session_state['categorisation_result']['previous_violations'], previous_details)
            pd.testing.assert_frame_equal(app.session_state['_test_export_frame'], score_table)
            self.assertEqual(len(detail_csv()), 2)

            self.widget(app, 'selectbox', 'Période des violations affichées').select('Période filtrée (M)')
            self.run_app(app)
            self.assertEqual(len(detail_csv()), 1)
            self.assertEqual(app.session_state['_history_integration_posts'], requests_before)
            self.assertEqual(memory.archive, first_archive)
            for raw in memory.archive.values():
                self.assertNotIn('PRIVATE_', raw.decode('utf-8'))
                self.assertNotIn('fixture.invalid', raw.decode('utf-8'))


if __name__ == '__main__':
    unittest.main()

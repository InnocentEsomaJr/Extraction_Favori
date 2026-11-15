"""Historique Streamlit isolé : archives immuables, contrôles exacts et reprise.

Le petit wrapper exerce le vrai panneau sans importer le tableau de bord ni ses
secrets. Les appels GitHub sont remplacés par un dépôt JSON partagé en mémoire ;
le faux DHIS2 garde le vrai adaptateur POST et les vérifications de métadonnées.
"""

from contextlib import contextmanager
from copy import deepcopy
import io
import json
import unittest
from unittest import mock

import pandas as pd
import streamlit as st
from streamlit.runtime.scriptrunner import StopException
from streamlit.testing.v1 import AppTest

from github_history import HistoryConflictError, HistoryFile, HistorySaveResult
from quality_batch import run_quality_batch
from validation_api import fetch_validation_analysis
from validation_history import build_snapshot, snapshot_filename, validate_snapshot
from validation_job import INTERRUPTED_MESSAGE


COUNTRY = 'CountryAAAA'
PROVINCES = ('ProvinceAAA', 'ProvinceBBB')
ZONES = ('ZoneAAAAAAA', 'ZoneBBBBBBB', 'ZoneCCCCCCC')
OUTSIDE = 'ZoneDDDDDDD'
CENTRES = ('CentreAAAAA', 'CentreBBBBB', 'CentreCCCCC')
RULE, NEW_RULE = 'RuleAAAAAAA', 'RuleBBBBBBB'
GROUP, CURRENT_GROUP = 'GroupAAAAAA', 'GroupBBBBBB'
STAMP = '2026-10-01T10:00:00+00:00'
PREFIX = 'validation_history_'
RULES = {
    identifier: {
        'id': identifier, 'displayName': name, 'importance': 'HIGH',
        'operator': 'less_than_or_equal_to',
        'leftSide': {'expression': '#{PRIVATE_EXPRESSION_LEFT}'},
        'rightSide': {'expression': '#{PRIVATE_EXPRESSION_RIGHT}'},
    }
    for identifier, name in [(RULE, 'Règle ancienne'), (NEW_RULE, 'Règle nouvelle')]
}


def violation(position, rule=RULE, period='202609'):
    return {
        'organisationUnitId': CENTRES[position],
        'organisationUnitDisplayName': f'Centre {position + 1}',
        'organisationUnitPath': f'/{COUNTRY}/{PROVINCES[position // 2]}/{ZONES[position]}/{CENTRES[position]}',
        'validationRuleId': rule, 'validationRuleDescription': RULES[rule]['displayName'],
        'periodId': period, 'periodDisplayName': period,
        'leftSideValue': 8, 'rightSideValue': 4,
        '_request_unit_id': ZONES[position],
        'password': 'PRIVATE_ROW_PASSWORD',
        'access_token': 'PRIVATE_ROW_TOKEN',
    }


def batch_for(targets, rows, start, end, label='M'):
    def fetch(identifier, *args, **kwargs):
        frame = pd.DataFrame([deepcopy(row) for row in rows if row['_request_unit_id'] == identifier])
        frame.attrs = {
            'request_method': 'POST',
            'request_params': {'startDate': start, 'endDate': end, 'token': 'PRIVATE_FRAME_TOKEN'},
            'failed_periods': [], 'truncated_periods': [],
        }
        return frame

    batch = run_quality_batch(
        targets, [start[:7].replace('-', '')], GROUP, fetch,
        start_date=start, end_date=end, label=label,
    )
    batch['context']['token'] = 'PRIVATE_CONTEXT_TOKEN'
    return batch


class MemoryHistoryStore:
    def __init__(self, fixture, repository, token, *, branch, directory):
        self.fixture, self.directory = fixture, directory
        fixture.configurations.append((repository, token, branch, directory))

    def save_snapshot(self, name, payload, snapshot_id):
        self.fixture.save_calls.append((name, snapshot_id))
        raw = json.dumps(payload, ensure_ascii=False, sort_keys=True, allow_nan=False).encode('utf-8')
        created = name not in self.fixture.archive
        if not created and self.fixture.archive[name] != raw:
            raise HistoryConflictError()
        self.fixture.archive.setdefault(name, raw)
        return HistorySaveResult(name, self.directory + '/' + name, created, 'https://fixture.invalid/archive')

    def list_snapshots(self):
        self.fixture.list_calls += 1
        return [
            HistoryFile(name, self.directory + '/' + name, 'a' * 40, len(raw), 'https://fixture.invalid/archive')
            for name, raw in sorted(self.fixture.archive.items(), reverse=True)
        ]

    def read_snapshot(self, name):
        self.fixture.read_calls.append(name)
        return json.loads(self.fixture.archive[name])


class HistoryFixture:
    def __init__(self):
        self.source = 'https://PRIVATE_USER:PRIVATE_PASSWORD@fixture.invalid/dhis?token=PRIVATE_QUERY'
        self.account = 'PRIVATE_LOGIN'
        self.cache_user = (self.account, 'PRIVATE_REVISION')
        self.config = {
            'GITHUB_HISTORY_TOKEN': 'PRIVATE_GITHUB_TOKEN',
            'GITHUB_HISTORY_REPOSITORY': 'fixture/history',
            'GITHUB_HISTORY_BRANCH': 'fixture-history',
            'GITHUB_HISTORY_DIRECTORY': 'fixture-archive',
        }
        self.targets = pd.DataFrame([
            {'Zone_ID': identifier, 'Zone': f'ZS {position + 1}', 'level': 3,
             'Pays_ID': COUNTRY, 'Pays': 'RDC',
             'Province_ID': PROVINCES[min(position // 2, 1)],
             'Province': 'Province A' if position < 2 else 'Province B'}
            for position, identifier in enumerate([*ZONES, OUTSIDE])
        ])
        self.geography = self.targets.rename(columns={'Zone_ID': 'Organisation unit ID'}).copy()
        self.geography['Zone_ID'] = self.targets['Zone_ID']
        self.units = {
            centre: {
                'id': centre, 'displayName': f'Centre {position + 1}', 'level': 4,
                'path': violation(position)['organisationUnitPath'],
                'ancestors': [
                    {'id': COUNTRY, 'displayName': 'RDC', 'level': 1},
                    {'id': PROVINCES[position // 2], 'displayName': 'Province A' if position < 2 else 'Province B', 'level': 2},
                    {'id': ZONES[position], 'displayName': f'ZS {position + 1}', 'level': 3},
                ],
            }
            for position, centre in enumerate(CENTRES)
        }
        current_batch = batch_for(self.targets, [violation(0, NEW_RULE)], '2026-09-20', '2026-09-30')
        previous_batch = batch_for(self.targets, [violation(1, period='202608')], '2026-08-01', '2026-08-31', 'M-1')
        self.saved = {
            'current_batch': current_batch, 'previous_batch': previous_batch,
            'current_assigned': deepcopy(current_batch['raw']),
            'previous_assigned': deepcopy(previous_batch['raw']),
            'violations': pd.DataFrame(),
            'previous_violations': pd.DataFrame([{'Période': 'M-1', 'Valeur': 27}]),
            'comparison': pd.DataFrame([{'Zone_ID': ZONES[1], 'Violations_Corrigees': 19}]),
            'rule_metadata': deepcopy(RULES), 'unit_metadata': deepcopy(self.units),
            'analysed_at': STAMP, 'current_extraction_id': 'fixture-current-extraction',
        }
        self.original_saved = deepcopy(self.saved)
        self.current_group = GROUP
        self.current_busy = False
        self.pause_after_step = False
        self.archive, self.downloads = {}, {}
        self.configurations, self.save_calls, self.read_calls = [], [], []
        self.list_calls = 0
        self.post_calls, self.get_calls, self.fetch_calls, self.metadata_calls = [], [], [], []
        self.interrupt_target = None
        self.interrupted = False
        self.responses = {ZONES[0]: [violation(0)], ZONES[1]: [], ZONES[2]: [violation(2, NEW_RULE)]}

    def make_store(self, repository, token, **kwargs):
        return MemoryHistoryStore(self, repository, token, **kwargs)

    def seed(self, month='09', snapshot_id='a' * 32):
        start, end = f'2026-{month}-01', f'2026-{month}-20'
        rows = [violation(position, period='2026' + month) for position in range(3)]
        batch = batch_for(self.targets.iloc[:3], rows, start, end)
        snapshot = build_snapshot(
            source_url=self.source, account=self.account, start_date=start, end_date=end,
            parent_id=COUNTRY, parent_name='RDC', group_id=GROUP, group_name='Groupe enregistré',
            targets=self.targets.iloc[:3], batch=batch, raw=batch['raw'], details=pd.DataFrame(),
            rule_metadata=RULES, analysed_at=STAMP, captured_at=STAMP, snapshot_id=snapshot_id,
        )
        name = snapshot_filename(snapshot)
        self.archive[name] = json.dumps(snapshot, ensure_ascii=False, sort_keys=True, allow_nan=False).encode('utf-8')
        return name

    def get_client(self):
        return self

    def post(self, endpoint, data=None):
        if endpoint != 'dataAnalysis/validationRules':
            raise AssertionError('POST fixture inattendu : ' + endpoint)
        self.post_calls.append(deepcopy(data))
        if data['ou'] == self.interrupt_target and not self.interrupted:
            self.interrupted = True
            raise StopException()
        return deepcopy(self.responses[data['ou']])

    def get(self, endpoint, params=None):
        self.get_calls.append((endpoint, deepcopy(params)))
        if endpoint == 'validationRuleGroups/' + GROUP:
            return {'id': GROUP, 'validationRules': [{'id': identifier} for identifier in RULES]}
        if endpoint not in {'organisationUnits', 'validationRules'}:
            raise AssertionError('GET fixture inattendu : ' + endpoint)
        identifiers = params['filter'].removeprefix('id:in:[').removesuffix(']').split(',')
        records = self.units if endpoint == 'organisationUnits' else RULES
        return {endpoint: [deepcopy(records[identifier]) for identifier in identifiers if identifier in records]}

    def fetch_validation(self, identifier, periods, group_id, *, cache_user, **bounds):
        self.fetch_calls.append((identifier, list(periods), group_id, deepcopy(cache_user), deepcopy(bounds)))
        return fetch_validation_analysis(self, identifier, periods, group_id, **bounds)

    def get_metadata(self, rule_ids, unit_ids, *, cache_user):
        self.metadata_calls.append((rule_ids, unit_ids, cache_user))
        return ({identifier: deepcopy(RULES[identifier]) for identifier in rule_ids},
                {identifier: deepcopy(self.units[identifier]) for identifier in unit_ids}, [])


WRAPPER = '''
import streamlit as st
from validation_history_ui import render_history_panel
fixture = st.session_state['_history_fixture']
rerun = render_history_panel(
    saved=st.session_state['categorisation_result'], source_url=fixture.source,
    account=fixture.account, targets=fixture.targets, geography=fixture.geography,
    start_date='2026-09-20', end_date='2026-09-30', parent_id='CountryAAAA', parent_name='RDC',
    group_id=fixture.current_group, group_name='Groupe actuel', cache_user=fixture.cache_user,
    read_config=fixture.config.get, fetch_validation=fixture.fetch_validation,
    get_client=fixture.get_client, get_metadata=fixture.get_metadata,
    current_busy=fixture.current_busy,
)
st.caption('Tableaux M et M-1 conserves')
st.dataframe(st.session_state['categorisation_result']['current_assigned'])
st.dataframe(st.session_state['categorisation_result']['previous_assigned'])
if rerun:
    if fixture.pause_after_step:
        st.stop()
    st.rerun()
'''


class ValidationHistoryUiTests(unittest.TestCase):
    def setUp(self):
        self.fixture = HistoryFixture()

    @contextmanager
    def offline(self):
        real_download = st.download_button

        def capture(label, *args, **kwargs):
            self.fixture.downloads[label] = kwargs.get('data', args[0] if args else None)
            return real_download(label, *args, **kwargs)

        with mock.patch('requests.sessions.Session.request', side_effect=AssertionError('Réseau interdit')), \
                mock.patch('validation_history_ui.GitHubHistoryStore', side_effect=self.fixture.make_store), \
                mock.patch('streamlit.download_button', side_effect=capture):
            yield

    def new_app(self):
        app = AppTest.from_string(WRAPPER, default_timeout=15)
        app.session_state['_history_fixture'] = self.fixture
        app.session_state['categorisation_result'] = self.fixture.saved
        app.session_state['validation_job'] = {'status': 'complete', 'phase': 'M-1', 'sentinel': 'independent'}
        return app

    def run_app(self, app):
        app.run(timeout=15)
        self.assertEqual(len(app.exception), 0, [entry.message for entry in app.exception])

    def widget(self, app, kind, label):
        return next(item for item in getattr(app, kind) if item.label == label)

    def click(self, app, label):
        self.widget(app, 'button', label).click()
        self.run_app(app)

    def select(self, app, label, option):
        self.widget(app, 'selectbox', label).select(option)
        self.run_app(app)

    def open_history(self, app, name=None):
        self.click(app, 'Charger l’historique GitHub')
        if name is not None:
            self.select(app, 'Extraction enregistrée', name)
        self.click(app, 'Ouvrir cette extraction')

    def summary(self, app):
        return next(item.value for item in app.dataframe if 'Anciennes' in item.value)

    def exported_comparison(self):
        return pd.read_csv(io.BytesIO(self.fixture.downloads['Télécharger la comparaison historique en CSV']),
                           dtype=str, keep_default_na=False)

    def assert_saved_unchanged(self, app):
        def assert_equal(actual, expected):
            if isinstance(expected, pd.DataFrame):
                pd.testing.assert_frame_equal(actual, expected)
                self.assertEqual(actual.attrs, expected.attrs)
            elif isinstance(expected, dict):
                self.assertEqual(actual.keys(), expected.keys())
                for key in expected:
                    assert_equal(actual[key], expected[key])
            elif isinstance(expected, (list, tuple)):
                self.assertEqual(len(actual), len(expected))
                for left, right in zip(actual, expected):
                    assert_equal(left, right)
            elif isinstance(expected, float) and pd.isna(expected):
                self.assertTrue(pd.isna(actual))
            else:
                self.assertEqual(actual, expected)

        assert_equal(app.session_state['categorisation_result'], self.fixture.original_saved)
        self.assertEqual(app.session_state['validation_job'],
                         {'status': 'complete', 'phase': 'M-1', 'sentinel': 'independent'})
        # Le transport Arrow de Streamlit peut inférer le dtype string ; les
        # objets originaux de session sont contrôlés strictement ci-dessus.
        pd.testing.assert_frame_equal(app.dataframe[-2].value, self.fixture.original_saved['current_assigned'], check_dtype=False)
        pd.testing.assert_frame_equal(app.dataframe[-1].value, self.fixture.original_saved['previous_assigned'], check_dtype=False)

    def test_initial_render_is_offline_and_exposes_json_without_credentials(self):
        app = self.new_app()
        with self.offline():
            self.run_app(app)
            self.assertEqual(self.fixture.post_calls + self.fixture.get_calls + self.fixture.read_calls, [])
            self.assertEqual(self.fixture.list_calls, 0)
            self.assertEqual(self.fixture.save_calls, [])
            raw = self.fixture.downloads['Télécharger cette extraction JSON']
            snapshot = validate_snapshot(json.loads(raw))
            self.assertEqual(snapshot['filters']['startDate'], '2026-09-20')
            self.assertEqual(len(snapshot['violations']), 1)
            for secret in ('PRIVATE_', 'fixture.invalid', 'access_token', 'password', 'expression'):
                self.assertNotIn(secret, raw.decode('utf-8'))
            self.assert_saved_unchanged(app)

    def test_missing_token_keeps_local_json_available_and_disables_github_save(self):
        self.fixture.config.pop('GITHUB_HISTORY_TOKEN')
        app = self.new_app()
        with self.offline():
            self.run_app(app)
            self.assertTrue(self.widget(app, 'button', 'Enregistrer cette extraction dans GitHub').disabled)
            self.assertTrue(any('GITHUB_HISTORY_TOKEN' in item.value for item in app.info))
            validate_snapshot(json.loads(self.fixture.downloads['Télécharger cette extraction JSON']))
            self.assertEqual(self.fixture.save_calls + self.fixture.post_calls + self.fixture.get_calls, [])
            self.assert_saved_unchanged(app)

    def test_double_save_is_idempotent_and_preserves_both_current_tables(self):
        app = self.new_app()
        with self.offline():
            self.run_app(app)
            self.click(app, 'Enregistrer cette extraction dans GitHub')
            first = dict(self.fixture.archive)
            self.click(app, 'Enregistrer cette extraction dans GitHub')
            self.assertEqual(len(first), 1)
            self.assertEqual(self.fixture.archive, first)
            self.assertEqual(self.fixture.save_calls[0], self.fixture.save_calls[1])
            self.assertTrue(any('déjà enregistrée' in item.value for item in app.success))
            self.assertEqual(self.fixture.post_calls + self.fixture.get_calls, [])
            self.assertEqual(set(self.fixture.configurations), {
                ('fixture/history', 'PRIVATE_GITHUB_TOKEN', 'fixture-history', 'fixture-archive')})
            self.assert_saved_unchanged(app)

    def test_load_month_filter_and_open_are_explicit_and_keep_archive_immutable(self):
        september = self.fixture.seed()
        august = self.fixture.seed('08', 'b' * 32)
        originals = dict(self.fixture.archive)
        app = self.new_app()
        with self.offline():
            self.run_app(app)
            self.click(app, 'Charger l’historique GitHub')
            self.assertEqual(self.fixture.list_calls, 1)
            self.assertEqual(self.fixture.read_calls, [])
            month = self.widget(app, 'selectbox', 'Mois de l’historique')
            self.assertEqual([value.replace('-', '') for value in month.options], ['Tous', '202609', '202608'])
            august_option = next(value for value in month.options if value.replace('-', '') == '202608')
            september_option = next(value for value in month.options if value.replace('-', '') == '202609')
            self.select(app, 'Mois de l’historique', august_option)
            self.assertEqual(self.widget(app, 'selectbox', 'Extraction enregistrée').options, [august])
            self.click(app, 'Ouvrir cette extraction')
            self.assertEqual(self.fixture.read_calls, [august])
            self.assertEqual(len(app.dataframe[0].value), 3)
            self.assertEqual(self.fixture.post_calls + self.fixture.get_calls, [])
            self.select(app, 'Mois de l’historique', september_option)
            self.assertEqual(self.widget(app, 'selectbox', 'Extraction enregistrée').options, [september])
            self.assertFalse(any(item.label == 'Vérifier les corrections de cette extraction' for item in app.button))
            self.assertEqual(self.fixture.archive, originals)
            self.assert_saved_unchanged(app)

    def test_disjoint_current_dates_never_claim_a_correction_even_after_metadata_check(self):
        self.fixture.seed()
        app = self.new_app()
        with self.offline():
            self.run_app(app)
            self.open_history(app)
            summary = self.summary(app)
            self.assertEqual(int(summary['Anciennes'].sum()), 3)
            self.assertEqual(int(summary['Corrigées'].sum()), 0)
            self.assertEqual(int(summary['Non vérifiables'].sum()), 3)
            self.click(app, 'Vérifier la comparaison avec l’analyse actuelle')
            self.assertEqual(int(self.summary(app)['Corrigées'].sum()), 0)
            self.assertNotIn('Corrigée', self.exported_comparison()['État historique'].tolist())
            self.assertEqual(self.fixture.post_calls + self.fixture.get_calls, [])
            self.assert_saved_unchanged(app)

    def test_recheck_posts_archived_bounds_group_and_only_common_targets(self):
        self.fixture.seed()
        self.fixture.current_group = CURRENT_GROUP
        originals = dict(self.fixture.archive)
        app = self.new_app()
        with self.offline():
            self.run_app(app)
            self.open_history(app)
            self.click(app, 'Vérifier les corrections de cette extraction')
            self.assertEqual([call['ou'] for call in self.fixture.post_calls], list(ZONES))
            for call in self.fixture.post_calls:
                self.assertEqual(call, {
                    'ou': call['ou'], 'startDate': '2026-09-01', 'endDate': '2026-09-20',
                    'vrg': GROUP, 'persist': False, 'notification': False,
                })
            for identifier, periods, group, cache_user, bounds in self.fixture.fetch_calls:
                self.assertEqual(periods, ['202609'])
                self.assertEqual(group, GROUP)
                self.assertEqual(bounds, {'start_date': '2026-09-01', 'end_date': '2026-09-20'})
                self.assertEqual(tuple(cache_user[:2]), self.fixture.cache_user)
                self.assertEqual(len(cache_user), 3)
            self.assertEqual(len({tuple(call[3]) for call in self.fixture.fetch_calls}), 1)
            self.assertEqual(app.session_state[PREFIX + 'job']['status'], 'complete')
            summary = self.summary(app)
            self.assertEqual(int(summary['Persistantes'].sum()), 1)
            self.assertEqual(int(summary['Corrigées'].sum()), 2)
            self.assertEqual(int(summary['Nouvelles'].sum()), 1)
            self.assertEqual(int(summary['Non vérifiables'].sum()), 0)
            self.assertTrue(any(endpoint == 'validationRuleGroups/' + GROUP for endpoint, _ in self.fixture.get_calls))
            self.assertEqual(self.fixture.archive, originals)
            self.assert_saved_unchanged(app)
            posts = deepcopy(self.fixture.post_calls)
            self.select(app, 'Comparer avec', 'Analyse du filtre actuel')
            self.assertEqual(int(self.summary(app)['Corrigées'].sum()), 0)
            self.select(app, 'Comparer avec', 'Contrôle de l’extraction enregistrée')
            self.assertEqual(int(self.summary(app)['Corrigées'].sum()), 2)
            self.assertEqual(self.fixture.post_calls, posts)
            self.assert_saved_unchanged(app)

    def test_grouping_status_filter_and_csv_match_visible_comparison(self):
        self.fixture.seed()
        app = self.new_app()
        with self.offline():
            self.run_app(app)
            self.open_history(app)
            self.click(app, 'Vérifier les corrections de cette extraction')
            calls = deepcopy(self.fixture.post_calls)
            self.select(app, 'Regrouper l’historique par', 'Provinces / DPS')
            summary = self.summary(app)
            self.assertEqual(summary['Province'].tolist(), ['Province A', 'Province B'])
            self.assertEqual(summary['Anciennes'].tolist(), [2, 1])
            self.assertEqual(summary['Corrigées'].tolist(), [1, 1])
            self.select(app, 'Regrouper l’historique par', 'Pays')
            summary = self.summary(app)
            self.assertEqual(summary['Pays'].tolist(), ['RDC'])
            self.assertEqual(summary['Anciennes'].tolist(), [3])
            self.assertEqual(summary['Corrigées'].tolist(), [2])
            self.select(app, 'État dans l’historique', 'Corrigée')
            exported = self.exported_comparison()
            self.assertEqual(len(exported), 2)
            self.assertEqual(exported['État historique'].tolist(), ['Corrigée', 'Corrigée'])
            visible = next(item.value for item in app.dataframe if 'État historique' in item.value)
            pd.testing.assert_frame_equal(exported, visible.reset_index(drop=True).fillna('').astype(str), check_dtype=False)
            self.assertFalse(any(column.startswith('_') or column.endswith('_ID') for column in exported))
            self.assertNotIn('PRIVATE_', self.fixture.downloads['Télécharger la comparaison historique en CSV'].decode('utf-8-sig'))
            self.select(app, 'État dans l’historique', 'Toujours violée')
            self.assertEqual(self.exported_comparison()['État historique'].tolist(), ['Toujours violée'])
            self.assertEqual(self.fixture.post_calls, calls)
            self.assert_saved_unchanged(app)

    def test_interruption_checkpoints_first_response_and_resume_never_reposts_attempted_targets(self):
        self.fixture.seed()
        self.fixture.interrupt_target = ZONES[1]
        app = self.new_app()
        with self.offline():
            self.run_app(app)
            self.open_history(app)
            self.click(app, 'Vérifier les corrections de cette extraction')
            checkpoint = app.session_state[PREFIX + 'job']
            self.assertEqual([call['ou'] for call in self.fixture.post_calls], list(ZONES[:2]))
            self.assertEqual(checkpoint['batches']['M']['processed_ids'], [ZONES[0]])
            self.assertEqual(checkpoint['inflight_id'], ZONES[1])
            first_response = deepcopy(checkpoint['batches']['M']['frames_by_id'][ZONES[0]])
            self.run_app(app)
            recovered = app.session_state[PREFIX + 'job']
            self.assertEqual(recovered['status'], 'interrupted')
            self.assertEqual(recovered['batches']['M']['pending_ids'], [ZONES[2]])
            self.assertFalse(recovered['batches']['M']['succeeded_by_id'][ZONES[1]])
            self.assertIn(INTERRUPTED_MESSAGE, '; '.join(recovered['batches']['M']['errors_by_id'][ZONES[1]]))
            self.click(app, 'Reprendre le contrôle historique')
            final = app.session_state[PREFIX + 'job']
            self.assertEqual(final['status'], 'complete')
            self.assertEqual([call['ou'] for call in self.fixture.post_calls], list(ZONES))
            pd.testing.assert_frame_equal(final['batches']['M']['frames_by_id'][ZONES[0]], first_response)
            summary = self.summary(app).set_index('Zone de santé')
            self.assertEqual(int(summary.loc['ZS 2', 'Corrigées']), 0)
            self.assertEqual(int(summary.loc['ZS 2', 'Non vérifiables']), 1)
            self.assertFalse(bool(summary.loc['ZS 2', 'Comparaison complète']))
            self.assert_saved_unchanged(app)

    def test_busy_current_analysis_blocks_manual_and_automatic_history_work(self):
        self.fixture.seed()
        self.fixture.pause_after_step = True
        app = self.new_app()
        with self.offline():
            self.run_app(app)
            self.open_history(app)
            self.click(app, 'Vérifier les corrections de cette extraction')
            self.assertEqual([call['ou'] for call in self.fixture.post_calls], [ZONES[0]])
            self.fixture.current_busy = True
            self.run_app(app)
            for label in ('Enregistrer cette extraction dans GitHub', 'Charger l’historique GitHub',
                          'Ouvrir cette extraction', 'Vérifier les corrections de cette extraction'):
                self.assertTrue(self.widget(app, 'button', label).disabled, label)
            self.assertEqual([call['ou'] for call in self.fixture.post_calls], [ZONES[0]])
            self.assert_saved_unchanged(app)
            self.fixture.current_busy = False
            self.fixture.pause_after_step = False
            self.run_app(app)
            self.assertEqual([call['ou'] for call in self.fixture.post_calls], list(ZONES))
            self.assert_saved_unchanged(app)

    def test_changing_account_or_instance_discards_loaded_private_history(self):
        for field, value in [('account', 'PRIVATE_OTHER_LOGIN'),
                             ('source', 'https://other-fixture.invalid/dhis')]:
            with self.subTest(field=field):
                self.fixture = HistoryFixture()
                self.fixture.seed()
                app = self.new_app()
                with self.offline():
                    self.run_app(app)
                    self.open_history(app)
                    self.click(app, 'Vérifier les corrections de cette extraction')
                    remote_before = deepcopy((self.fixture.post_calls, self.fixture.get_calls,
                                              self.fixture.read_calls, self.fixture.list_calls))
                    setattr(self.fixture, field, value)
                    self.run_app(app)
                    for key in ('files', 'reference', 'job', 'rechecked', 'verified'):
                        self.assertNotIn(PREFIX + key, app.session_state)
                    self.assertFalse(any('Anciennes' in item.value for item in app.dataframe))
                    self.assertEqual((self.fixture.post_calls, self.fixture.get_calls,
                                      self.fixture.read_calls, self.fixture.list_calls), remote_before)
                    self.assert_saved_unchanged(app)


if __name__ == '__main__':
    unittest.main()

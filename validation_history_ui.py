"""Historique volontaire des extractions, distinct des analyses M et M-1."""

from datetime import datetime, timezone
import json
import re
from uuid import uuid4

import pandas as pd
import streamlit as st

from github_history import GitHubHistoryError, GitHubHistoryStore
from quality_analysis import analyse_quality_period
from validation_details import build_validation_details, validation_detail_identifiers
from validation_detail_view import enrich_detail_geography
from validation_filters import date_range_months
from validation_history import (
    account_fingerprint, build_snapshot, compare_snapshots, snapshot_filename,
    source_fingerprint, validate_snapshot,
)
from validation_history_checks import verify_history_absences
from validation_job import (
    advance_validation_job, create_validation_job, recover_validation_job, resume_validation_job,
)


DEFAULT_REPOSITORY = 'InnocentEsomaJr/Extraction_Favori'
PREFIX = 'validation_history_'
COUNTS = ['Anciennes', 'Persistantes', 'Corrigées', 'Nouvelles',
          'Nouvelles dans le filtre actuel', 'Non vérifiables']


def _now():
    return datetime.now(timezone.utc).isoformat(timespec='microseconds')


def clear_history_session():
    for key in list(st.session_state):
        if isinstance(key, str) and key.startswith(PREFIX):
            st.session_state.pop(key, None)


def _with_scope_errors(batch, errors):
    result = dict(batch)
    issues = dict(batch.get('errors_by_id', {}))
    for identifier, reasons in (errors or {}).items():
        if reasons:
            previous = issues.get(identifier, [])
            previous = [previous] if isinstance(previous, str) else list(previous)
            reasons = [reasons] if isinstance(reasons, str) else list(reasons)
            issues[identifier] = list(dict.fromkeys(previous + reasons))
    result['errors_by_id'] = issues
    return result


def _verify(get_client, previous, current, allowed):
    try:
        return verify_history_absences(get_client(), previous, current, allowed)
    except Exception:
        return set(), ['Les métadonnées nécessaires au contrôle des corrections sont indisponibles.']


def _snapshot_from_current(saved, source_url, account, targets, start_date, end_date,
                           parent_id, parent_name, group_id, group_name):
    batch = (saved or {}).get('current_batch')
    if not batch or not batch.get('processed_ids'):
        return None
    anchor = (saved.get('current_extraction_id', saved.get('analysed_at')),
              tuple(batch.get('processed_ids', [])), tuple(batch.get('pending_ids', [])))
    cached = st.session_state.get(PREFIX + 'pending_snapshot')
    if cached and cached['anchor'] == anchor:
        return cached['payload']
    payload = build_snapshot(
        source_url=source_url, account=account, start_date=start_date, end_date=end_date,
        parent_id=parent_id, parent_name=parent_name, group_id=group_id, group_name=group_name,
        targets=targets, batch=_with_scope_errors(batch, saved.get('current_request_errors')),
        raw=saved.get('current_assigned', pd.DataFrame()),
        details=saved.get('violations', pd.DataFrame()), rule_metadata=saved.get('rule_metadata', {}),
        analysed_at=saved['analysed_at'],
    )
    st.session_state[PREFIX + 'pending_snapshot'] = {'anchor': anchor, 'payload': payload}
    return payload


def _saved_rows(snapshot, allowed):
    return pd.DataFrame([
        {'Pays': row['countryName'], 'Province': row['provinceName'],
         'Zone de santé': row['zoneName'], 'Unité d’organisation': row['organisationUnitName'],
         'Période': row['periodName'] or row['periodID'], 'Règle de validation': row['validationRuleName'],
         'Importance': row['importance'], 'Valeur gauche': row['leftValue'],
         'Opérateur': row['operator'], 'Valeur droite': row['rightValue']}
        for row in snapshot['violations'] if row['targetID'] in allowed
    ])


def _render_comparison(previous, current, allowed, verified, warnings):
    compared = compare_snapshots(previous, current, allowed, verified_absences=verified)
    summary, details = compared['summary'].copy(), compared['details'].copy()
    st.caption(
        f"Extraction enregistrée : du {previous['filters']['startDate']} au {previous['filters']['endDate']} "
        f"(lecture du {previous['analysedAt']}). "
        f"Analyse comparée : du {current['filters']['startDate']} au {current['filters']['endDate']} "
        f"(lecture du {current['analysedAt']})."
    )
    issues = list(dict.fromkeys([*warnings, *compared['warnings']]))
    if issues:
        with st.expander('Précisions sur la comparaison historique'):
            for message in issues:
                st.text(message)
    if summary.empty:
        st.info("Aucune zone commune avec le filtre géographique actuel.")
        return
    old_status = {target['ID']: target['status'] for target in previous['targets']}
    new_status = {target['ID']: target['status'] for target in current['targets']}
    summary['Extraction complète'] = summary['Zone_ID'].map(old_status).eq('complete')
    summary['Comparaison complète'] = summary['Zone_ID'].map(new_status).eq('complete')
    level = st.selectbox('Regrouper l’historique par', ['Zones de santé', 'Provinces / DPS', 'Pays'],
                         key=PREFIX + 'summary_level')
    if level == 'Zones de santé':
        shown = summary.drop(columns=['Zone_ID', 'Pays_ID', 'Province_ID'])
    else:
        keys = ['Pays_ID', 'Pays'] + (['Province_ID', 'Province'] if level == 'Provinces / DPS' else [])
        shown = summary.groupby(keys, dropna=False, sort=True).agg({
            **{column: 'sum' for column in COUNTS},
            'Extraction complète': 'all', 'Comparaison complète': 'all',
        }).reset_index().drop(columns=[column for column in keys if column.endswith('_ID')])
    st.dataframe(shown, use_container_width=True, hide_index=True)
    st.caption(
        'Les compteurs portent sur les violations distinctes observées. Une analyse partielle ne prouve pas '
        'l’absence d’erreurs. « Corrigée » constate une absence au contrôle comparable, sans préciser '
        'la date exacte ni l’auteur du changement.'
    )
    status = st.selectbox('État dans l’historique',
                         ['Tous', 'Corrigée', 'Toujours violée', 'Nouvelle',
                          'Nouvelle dans le filtre actuel', 'Non vérifiable'], key=PREFIX + 'status')
    if status != 'Tous':
        details = details.loc[details['État historique'].eq(status)]
    public = [column for column in details if not column.startswith('_') and column not in {
        'Zone_ID', 'dayInPeriod', 'attributeOptionComboID', 'categoryOptionComboID', 'dataSetID',
    }]
    visible = details[public]
    st.dataframe(visible, use_container_width=True, hide_index=True)
    st.download_button('Télécharger la comparaison historique en CSV',
                       visible.to_csv(index=False).encode('utf-8-sig'),
                       file_name='comparaison_historique.csv', mime='text/csv', key=PREFIX + 'csv')


def render_history_panel(*, saved, source_url, account, targets, geography, start_date, end_date,
                         parent_id, parent_name, group_id, group_name, cache_user,
                         read_config, fetch_validation, get_client, get_metadata, current_busy=False):
    """Affiche les sauvegardes et avance au plus une ZS du contrôle historique.

    Aucun appel GitHub ni DHIS2 n'est déclenché à l'ouverture initiale : les
    lectures distantes et les écritures nécessitent leur bouton explicite.
    ``True`` demande un rerun après affichage des tableaux et exports courants.
    """
    access = (source_fingerprint(source_url), account_fingerprint(account))
    if st.session_state.get(PREFIX + 'access') != access:
        clear_history_session()
        st.session_state[PREFIX + 'access'] = access
    allowed = set(targets['Zone_ID']) if 'Zone_ID' in targets else set()
    rerun = False
    active_job = st.session_state.get(PREFIX + 'job')
    reference_entry = st.session_state.get(PREFIX + 'reference')
    if active_job and active_job.get('status') == 'running':
        reference = reference_entry['payload'] if reference_entry else None
        expected = ((reference['snapshotID'], access,
                     tuple(sorted(allowed.intersection(target['ID'] for target in reference['targets']))),
                     tuple(cache_user)) if reference else None)
        if active_job.get('context') != expected:
            active_job['status'], active_job['auto_continue'] = 'interrupted', False
            st.session_state[PREFIX + 'job'] = active_job
            rerun = True
    history_busy = bool(active_job and active_job.get('status') == 'running')
    with st.expander('Historique des règles de validation'):
        st.caption(
            'Enregistrez une extraction avant le feedback, puis contrôlez à nouveau ses mêmes dates. '
            'Les fichiers enregistrés restent intacts. Les résultats du filtre actuel et M-1 sont conservés.'
        )
        repository = read_config('GITHUB_HISTORY_REPOSITORY') or DEFAULT_REPOSITORY
        branch = read_config('GITHUB_HISTORY_BRANCH') or 'validation-history'
        directory = read_config('GITHUB_HISTORY_DIRECTORY') or 'validation-history'
        token = read_config('GITHUB_HISTORY_TOKEN') or read_config('GITHUB_TOKEN') or None
        try:
            store = GitHubHistoryStore(repository, token, branch=branch, directory=directory)
        except GitHubHistoryError:
            store = None
        if not token or store is None:
            st.info('Pour enregistrer depuis Streamlit, configurez GITHUB_HISTORY_TOKEN dans les Secrets '
                    'avec la permission Contents: Read and write sur le dépôt choisi.')
            with st.expander('Configuration de l’historique GitHub'):
                st.code('GITHUB_HISTORY_TOKEN = "votre-jeton"\n'
                        f'GITHUB_HISTORY_REPOSITORY = "{DEFAULT_REPOSITORY}"\n'
                        'GITHUB_HISTORY_BRANCH = "validation-history"\n'
                        'GITHUB_HISTORY_DIRECTORY = "validation-history"', language='toml')
                st.markdown('[Guide de configuration du jeton et des Secrets Streamlit]('
                            f'https://github.com/{DEFAULT_REPOSITORY}/blob/main/docs/HISTORIQUE_VALIDATION.md'
                            '#configuration-dans-streamlit-cloud)')
        st.caption('Les fichiers sont lisibles par les personnes ayant accès au dépôt GitHub configuré. '
                   'Aucun mot de passe, jeton ou identifiant de connexion DHIS2 n’est enregistré dans les fichiers.')
        try:
            current = (None if current_busy else
                       _snapshot_from_current(saved, source_url, account, targets, start_date, end_date,
                                              parent_id, parent_name, group_id, group_name))
        except (ValueError, TypeError, KeyError):
            current = None
            st.info('Cette extraction ne peut pas encore être préparée pour l’historique.')
        if st.button('Enregistrer cette extraction dans GitHub',
                     disabled=current is None or not token or store is None or current_busy or history_busy,
                     key=PREFIX + 'save'):
            try:
                filename = snapshot_filename(current)
                saved_file = store.save_snapshot(filename, current, current['snapshotID'])
                st.session_state[PREFIX + 'files'] = sorted(set(
                    st.session_state.get(PREFIX + 'files', []) + [saved_file.name]), reverse=True)
                st.success('Extraction enregistrée dans GitHub.' if saved_file.created else
                           'Cette extraction est déjà enregistrée ; elle n’a pas été remplacée.')
                st.caption(saved_file.name)
            except GitHubHistoryError as error:
                st.error(str(error))
        if current is not None:
            st.caption(f"Extraction à enregistrer : M du {current['filters']['startDate']} "
                       f"au {current['filters']['endDate']}.")
            if any(target['status'] != 'complete' for target in current['targets']):
                st.caption('Extraction partielle : les résultats reçus peuvent être archivés, '
                           'mais les unités incomplètes ne prouvent aucune correction.')
            download = st.session_state.get(PREFIX + 'download')
            if not download or download['ID'] != current['snapshotID']:
                download = {'ID': current['snapshotID'], 'data':
                            json.dumps(current, ensure_ascii=False, allow_nan=False, indent=2).encode('utf-8')}
                st.session_state[PREFIX + 'download'] = download
            st.download_button('Télécharger cette extraction JSON', download['data'],
                               file_name=snapshot_filename(current), mime='application/json', key=PREFIX + 'json')
        if st.button('Charger l’historique GitHub', disabled=store is None or current_busy or history_busy,
                     key=PREFIX + 'load'):
            try:
                st.session_state[PREFIX + 'files'] = [entry.name for entry in store.list_snapshots()]
            except GitHubHistoryError as error:
                st.error(str(error))
        files = st.session_state.get(PREFIX + 'files', [])
        if not files:
            return rerun
        file_months = {}
        for name in files:
            match = re.match(r'^[a-z]+_(\d{4}-\d{2}-\d{2})_(\d{4}-\d{2}-\d{2})_', name)
            try:
                file_months[name] = set(date_range_months(match.group(1), match.group(2))) if match else set()
            except ValueError:
                file_months[name] = set()
        months = sorted(set().union(*file_months.values()), reverse=True)
        month_options = ['Tous', *months]
        if st.session_state.get(PREFIX + 'month') not in month_options:
            st.session_state.pop(PREFIX + 'month', None)
        month = st.selectbox('Mois de l’historique', month_options, key=PREFIX + 'month', disabled=history_busy)
        options = [name for name in files if month == 'Tous' or month in file_months[name]]
        if not options:
            return rerun
        if st.session_state.get(PREFIX + 'file') not in options:
            st.session_state.pop(PREFIX + 'file', None)
        selected = st.selectbox('Extraction enregistrée', options, key=PREFIX + 'file', disabled=history_busy)
        if st.button('Ouvrir cette extraction', disabled=store is None or current_busy or history_busy, key=PREFIX + 'open'):
            try:
                reference = validate_snapshot(store.read_snapshot(selected))
                st.session_state[PREFIX + 'reference'] = {'name': selected, 'payload': reference}
                for key in ('job', 'rechecked', 'verified', 'warnings'):
                    st.session_state.pop(PREFIX + key, None)
            except (GitHubHistoryError, ValueError, TypeError, KeyError):
                st.error('Cette extraction est indisponible ou son format historique n’est pas valide.')
        entry = st.session_state.get(PREFIX + 'reference')
        if not entry or entry['name'] != selected:
            return rerun
        previous = entry['payload']
        shared = allowed.intersection(target['ID'] for target in previous['targets'])
        if previous['sourceID'] != access[0] or previous['accountID'] != access[1] or not shared:
            st.info('Cette extraction ne correspond pas à l’instance, au compte et aux zones du filtre actuel.')
            return rerun
        st.caption(f"Période enregistrée : {previous['filters']['startDate']} → {previous['filters']['endDate']}. "
                   f"{len(shared)} zone(s) commune(s) avec le filtre géographique actuel.")
        with st.expander('Voir les violations de l’extraction enregistrée'):
            st.dataframe(_saved_rows(previous, shared), use_container_width=True, hide_index=True)
        context = (previous['snapshotID'], access, tuple(sorted(shared)), tuple(cache_user))
        job = st.session_state.get(PREFIX + 'job')
        historical_targets = targets.loc[targets['Zone_ID'].isin(shared)].copy()
        months = date_range_months(previous['filters']['startDate'], previous['filters']['endDate'])
        if job and job.get('context') != context:
            job['status'] = 'interrupted'
            job['auto_continue'] = False
            st.session_state[PREFIX + 'job'] = job
        elif job and job.get('status') == 'running':
            job = recover_validation_job(job, historical_targets, months, [])
        start = st.button('Vérifier les corrections de cette extraction', key=PREFIX + 'recheck',
                          disabled=current_busy or previous['accountID'] != access[1]
                          or bool(job and job.get('status') == 'running'))
        resume = False
        if job and job.get('context') == context and job.get('status') == 'interrupted':
            st.info('Contrôle historique interrompu : les réponses reçues sont conservées.')
            resume = st.button('Reprendre le contrôle historique', disabled=current_busy, key=PREFIX + 'resume')
        if start:
            job = create_validation_job(context, 'M', (*cache_user, uuid4().hex))
            job['id'], job['analysed_at'] = uuid4().hex, _now()
            st.session_state.pop(PREFIX + 'rechecked', None)
        elif resume:
            job = resume_validation_job(job)
        if job and job.get('context') == context:
            st.session_state[PREFIX + 'job'] = job
            if job['status'] == 'running' and not current_busy and (start or resume or job.get('auto_continue')):
                job['auto_continue'] = False
                st.session_state[PREFIX + 'job'] = job
                progress = st.progress(0.0, text='Contrôle de l’extraction enregistrée')
                def checkpoint(value):
                    st.session_state[PREFIX + 'job'] = value
                job = advance_validation_job(
                    job, historical_targets, months, [], previous['filters']['groupID'], fetch_validation,
                    start_date=previous['filters']['startDate'], end_date=previous['filters']['endDate'],
                    on_checkpoint=checkpoint,
                    on_progress=lambda position, total, uid, name: progress.progress(
                        (position - 1) / max(1, total), text=f'Contrôle historique : {position}/{total} — {name}'),
                )
                progress.empty()
                job['auto_continue'] = job['status'] == 'running'
                st.session_state[PREFIX + 'job'] = job
                rerun = True
            batch = job['batches'].get('M')
            if batch and batch.get('processed_ids'):
                count = len(batch['processed_ids'])
                st.caption(f'Contrôle historique : {count}/{len(historical_targets)} zone(s) traitée(s).')
                anchor = (job['id'], tuple(batch['processed_ids']), job['status'])
                checked = st.session_state.get(PREFIX + 'rechecked')
                if not checked or checked['anchor'] != anchor:
                    raw = batch['raw']
                    identifiers = validation_detail_identifiers(raw)
                    rule_meta = dict(job.get('history_rules', {}))
                    unit_meta = dict(job.get('history_units', {}))
                    missing_rules = tuple(uid for uid in identifiers['rule_ids'] if uid not in rule_meta)
                    missing_units = tuple(uid for uid in identifiers['org_ids'] if uid not in unit_meta)
                    if missing_rules or missing_units:
                        new_rules, new_units, _ = get_metadata(missing_rules, missing_units, cache_user=job['cache_user'])
                        rule_meta.update(new_rules)
                        unit_meta.update(new_units)
                    job.update(history_rules=rule_meta, history_units=unit_meta)
                    st.session_state[PREFIX + 'job'] = job
                    assigned, scope_errors, _ = analyse_quality_period(batch, historical_targets, unit_meta)
                    details = enrich_detail_geography(build_validation_details(assigned, rule_meta, unit_meta),
                                                      assigned, geography, unit_meta)
                    latest = build_snapshot(
                        source_url=source_url, account=account,
                        start_date=previous['filters']['startDate'], end_date=previous['filters']['endDate'],
                        parent_id=parent_id, parent_name=parent_name,
                        group_id=previous['filters']['groupID'], group_name=previous['filters']['groupName'],
                        targets=historical_targets, batch=_with_scope_errors(batch, scope_errors),
                        raw=assigned, details=details,
                        rule_metadata=rule_meta, analysed_at=job['analysed_at'],
                    )
                    verified, warnings = (_verify(get_client, previous, latest, shared)
                                          if job['status'] != 'running' else
                                          (set(), ['Le contrôle historique est encore en cours.']))
                    checked = {'anchor': anchor, 'payload': latest, 'verified': verified, 'warnings': warnings}
                    st.session_state[PREFIX + 'rechecked'] = checked
                if job['status'] == 'complete':
                    st.success('Contrôle historique terminé. Les résultats du filtre actuel sont conservés.')
        checked = st.session_state.get(PREFIX + 'rechecked')
        choices = (['Contrôle de l’extraction enregistrée'] if checked else []) + (['Analyse du filtre actuel'] if current else [])
        if choices:
            key = PREFIX + 'comparison_source'
            if st.session_state.get(key) not in choices:
                st.session_state.pop(key, None)
            mode = st.selectbox('Comparer avec', choices, key=key)
            if mode == 'Contrôle de l’extraction enregistrée':
                latest, verified, warnings = checked['payload'], checked['verified'], checked['warnings']
            else:
                latest = current
                anchor = (previous['snapshotID'], current['snapshotID'], tuple(sorted(shared)))
                proof = st.session_state.get(PREFIX + 'verified')
                if st.button('Vérifier la comparaison avec l’analyse actuelle', disabled=current_busy,
                             key=PREFIX + 'verify_current'):
                    verified, warnings = _verify(get_client, previous, current, shared)
                    proof = {'anchor': anchor, 'verified': verified, 'warnings': warnings}
                    st.session_state[PREFIX + 'verified'] = proof
                verified, warnings = ((proof['verified'], proof['warnings']) if proof and proof['anchor'] == anchor else
                                      (set(), ['Une disparition nécessite un contrôle comparable pour confirmer une correction.']))
            _render_comparison(previous, latest, shared, verified, warnings)
    return rerun

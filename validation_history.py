"""Instantanés JSON publics et comparaison prudente des extractions de validation.

Les lignes serveur restent intactes en nombre dans l'instantané. Les compteurs
de comparaison portent sur les identités distinctes, jamais sur les doublons.
Une disparition n'est une correction qu'après une preuve externe explicite.
"""

from collections import Counter
from datetime import date, datetime, timezone
import hashlib
import json
import math
import re
from urllib.parse import urlsplit, urlunsplit
from uuid import uuid4

import pandas as pd


UID_RE = re.compile(r"[A-Za-z][A-Za-z0-9]{10}\Z")
HASH_RE = re.compile(r"[0-9a-f]{64}\Z")
MONTHS = ('janvier', 'fevrier', 'mars', 'avril', 'mai', 'juin',
          'juillet', 'aout', 'septembre', 'octobre', 'novembre', 'decembre')
FILTER_FIELDS = {'startDate', 'endDate', 'parentID', 'parentName', 'groupID', 'groupName'}
TARGET_FIELDS = {'ID', 'name', 'countryID', 'countryName', 'provinceID', 'provinceName',
                 'zoneID', 'zoneName', 'status'}
VIOLATION_FIELDS = {
    'targetID', 'organisationUnitID', 'validationRuleID', 'periodID', 'dayInPeriod',
    'attributeOptionComboID', 'categoryOptionComboID', 'dataSetID',
    'countryID', 'countryName', 'provinceID', 'provinceName', 'zoneID', 'zoneName',
    'organisationUnitName', 'validationRuleName', 'periodName', 'importance',
    'leftValue', 'operator', 'rightValue', 'ruleFingerprint',
}
TOP_FIELDS = {'schemaVersion', 'snapshotID', 'sourceID', 'accountID', 'capturedAt',
              'analysedAt', 'filters', 'targets', 'violations'}
PERSISTENT = 'Toujours violée'
CORRECTED = 'Corrigée'
NEW = 'Nouvelle'
NEW_FILTER = 'Nouvelle dans le filtre actuel'
UNVERIFIABLE = 'Non vérifiable'
SUMMARY_COLUMNS = ['Zone_ID', 'Pays_ID', 'Pays', 'Province_ID', 'Province',
                   'Zone de santé', 'Anciennes', 'Persistantes', 'Corrigées',
                   'Nouvelles', 'Nouvelles dans le filtre actuel', 'Non vérifiables']
DETAIL_COLUMNS = [
    'Pays', 'Province', 'Zone de santé', 'Unité d’organisation', 'Période',
    'Importance', 'Règle de validation', 'Valeur gauche', 'Opérateur', 'Valeur droite',
    'État historique', 'Source', 'Occurrences anciennes', 'Occurrences actuelles',
    'Zone_ID', '_country_id', '_province_id', '_health_zone_id', '_ou_id', '_rule_id',
    '_period_id', 'dayInPeriod', 'attributeOptionComboID', 'categoryOptionComboID', 'dataSetID',
]


def _text(value):
    if value is None or isinstance(value, (dict, list, tuple, set)):
        return ''
    try:
        if pd.isna(value):
            return ''
    except (TypeError, ValueError):
        return ''
    return str(value).strip()


def _uid(value):
    if isinstance(value, dict):
        return _uid(value.get('id')) or _uid(value.get('uid'))
    value = _text(value)
    return value if UID_RE.fullmatch(value) else None


def _first_text(*values):
    return next((text for value in values if (text := _text(value))), '')


def _row_uid(row, key, *fallbacks):
    return next((uid for value in [row.get(key + 'ID'), row.get(key + 'Id'), row.get(key)]
                 + [row.get(name) for name in fallbacks] if (uid := _uid(value))), None)


def _name(value):
    return _first_text(value.get('displayName'), value.get('name')) if isinstance(value, dict) else ''


def _hash(value):
    return hashlib.sha256(value.encode('utf-8')).hexdigest()


def source_fingerprint(source_url):
    """Identifie l'instance sans publier son URL ou ses éventuels identifiants."""
    parts = urlsplit(_text(source_url))
    if parts.scheme.lower() not in {'http', 'https'} or not parts.hostname:
        raise ValueError("URL source invalide.")
    host = parts.hostname.lower()
    if ':' in host:
        host = '[' + host + ']'
    port = parts.port
    if port is not None and (parts.scheme.lower(), port) not in {('http', 80), ('https', 443)}:
        host += ':' + str(port)
    canonical = urlunsplit((parts.scheme.lower(), host, parts.path.rstrip('/'), '', ''))
    return _hash(canonical)


def account_fingerprint(account):
    value = _text(account)
    if not value:
        raise ValueError("Le compte source est requis.")
    return _hash(value)


def rule_fingerprint(metadata):
    """Empreinte de la définition ; ne conserve aucune expression en clair."""
    if not isinstance(metadata, dict):
        return None
    left, right = metadata.get('leftSide'), metadata.get('rightSide')
    if not isinstance(left, dict) or not isinstance(right, dict):
        return None
    values = {'operator': _text(metadata.get('operator')).casefold(),
              'leftExpression': _text(left.get('expression')),
              'rightExpression': _text(right.get('expression'))}
    if not all(values.values()):
        return None
    return _hash(json.dumps(values, sort_keys=True, ensure_ascii=False, separators=(',', ':')))


def _date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = _text(value)
    if not re.fullmatch(r'\d{4}-\d{2}-\d{2}', text):
        raise ValueError("Date ISO invalide.")
    return date.fromisoformat(text)


def _timestamp(value):
    if not isinstance(value, datetime):
        value = datetime.fromisoformat(_text(value).replace('Z', '+00:00'))
    if value.tzinfo is None or value.utcoffset() is None:
        raise ValueError("Une date de capture ou d'analyse doit avoir un fuseau horaire.")
    return value.isoformat(timespec='microseconds')


def _period(value):
    value = _text(value)
    if re.fullmatch(r'\d{4}', value):
        return value if 1 <= int(value) <= 9999 else None
    if re.fullmatch(r'\d{6}', value):
        return value if 1 <= int(value[4:]) <= 12 and int(value[:4]) > 0 else None
    if re.fullmatch(r'\d{8}', value):
        try:
            datetime.strptime(value, '%Y%m%d')
            return value
        except ValueError:
            return None
    suffix = (r'(?:(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)?W(?:0?[1-9]|[1-4][0-9]|5[0-3])'
              r'|Q[1-4]|S[12]|(?:0[1-9]|1[0-2])B|(?:April|July|Oct|Nov|August)(?:S[12])?)')
    return value if re.fullmatch(r'[1-9]\d{3}' + suffix, value) else None


def _row_period(row, detail):
    embedded = row.get('period')
    embedded = embedded if isinstance(embedded, dict) else {}
    return next((period for value in (
        row.get('periodID'), row.get('periodId'), embedded.get('iso'), embedded.get('isoDate'),
        embedded.get('id'), row.get('period'), row.get('_requested_period'), detail.get('_period_id'),
    ) if (period := _period(value))), None)


def _number(value):
    if isinstance(value, bool) or not _text(value):
        return None
    try:
        number = float(value)
        return number if math.isfinite(number) else None
    except (TypeError, ValueError, OverflowError):
        return None


def _day(value):
    number = _number(value)
    return int(number) if number is not None and number.is_integer() and 0 <= number <= 366 else None


def violation_key(row):
    """Identité stricte d'une observation ; les valeurs et libellés ne comptent pas."""
    if not isinstance(row, dict):
        return None
    ou, rule, period = (_uid(row.get('organisationUnitID')), _uid(row.get('validationRuleID')),
                        _period(row.get('periodID')))
    if not ou or not rule or not isinstance(row.get('periodID'), str) or period != row['periodID']:
        return None
    day = row.get('dayInPeriod')
    if day is not None and (type(day) is not int or not 0 <= day <= 366):
        return None
    for field in ('organisationUnitID', 'validationRuleID', 'attributeOptionComboID',
                  'categoryOptionComboID', 'dataSetID'):
        value = row.get(field)
        if value is not None and (not isinstance(value, str) or _uid(value) != value):
            return None
    return (ou, rule, period, day,
            _uid(row.get('attributeOptionComboID')), _uid(row.get('categoryOptionComboID')),
            _uid(row.get('dataSetID')))


def _target_status(identifier, batch):
    if identifier in batch.get('pending_ids', []) or identifier not in batch.get('processed_ids', []):
        return 'pending'
    frame = batch.get('frames_by_id', {}).get(identifier)
    attrs = frame.attrs if isinstance(frame, pd.DataFrame) else {}
    if attrs.get('truncated_periods'):
        return 'truncated'
    if (batch.get('succeeded_by_id', {}).get(identifier) is not True
            or batch.get('errors_by_id', {}).get(identifier) or attrs.get('failed_periods')):
        return 'failed'
    return 'complete'


def build_snapshot(*, source_url, account, start_date, end_date, parent_id, parent_name,
                   group_id, group_name, targets, batch, raw, details, rule_metadata,
                   analysed_at, captured_at=None, snapshot_id=None):
    """Construit une whitelist JSON : aucun contexte, secret ou attribut brut."""
    first, last = _date(start_date), _date(end_date)
    if first > last:
        raise ValueError("L'intervalle de l'instantané est inversé.")
    if not isinstance(targets, pd.DataFrame) or not isinstance(raw, pd.DataFrame):
        raise ValueError("Les cibles et violations doivent être des tableaux.")
    if not isinstance(batch, dict):
        raise ValueError("L'état des requêtes est requis.")
    detail_rows = details.to_dict('records') if isinstance(details, pd.DataFrame) else []
    if detail_rows and len(detail_rows) != len(raw):
        raise ValueError("Les détails ne correspondent pas aux lignes source.")
    normalized_targets = []
    for target in targets.to_dict('records'):
        identifier = _uid(target.get('Zone_ID')) or _uid(target.get('ID'))
        if not identifier:
            raise ValueError("Une cible n'a pas d'UID valide.")
        normalized_targets.append({
            'ID': identifier, 'name': _first_text(target.get('Zone'), target.get('name')),
            'countryID': _uid(target.get('Pays_ID')), 'countryName': _text(target.get('Pays')),
            'provinceID': _uid(target.get('Province_ID')), 'provinceName': _text(target.get('Province')),
            'zoneID': identifier, 'zoneName': _first_text(target.get('Zone'), target.get('name')),
            'status': _target_status(identifier, batch),
        })
    target_by_id = {target['ID']: target for target in normalized_targets}
    if len(target_by_id) != len(normalized_targets):
        raise ValueError("Les cibles de l'instantané sont répétées.")
    violations = []
    for index, row in enumerate(raw.to_dict('records')):
        detail = detail_rows[index] if detail_rows else {}
        ou = _row_uid(row, 'organisationUnit', '_ou_id') or _uid(detail.get('_ou_id'))
        rule = _row_uid(row, 'validationRule', '_rule_id') or _uid(detail.get('_rule_id'))
        target_id = next((uid for value in (
            row.get('_request_unit_id'), row.get('targetID'), detail.get('_health_zone_id'),
            row.get('_analysis_target_id'), row.get('_zone_id'), ou,
        ) if (uid := _uid(value)) in target_by_id), None)
        target = target_by_id.get(target_id, {})
        meta = (rule_metadata or {}).get(rule)
        if not isinstance(meta, dict):
            meta = row.get('validationRule') if isinstance(row.get('validationRule'), dict) else {}
        record = {
            'targetID': target_id, 'organisationUnitID': ou, 'validationRuleID': rule,
            'periodID': _row_period(row, detail), 'dayInPeriod': _day(row.get('dayInPeriod')),
            'attributeOptionComboID': _row_uid(row, 'attributeOptionCombo'),
            'categoryOptionComboID': _row_uid(row, 'categoryOptionCombo'),
            'dataSetID': _row_uid(row, 'dataSet'),
            'organisationUnitName': _first_text(row.get('organisationUnitDisplayName'),
                                                _name(row.get('organisationUnit')), detail.get('Unité d’organisation')),
            'validationRuleName': _first_text(row.get('validationRuleDescription'), row.get('validationRuleDisplayName'),
                                              _name(row.get('validationRule')), detail.get('Règle de validation')),
            'periodName': _first_text(row.get('periodDisplayName'), _name(row.get('period')), detail.get('Période')),
            'importance': _first_text(row.get('importance'), detail.get('Importance')),
            'leftValue': _number(row.get('leftSideValue', row.get('leftsideValue', row.get('leftValue')))),
            'operator': _first_text(row.get('operator'), detail.get('Opérateur')),
            'rightValue': _number(row.get('rightSideValue', row.get('rightsideValue', row.get('rightValue')))),
            'ruleFingerprint': rule_fingerprint(meta),
        }
        for prefix, detail_id, detail_name in (
            ('country', '_country_id', 'Pays'), ('province', '_province_id', 'Province'),
            ('zone', '_health_zone_id', 'Zone de santé'),
        ):
            record[prefix + 'ID'] = _uid(detail.get(detail_id)) or target.get(prefix + 'ID')
            record[prefix + 'Name'] = _first_text(detail.get(detail_name), target.get(prefix + 'Name'))
        violations.append(record)
    result = {
        'schemaVersion': 1, 'snapshotID': snapshot_id if snapshot_id is not None else uuid4().hex,
        'sourceID': source_fingerprint(source_url), 'accountID': account_fingerprint(account),
        'capturedAt': _timestamp(captured_at if captured_at is not None else datetime.now(timezone.utc)),
        'analysedAt': _timestamp(analysed_at),
        'filters': {'startDate': first.isoformat(), 'endDate': last.isoformat(),
                    'parentID': _uid(parent_id), 'parentName': _text(parent_name),
                    'groupID': _uid(group_id), 'groupName': _text(group_name)},
        'targets': normalized_targets, 'violations': violations,
    }
    if not result['filters']['parentID'] or (_text(group_id) and not result['filters']['groupID']):
        raise ValueError("Le parent ou groupe n'a pas d'UID valide.")
    validate_snapshot(result)
    return result


def _shape(value, fields, label):
    if not isinstance(value, dict) or set(value) != fields:
        raise ValueError("Structure d'instantané invalide : " + label)


def _validate_text(value):
    if not isinstance(value, str) or len(value) > 4000 or any(ord(char) < 32 and char not in '\n\t\r' for char in value):
        raise ValueError("Texte d'instantané invalide.")


def validate_snapshot(payload):
    """Valide strictement le schéma importé, les nombres et les références."""
    _shape(payload, TOP_FIELDS, 'racine')
    if type(payload['schemaVersion']) is not int or payload['schemaVersion'] != 1:
        raise ValueError("Version d'instantané non prise en charge.")
    if not isinstance(payload['snapshotID'], str) or not re.fullmatch(r'[0-9a-f]{32}', payload['snapshotID']):
        raise ValueError("Identifiant d'instantané invalide.")
    for name in ('sourceID', 'accountID'):
        if not isinstance(payload[name], str) or not HASH_RE.fullmatch(payload[name]):
            raise ValueError("Empreinte d'instantané invalide.")
    for name in ('capturedAt', 'analysedAt'):
        if not isinstance(payload[name], str):
            raise ValueError("Horodatage d'instantané invalide.")
        _timestamp(payload[name])
    if (datetime.fromisoformat(payload['capturedAt'].replace('Z', '+00:00'))
            < datetime.fromisoformat(payload['analysedAt'].replace('Z', '+00:00'))):
        raise ValueError("La capture précède la date de l'analyse.")
    filters = payload['filters']
    _shape(filters, FILTER_FIELDS, 'filtres')
    if not isinstance(filters['startDate'], str) or not isinstance(filters['endDate'], str):
        raise ValueError("Intervalle d'instantané invalide.")
    if _date(filters['startDate']) > _date(filters['endDate']):
        raise ValueError("Intervalle d'instantané inversé.")
    if _uid(filters['parentID']) != filters['parentID'] or filters['parentID'] is None:
        raise ValueError("Parent d'instantané invalide.")
    if filters['groupID'] is not None and _uid(filters['groupID']) != filters['groupID']:
        raise ValueError("Groupe d'instantané invalide.")
    for field in ('parentName', 'groupName'):
        _validate_text(filters[field])
    if not isinstance(payload['targets'], list) or not isinstance(payload['violations'], list):
        raise ValueError("Les cibles et violations doivent être des listes JSON.")
    target_ids = set()
    for target in payload['targets']:
        _shape(target, TARGET_FIELDS, 'cible')
        if not isinstance(target['ID'], str) or not _uid(target['ID']) or target['ID'] in target_ids:
            raise ValueError("UID de cible invalide ou répété.")
        target_ids.add(target['ID'])
        if not isinstance(target['status'], str) or target['status'] not in {'complete', 'failed', 'pending', 'truncated'}:
            raise ValueError("État de cible invalide.")
        for field, value in target.items():
            if field.endswith('ID'):
                if value is not None and (not isinstance(value, str) or _uid(value) != value):
                    raise ValueError("UID géographique invalide.")
            elif field != 'status':
                _validate_text(value)
    for row in payload['violations']:
        _shape(row, VIOLATION_FIELDS, 'violation')
        for field, value in row.items():
            if field == 'periodID':
                if value is not None and (not isinstance(value, str) or _period(value) != value):
                    raise ValueError("Période de violation invalide.")
            elif field.endswith('ID'):
                if value is not None and (not isinstance(value, str) or _uid(value) != value):
                    raise ValueError("UID de violation invalide.")
            elif field == 'dayInPeriod':
                if value is not None and (type(value) is not int or not 0 <= value <= 366):
                    raise ValueError("Jour de violation invalide.")
            elif field in {'leftValue', 'rightValue'}:
                if value is not None:
                    try:
                        finite = type(value) in {int, float} and math.isfinite(value)
                    except OverflowError:
                        finite = False
                    if not finite:
                        raise ValueError("Valeur de violation invalide.")
            elif field == 'ruleFingerprint':
                if value is not None and (not isinstance(value, str) or not HASH_RE.fullmatch(value)):
                    raise ValueError("Empreinte de règle invalide.")
            else:
                _validate_text(value)
        if row['targetID'] is not None and row['targetID'] not in target_ids:
            raise ValueError("La violation référence une cible absente.")
    json.dumps(payload, allow_nan=False)
    return payload


def snapshot_filename(snapshot):
    validate_snapshot(snapshot)
    filters = snapshot['filters']
    first = _date(filters['startDate'])
    stamp = datetime.fromisoformat(snapshot['capturedAt'].replace('Z', '+00:00')).astimezone(timezone.utc)
    return (f"{MONTHS[first.month - 1]}_{filters['startDate']}_{filters['endDate']}_"
            f"{filters['parentID']}__{stamp:%Y%m%dT%H%M%S%fZ}_{snapshot['snapshotID']}.json")


def snapshot_comparison_filters(snapshot):
    validate_snapshot(snapshot)
    return dict(snapshot['filters'], sourceID=snapshot['sourceID'], accountID=snapshot['accountID'])


def _covers(first, second):
    return (first['filters']['startDate'] <= second['filters']['startDate']
            and first['filters']['endDate'] >= second['filters']['endDate'])


def _detail(row, status, origin, old_count, new_count):
    return {
        'Pays': row['countryName'], 'Province': row['provinceName'], 'Zone de santé': row['zoneName'],
        'Unité d’organisation': row['organisationUnitName'], 'Période': row['periodName'] or row['periodID'],
        'Importance': row['importance'], 'Règle de validation': row['validationRuleName'],
        'Valeur gauche': row['leftValue'], 'Opérateur': row['operator'], 'Valeur droite': row['rightValue'],
        'État historique': status, 'Source': origin, 'Occurrences anciennes': old_count,
        'Occurrences actuelles': new_count, 'Zone_ID': row['targetID'],
        '_country_id': row['countryID'], '_province_id': row['provinceID'],
        '_health_zone_id': row['zoneID'], '_ou_id': row['organisationUnitID'],
        '_rule_id': row['validationRuleID'], '_period_id': row['periodID'],
        'dayInPeriod': row['dayInPeriod'], 'attributeOptionComboID': row['attributeOptionComboID'],
        'categoryOptionComboID': row['categoryOptionComboID'], 'dataSetID': row['dataSetID'],
    }


def compare_snapshots(old, new, allowed_target_ids, verified_absences=None):
    """Compare dans le périmètre autorisé ; une absence seule reste inconnue."""
    validate_snapshot(old)
    validate_snapshot(new)
    allowed = set(allowed_target_ids)
    if any(not isinstance(identifier, str) or not _uid(identifier) for identifier in allowed):
        raise ValueError("Le périmètre autorisé contient un UID invalide.")
    verified = set(verified_absences or ())
    old_targets = {target['ID']: target for target in old['targets']}
    new_targets = {target['ID']: target for target in new['targets']}
    selected = sorted(allowed & (set(old_targets) | set(new_targets)))
    same_context = all(old[key] == new[key] for key in ('sourceID', 'accountID')) and (
        old['filters']['groupID'] == new['filters']['groupID'])
    chronological = (datetime.fromisoformat(new['analysedAt'].replace('Z', '+00:00'))
                     >= datetime.fromisoformat(old['analysedAt'].replace('Z', '+00:00')))
    compatible = same_context and chronological
    warnings = []
    if not same_context:
        warnings.append("La source, le compte ou le groupe diffère : les statuts ne sont pas vérifiables.")
    if not chronological:
        warnings.append("L'analyse actuelle est antérieure à l'analyse historique : les statuts ne sont pas vérifiables.")
    if not _covers(new, old):
        warnings.append("L'extraction actuelle ne couvre pas toutes les dates historiques ; aucune disparition hors couverture n'est une correction.")
    old_by_target, new_by_target = {}, {}
    for snapshot, grouped in ((old, old_by_target), (new, new_by_target)):
        for row in snapshot['violations']:
            if row['targetID'] in allowed:
                grouped.setdefault(row['targetID'], []).append(row)
    summaries, details = [], []
    status_columns = {PERSISTENT: 'Persistantes', CORRECTED: 'Corrigées', NEW: 'Nouvelles',
                      NEW_FILTER: 'Nouvelles dans le filtre actuel', UNVERIFIABLE: 'Non vérifiables'}
    for identifier in selected:
        target = new_targets.get(identifier) or old_targets[identifier]
        before = old_by_target.get(identifier, [])
        now = new_by_target.get(identifier, [])
        old_counts = Counter(key for row in before if (key := violation_key(row)) is not None)
        new_counts = Counter(key for row in now if (key := violation_key(row)) is not None)
        old_fingerprints, new_fingerprints = {}, {}
        for rows, fingerprints in ((before, old_fingerprints), (now, new_fingerprints)):
            for row in rows:
                key = violation_key(row)
                if key is not None:
                    fingerprints.setdefault(key, set()).add(row['ruleFingerprint'])
        old_complete = old_targets.get(identifier, {}).get('status') == 'complete'
        new_complete = new_targets.get(identifier, {}).get('status') == 'complete'
        summary = {'Zone_ID': identifier, 'Pays_ID': target['countryID'], 'Pays': target['countryName'],
                   'Province_ID': target['provinceID'], 'Province': target['provinceName'],
                   'Zone de santé': target['zoneName'] or target['name'],
                   'Anciennes': len(old_counts) + sum(violation_key(row) is None for row in before),
                   **{column: 0 for column in status_columns.values()}}
        counted = set()
        for index, row in enumerate(before):
            key = violation_key(row)
            status = UNVERIFIABLE
            if compatible and key is not None:
                if key in new_counts:
                    previous_hashes = old_fingerprints[key] - {None}
                    current_hashes = new_fingerprints[key] - {None}
                    changed_rule = (len(previous_hashes) > 1 or len(current_hashes) > 1
                                    or bool(previous_hashes and current_hashes and previous_hashes != current_hashes))
                    status = UNVERIFIABLE if changed_rule else PERSISTENT
                elif (old_complete and new_complete and _covers(new, old)
                      and None not in old_fingerprints[key] and len(old_fingerprints[key]) == 1
                      and key in verified):
                    status = CORRECTED
            count_key = key if key is not None else ('unidentified-old', index)
            if count_key not in counted:
                summary[status_columns[status]] += 1
                counted.add(count_key)
            details.append(_detail(row, status, 'Historique', old_counts.get(key, 1), new_counts.get(key, 0)))
        for index, row in enumerate(now):
            key = violation_key(row)
            if key is not None and key in old_counts:
                continue
            status = UNVERIFIABLE
            if compatible and key is not None and old_complete:
                status = NEW if _covers(old, new) else NEW_FILTER
            elif compatible and key is not None and identifier not in old_targets:
                status = NEW_FILTER
            count_key = key if key is not None else ('unidentified-new', index)
            if count_key not in counted:
                summary[status_columns[status]] += 1
                counted.add(count_key)
            details.append(_detail(row, status, 'Actuelle', 0, new_counts.get(key, 1)))
        if not old_complete or not new_complete:
            warnings.append(f"{target['name'] or identifier} : au moins une extraction est incomplète ou absente.")
        summaries.append(summary)
    if any(row['targetID'] is None for snapshot in (old, new) for row in snapshot['violations']):
        warnings.append("Des lignes sans cible vérifiable sont exclues du comparatif géographique.")
    return {'summary': pd.DataFrame(summaries, columns=SUMMARY_COLUMNS),
            'details': pd.DataFrame(details, columns=DETAIL_COLUMNS),
            'warnings': list(dict.fromkeys(warnings))}

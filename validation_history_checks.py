"""Vérifications de métadonnées, en lecture seule, avant de qualifier une absence.

L'absence doit déjà être observée dans une nouvelle analyse complète du même
intervalle. Ces contrôles ne rejouent jamais une analyse et ne modifient pas DHIS2.
Un accès, un rattachement ou une règle invérifiable laisse la ligne non vérifiée.
"""

from collections import defaultdict
from datetime import date
import re

from validation_history import rule_fingerprint, violation_key


_UID = re.compile(r"[A-Za-z][A-Za-z0-9]{10}\Z")
_HASH = re.compile(r"[0-9a-f]{64}\Z")
_BATCH_SIZE = 60

_CONTEXT_WARNING = (
    "Absences non vérifiées : l'instance, le compte, le groupe ou l'intervalle "
    "de comparaison ne correspond pas à l'analyse enregistrée."
)
_SNAPSHOT_WARNING = "Absences non vérifiées : l'analyse enregistrée ou le nouveau résultat est incomplet."
_TARGET_WARNING = "Certaines absences restent non vérifiées : leur cible est incomplète ou non autorisée."
_UNIT_WARNING = (
    "Certaines absences restent non vérifiées : l'accès aux unités ou leur "
    "rattachement actuel n'a pas pu être confirmé."
)
_RULE_WARNING = (
    "Certaines absences restent non vérifiées : les règles sont modifiées "
    "ou leurs métadonnées ne sont pas accessibles."
)
_GROUP_WARNING = (
    "Certaines absences restent non vérifiées : l'appartenance actuelle "
    "des règles au groupe n'a pas pu être confirmée."
)


def _uid(value):
    return value if isinstance(value, str) and _UID.fullmatch(value) else None


def _hash(value):
    return value if isinstance(value, str) and _HASH.fullmatch(value) else None


def _dates(filters):
    try:
        first, last = filters["startDate"], filters["endDate"]
        # fromisoformat also accepts compact dates; snapshots use ISO dates only.
        if not isinstance(first, str) or not isinstance(last, str):
            return None
        if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", first) or not re.fullmatch(r"\d{4}-\d{2}-\d{2}", last):
            return None
        first, last = date.fromisoformat(first), date.fromisoformat(last)
        return (first, last) if first <= last else None
    except (KeyError, TypeError, ValueError):
        return None


def _compatible(previous, current):
    if not isinstance(previous, dict) or not isinstance(current, dict):
        return False
    for field in ("sourceID", "accountID"):
        if not _hash(previous.get(field)) or previous.get(field) != current.get(field):
            return False
    old_filters, new_filters = previous.get("filters"), current.get("filters")
    if not isinstance(old_filters, dict) or not isinstance(new_filters, dict):
        return False
    old_dates, new_dates = _dates(old_filters), _dates(new_filters)
    if not old_dates or not new_dates:
        return False
    if not (new_dates[0] <= old_dates[0] <= old_dates[1] <= new_dates[1]):
        return False
    old_id, new_id = old_filters.get("groupID"), new_filters.get("groupID")
    if old_id != new_id or (old_id is not None and not _uid(old_id)):
        return False
    return True


def _complete_targets(snapshot):
    targets = snapshot.get("targets")
    if not isinstance(targets, list):
        return set()
    occurrences = defaultdict(list)
    for target in targets:
        if isinstance(target, dict) and _uid(target.get("ID")):
            occurrences[target["ID"]].append(target.get("status"))
    # Duplicate declarations are ambiguous, even when their current status agrees.
    return {identifier for identifier, states in occurrences.items() if states == ["complete"]}


def _metadata(client, endpoint, identifiers, fields):
    """Rend uniquement les objets demandés et non ambigus, par lots bornés."""
    found, incomplete = {}, False
    identifiers = sorted(identifiers)
    for offset in range(0, len(identifiers), _BATCH_SIZE):
        batch = identifiers[offset:offset + _BATCH_SIZE]
        params = {"fields": fields, "filter": f"id:in:[{','.join(batch)}]", "paging": "false"}
        try:
            payload = client.get(endpoint, params=params)
        except Exception:
            incomplete = True
            continue
        records = payload.get(endpoint) if isinstance(payload, dict) else None
        if not isinstance(records, list) or any(not isinstance(row, dict) for row in records):
            incomplete = True
            continue
        pager = payload.get("pager")
        if isinstance(pager, dict):
            try:
                if int(pager.get("pageCount", 1)) > 1 or int(pager.get("total", len(records))) > len(records):
                    incomplete = True
                    continue
            except (TypeError, ValueError, OverflowError):
                incomplete = True
                continue
        requested, by_id = set(batch), defaultdict(list)
        for row in records:
            identifier = _uid(row.get("id"))
            if identifier not in requested:
                incomplete = True
                continue
            by_id[identifier].append(row)
        for identifier in batch:
            matches = by_id.get(identifier, [])
            if len(matches) == 1:
                found[identifier] = matches[0]
            else:
                incomplete = True
    return found, incomplete


def _group_rules(client, group_id):
    try:
        payload = client.get(
            f"validationRuleGroups/{group_id}", params={"fields": "id,validationRules[id]"},
        )
    except Exception:
        return None
    if not isinstance(payload, dict) or payload.get("id") != group_id:
        return None
    rows = payload.get("validationRules")
    if not isinstance(rows, list) or any(not isinstance(row, dict) or not _uid(row.get("id")) for row in rows):
        return None
    return {row["id"] for row in rows}


def verify_history_absences(client, previous_snapshot, current_snapshot, allowed_target_ids):
    """Retourne ``(clés vérifiées, avertissements)`` sans aucun appel d'écriture.

    Une clé n'est rendue que si toutes ses occurrences anciennes ont une même
    cible complète et une même empreinte de règle, si elle est absente du nouveau
    résultat couvrant les anciennes dates, et si les métadonnées actuelles
    confirment encore l'accès, le rattachement et la règle. Les messages restent
    fixes : aucun corps HTTP, identifiant de compte ou secret n'est reproduit.
    """
    warnings = []

    def warn(message):
        if message not in warnings:
            warnings.append(message)

    if not _compatible(previous_snapshot, current_snapshot):
        return set(), [_CONTEXT_WARNING]
    previous_rows = previous_snapshot.get("violations")
    current_rows = current_snapshot.get("violations")
    if not isinstance(previous_rows, list) or not isinstance(current_rows, list):
        return set(), [_SNAPSHOT_WARNING]
    if any(not isinstance(row, dict) or violation_key(row) is None for row in current_rows):
        return set(), [_SNAPSHOT_WARNING]
    current_keys = {violation_key(row) for row in current_rows}
    old_by_key = defaultdict(list)
    for row in previous_rows:
        key = violation_key(row) if isinstance(row, dict) else None
        if key is None:
            warn(_SNAPSHOT_WARNING)
        elif key not in current_keys:
            old_by_key[key].append(row)
    if not old_by_key:
        return set(), warnings

    if isinstance(allowed_target_ids, str):
        allowed_target_ids = [allowed_target_ids]
    try:
        allowed = {identifier for identifier in allowed_target_ids if _uid(identifier)}
    except TypeError:
        allowed = set()
    complete = _complete_targets(previous_snapshot) & _complete_targets(current_snapshot) & allowed
    candidates = {}
    for key, rows in old_by_key.items():
        target_ids = {_uid(row.get("targetID")) for row in rows}
        fingerprints = {_hash(row.get("ruleFingerprint")) for row in rows}
        if len(target_ids) != 1 or not target_ids.issubset(complete):
            warn(_TARGET_WARNING)
            continue
        if len(fingerprints) != 1 or not all(_hash(value) for value in fingerprints):
            warn(_RULE_WARNING)
            continue
        if any(not _uid(row.get(field)) for row in rows for field in ("organisationUnitID", "validationRuleID")):
            warn(_SNAPSHOT_WARNING)
            continue
        candidates[key] = rows[0]
    if not candidates:
        return set(), warnings

    group_id = previous_snapshot["filters"].get("groupID")
    if group_id:
        members = _group_rules(client, group_id)
        if members is None:
            warn(_GROUP_WARNING)
            return set(), warnings
        grouped = {key: row for key, row in candidates.items() if row["validationRuleID"] in members}
        if len(grouped) != len(candidates):
            warn(_GROUP_WARNING)
        candidates = grouped
    if not candidates:
        return set(), warnings

    units, units_incomplete = _metadata(
        client, "organisationUnits", {row["organisationUnitID"] for row in candidates.values()}, "id,path",
    )
    rules, rules_incomplete = _metadata(
        client, "validationRules", {row["validationRuleID"] for row in candidates.values()},
        "id,operator,leftSide[expression],rightSide[expression]",
    )
    if units_incomplete:
        warn(_UNIT_WARNING)
    if rules_incomplete:
        warn(_RULE_WARNING)
    verified = set()
    for key, row in candidates.items():
        unit = units.get(row["organisationUnitID"], {})
        path = unit.get("path")
        segments = path.split("/") if isinstance(path, str) else []
        if not segments or segments[-1] != row["organisationUnitID"] or row["targetID"] not in segments:
            warn(_UNIT_WARNING)
            continue
        fingerprint = rule_fingerprint(rules.get(row["validationRuleID"], {}))
        if fingerprint is None or fingerprint != row["ruleFingerprint"]:
            warn(_RULE_WARNING)
            continue
        verified.add(key)
    return verified, warnings

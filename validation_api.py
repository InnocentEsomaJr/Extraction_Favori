"""Lecture des violations calculées, sans persistance ni notification DHIS2.

Contrat officiel GET /dataAnalysis/validationRules :
https://docs.dhis2.org/en/develop/using-the-api/dhis-core-version-242/data-validation.html#validation-rule-analysis

Ce point d'entrée possède une limite de résultats, sans pagination documentée.
Une réponse atteignant la limite reste consultable mais son décompte est incomplet.
"""

import calendar
from datetime import date, timedelta
from decimal import Decimal, InvalidOperation
import re

import pandas as pd


def _identifier(value):
    if isinstance(value, dict):
        value = value.get("id") or value.get("uid") or value.get("iso")
    if value is None:
        return ""
    text = str(value).strip()
    return "" if text.lower() in {"", "none", "null", "nan", "<na>"} else text


def _month_dates(period):
    if not re.fullmatch(r"\d{6}", period):
        raise ValueError("Une période mensuelle AAAAMM est nécessaire.")
    year, month = int(period[:4]), int(period[4:])
    return date(year, month, 1), date(year, month, calendar.monthrange(year, month)[1])


def _period_dates(period):
    """Intervalles ISO usuels, pour filtrer aussi les règles d'autres périodicités."""
    if re.fullmatch(r"\d{6}", period):
        return _month_dates(period)
    if re.fullmatch(r"\d{8}", period):
        day = date(int(period[:4]), int(period[4:6]), int(period[6:]))
        return day, day
    if re.fullmatch(r"\d{4}", period):
        return date(int(period), 1, 1), date(int(period), 12, 31)
    match = re.fullmatch(r"(\d{4})Q([1-4])", period)
    if match:
        year, quarter = int(match[1]), int(match[2])
        end_month = quarter * 3
        return date(year, end_month - 2, 1), date(year, end_month, calendar.monthrange(year, end_month)[1])
    match = re.fullmatch(r"(\d{4})W(\d{1,2})", period)
    if match:
        start = date.fromisocalendar(int(match[1]), int(match[2]), 1)
        return start, start + timedelta(days=6)
    raise ValueError("La période de la violation n'est pas vérifiable.")


def _value(record, flat, nested):
    flat_value = _identifier(record.get(flat))
    if flat_value:
        return flat_value
    nested_value = record.get(nested)
    if nested == "period" and isinstance(nested_value, dict):
        return (
            _identifier(nested_value.get("iso"))
            or _identifier(nested_value.get("isoDate"))
            or _identifier(nested_value)
        )
    return _identifier(nested_value)


def _scope_matches(record, requested_ou):
    own_id = _value(record, "organisationUnitId", "organisationUnit")
    if not own_id:
        return None
    if own_id == requested_ou:
        return True
    nested = record.get("organisationUnit")
    nested_path = nested.get("path", "") if isinstance(nested, dict) else ""
    path = str(record.get("organisationUnitPath") or nested_path).strip()
    if not path:
        return None
    segments = [segment for segment in path.split("/") if segment]
    if not segments or segments[-1] != own_id:
        return None
    return requested_ou in segments


def _with_legacy_aliases(record):
    result = dict(record)
    for flat_id, flat_name, nested in [
        ("organisationUnitId", "organisationUnitDisplayName", "organisationUnit"),
        ("validationRuleId", "validationRuleDescription", "validationRule"),
        ("periodId", "periodDisplayName", "period"),
        ("attributeOptionComboId", "attributeOptionComboDisplayName", "attributeOptionCombo"),
    ]:
        identifier = _value(record, flat_id, nested)
        if identifier:
            result[nested] = dict(result[nested]) if isinstance(result.get(nested), dict) else {}
            result[nested]["id"] = identifier
            if record.get(flat_name):
                result[nested]["displayName"] = record[flat_name]
    for flat, legacy in [("leftSideValue", "leftsideValue"), ("rightSideValue", "rightsideValue")]:
        if flat in result and legacy not in result:
            result[legacy] = result[flat]
    return result


def _observation_value(record, flat, legacy):
    """Clé numérique stable : zéro reste distinct d'une valeur absente."""
    value = record[flat] if flat in record else record.get(legacy)
    if value is None:
        return None
    text = str(value).strip()
    if text.lower() in {"", "none", "null", "nan", "<na>"}:
        return None
    try:
        number = Decimal(text)
        return number.normalize() if number.is_finite() else text
    except InvalidOperation:
        return text


def fetch_validation_analysis(dhis, ou_id, period_list, group_id=None, limit=500):
    """Calcule les violations du périmètre et de ses descendants, par mois.

    Retourne les lignes disponibles avec ``failed_periods`` et
    ``truncated_periods`` dans ``DataFrame.attrs``. Toute période incomplète
    est signalée, même lorsqu'une partie de ses lignes reste affichable.
    """
    if isinstance(limit, bool) or not isinstance(limit, int) or limit < 1:
        raise ValueError("La limite de résultats doit être un entier strictement positif.")
    requested_ou = _identifier(ou_id)
    periods = list(dict.fromkeys(str(value).strip() for value in (
        [period_list] if isinstance(period_list, str) else (period_list or [])
    )))
    failed, truncated, errors, results, seen = [], [], {}, [], set()

    def incomplete(period, reason):
        if period not in failed:
            failed.append(period)
        errors.setdefault(period, []).append(reason)

    for period in periods:
        try:
            start, end = _month_dates(period)
        except (ValueError, OverflowError):
            incomplete(period, "Période mensuelle invalide")
            continue
        if not requested_ou:
            incomplete(period, "Identifiant de l'unité indisponible")
            continue
        params = {
            "ou": requested_ou, "startDate": start.isoformat(), "endDate": end.isoformat(),
            "persist": "false", "notification": "false", "maxResults": limit,
        }
        if _identifier(group_id):
            params["vrg"] = _identifier(group_id)
        try:
            payload = dhis.get("dataAnalysis/validationRules", params=params)
            if isinstance(payload, list):
                records = payload
            elif isinstance(payload, dict):
                records = payload.get("validationResults", payload.get("validationRuleViolations"))
            else:
                records = None
            if not isinstance(records, list):
                raise ValueError("Réponse invalide")
        except Exception:
            incomplete(period, "Analyse DHIS2 indisponible ou réponse invalide")
            continue
        if len(records) >= limit:
            truncated.append(period)
            incomplete(period, "Limite des résultats atteinte ; décompte incomplet")
        for record in records:
            if not isinstance(record, dict):
                incomplete(period, "Une violation retournée est invalide")
                continue
            matches = _scope_matches(record, requested_ou)
            if matches is None:
                incomplete(period, "Le périmètre d'une violation n'est pas vérifiable")
                continue
            if not matches:
                continue
            record_period = _value(record, "periodId", "period")
            try:
                record_start, record_end = _period_dates(record_period)
            except (ValueError, OverflowError):
                incomplete(period, "La période d'une violation n'est pas vérifiable")
                continue
            if record_start > end or record_end < start:
                continue
            rule = _value(record, "validationRuleId", "validationRule")
            if not rule:
                incomplete(period, "L'identifiant d'une règle est indisponible")
                continue
            normalized = _with_legacy_aliases(record)
            key = (
                _value(record, "organisationUnitId", "organisationUnit"), rule, record_period,
                _value(record, "attributeOptionComboId", "attributeOptionCombo"),
                _identifier(record.get("categoryOptionCombo")), record.get("dayInPeriod"),
                _observation_value(record, "leftSideValue", "leftsideValue"),
                _observation_value(record, "rightSideValue", "rightsideValue"),
            )
            if key not in seen:
                seen.add(key)
                results.append(normalized)
    result = pd.DataFrame(results)
    result.attrs.update(
        source="calculated", failed_periods=failed, truncated_periods=truncated,
        errors=errors, scope_ou_id=requested_ou, max_results=limit,
    )
    return result

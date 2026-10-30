"""Analyse des violations avec le même contrat que l'application Data Quality.

Source : dhis2/data-quality-app, ValidationRulesAnalysis.jsx (POST JSON unique).
Le repli GET est réservé aux anciennes versions qui refusent cette méthode.
Les périodes et les lignes retournées sont choisies par DHIS2, sans redécoupage
ni déduplication locale. Une limite atteinte laisse le décompte incomplet.
"""

import calendar
from datetime import date
import re

import pandas as pd

from validation_filters import date_range_months, normalize_date


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
    nested = record.get("organisationUnit")
    nested_path = nested.get("path", "") if isinstance(nested, dict) else ""
    path = str(record.get("organisationUnitPath") or nested_path).strip()
    if not path:
        # Certaines versions ne renvoient pas de chemin. L'API a déjà appliqué
        # le parent et ses descendants ; cette absence n'est pas une erreur.
        return True
    segments = [segment for segment in path.split("/") if segment]
    if not segments or segments[-1] != own_id:
        return False
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


def _http_status(error):
    response = getattr(error, "response", None)
    value = getattr(response, "status_code", None) if response is not None else None
    if value is None:
        value = getattr(error, "status_code", None)
    try:
        return int(value) if value is not None else None
    except (TypeError, ValueError):
        return None


def _request_error(error, method, endpoint):
    """Diagnostic utile sans reprendre un corps de réponse ou une URL sensible."""
    status = _http_status(error)
    prefix = f"{method} /api/{endpoint}"
    if status is not None:
        reasons = {
            401: "authentification refusée", 403: "autorisation refusée",
            404: "ressource introuvable", 405: "méthode non prise en charge",
            501: "méthode non implémentée",
        }
        return f"{prefix} — HTTP {status} : {reasons.get(status, 'échec de l’analyse DHIS2')}"
    if isinstance(error, TimeoutError) or "timeout" in type(error).__name__.lower():
        return f"{prefix} — délai d’attente dépassé"
    if isinstance(error, ValueError):
        return f"{prefix} — réponse non JSON ou invalide (vérifiez la connexion et l'accès à l'API)"
    return f"{prefix} — connexion ou analyse DHIS2 indisponible"


def fetch_validation_analysis(
    dhis, ou_id, period_list, group_id=None, limit=None, *, start_date=None, end_date=None
):
    """Une analyse DHIS2 sur le parent et l'intervalle complet sélectionnés.

    Sans dates explicites, les mois fournis déterminent les deux bornes.
    Comme l'application officielle, ``maxResults`` est omis par défaut ;
    500 lignes déclenchent alors un avertissement prudent de limite atteinte.
    Les résultats conservent l'ordre, les périodes et les doublons du serveur.
    """
    if limit is not None and (isinstance(limit, bool) or not isinstance(limit, int) or limit < 1):
        raise ValueError("La limite de résultats doit être un entier strictement positif.")
    effective_limit = 500 if limit is None else limit
    exact_start = exact_end = None
    if start_date is not None or end_date is not None:
        if start_date is None or end_date is None:
            raise ValueError("Les dates de début et de fin doivent être fournies ensemble.")
        exact_start, exact_end = normalize_date(start_date), normalize_date(end_date)
        if exact_start > exact_end:
            raise ValueError("La date de début doit précéder ou égaler la date de fin.")
    requested_ou = _identifier(ou_id)
    periods = list(dict.fromkeys(str(value).strip() for value in (
        [period_list] if isinstance(period_list, str) else (period_list or [])
    )))
    failed, truncated, errors, results = [], [], {}, []
    method = "POST"
    endpoint = "dataAnalysis/validationRules"
    params = {}

    def incomplete(period, reason):
        if period not in failed:
            failed.append(period)
        errors.setdefault(period, []).append(reason)

    if exact_start is not None:
        periods = date_range_months(exact_start, exact_end)
    else:
        valid_bounds = []
        for period in periods:
            try:
                valid_bounds.append(_month_dates(period))
            except (ValueError, OverflowError):
                incomplete(period, "Période mensuelle invalide")
        if valid_bounds:
            exact_start = min(first for first, _ in valid_bounds)
            exact_end = max(last for _, last in valid_bounds)
            periods = date_range_months(exact_start, exact_end)

    def incomplete_interval(reason):
        for period in periods:
            incomplete(period, reason)

    if not requested_ou:
        incomplete_interval("Identifiant de l'unité indisponible")
    elif exact_start is not None and exact_end is not None:
        params = {
            "ou": requested_ou, "startDate": exact_start.isoformat(), "endDate": exact_end.isoformat(),
            "persist": False, "notification": False,
        }
        if limit is not None:
            params["maxResults"] = limit
        if _identifier(group_id):
            params["vrg"] = _identifier(group_id)
        try:
            try:
                payload = dhis.post(endpoint, data=params)
            except Exception as error:
                if _http_status(error) not in {404, 405, 501}:
                    raise
                method = "GET"
                params = {**params, "persist": "false", "notification": "false"}
                payload = dhis.get(endpoint, params=params)
            if isinstance(payload, list):
                records = payload
            elif isinstance(payload, dict):
                records = payload.get("validationResults", payload.get("validationRuleViolations"))
            else:
                records = None
            if not isinstance(records, list):
                incomplete_interval(f"{method} /api/{endpoint} — réponse DHIS2 invalide (liste de violations attendue)")
                records = []
        except Exception as error:
            incomplete_interval(_request_error(error, method, endpoint))
            records = []
        if len(records) >= effective_limit:
            truncated.extend(periods)
            incomplete_interval("Limite des résultats atteinte ; décompte incomplet")
        for record in records:
            if not isinstance(record, dict):
                incomplete_interval("Une violation retournée est invalide")
                continue
            matches = _scope_matches(record, requested_ou)
            if matches is None:
                incomplete_interval("L'identifiant de l'unité d'une violation est indisponible")
                continue
            if not matches:
                incomplete_interval("DHIS2 a retourné une violation dont le chemin ne correspond pas au parent sélectionné")
                continue
            record_period = _value(record, "periodId", "period")
            if not record_period:
                incomplete_interval("L'identifiant de la période d'une violation est indisponible")
                continue
            rule = _value(record, "validationRuleId", "validationRule")
            if not rule:
                incomplete_interval("L'identifiant d'une règle est indisponible")
                continue
            results.append(_with_legacy_aliases(record))
    result = pd.DataFrame(results)
    result.attrs.update(
        source="calculated", failed_periods=failed, truncated_periods=truncated,
        errors=errors, scope_ou_id=requested_ou, max_results=effective_limit,
        request_method=method, request_params=params,
        start_date=exact_start.isoformat() if exact_start is not None else None,
        end_date=exact_end.isoformat() if exact_end is not None else None,
    )
    return result

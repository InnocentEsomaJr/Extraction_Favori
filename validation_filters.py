"""Bornes calendaires des filtres de validation, indépendantes de Streamlit."""

import calendar
from datetime import date, datetime, timedelta
import re


def normalize_date(value):
    """Retourne une date civile depuis date, datetime ou une chaîne ISO."""
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, str):
        text = value.strip()
        try:
            return date.fromisoformat(text)
        except ValueError:
            try:
                return datetime.fromisoformat(text).date()
            except ValueError:
                pass
    raise ValueError("Une date valide au format ISO est nécessaire.")


def date_range_months(start, end):
    """Liste les mois AAAAMM intersectant l'intervalle, bornes incluses."""
    first, last = normalize_date(start), normalize_date(end)
    if first > last:
        raise ValueError("La date de début doit précéder ou égaler la date de fin.")
    first_month = first.year * 12 + first.month - 1
    last_month = last.year * 12 + last.month - 1
    return [f"{number // 12:04d}{number % 12 + 1:02d}" for number in range(first_month, last_month + 1)]


def reference_month(start):
    """Mois précédant celui de la date de début sélectionnée."""
    first = normalize_date(start).replace(day=1)
    try:
        previous = first - timedelta(days=1)
    except OverflowError as error:
        raise ValueError("Le mois de référence précède les dates prises en charge.") from error
    return f"{previous.year:04d}{previous.month:02d}"


def monthly_bounds(period_list):
    """Première et dernière date des mois fournis, quel que soit leur ordre."""
    try:
        periods = [period_list] if isinstance(period_list, str) else list(period_list or [])
    except TypeError as error:
        raise ValueError("Une liste de périodes mensuelles AAAAMM est nécessaire.") from error
    if not periods:
        raise ValueError("Au moins une période mensuelle est nécessaire.")
    bounds = []
    for period in periods:
        value = str(period).strip()
        if not re.fullmatch(r"\d{6}", value):
            raise ValueError("Une période mensuelle AAAAMM est nécessaire.")
        try:
            year, month = int(value[:4]), int(value[4:])
            bounds.append((date(year, month, 1), date(year, month, calendar.monthrange(year, month)[1])))
        except (ValueError, OverflowError) as error:
            raise ValueError("Une période mensuelle AAAAMM est nécessaire.") from error
    return min(first for first, _ in bounds), max(last for _, last in bounds)

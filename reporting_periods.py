"""Périodes calendaires du rapport, indépendantes de Streamlit et de DHIS2.

Identifiants natifs : AAAAMM, AAAAQn, AAAASn, AAAA et AAAAWn.
Voir https://docs.dhis2.org/archive/en/2.30/developer/html/webapi_date_perid_format.html
Les semaines ISO commencent le lundi et se terminent le dimanche, même
quand leurs bornes appartiennent à deux années civiles différentes.
"""

import calendar
from dataclasses import dataclass
from datetime import date
import re

from validation_filters import date_range_months


PERIODICITIES = ("Mensuel", "Trimestriel", "Semestriel", "Annuel", "Hebdomadaire")
MONTH_NAMES = (
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
)


@dataclass(frozen=True)
class ReportingPeriod:
    """Une option du filtre avec son identifiant DHIS2 et ses bornes incluses."""

    id: str
    label: str
    start_date: date
    end_date: date


def _year_number(year):
    if isinstance(year, bool) or not re.fullmatch(r"[0-9]{1,4}", str(year).strip()):
        raise ValueError("Une année civile valide est nécessaire.")
    number = int(year)
    if not 1 <= number <= 9999:
        raise ValueError("Une année civile valide est nécessaire.")
    return number


def _month_period(year, first_month, last_month, period_id, label):
    return ReportingPeriod(
        period_id,
        label,
        date(year, first_month, 1),
        date(year, last_month, calendar.monthrange(year, last_month)[1]),
    )


def period_options(periodicity, year):
    """Toutes les périodes de l'année, ordonnées et sans exclusion de dates futures.

    L'année hebdomadaire est l'année ISO : sa première semaine peut commencer
    en décembre de l'année précédente et sa dernière finir en janvier suivant.
    """
    if periodicity not in PERIODICITIES:
        raise ValueError("Choisissez une périodicité valide.")
    year = _year_number(year)
    year_id = f"{year:04d}"
    if periodicity == "Hebdomadaire":
        options = []
        week_count = date(year, 12, 28).isocalendar().week
        try:
            for week in range(1, week_count + 1):
                first = date.fromisocalendar(year, week, 1)
                last = date.fromisocalendar(year, week, 7)
                options.append(ReportingPeriod(
                    f"{year_id}W{week}",
                    f"Semaine {week} — {first:%d/%m/%Y} au {last:%d/%m/%Y}",
                    first,
                    last,
                ))
        except (ValueError, OverflowError) as error:
            raise ValueError("Les bornes de cette année ISO dépassent les dates prises en charge.") from error
        return tuple(options)

    month_count = {"Mensuel": 1, "Trimestriel": 3, "Semestriel": 6, "Annuel": 12}[periodicity]
    options = []
    for index, first_month in enumerate(range(1, 13, month_count), start=1):
        last_month = first_month + month_count - 1
        if periodicity == "Mensuel":
            period_id, label = f"{year_id}{first_month:02d}", MONTH_NAMES[first_month - 1]
        elif periodicity == "Trimestriel":
            period_id = f"{year_id}Q{index}"
            label = f"T{index} — {MONTH_NAMES[first_month - 1]} à {MONTH_NAMES[last_month - 1]}"
        elif periodicity == "Semestriel":
            period_id = f"{year_id}S{index}"
            label = f"S{index} — {MONTH_NAMES[first_month - 1]} à {MONTH_NAMES[last_month - 1]}"
        else:
            period_id, label = year_id, year_id
        options.append(_month_period(year, first_month, last_month, period_id, label))
    return tuple(options)


def select_period_range(periodicity, year, start_id, end_id):
    """Sélectionne les périodes de début à fin incluses et refuse l'ordre inversé."""
    options = period_options(periodicity, year)
    positions = {period.id: index for index, period in enumerate(options)}
    if start_id not in positions or end_id not in positions:
        raise ValueError("Choisissez des périodes appartenant à l'année et à la périodicité sélectionnées.")
    first, last = positions[start_id], positions[end_id]
    if first > last:
        raise ValueError("La période de début doit précéder ou égaler la période de fin.")
    return options[first:last + 1]


def _selected_periods(periods):
    try:
        selected = tuple(periods)
    except TypeError as error:
        raise ValueError("Au moins une période est nécessaire.") from error
    if not selected or any(not isinstance(period, ReportingPeriod) for period in selected):
        raise ValueError("Au moins une période valide est nécessaire.")
    return selected


def selection_bounds(periods):
    """Première et dernière date exactes des périodes sélectionnées."""
    selected = _selected_periods(periods)
    return min(period.start_date for period in selected), max(period.end_date for period in selected)


def coverage_months(periods):
    """Mois AAAAMM intersectant les périodes, dédoublonnés et ordonnés.

    Ces mois servent aux filtres mensuels existants ; pour une semaine, leurs
    bornes mensuelles ne remplacent pas les dates exactes de selection_bounds.
    """
    selected = _selected_periods(periods)
    months = set()
    for period in selected:
        months.update(date_range_months(period.start_date, period.end_date))
    return sorted(months)


def analytics_periods(periods):
    """Identifiants natifs sélectionnés, dédoublonnés dans l'ordre du filtre.

    DHIS2 effectue l'agrégation à la périodicité demandée. Les mois couverts
    restent disponibles séparément pour initialiser les dates de validation.
    """
    selected = _selected_periods(periods)
    return list(dict.fromkeys(period.id for period in selected))

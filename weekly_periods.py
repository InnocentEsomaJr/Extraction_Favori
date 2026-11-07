"""Semaines ISO de surveillance rattachées aux mois évalués.

Une semaine est repérée par son dimanche. Avec le report, le cinquième
dimanche d'un mois appartient au mois suivant, où il s'ajoute aux quatre
premières semaines. Un mois évalué peut donc contenir cinq semaines ; aucune
semaine n'est supprimée pour imposer artificiellement une limite de quatre.
"""

import calendar
from datetime import date, timedelta
import re


SCHEDULE_COLUMNS = ['Mois', 'Semaine', 'Début', 'Fin', 'Reportée']


def _selected_months(months):
    if isinstance(months, str):
        values = months.split(';')
    else:
        try:
            values = list(months or [])
        except TypeError as error:
            raise ValueError('Une liste de périodes mensuelles AAAAMM est nécessaire.') from error
    if not values:
        raise ValueError('Au moins une période mensuelle AAAAMM est nécessaire.')
    selected = []
    for value in values:
        period = str(value).strip()
        if not re.fullmatch(r'[0-9]{6}', period):
            raise ValueError('Une période mensuelle AAAAMM est nécessaire.')
        try:
            first = date(int(period[:4]), int(period[4:]), 1)
        except (ValueError, OverflowError) as error:
            raise ValueError('Une période mensuelle AAAAMM valide est nécessaire.') from error
        selected.append((period, first))
    return sorted(dict(selected).items())


def _sundays(first):
    first_day = 1 + (6 - first.weekday()) % 7
    last_day = calendar.monthrange(first.year, first.month)[1]
    return [date(first.year, first.month, day) for day in range(first_day, last_day + 1, 7)]


def weekly_period_schedule(months, carry_fifth=True):
    """Retourne les semaines et leurs bornes ISO, sans doublon entre mois.

    La sélection est indépendante de l'ordre et des répétitions de ``months``.
    Par défaut, chaque mois reçoit ses quatre premiers dimanches ainsi que le
    cinquième dimanche du mois précédent, s'il existe. Son propre cinquième
    dimanche est réservé au mois suivant. En mode ``carry_fifth=False``, les
    quatre ou cinq dimanches sont tous rattachés à leur mois civil.

    ``Mois`` désigne le mois auquel la semaine est rattachée pour l'évaluation.
    ``Début`` et ``Fin`` sont les dates civiles réelles, y compris lors du report.
    Les identifiants DHIS2 sont de la forme ``2026W1`` sans zéro de remplissage.
    """
    if not isinstance(carry_fifth, bool):
        raise ValueError('Le report de la cinquième semaine doit être un booléen.')
    result, seen = [], set()
    for period, first in _selected_months(months):
        endings = _sundays(first)
        assigned = [(ending, False) for ending in (endings[:4] if carry_fifth else endings)]
        if carry_fifth and first > date.min:
            previous_endings = _sundays((first - timedelta(days=1)).replace(day=1))
            if len(previous_endings) == 5:
                assigned.insert(0, (previous_endings[-1], True))
        for ending, transferred in assigned:
            iso_year, iso_week, _ = ending.isocalendar()
            identifier = f'{iso_year:04d}W{iso_week}'
            if identifier in seen:
                continue
            seen.add(identifier)
            result.append({
                'Mois': period, 'Semaine': identifier,
                'Début': (ending - timedelta(days=6)).isoformat(),
                'Fin': ending.isoformat(), 'Reportée': transferred,
            })
    return sorted(result, key=lambda row: (row['Fin'], row['Mois']))


def weekly_periods_for_months(months, carry_fifth=True):
    """Identifiants uniques de surveillance dans l'ordre des dates réelles."""
    return tuple(row['Semaine'] for row in weekly_period_schedule(months, carry_fifth=carry_fifth))

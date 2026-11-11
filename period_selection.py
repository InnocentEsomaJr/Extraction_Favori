"""Périodes fixes et intervalles DHIS2, sans limite arbitraire de semaines."""

import calendar
from datetime import date, timedelta
import re
import unicodedata

from validation_filters import date_range_months, normalize_date


MONTH_NAMES = (
    'Janvier', 'Février', 'Mars', 'Avril', 'Mai', 'Juin',
    'Juillet', 'Août', 'Septembre', 'Octobre', 'Novembre', 'Décembre',
)
PERIOD_KINDS = ('Daily', 'Monthly', 'Quarterly', 'SixMonthly', 'Yearly', 'Weekly')


def _kind(value):
    text = unicodedata.normalize('NFKD', str(value)).casefold()
    text = ''.join(character for character in text if not unicodedata.combining(character))
    text = re.sub(r'[^a-z]', '', text)
    kinds = {
        'daily': 'Daily', 'journalier': 'Daily', 'journaliere': 'Daily', 'jour': 'Daily',
        'monthly': 'Monthly', 'mensuel': 'Monthly', 'mois': 'Monthly',
        'quarterly': 'Quarterly', 'trimestriel': 'Quarterly', 'trimestre': 'Quarterly',
        'sixmonthly': 'SixMonthly', 'semestriel': 'SixMonthly', 'semestre': 'SixMonthly',
        'yearly': 'Yearly', 'annuel': 'Yearly', 'annee': 'Yearly',
        'weekly': 'Weekly', 'hebdomadaire': 'Weekly', 'semaine': 'Weekly',
    }
    if text not in kinds:
        raise ValueError('Choisissez une périodicité journalière, mensuelle, trimestrielle, semestrielle, annuelle ou hebdomadaire.')
    return kinds[text]


def _year(value):
    text = str(value).strip()
    if isinstance(value, bool) or not re.fullmatch(r'[0-9]{1,4}', text):
        raise ValueError('Une année civile valide est nécessaire.')
    year = int(text)
    if not 1 <= year <= 9999:
        raise ValueError('Une année civile valide est nécessaire.')
    return year


def period_info(identifier):
    """Décrit une période fixe ; les bornes sont des dates civiles inclusives."""
    value = str(identifier).strip().upper()
    if re.fullmatch(r'[0-9]{8}', value):
        try:
            day = date(_year(value[:4]), int(value[4:6]), int(value[6:]))
        except (ValueError, OverflowError) as error:
            raise ValueError('Une période DHIS2 fixe valide est nécessaire.') from error
        return {
            'id': value, 'kind': 'Daily', 'label': f'Jour {day:%d/%m/%Y}',
            'start': day, 'end': day,
        }
    match = re.fullmatch(r'([0-9]{4})(?:([QS])([1-4])|W([0-9]{1,2})|([0-9]{2}))?', value)
    if not match:
        raise ValueError('Une période DHIS2 fixe valide est nécessaire.')
    year = _year(match.group(1))
    marker, index, week, month = match.group(2, 3, 4, 5)
    try:
        if week is not None:
            number = int(week)
            first = date.fromisocalendar(year, number, 1)
            last = first + timedelta(days=6)
            kind, value = 'Weekly', f'{year:04d}W{number}'
            label = f'Semaine {number} {year:04d} — {first.isoformat()} au {last.isoformat()}'
        else:
            if month is not None:
                first_month, last_month = int(month), int(month)
                kind, label = 'Monthly', f'{MONTH_NAMES[first_month - 1]} {year:04d}' if 1 <= first_month <= 12 else ''
            elif marker == 'Q':
                first_month, last_month = (int(index) - 1) * 3 + 1, int(index) * 3
                kind, label = 'Quarterly', f'Trimestre {index} {year:04d}'
            elif marker == 'S' and index in ('1', '2'):
                first_month, last_month = (int(index) - 1) * 6 + 1, int(index) * 6
                kind, label = 'SixMonthly', f'Semestre {index} {year:04d}'
            elif marker is not None:
                raise ValueError('Semestre invalide')
            else:
                first_month, last_month = 1, 12
                kind, label = 'Yearly', f'Année {year:04d}'
            first = date(year, first_month, 1)
            last = date(year, last_month, calendar.monthrange(year, last_month)[1])
    except (ValueError, OverflowError) as error:
        raise ValueError('Une période DHIS2 fixe valide est nécessaire.') from error
    return {'id': value, 'kind': kind, 'label': label, 'start': first, 'end': last}


def period_choices(kind, year):
    """Périodes d'une année ; l'année des semaines est leur année ISO."""
    kind, year = _kind(kind), _year(year)
    prefix = f'{year:04d}'
    if kind == 'Daily':
        return [
            period_info(f'{day.year:04d}{day.month:02d}{day.day:02d}')
            for day in (date(year, 1, 1) + timedelta(days=offset)
                        for offset in range(366 if calendar.isleap(year) else 365))
        ]
    identifiers = {
        'Monthly': [f'{prefix}{number:02d}' for number in range(1, 13)],
        'Quarterly': [f'{prefix}Q{number}' for number in range(1, 5)],
        'SixMonthly': [f'{prefix}S{number}' for number in range(1, 3)],
        'Yearly': [prefix],
        'Weekly': [f'{prefix}W{number}' for number in range(1, date(year, 12, 28).isocalendar().week + 1)],
    }
    return [period_info(identifier) for identifier in identifiers[kind]]


def canonical_period_ids(identifiers):
    """Valide, déduplique et trie les sélections sans élargir les périodes."""
    if isinstance(identifiers, str):
        values = identifiers.split(';')
    else:
        try:
            values = list(identifiers or [])
        except TypeError as error:
            raise ValueError('Une liste de périodes fixes est nécessaire.') from error
    if not values:
        raise ValueError('Sélectionnez au moins une période fixe.')
    records = {record['id']: record for record in map(period_info, values)}
    return tuple(sorted(records, key=lambda identifier: (records[identifier]['start'], records[identifier]['end'], identifier)))


def selection_bounds(identifiers):
    """Bornes inclusives de la sélection, y compris entre plusieurs années."""
    records = [period_info(identifier) for identifier in canonical_period_ids(identifiers)]
    return min(record['start'] for record in records), max(record['end'] for record in records)


def selection_months(identifiers):
    """Mois de reporting entiers, sans prorata des rapports mensuels.

    Une semaine est rattachée au mois de son dimanche : son lundi dans le
    mois précédent n'ajoute pas implicitement ce mois aux rapports mensuels.
    Les autres périodes couvrent leurs mois civils, sans combler les trous.
    """
    months = set()
    for identifier in canonical_period_ids(identifiers):
        record = period_info(identifier)
        if record['kind'] == 'Weekly':
            months.add(f"{record['end'].year:04d}{record['end'].month:02d}")
        else:
            months.update(date_range_months(record['start'], record['end']))
    return tuple(sorted(months))


def period_interval(kind, start_id, end_id):
    """Périodes fixes d'un intervalle inclusif, sans mélanger les périodicités."""
    kind = _kind(kind)
    first, last = period_info(start_id), period_info(end_id)
    if first['kind'] != kind or last['kind'] != kind:
        raise ValueError("Les deux bornes doivent avoir la périodicité choisie.")
    if first['start'] > last['start']:
        raise ValueError("La période de début doit précéder ou égaler la période de fin.")
    if kind == 'Daily':
        return [
            period_info(f'{day.year:04d}{day.month:02d}{day.day:02d}')
            for day in (first['start'] + timedelta(days=offset)
                        for offset in range((last['start'] - first['start']).days + 1))
        ]
    if kind == 'Weekly':
        records, current = [], first['start']
        while current <= last['start']:
            iso_year, iso_week, _ = current.isocalendar()
            records.append(period_info(f'{iso_year:04d}W{iso_week}'))
            if current == last['start']:
                break
            current += timedelta(days=7)
        return records
    return [
        record for year in range(first['start'].year, last['start'].year + 1)
        for record in period_choices(kind, year)
        if first['start'] <= record['start'] <= last['start']
    ]


def weekly_choices_for_range(start, end, boundary='ending'):
    """Semaines selon des bornes civiles inclusives, sans découper les semaines.

    ``ending`` retient les semaines dont le dimanche est dans [début, fin].
    ``contained`` exige également que leur lundi soit dans cet intervalle.
    Aucun prorata ni report de semaine n'est appliqué.
    """
    if boundary not in ('ending', 'contained'):
        raise ValueError('Choisissez une règle de bornes hebdomadaires valide.')
    first, last = normalize_date(start), normalize_date(end)
    if first > last:
        raise ValueError('La date de début doit précéder ou égaler la date de fin.')
    delta = (6 - first.weekday()) % 7
    if (last - first).days < delta:
        return []
    ending, records = first + timedelta(days=delta), []
    while ending <= last:
        iso_year, iso_week, _ = ending.isocalendar()
        record = period_info(f'{iso_year:04d}W{iso_week}')
        if boundary == 'ending' or record['start'] >= first:
            records.append(record)
        if (last - ending).days < 7:
            break
        ending += timedelta(days=7)
    return records

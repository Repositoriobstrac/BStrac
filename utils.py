import csv
import re
from datetime import timedelta
from typing import List, Dict

from driver_work_report_generator.constants import DATE_REGEX, DAYS_OF_THE_WEEK, TRAVEL_MESSAGES


def is_date(value):
    return re.search(DATE_REGEX, value)


def is_vehicle_plate(value):
    if value and not is_date(value) and not is_shift(value) and not is_day_of_the_week(value) and not is_field(value):
        return True
    return False


def is_shift(value):
    return value.lower() in ('diurno', 'noturno')


def is_day_of_the_week(value):
    return value.lower() in DAYS_OF_THE_WEEK


def is_field(value: str) -> bool:
    return 'agrupamento' in value.lower()


def is_trip(message, end):
    return message.lower() in TRAVEL_MESSAGES and len(end.split(' ')) == 1


def is_blank(row: List) -> bool:
    if not [i for i in row if i and i != '""']:
        return True
    return False


def is_new_table(row: List) -> bool:
    return row[0] and not any(row[1:]) and row[0] != '""'


def is_field_names(cache_: str, data_: Dict) -> bool:
    return cache_ in data_ and data_[cache_] == []


def is_night_shift(value):
    return value.lower() == 'noturno'


def convert_time_to_seconds(value: str) -> int:
    if 'dia' in value:
        mo = re.search(r'[-]?(\d+) dia[s] ([0-9:]+)', value)
        if mo:
            hours_from_days = int(mo.group(1)) * 24
            hours, minutes, seconds = map(lambda x: int(x), mo.group(2).split(':'))
            value = f'{hours+hours_from_days}:{minutes}:{seconds}'

    is_negative = value.startswith('-')
    if is_negative:
        value = value[1:]

    hours, minutes, seconds = map(lambda x: int(x), value.split(':'))
    total_seconds = (hours * 3600) + (minutes * 60) + seconds
    return total_seconds if not is_negative else -total_seconds


def convert_to_time(value: str):
    if not value:
        value = '0:0:0'
    elif 'dia' in value:
        value = format_to_day_min_sec(convert_time_to_seconds(value))
    h, m, s = map(lambda x: int(x), value.split(':'))

    if h < 0:
        value_in_time = -timedelta(hours=h*-1, minutes=m, seconds=s)
    else:
        value_in_time = timedelta(hours=h, minutes=m, seconds=s)

    return value_in_time


def get_number_hours(duration):
    hour = '00'
    remaining_duration = 0
    number_hours = duration // 3600

    if number_hours > 0:
        hour = number_hours
        remaining_duration = duration - (number_hours*3600)

    return hour, remaining_duration


def get_number_minutes(duration):
    minute = '00'
    number_minutes = duration // 60
    if number_minutes > 0:
        minute = number_minutes
        remaining_duration = duration - (number_minutes * 60)
    else:
        remaining_duration = duration
    return minute, remaining_duration


def get_data_from_csv_file(filename: str) -> Dict:
    with open(filename, newline='') as csvfile:
        data = {}
        cache = None
        last_field = None

        spam_reader = csv.reader(csvfile, delimiter=',', quotechar='|')
        for row in spam_reader:
            if is_blank(row):
                continue
            key = row[0]
            if is_new_table(row):
                last_field = table_name = key
                data[table_name] = []
            elif is_field_names(cache, data):
                fields = clean_fields(row)
                data[last_field].append(fields)
            else:
                data[last_field].append(clean_meal_time(row, last_field))
            cache = key
    return data


def clean_fields(fields: List) -> List:
    ungroup = 'Agrupamento Início'
    if fields[0] == 'Agrupamento Início':
        return ungroup.split(' ') + fields[1:]
    elif fields[-2] == 'Jornada Diária Motorista':
        return fields[:-2] + ['Jornada Diária', 'Motorista']
    return fields


def clean_meal_time(row: List, key: str) -> List:
    if key == 'Horário de Refeição':
        mo = re.search(r'(.*?) ([0-9:]+)', row[0])
        if mo:
            row = [mo.group(1), mo.group(2)] + row[1:]

        if 'Código' in row[3]:
            mo = re.search(r'(.*?\d+:\d+:\d+) (.*? - Código \d+)', row[3])
            if mo:
                row = row[:-2] + [mo.group(1), mo.group(1)]
    return row


def format_number_with_two_digits(value: str | int) -> str:
    value = str(value)
    return value if len(value) >= 2 else '0' + value


def format_to_day_min_sec(duration: int, full_hours: int = 0) -> str:
    is_negative = duration < 0
    if is_negative:
        duration = duration * -1

    hour, remaining_duration = get_number_hours(duration)

    if remaining_duration:
        minute, remaining_duration = get_number_minutes(remaining_duration)
        second = str(remaining_duration)

    elif hour != '00' and not remaining_duration:
        minute = '0'
        second = '0'

    else:
        minute, remaining_duration = get_number_minutes(duration)
        second = str(remaining_duration)

    hour = full_hours + int(hour)
    hour = format_number_with_two_digits(hour)
    minute = format_number_with_two_digits(minute)
    second = format_number_with_two_digits(second)

    return f'{hour}:{minute}:{second}' if not is_negative else f'-{hour}:{minute}:{second}'

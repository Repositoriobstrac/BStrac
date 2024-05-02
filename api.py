import datetime
from typing import List, Dict

from driver_work_report_generator.constants import DAYS_OF_THE_WEEK
from driver_work_report_generator.utils import (is_date, is_vehicle_plate, is_night_shift)


class Tabela:
    key = None

    def execute(self, data: Dict) -> List:
        return data[self.key]

    def _insert_date(self, data):
        COL_DATE = ['Data', 'Dia']
        new_data = []
        date = COL_DATE
        for row in data:
            agrouping = row[0]
            if date == COL_DATE:
                value = COL_DATE
                date = None
            elif is_date(agrouping):
                value = [agrouping, self._get_day(agrouping)]
                date = [agrouping, self._get_day(agrouping)]
            elif date:
                value = date
            else:
                value = ['', '']
            new_data.append(value + row)
        return new_data

    def _get_day(self, date):
        d, m, y = date.split('.')
        data = datetime.date(year=int(y), month=int(m), day=int(d))
        indice_da_semana = data.weekday()
        return DAYS_OF_THE_WEEK[indice_da_semana]

    def remove_duplicate_itens(self, data):
        new_data = []
        container = []
        for row in data:
            compare = row[3:]
            if compare not in container:
                new_data.append(row)
                container.append(compare)
        return new_data

    def filter_plate_itens(self, data):
        return list(filter(lambda x: is_vehicle_plate(x[2]), data))

    def filter_night_shift_itens(self, data):
        return list(filter(lambda x: is_night_shift(x[2]), data))

    def filter_driver(self, data, driver, index):
        return list(filter(lambda x: x[index] == driver, data))


class TabelaMensagens(Tabela):
    key = 'Mensagens'

    def get_dates(self, data):
        dates = []
        for item in data:
            if item[:2] not in dates:
                dates.append(item[:2])
        return dates


class TabelaJornadaDiaria(Tabela):
    key = 'Jornada Diária'


class TabelaHorarioRefeicao(Tabela):
    key = 'Horário de Refeição'

    @classmethod
    def clean(cls, row):
        count = len(row[0].strip().split(' '))
        if len(row) == 4 and count > 1:
            row = row[0].strip().split(' ') + row[1:]
        return row

    def execute(self, data: Dict) -> List:
        return [row for row in data[self.key] if len(row) == 5]


class TabelaCargaDecarga(Tabela):
    key = 'Carga e Descarga'


class TabelaParadaEspera(Tabela):
    key = 'Parada Espera'


class TabelaDescansoIntrajornada(Tabela):
    key = 'Descanso Intrajornada'


class TabelaViagens(Tabela):
    key = 'Viagens'

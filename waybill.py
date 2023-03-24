import pandas as pd
from docxtpl import DocxTemplate
import datetime
from pathlib import Path
import os
from settings import company_data

def get_df_from_excel_file(filename: str, sheetname='данные') -> pd.DataFrame:
    """Принимает имя файла в виде текста. Возвращает датафрейм."""
    return pd.read_excel(filename, na_values=['(пусто)'], sheet_name=sheetname)


def filter_df(data: dict, df_data: pd.DataFrame) -> pd.DataFrame:
    """Возвращает датафрейм по маске. Маска формируются на основе словаря."""
    if isinstance(data, dict):
        __df = pd.DataFrame(data)
        user_vin = __df.columns[0]
        user_start_date = transform_text_to_datetime(__df.iat[0, 0])
        user_end_date = transform_text_to_datetime(__df.iat[1, 0])
        
        # Смена
        if __df.iat[2, 0] != 'все':
            user_shift = [__df.iat[2, 0]]
        else:
            user_shift = ['день', 'ночь']

        mask = (df_data['краткий VIN'] == user_vin) &\
            ((df_data['Дата'] >= user_start_date) & (df_data['Дата'] <= user_end_date)) &\
            (df_data['Смена: день / ночь'].isin(user_shift)) &\
            pd.notnull(df_data['№ путевого листа']
                       )  # Если номер путевого листа пустой, то в выборку не попадет

        return df_data[mask]


def transform_text_to_datetime(datestr: str):
    if isinstance(datestr, str):
        date_from_text = datetime.datetime.strptime(datestr, '%d.%m.%Y')
        return date_from_text


def generate_filename(wb: dict) -> str:
    filename = f"{wb.get('site')}_" \
        f"{wb.get('vehicle')}_" \
        f"{wb.get('vin')}_" \
        f"{wb.get('waybill_number')}_" \
        f"{datetime.datetime.now().strftime('%d.%m.%Y.%H.%M.%S.%f')}" \
        f".docx"
    return filename


def number_is_integer(number: float) -> int | float:
    if isinstance(number, float) and number.is_integer():
        return int(number)
    else:
        return number


class Waybill():

    def __init__(self, index, _df):

        # Общая информация
        self.site = _df.at[index, 'Карьер']
        self.waybill_number = _df.at[index, '№ путевого листа']
        self.driver = _df.at[index, 'Машинист (ФИО)  из ПУТЕВОГО ЛИСТА']
        self.master_name = _df.at[index, 'ФИО мастера']

        # Техника
        self.vehicle = _df.at[index, 'Наименование СТ']
        self.vehicle_model = _df.at[index, 'Модель']
        self.vin = _df.at[index, 'краткий VIN']

        # Дата
        self.date = _df.at[index, 'Дата']
        self.shift = _df.at[index, 'Смена: день / ночь']

        # Моточасы
        self.hours_start = number_is_integer(
            _df.at[index, 'Моточасы на начало смены'])
        self.hours_end = number_is_integer(
            _df.at[index, 'Моточасы на конец смены'])
        self.hours_shift = number_is_integer(
            _df.at[index, 'Моточасы - наработка за смену  из ПУТЕВОГО ЛИСТА'])

        # Топливо
        self.fuel_start = number_is_integer(
            _df.at[index, 'Топливо -остаток на начало смены'])
        self.fuel_end = number_is_integer(
            _df.at[index, 'Топливо - остаток на конец смены  из ПУТЕВОГО ЛИСТА'])
        self.fuel_consumption = number_is_integer(
            _df.at[index, 'Топливо - расход - факт'])
        self.fuel_refill = number_is_integer(
            _df.at[index, 'Топливо - заправлено из ПУТЕВОГО ЛИСТА'])
        self.fuel_drain = number_is_integer(
            _df.at[index, 'Топливо - СЛИВ В ЕМКОСТЬ'])


def print_number_of_waybills(df: pd.DataFrame) -> None:
    print(f'Сгенерировано {len(df)} путевых листов.')


def generate_two_waybills(_df2: pd.DataFrame) -> None:
    doc = DocxTemplate(Path(__file__).parent.joinpath(
        'templates/waybill_template_two.docx'))
    for i in range(0, len(_df2), 2):

        wb1 = Waybill(_df2.index[i], _df2).__dict__
        wb2 = Waybill(_df2.index[i + 1], _df2).__dict__
        context = {'company_data': company_data, 'wb1': wb1, 'wb2': wb2}

        doc.render(context)

        doc.save(Path(__file__).parent.joinpath(
            f'output/{generate_filename(wb1)}'))


def generate_one_waybill(_df1: pd.DataFrame) -> None:
    doc = DocxTemplate(Path(__file__).parent.joinpath(
        'templates/waybill_template_one.docx'))
    wb1 = Waybill(_df1.index[0], _df1).__dict__

    context = {'company_data': company_data, 'wb1': wb1}

    doc.render(context)

    doc.save(Path(__file__).parent.joinpath(
        f'output/{generate_filename(wb1)}'))


def generate_waybills(_df):
    if len(_df) != 0:
        if len(_df) % 2 != 0:
            generate_two_waybills(_df[:-1])
            generate_one_waybill(_df[-1:])
            print_number_of_waybills(_df)
        else:
            generate_two_waybills(_df)
            print_number_of_waybills(_df)
    else:
        print('Нет данных.')


def compare_dates(date1, date2):
    date_format = '%d.%m.%Y'
    if check_date(date1) and check_date(date2):

        return datetime.datetime.strptime(date1, date_format) <= datetime.datetime.strptime(date2, date_format)


def check_date(date_str: str) -> True | False:
    """Преобразовывает строку в дату. Возвращает False, если преобразование не удалось."""
    date_format = '%d.%m.%Y'
    try:
        datetime.datetime.strptime(date_str, date_format)
    except ValueError:
        return False
    else:
        return True


def check_files():
    template1 = Path(__file__).parent.joinpath(
        'templates/waybill_template_one.docx')
    template2 = Path(__file__).parent.joinpath(
        'templates/waybill_template_two.docx')
    excel_file = Path(__file__).parent.joinpath('waybill_data.xlsx')
    py_script = Path(__file__).parent.joinpath('waybill.py')
    settings = Path(__file__).parent.joinpath('settings.py')

    paths = [template1, template2, excel_file, py_script, settings]

    for path in paths:
        if not path.is_file():
            print(f'Нет файла: {path.name}')
            return False

    print('Проверка файлов успешно завершена.')
    return True


def main(input_data: dict):
    start = list(input_data.values())[0][0]
    end = list(input_data.values())[0][1]

    if bool(input_data) and compare_dates(start, end):

        # Изменить путь на тот, где находится файл скрипта
        script_path = Path(__file__).parent.absolute()
        os.chdir(script_path)

        df = get_df_from_excel_file('waybill_data.xlsx')

        df_filtered = filter_df(input_data, df)

        generate_waybills(df_filtered)
    else:
        print('Неверные данные! Проверьте даты!')


if __name__ == '__main__':

    main()

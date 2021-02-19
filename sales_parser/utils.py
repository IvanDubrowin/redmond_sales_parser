from datetime import datetime
from string import punctuation
from typing import List

import pandas as pd
from pandas.core.generic import FrameOrSeries

from sales_parser import constants

VALID_EXTENSIONS = ['.xls', '.xlsx']


def get_shop(file_path: str) -> List[str]:
    dataframe: FrameOrSeries = pd.read_excel(file_path, sheet_name=constants.TARGET_SHEET_NAME)
    choices_set = set(dataframe.filter(items=[constants.SHOP_ADDRESS_COL])[constants.SHOP_ADDRESS_COL])
    choices_list = sorted([item for item in choices_set if isinstance(item, str)])
    return choices_list


def validate_report_path(file_path: str) -> bool:
    is_valid = False

    for ext in VALID_EXTENSIONS:
        if file_path.endswith(ext):
            is_valid = True

    return is_valid


def write_to_excel(data: FrameOrSeries) -> None:
    filename = f'Отчет{datetime.strftime(datetime.now(), "%Y.%m.%d")}.xlsx'
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    data.to_excel(writer, sheet_name='Отчет')

    workbook = writer.book
    worksheet = writer.sheets['Отчет']

    f = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    worksheet.set_column('D:F', None, f)
    writer.save()


def normalize_product_code(code: str) -> str:
    stripped_chars = list(punctuation) + [' ', '']
    eng = [
        "a", "b", "c", "e", "k", "m", "n", "h", "o", "p", "t", "u", "y",
        "A", "B", "E", "K", "M", "O", "P", "C", "T", "H", "Y"
    ]
    rus = [
        "а", "в", "с", "е", "к", "м", "н", "н", "о", "р", "т", "и", "у",
        "А", "В", "Е", "К", "М", "О", "Р", "С", "Т", "Н", "У"
    ]
    code_chars = list(code)

    for char in code_chars:
        if char in rus:
            rus_idx = rus.index(char)
            idx = code_chars.index(char)
            code_chars[idx] = eng[rus_idx]
        if char in stripped_chars:
            code_chars.remove(char)

    return ''.join(code_chars).upper()

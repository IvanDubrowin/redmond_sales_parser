from datetime import datetime
from string import punctuation
from typing import List

import pandas as pd
from pandas.core.frame import DataFrame
from pandas.core.series import Series

from sales_parser import constants

VALID_EXTENSIONS = ['.xls', '.xlsx']


def get_shop(file_path: str) -> List[str]:
    dataframe: DataFrame = pd.read_excel(file_path, sheet_name=constants.TARGET_SHEET_NAME)
    choices_set = set(dataframe.filter(items=[constants.SHOP_ADDRESS_COL])[constants.SHOP_ADDRESS_COL])
    choices_list = sorted([item for item in choices_set if isinstance(item, str)])
    return choices_list


def validate_report_path(file_path: str) -> bool:
    for ext in VALID_EXTENSIONS:
        if file_path.endswith(ext):
            return True
    return False


def write_to_excel(dataframe: DataFrame) -> None:
    filename = f'Отчет{datetime.strftime(datetime.now(), "%Y.%m.%d")}.xlsx'
    writer = pd.ExcelWriter(filename, engine='xlsxwriter')
    dataframe.to_excel(writer, sheet_name='Отчет')

    workbook = writer.book
    worksheet = writer.sheets['Отчет']

    f = workbook.add_format({'align': 'center', 'valign': 'vcenter'})
    worksheet.set_column('D:F', None, f)

    for idx, col in enumerate(dataframe, 1):
        series: Series = dataframe[col]
        max_len = max([
            series.astype(str).map(len).max(),
            len(series.name or '')
        ])
        worksheet.set_column(idx, idx, max_len)
    writer.save()


def normalize_product_code(code: str) -> str:
    stripped_chars = list(punctuation) + [' ', '']
    rus_to_eng = {
        "а": "a",
        "в": "b",
        "с": "c",
        "е": "e",
        "к": "k",
        "м": "m",
        "н": "h",
        "о": "o",
        "р": "p",
        "т": "t",
        "и": "u",
        "у": "y",
        "А": "A",
        "В": "B",
        "Е": "E",
        "К": "K",
        "М": "M",
        "О": "O",
        "Р": "P",
        "С": "C",
        "Т": "T",
        "Н": "H",
        "У": "Y"
    }

    code_chars = list(code)

    for char in code_chars:
        if char in stripped_chars:
            code_chars.remove(char)

    for idx, char in enumerate(code_chars):
        char_to_replace = rus_to_eng.get(char)
        if char_to_replace is not None:
            code_chars[idx] = char_to_replace

    return ''.join(code_chars).upper()

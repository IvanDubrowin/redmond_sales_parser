import pandas as pd
import os

def get_shop(filepath):

    sales_report = pd.read_excel(filepath, sheet_name='Портянка').filter(items=['Адрес т.т. '])
    choices_set = set(sales_report['Адрес т.т. '])

    return choices_set

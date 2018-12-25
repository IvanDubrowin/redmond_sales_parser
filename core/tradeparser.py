import pandas as pd
from numpy import mean, nan
import re
from datetime import datetime
from string import punctuation
from openpyxl import load_workbook


class TradeParser:

    def __init__(self, path_realization=None, path_stock=None, sales_report=None, shop=None):
        self.path_realization = path_realization
        self.stock = path_stock
        self.sales_report = sales_report
        self.shop = shop

    def parse_stock_and_realization(self):
        realization = pd.read_excel(self.path_realization)
        stock = pd.read_excel(self.stock)
        cols_rl = ['Наименование', 'Кол-во', 'Сумма']
        if 'Остаток на дату' in stock:
            cols_st = ['Наименование', 'Остаток на дату']
        else:
            cols_st = ['Наименование', 'Тек. Остаток']

        try:
            data_rl = realization[cols_rl].groupby('Наименование').apply(
                        lambda x: x.set_index('Наименование').to_dict('list')).to_dict()

            data_st = stock[cols_st].groupby('Наименование').mean().to_dict('int')

            for k_rl in data_rl.keys():
                data_rl[k_rl]['Кол-во'] = sum(data_rl[k_rl]['Кол-во'])
                data_rl[k_rl]['Сумма'] = int(mean(data_rl[k_rl]['Сумма']))
                data_rl[k_rl].update({'Тек. Остаток': 0})

            for k_st in data_st.keys():
                if k_st in data_rl.keys():
                    data_rl[k_st]['Тек. Остаток'] = int(data_st[k_st][cols_st[1]])
                else:
                    data_rl.update(dict([(k_st, data_st[k_st])]))
        except KeyError as e:
            return e
        return data_rl

    def parse_sales_report(self):
        get_product_code = r'(R[A-ZА-Я]{1,4}[- ]{1}[А-ЯA-Z0-9/-]{1,8})\w+|(MSC[- ]{1}[А-ЯA-Z0-9/-]{1,8})\w+'
        data_dict = self.parse_stock_and_realization()
        count_data_dict_keys = len(data_dict)
        not_found_keys = set()
        found_keys = set()
        append_col = pd.DataFrame({'Кол-во': [nan], 'Тек. Остаток': [nan], 'Сумма': [nan]})

        eng = ["a", "b", "c", "e", "k", "m", "n", "h", "o", "p", "t", "u", "y", "A", "B", "E", "K", "M", "O", "P", "C", "T", "H", "Y"]
        rus = ["а", "в", "с", "е", "к", "м", "н", "н", "о", "р", "т", "и", "у", "А", "В", "Е", "К", "М", "О", "Р", "С", "Т", "Н", "У"]
        stripped_chars = list(punctuation)
        tables = [' ', '']
        stripped_chars.extend(tables)

        sales_report = pd.read_excel(self.sales_report, sheet_name='Портянка').filter(items=['Адрес т.т. ', 'Название SKU'])
        sales_report = sales_report[sales_report['Адрес т.т. '].str.contains(self.shop, regex=False)]
        sales_report = sales_report.join(append_col)
        sales_report = sales_report.fillna(0)

        for key in data_dict.keys():
            try:
                product_code = re.search(get_product_code, key).group()
            except AttributeError:
                product_code = key.split(' ')[-1]

            product_code.upper()
            product_code = list(product_code)

            for char in product_code:
                if char in rus:
                    rus_ind = rus.index(char)
                    ind = product_code.index(char)
                    product_code[ind] = eng[rus_ind]
                if char in stripped_chars:
                    product_code.remove(char)

            product_code = ''.join(product_code)

            for item in sales_report['Название SKU'].items():
                list_item = list(item[1].upper())
                for char in list_item:
                    if char in rus:
                        rus_ind = rus.index(char)
                        ind = list_item.index(char)
                        list_item[ind] = eng[rus_ind]
                    if char in stripped_chars:
                        list_item.remove(char)
                stripped_item = ''.join(list_item)

                if product_code in stripped_item:
                    sales_report.at[item[0], 'Кол-во'] = int(data_dict[key].get('Кол-во', 0))
                    sales_report.at[item[0], 'Тек. Остаток'] = int(data_dict[key].get('Тек. Остаток', 0))
                    sales_report.at[item[0], 'Сумма'] =  int(data_dict[key].get('Сумма', 0))
                    found_keys.add(key)
                else:
                    not_found_keys.add(key)

        count_matches_keys = len(found_keys)

        if len(not_found_keys) > 0:
            not_found_keys = list(found_keys.symmetric_difference(not_found_keys))
            not_found_keys = '\n'.join(not_found_keys)
            report_info = f'Найдено {count_matches_keys} товаров из {count_data_dict_keys}, не найдено: {not_found_keys} \nНеобходимо исправить данные!'
        else:
            report_info = f'Найдено {count_matches_keys} товаров из {count_data_dict_keys}'
        return sales_report, report_info

    def write_to_excel(self):
        filename = f'Отчет{datetime.strftime(datetime.now(), "%Y.%m.%d")}.xlsx'
        sales = self.parse_sales_report()
        writer = pd.ExcelWriter(filename, engine = 'xlsxwriter')
        sales[0].to_excel(writer, sheet_name = 'Отчет')
        writer.save()
        writer.close()
        return sales[1]

def get_shop(filepath):
    sales_report = pd.read_excel(filepath, sheet_name='Портянка').filter(items=['Адрес т.т. '])
    choices_set = set(sales_report['Адрес т.т. '])
    choices_set = sorted(list(choices_set))
    return choices_set

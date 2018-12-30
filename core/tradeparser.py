import pandas as pd
import re
from numpy import mean, nan
from datetime import datetime
from string import punctuation
from fuzzywuzzy import fuzz


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
                    val = dict(zip(data_st[k_st].keys(),
                                  [int(value) for value in data_st[k_st].values()])
                                )
                    data_rl.update({k_st: val})
        except KeyError as e:
            return e
        return data_rl

    def parse_sales_report(self):
        data_dict = self.parse_stock_and_realization()
        count_data_dict_keys = len(data_dict)
        not_found_keys = set()
        found_keys = set()
        append_col = pd.DataFrame({'Кол-во': [nan], 'Тек. Остаток': [nan], 'Сумма': [nan]})

        sales_report = pd.read_excel(self.sales_report, sheet_name='Портянка').filter(items=['Адрес т.т. ', 'Название SKU'])
        sales_report = sales_report[sales_report['Адрес т.т. '].str.contains(self.shop, regex=False)]
        sales_report.join(append_col)

        sales_report_keys_and_code = self.search_code(sales_report['Название SKU'].items())
        data_dict_keys_and_code = self.search_code(data_dict.keys())

        for key in data_dict_keys_and_code:
            match = [item for item in sales_report_keys_and_code if fuzz.ratio(key[1], item[1]) == 100]
            if match:
                sales_report.at[match[0][0][0], 'Кол-во'] = str(data_dict[key[0]].get('Кол-во', ''))
                sales_report.at[match[0][0][0], 'Тек. Остаток'] = str(data_dict[key[0]].get('Тек. Остаток', ''))
                sales_report.at[match[0][0][0], 'Сумма'] =  str(data_dict[key[0]].get('Сумма', ''))
                found_keys.add(key[0])
            else:
                not_found_keys.add(key[0])

        count_matches_keys = len(found_keys)

        if not_found_keys:
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

    @staticmethod
    def search_code(product_keys):
        search_keys = []
        get_product_code = r'(R[A-ZА-Я]{1,4}[- ]{1}[А-ЯA-Z0-9/-]{1,8})\w+|(MSC[- ]{1}[А-ЯA-Z0-9/-]{1,8})\w+'
        stripped_chars = list(punctuation) + [' ', '']

        eng = ["a", "b", "c", "e", "k", "m", "n", "h", "o", "p", "t", "u", "y",
               "A", "B", "E", "K", "M", "O", "P", "C", "T", "H", "Y"]
        rus = ["а", "в", "с", "е", "к", "м", "н", "н", "о", "р", "т", "и", "у",
               "А", "В", "Е", "К", "М", "О", "Р", "С", "Т", "Н", "У"]

        for item in product_keys:
            try:
                if isinstance(item, tuple):
                    product_code = re.search(get_product_code, item[1]).group()
                else:
                    product_code = re.search(get_product_code, item).group()
            except AttributeError:
                product_code = item.split(' ')[-1]

            product_code = list(product_code)

            for char in product_code:
                if char in rus:
                    rus_ind = rus.index(char)
                    ind = product_code.index(char)
                    product_code[ind] = eng[rus_ind]
                if char in stripped_chars:
                    product_code.remove(char)

            product_code = ''.join(product_code)
            search_keys.append((item, product_code.upper()))

        return search_keys

def get_shop(filepath):
    sales_report = pd.read_excel(filepath, sheet_name='Портянка').filter(items=['Адрес т.т. '])
    choices_set = set(sales_report['Адрес т.т. '])
    choices_set = sorted(list(choices_set))
    return choices_set

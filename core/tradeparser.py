import pandas as pd
import re
from datetime import datetime
from numpy import mean, nan
from openpyxl import load_workbook


class TradeParser:

    def __init__(self, path_realization=None, path_stock=None, sales_report=None):
        self.path_realization = path_realization
        self.stock = path_stock
        self.sales_report = sales_report

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
        append_col = pd.DataFrame({'Кол-во': [nan], 'Тек. Остаток': [nan], 'Сумма': [nan]})

        eng = ["a", "b", "c", "e", "k", "m", "n", "h", "o", "p", "t", "u", "y", "A", "B", "E", "K", "M", "O", "P", "C", "T", "H", "Y"]
        rus = ["а", "в", "с", "е", "к", "м", "н", "н", "о", "р", "т", "и", "у", "А", "В", "Е", "К", "М", "О", "Р", "С", "Т", "Н", "У"]

        sales_report = pd.read_excel(self.sales_report, sheet_name='Портянка').filter(items=['Код ТТ', 'Название SKU'])
        sales_report = sales_report[sales_report['Код ТТ'].isin([77056])]
        sales_report = sales_report.join(append_col)
        sales_report = sales_report.fillna(0)
        for item in sales_report['Название SKU'].items():
            try:
                product_code = re.search(get_product_code, item[1]).group()

                product_code = list(product_code)
                for char in product_code:
                    if char in rus:
                        rus_ind = rus.index(char)
                        ind = product_code.index(char)
                        product_code[ind] = eng[rus_ind]

                product_code = ''.join(product_code)

            except AttributeError:
                pass
            for key in data_dict.keys():
                try:
                    code_in_data_dict = re.search(get_product_code, key).group()
                except AttributeError:
                    pass

                if product_code == code_in_data_dict:
                    sales_report.at[item[0], 'Кол-во'] = int(data_dict[key].get('Кол-во', 0))
                    sales_report.at[item[0], 'Тек. Остаток'] = int(data_dict[key].get('Тек. Остаток', 0))
                    sales_report.at[item[0], 'Сумма'] =  int(data_dict[key].get('Сумма', 0))

        return sales_report

    def write_to_excel(self):
        filename = f'Отчет{datetime.strftime(datetime.now(), "%Y.%m.%d %H:%M:%S")}.xlsx'
        sales = self.parse_sales_report()
        writer = pd.ExcelWriter(filename, engine = 'xlsxwriter')
        sales.to_excel(writer, sheet_name = 'Отчет')
        writer.save()
        writer.close()

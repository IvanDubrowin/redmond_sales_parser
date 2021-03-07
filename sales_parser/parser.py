import re
from typing import Set, Tuple, List, Union

import pandas as pd
from numpy import nan
from pandas.core.series import Series
from pandas.core.frame import DataFrame

from sales_parser import constants
from sales_parser.utils import normalize_product_code


class SalesParser:
    def __init__(
            self,
            realization_path: str,
            stock_path: str,
            report_path: str,
            shop: str
    ) -> None:
        self.realization_path = realization_path
        self.stock_path = stock_path
        self.report_path = report_path
        self.shop = shop

    PRODUCT_CODE_REGEX = re.compile(r'(R[A-ZА-Я]{1,4}[- ][А-ЯA-Z0-9/-]{1,8})\w+|(MSC[- ][А-ЯA-Z0-9/-]{1,8})\w+')

    def _parse_stock(self) -> DataFrame:
        stock: DataFrame = pd.read_excel(self.stock_path)
        stock_cols = self._get_stock_cols()

        frame = stock[stock_cols] \
            .groupby(constants.PRODUCT_NAME_IN_STOCK) \
            .mean() \
            .reset_index() \
            .set_axis(
                [
                    constants.PRODUCT_NAME_IN_REAL,
                    constants.CURRENT_BALANCE,
                    constants.PRODUCT_PRICE_IN_STOCK
                ],
                axis='columns'
            )

        frame[constants.CURRENT_BALANCE] = pd.to_numeric(
            frame[constants.CURRENT_BALANCE], downcast='unsigned'
        )
        frame[constants.PRODUCT_PRICE_IN_STOCK] = pd.to_numeric(
            frame[constants.PRODUCT_PRICE_IN_STOCK], downcast='unsigned'
        )
        return frame

    def _parse_realization(self) -> DataFrame:
        realization: DataFrame = pd.read_excel(self.realization_path)
        realization_cols = self._get_realization_cols()

        return realization[realization_cols] \
            .groupby(constants.PRODUCT_NAME_IN_REAL) \
            .sum() \
            .apply(self._calc_avg_product_price, axis='columns') \
            .reset_index()

    def _get_stock_and_realization(self) -> DataFrame:
        realization = self._parse_realization()
        stock = self._parse_stock()

        return pd.merge(
            stock,
            realization,
            on=constants.PRODUCT_NAME_IN_REAL,
            how='outer'
        )

    def _get_sales_report(self) -> DataFrame:
        sales_report: DataFrame = pd.read_excel(
            self.report_path,
            sheet_name=constants.TARGET_SHEET_NAME
        ) \
            .filter(
            items=[
                constants.SHOP_ADDRESS_COL,
                constants.PRODUCT_NAME_IN_REPORT
            ]
        )
        sales_report = sales_report[
            sales_report[constants.SHOP_ADDRESS_COL].str.contains(
                self.shop,
                regex=False,
                na=False
            )
        ]
        sales_report.join(
            pd.DataFrame({
                constants.PRODUCT_COUNT: [nan],
                constants.CURRENT_BALANCE: [nan],
                constants.PRODUCT_SUM: [nan]
            })
        )
        return sales_report

    def _update_sales_report(self) -> Tuple[DataFrame, str]:
        stock_and_realization = self._get_stock_and_realization()
        sales_report = self._get_sales_report()
        found_products = set()

        for _, product_attrs in stock_and_realization.iterrows():
            sales_report = sales_report.apply(
                func=self.__update_sales_report_row,
                axis='columns',
                args=(product_attrs, found_products)
            )

        raise AssertionError

        # for product, attrs in stock_and_realization.items():
        #     for label, row_data in sales_report[constants.PRODUCT_NAME_IN_REPORT].items():
        #         if self._is_equal_product_codes(product, row_data):
        #             product_count = attrs.get(constants.PRODUCT_COUNT)
        #
        #             if product_count is not None:
        #                 sales_report.at[label, constants.PRODUCT_COUNT] = product_count
        #
        #             product_stock = attrs.get(constants.CURRENT_BALANCE)
        #
        #             if product_stock is not None:
        #                 sales_report.at[label, constants.CURRENT_BALANCE] = product_stock
        #
        #             product_sum = attrs.get(constants.PRODUCT_SUM)
        #
        #             if product_sum is not None:
        #                 sales_report.at[label, constants.PRODUCT_SUM] = product_sum
        #
        #             found_products.add(product)

        return sales_report, self._get_report_info(stock_and_realization, found_products)

    def __update_sales_report_row(
            self,
            row: Series,
            product_attrs: Series,
            found_products: set
    ) -> Series:
        if self._is_equal_product_codes(
                row[constants.PRODUCT_NAME_IN_REPORT],
                product_attrs[constants.PRODUCT_NAME_IN_REAL]
        ):
            print('Equal')
        return row

    def _is_equal_product_codes(self, left: str, right: str) -> bool:
        left_match = self.PRODUCT_CODE_REGEX.search(left)
        right_match = self.PRODUCT_CODE_REGEX.search(right)

        if not left_match or not right_match:
            return False

        return normalize_product_code(left_match.group()) == normalize_product_code(right_match.group())

    @staticmethod
    def _get_report_info(stock_and_real: dict, found: Set[str]) -> str:
        count_all_products = len(stock_and_real)
        not_found = set(stock_and_real) - found
        report_info = f'Найдено {len(found)} товаров из {count_all_products}'

        if not_found:
            not_found_to_str = "\n\n".join(not_found)
            report_info += (
                f', не найдено: \n\n{not_found_to_str} \n\nНеобходимо исправить данные!'
            )
        return report_info

    @staticmethod
    def _get_realization_cols() -> List[str]:
        return [
            constants.PRODUCT_NAME_IN_REAL,
            constants.PRODUCT_COUNT,
            constants.PRODUCT_PRICE_IN_REAL
        ]

    @staticmethod
    def _get_stock_cols() -> List[str]:
        return [
            constants.PRODUCT_NAME_IN_STOCK,
            constants.PRODUCT_COUNT,
            constants.PRODUCT_PRICE_IN_STOCK
        ]

    @staticmethod
    def _calc_avg_product_price(series: Series) -> Series:
        if series[constants.PRODUCT_PRICE_IN_REAL] > 0:
            series[constants.PRODUCT_PRICE_IN_REAL] = int(
                series[constants.PRODUCT_PRICE_IN_REAL] / series[constants.PRODUCT_COUNT]
            )
        return series

    def get_report(self) -> Tuple[DataFrame, str]:
        return self._update_sales_report()

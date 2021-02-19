import re
from typing import List, Set, Tuple

import pandas as pd
from numpy import nan
from pandas.core.frame import DataFrame
from pandas.core.generic import FrameOrSeries

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

    def _parse_stock_and_realization(self) -> dict:
        realization: DataFrame = pd.read_excel(self.realization_path)
        stock: DataFrame = pd.read_excel(self.stock_path)

        realization_cols = self._get_realization_cols()
        stock_cols = self._get_stock_cols(stock)

        realization_data = realization[realization_cols] \
            .groupby(constants.PRODUCT_NAME_IN_STOCK_AND_REAL) \
            .apply(lambda x: x.set_index(constants.PRODUCT_NAME_IN_STOCK_AND_REAL).to_dict('list')) \
            .to_dict()

        stock_data = stock[stock_cols] \
            .groupby(constants.PRODUCT_NAME_IN_STOCK_AND_REAL) \
            .mean() \
            .to_dict('int')

        for key in realization_data:
            realization_data[key][constants.PRODUCT_COUNT] = sum(realization_data[key][constants.PRODUCT_COUNT])

            if sum(realization_data[key][constants.PRODUCT_SUM]) != 0:
                realization_data[key][constants.PRODUCT_SUM] = int(
                    sum(
                        realization_data[key][constants.PRODUCT_SUM]
                    ) / realization_data[key][constants.PRODUCT_COUNT]
                )
            else:
                realization_data[key][constants.PRODUCT_SUM] = 0

            realization_data[key].update({constants.CURRENT_BALANCE: None})

        for key in stock_data:
            balance = int(stock_data[key][stock_cols[1]])

            if key in realization_data:
                realization_data[key][constants.CURRENT_BALANCE] = balance
            else:
                realization_data.update({key: {constants.CURRENT_BALANCE: balance}})

        return realization_data

    def _parse_sales_report(self) -> Tuple[FrameOrSeries, str]:
        stock_and_realization = self._parse_stock_and_realization()
        found_products = set()

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

        for product, attrs in stock_and_realization.items():
            for label, row_data in sales_report[constants.PRODUCT_NAME_IN_REPORT].items():
                if self._compare_product_code(product, row_data):
                    product_count = attrs.get(constants.PRODUCT_COUNT)

                    if product_count is not None:
                        sales_report.at[label, constants.PRODUCT_COUNT] = product_count

                    product_stock = attrs.get(constants.CURRENT_BALANCE)

                    if product_stock is not None:
                        sales_report.at[label, constants.CURRENT_BALANCE] = product_stock

                    product_sum = attrs.get(constants.PRODUCT_SUM)

                    if product_sum is not None:
                        sales_report.at[label, constants.PRODUCT_SUM] = product_sum

                    found_products.add(product)

        return sales_report, self._get_report_info(stock_and_realization, found_products)

    def _compare_product_code(self, left: str, right: str) -> bool:
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
            constants.PRODUCT_NAME_IN_STOCK_AND_REAL,
            constants.PRODUCT_COUNT,
            constants.PRODUCT_SUM
        ]

    @staticmethod
    def _get_stock_cols(stock: DataFrame) -> List[str]:
        if constants.BALANCE_PER_DATE in stock:
            return [
                constants.PRODUCT_NAME_IN_STOCK_AND_REAL,
                constants.BALANCE_PER_DATE
            ]
        return [
            constants.PRODUCT_NAME_IN_STOCK_AND_REAL,
            constants.CURRENT_BALANCE
        ]

    def parse(self) -> Tuple[FrameOrSeries, str]:
        return self._parse_sales_report()

import time
import traceback
from typing import Optional

from PyQt5 import QtCore, QtWidgets

from sales_parser.parser import SalesParser
from sales_parser.utils import get_shop, validate_report_path, write_to_excel


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self.setFixedSize(640, 440)
        self._setupUi()
        self.sales_path.setReadOnly(True)
        self.report_path.setReadOnly(True)
        self.stock_path.setReadOnly(True)
        self.download_report_button.setEnabled(False)
        self.sales_button.setEnabled(False)
        self.stock_button.setEnabled(False)
        self.generate_report_button.setEnabled(False)
        self.choice_shop.setEnabled(False)
        self.progressBar.setFormat('Индикатор выполнения программы')
        self._centerOnScreen()
        self.shop: Optional[str] = None

    def _setupUi(self) -> None:
        self.setObjectName("Generator")
        self.resize(641, 461)
        self.centralwidget = QtWidgets.QWidget(self)
        self.centralwidget.setObjectName("centralwidget")
        self.report_path = QtWidgets.QTextEdit(self.centralwidget)
        self.report_path.setGeometry(QtCore.QRect(30, 30, 401, 31))
        self.report_path.setObjectName("report_path")
        self.report_button = QtWidgets.QPushButton(self.centralwidget)
        self.report_button.setGeometry(QtCore.QRect(450, 30, 161, 29))
        self.report_button.setObjectName("report_button")
        self.report_button.clicked.connect(self._open_report)
        self.download_report_button = QtWidgets.QPushButton(self.centralwidget)
        self.download_report_button.setGeometry(QtCore.QRect(238, 80, 191, 29))
        self.download_report_button.setObjectName("download_report_button")
        self.download_report_button.clicked.connect(self._download_report)
        self.choice_shop = QtWidgets.QComboBox(self.centralwidget)
        self.choice_shop.setGeometry(QtCore.QRect(30, 130, 580, 25))
        self.choice_shop.setObjectName("choice_shop")
        self.choice_shop.currentIndexChanged[str].connect(self._set_shop)
        self.sales_path = QtWidgets.QTextEdit(self.centralwidget)
        self.sales_path.setGeometry(QtCore.QRect(30, 190, 401, 31))
        self.sales_path.setObjectName("sales_path")
        self.stock_button = QtWidgets.QPushButton(self.centralwidget)
        self.stock_button.setGeometry(QtCore.QRect(450, 260, 161, 29))
        self.stock_button.setObjectName("stock_button")
        self.stock_button.clicked.connect(self._open_stock)
        self.stock_path = QtWidgets.QTextEdit(self.centralwidget)
        self.stock_path.setGeometry(QtCore.QRect(30, 260, 401, 31))
        self.stock_path.setObjectName("stock_path")
        self.sales_button = QtWidgets.QPushButton(self.centralwidget)
        self.sales_button.setGeometry(QtCore.QRect(450, 190, 161, 29))
        self.sales_button.setObjectName("sales_button")
        self.sales_button.clicked.connect(self._open_sales_report)
        self.generate_report_button = QtWidgets.QPushButton(self.centralwidget)
        self.generate_report_button.setGeometry(QtCore.QRect(240, 330, 191, 29))
        self.generate_report_button.setObjectName("generate_report_button")
        self.generate_report_button.clicked.connect(self._write_to_report)
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(10, 390, 622, 25))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)

        self._retranslateUi()
        QtCore.QMetaObject.connectSlotsByName(self)

    def _retranslateUi(self) -> None:
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("Generator", "Генератор отчетов"))
        self.report_button.setText(_translate("Generator", "Отчет для записи"))
        self.download_report_button.setText(_translate("Generator", "Загрузить"))
        self.stock_button.setText(_translate("Generator", "Остатки"))
        self.sales_button.setText(_translate("Generator", "Продажи"))
        self.generate_report_button.setText(_translate("Generator", "Создать отчет"))

    def _centerOnScreen(self) -> None:
        resolution: QtCore.QRect = QtWidgets.QDesktopWidget().screenGeometry()
        self.move(
            (resolution.width() / 2) - (self.frameSize().width() / 2),
            (resolution.height() / 2) - (self.frameSize().height() / 2)
        )

    def _open_sales_report(self) -> None:
        filename, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            'Вложите отчет с продажами',
            '',
            'Excel (*.xls *.xlsx)',
            options=QtWidgets.QFileDialog.DontUseNativeDialog
        )
        self.sales_path.setText(filename)

    def _open_stock(self) -> None:
        filename, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            'Вложите отчет с остатками',
            '',
            'Excel (*.xls *.xlsx)',
            options=QtWidgets.QFileDialog.DontUseNativeDialog
        )
        self.stock_path.setText(filename)

    def _open_report(self) -> None:
        filename, _ = QtWidgets.QFileDialog.getOpenFileName(
            self,
            'Вложите отчет для записи',
            '',
            'Excel (*.xls *.xlsx)',
            options=QtWidgets.QFileDialog.DontUseNativeDialog
        )
        self.report_path.setText(filename)
        text = self.report_path.toPlainText()

        if text:
            self.download_report_button.setEnabled(True)

    def _download_report(self) -> None:
        report_path = self.report_path.toPlainText()
        is_valid = validate_report_path(report_path)

        if not is_valid:
            self._set_error_messagebox("Вы загрузили не Excel файл")
            return None

        try:
            self.progressBar.setFormat('Идет выполнение программы, ждите')
            self._set_progress_percent(100)

            choices = get_shop(report_path)

            self.choice_shop.addItems(choices)
            self.choice_shop.setEnabled(True)
            self._set_progress_percent(0)
            self.progressBar.setFormat('Индикатор выполнения программы')
        except Exception as e:
            self._set_error_messagebox(self._format_error(e))

    def _set_shop(self, shop: str) -> None:
        self.shop = shop
        self.sales_button.setEnabled(True)
        self.stock_button.setEnabled(True)
        self.generate_report_button.setEnabled(True)

    def _write_to_report(self) -> None:
        realization = self.sales_path.toPlainText()
        stock = self.stock_path.toPlainText()
        report = self.report_path.toPlainText()

        is_valid_paths = all(validate_report_path(path) for path in [realization, stock, report])

        if not is_valid_paths or not self.shop:
            self._set_error_messagebox("Вы загрузили не все файлы!")
            return None

        try:
            parser = SalesParser(
                realization_path=realization,
                stock_path=stock,
                report_path=report,
                shop=self.shop
            )
            self.progressBar.setFormat('Идет выполнение программы, ждите')

            self._set_progress_percent(50)

            data, info = parser.parse()

            self._set_progress_percent(100, use_current_value=True)

            write_to_excel(data)

            self._set_progress_percent(0)
            self.progressBar.setFormat('Индикатор выполнения программы')
            success = QtWidgets.QMessageBox()
            success.setIcon(QtWidgets.QMessageBox.Information)
            success.setInformativeText(info)
            success.setWindowTitle("Загрузка прошла успешно!")
            success.setStandardButtons(QtWidgets.QMessageBox.Ok)
            success.exec_()
        except Exception as e:
            self._set_error_messagebox(self._format_error(e))

    def _set_error_messagebox(self, err_text: str) -> None:
        self._set_progress_percent(0)
        self.progressBar.setFormat('Индикатор выполнения программы')
        errmsg = QtWidgets.QMessageBox()
        errmsg.setIcon(QtWidgets.QMessageBox.Warning)
        errmsg.setInformativeText(err_text)
        errmsg.setWindowTitle("Ошибка")
        errmsg.setStandardButtons(QtWidgets.QMessageBox.Ok)
        errmsg.exec_()

    def _set_progress_percent(self, percent: int, use_current_value: bool = False) -> None:
        value = self.progressBar.value() if use_current_value else 0

        if percent < 1:
            self.progressBar.setValue(value)
            return None

        for i in range(percent):
            value += 1
            time.sleep(0.1)
            self.progressBar.setValue(value)

    @staticmethod
    def _format_error(err: Exception) -> str:
        fmt_traceback = '\n'.join(traceback.format_tb(err.__traceback__))
        return f'{str(err)}\n {fmt_traceback}'

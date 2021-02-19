from typing import Optional

from PyQt5 import QtCore, QtWidgets

from sales_parser.parser import SalesParser
from sales_parser.utils import get_shop, validate_report_path, write_to_excel


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        self.setFixedSize(640, 440)
        self.setupUi()
        self.sales_path.setReadOnly(True)
        self.report_path.setReadOnly(True)
        self.stock_path.setReadOnly(True)
        self.download_report_button.setEnabled(False)
        self.sales_button.setEnabled(False)
        self.stock_button.setEnabled(False)
        self.generate_report_button.setEnabled(False)
        self.choice_shop.setEnabled(False)
        self.progressBar.setFormat('Индикатор выполнения программы')
        self.centerOnScreen()
        self.shop: Optional[str] = None

    def setupUi(self) -> None:
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
        self.report_button.clicked.connect(self.open_report)
        self.download_report_button = QtWidgets.QPushButton(self.centralwidget)
        self.download_report_button.setGeometry(QtCore.QRect(238, 80, 191, 29))
        self.download_report_button.setObjectName("download_report_button")
        self.download_report_button.clicked.connect(self.download_report)
        self.choice_shop = QtWidgets.QComboBox(self.centralwidget)
        self.choice_shop.setGeometry(QtCore.QRect(30, 130, 580, 25))
        self.choice_shop.setObjectName("choice_shop")
        self.choice_shop.currentIndexChanged[str].connect(self.set_shop)
        self.sales_path = QtWidgets.QTextEdit(self.centralwidget)
        self.sales_path.setGeometry(QtCore.QRect(30, 190, 401, 31))
        self.sales_path.setObjectName("sales_path")
        self.stock_button = QtWidgets.QPushButton(self.centralwidget)
        self.stock_button.setGeometry(QtCore.QRect(450, 260, 161, 29))
        self.stock_button.setObjectName("stock_button")
        self.stock_button.clicked.connect(self.open_stock)
        self.stock_path = QtWidgets.QTextEdit(self.centralwidget)
        self.stock_path.setGeometry(QtCore.QRect(30, 260, 401, 31))
        self.stock_path.setObjectName("stock_path")
        self.sales_button = QtWidgets.QPushButton(self.centralwidget)
        self.sales_button.setGeometry(QtCore.QRect(450, 190, 161, 29))
        self.sales_button.setObjectName("sales_button")
        self.sales_button.clicked.connect(self.open_sales_report)
        self.generate_report_button = QtWidgets.QPushButton(self.centralwidget)
        self.generate_report_button.setGeometry(QtCore.QRect(240, 330, 191, 29))
        self.generate_report_button.setObjectName("generate_report_button")
        self.generate_report_button.clicked.connect(self.write_to_report)
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setGeometry(QtCore.QRect(10, 390, 622, 25))
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(self)
        self.statusbar.setObjectName("statusbar")
        self.setStatusBar(self.statusbar)

        self.retranslateUi()
        QtCore.QMetaObject.connectSlotsByName(self)

    def retranslateUi(self) -> None:
        _translate = QtCore.QCoreApplication.translate
        self.setWindowTitle(_translate("Generator", "Генератор отчетов"))
        self.report_button.setText(_translate("Generator", "Отчет для записи"))
        self.download_report_button.setText(_translate("Generator", "Загрузить"))
        self.stock_button.setText(_translate("Generator", "Остатки"))
        self.sales_button.setText(_translate("Generator", "Продажи"))
        self.generate_report_button.setText(_translate("Generator", "Создать отчет"))

    def centerOnScreen(self) -> None:
        resolution: QtCore.QRect = QtWidgets.QDesktopWidget().screenGeometry()
        self.move(
            (resolution.width() / 2) - (self.frameSize().width() / 2),
            (resolution.height() / 2) - (self.frameSize().height() / 2)
        )

    def open_sales_report(self) -> None:
        filename, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, 'Вложите отчет с продажами', '', 'Excel (*.xls *.xlsx)'
        )
        self.sales_path.setText(filename)

    def open_stock(self) -> None:
        filename, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, 'Вложите отчет с остатками', '', 'Excel (*.xls *.xlsx)'
        )
        self.stock_path.setText(filename)

    def open_report(self) -> None:
        filename, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, 'Вложите отчет для записи', '', 'Excel (*.xls *.xlsx)'
        )
        self.report_path.setText(filename)
        text = self.report_path.toPlainText()

        if text:
            self.download_report_button.setEnabled(True)

    def download_report(self) -> None:
        report_path = self.report_path.toPlainText()
        is_valid = validate_report_path(report_path)

        if not is_valid:
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Warning)
            msg.setInformativeText("Вы загрузили не Excel файл")
            msg.setWindowTitle("Ошибка")
            msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
            msg.exec_()
            return None

        try:
            percent = 0
            self.progressBar.setFormat('Идет выполнение программы, ждите')

            for i in range(100):
                percent += 1
                self.progressBar.setValue(percent)

            choices = get_shop(report_path)

            self.choice_shop.addItems(choices)
            self.choice_shop.setEnabled(True)
            self.progressBar.setValue(0)
            self.progressBar.setFormat('Индикатор выполнения программы')
        except Exception as e:
            errmsg = QtWidgets.QMessageBox()
            errmsg.setIcon(QtWidgets.QMessageBox.Warning)
            errmsg.setInformativeText(str(e))
            errmsg.setWindowTitle("Ошибка")
            errmsg.setStandardButtons(QtWidgets.QMessageBox.Ok)
            errmsg.exec_()

    def set_shop(self, shop: str) -> None:
        self.shop = shop
        self.sales_button.setEnabled(True)
        self.stock_button.setEnabled(True)
        self.generate_report_button.setEnabled(True)

    def write_to_report(self) -> None:
        realization = self.sales_path.toPlainText()
        stock = self.stock_path.toPlainText()
        report = self.report_path.toPlainText()

        is_valid_paths = all(validate_report_path(path) for path in [realization, stock, report])

        if not is_valid_paths or not self.shop:
            msg = QtWidgets.QMessageBox()
            msg.setIcon(QtWidgets.QMessageBox.Warning)
            msg.setInformativeText("Вы загрузили не все файлы!")
            msg.setWindowTitle("Ошибка")
            msg.setStandardButtons(QtWidgets.QMessageBox.Ok)
            msg.exec_()
            return None

        try:
            parser = SalesParser(
                realization_path=realization,
                stock_path=stock,
                report_path=report,
                shop=self.shop
            )
            self.progressBar.setFormat('Идет выполнение программы, ждите')

            percent = 0
            for i in range(100):
                percent += 1
                self.progressBar.setValue(percent)

            data, info = parser.parse()
            write_to_excel(data)

            self.progressBar.setValue(0)
            self.progressBar.setFormat('Индикатор выполнения программы')
            success = QtWidgets.QMessageBox()
            success.setIcon(QtWidgets.QMessageBox.Information)
            success.setInformativeText(info)
            success.setWindowTitle("Загрузка прошла успешно!")
            success.setStandardButtons(QtWidgets.QMessageBox.Ok)
            success.exec_()
        except OSError as e:
            self.progressBar.setValue(0)
            self.progressBar.setFormat('Индикатор выполнения программы')
            errmsg = QtWidgets.QMessageBox()
            errmsg.setIcon(QtWidgets.QMessageBox.Warning)
            errmsg.setInformativeText(str(e))
            errmsg.setWindowTitle("Ошибка")
            errmsg.setStandardButtons(QtWidgets.QMessageBox.Ok)
            errmsg.exec_()

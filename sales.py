from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from core.tradeparser import TradeParser, get_shop
import sys

class Ui_Generator(object):
    def setupUi(self, Generator):
        Generator.setObjectName("Generator")
        Generator.resize(641, 461)
        self.centralwidget = QtWidgets.QWidget(Generator)
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
        self.choice_shop = QComboBox(self.centralwidget)
        self.choice_shop.setGeometry(QtCore.QRect(30, 130, 580, 25))
        self.choice_shop.setObjectName("choice_shop")
        self.choice_shop.currentIndexChanged[str].connect(self.pick_shop)
        self.sales_path = QTextEdit(self.centralwidget)
        self.sales_path.setGeometry(QtCore.QRect(30, 190, 401, 31))
        self.sales_path.setObjectName("sales_path")
        self.stock_button = QtWidgets.QPushButton(self.centralwidget)
        self.stock_button.setGeometry(QtCore.QRect(450, 260, 161, 29))
        self.stock_button.setObjectName("stock_button")
        self.stock_button.clicked.connect(self.open_stock)
        self.stock_path = QTextEdit(self.centralwidget)
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
        Generator.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(Generator)
        self.statusbar.setObjectName("statusbar")
        Generator.setStatusBar(self.statusbar)

        self.retranslateUi(Generator)
        QtCore.QMetaObject.connectSlotsByName(Generator)

    def retranslateUi(self, Generator):
        _translate = QtCore.QCoreApplication.translate
        Generator.setWindowTitle(_translate("Generator", "Генератор отчетов"))
        self.report_button.setText(_translate("Generator", "Отчет для записи"))
        self.download_report_button.setText(_translate("Generator", "Загрузить"))
        self.stock_button.setText(_translate("Generator", "Остатки"))
        self.sales_button.setText(_translate("Generator", "Продажи"))
        self.generate_report_button.setText(_translate("Generator", "Создать отчет"))


class Main(QMainWindow, Ui_Generator):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setFixedSize(640, 440)
        self.setupUi(self)
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

    def centerOnScreen(self):
        resolution = QDesktopWidget().screenGeometry()
        self.move((resolution.width() / 2) - (self.frameSize().width() / 2),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))

    def open_sales_report(self):
        self.sales_path.setText(str(QFileDialog.getOpenFileName(self,'Вложите отчет с продажами', "","Excel (*.xls *.xlsx)")[0]))

    def open_stock(self):
        self.stock_path.setText(str(QFileDialog.getOpenFileName(self, 'Вложите отчет с остатками',"","Excel (*.xls *.xlsx)")[0]))

    def open_report(self):
        self.report_path.setText(str(QFileDialog.getOpenFileName(self,'Вложите отчет для записи' ,"","Excel (*.xls *.xlsx)")[0]))
        text = self.report_path.toPlainText()
        if len(text) > 0:
            self.download_report_button.setEnabled(True)

    def download_report(self):
        extensions = ['xls', 'xlsx']
        report = self.report_path.toPlainText()
        choices = None

        for extension in extensions:
            if report.endswith(extension):
                try:
                    percent = 0
                    self.progressBar.setFormat('Идет выполнение программы, ждите')
                    for i in range(100):
                        percent += 1
                        self.progressBar.setValue(percent)
                    choices = get_shop(report)
                    self.choice_shop.addItems(list(choices))
                    self.choice_shop.setEnabled(True)
                    self.progressBar.setValue(0)
                    self.progressBar.setFormat('Индикатор выполнения программы')
                except Exception as e:
                    errmsg = QMessageBox()
                    errmsg.setIcon(QMessageBox.Warning)
                    errmsg.setInformativeText(str(e))
                    errmsg.setWindowTitle("Ошибка")
                    errmsg.setStandardButtons(QMessageBox.Ok)
                    retval = errmsg.exec_()
        if choices is None:
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)
            msg.setInformativeText("Вы загрузили не Excel файл")
            msg.setWindowTitle("Ошибка")
            msg.setStandardButtons(QMessageBox.Ok)
            retval = msg.exec_()

    def pick_shop(self, shop):
        self.shop = shop
        self.sales_button.setEnabled(True)
        self.stock_button.setEnabled(True)
        self.generate_report_button.setEnabled(True)

    def write_to_report(self):
        realization = self.sales_path.toPlainText()
        stock = self.stock_path.toPlainText()
        report = self.report_path.toPlainText()

        tradeparser_args = {'realization': None, 'stock': None, 'report': None}
        extensions = ['xls', 'xlsx']

        for extension in extensions:
            if realization.endswith(extension):
                tradeparser_args['realization'] = realization
            if stock.endswith(extension):
                tradeparser_args['stock'] = stock
            if report.endswith(extension):
                tradeparser_args['report'] = report

        if None in tradeparser_args.values():
            msg = QMessageBox()
            msg.setIcon(QMessageBox.Warning)
            msg.setInformativeText("Вы загрузили не все файлы!")
            msg.setWindowTitle("Ошибка")
            msg.setStandardButtons(QMessageBox.Ok)
            retval = msg.exec_()
        else:
            try:
                parser = TradeParser(path_realization=realization, path_stock=stock, sales_report=report, shop=self.shop)
                percent = 0
                self.progressBar.setFormat('Идет выполнение программы, ждите')
                for i in range(100):
                    percent += 1
                    self.progressBar.setValue(percent)
                write_rep = parser.write_to_excel()
                self.progressBar.setValue(0)
                self.progressBar.setFormat('Индикатор выполнения программы')
                success = QMessageBox()
                success.setIcon(QMessageBox.Information)
                success.setInformativeText(write_rep)
                success.setWindowTitle("Загрузка прошла успешно!")
                success.setStandardButtons(QMessageBox.Ok)
                retval = success.exec_()
            except Exception as e:
                self.progressBar.setValue(0)
                self.progressBar.setFormat('Индикатор выполнения программы')
                errmsg = QMessageBox()
                errmsg.setIcon(QMessageBox.Warning)
                errmsg.setInformativeText(str(e))
                errmsg.setWindowTitle("Ошибка")
                errmsg.setStandardButtons(QMessageBox.Ok)
                retval = errmsg.exec_()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Main()
    window.show()
    sys.exit(app.exec_())

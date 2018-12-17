from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from core.tradeparser import TradeParser
import sys

class Ui_Generator(object):
    def setupUi(self, Generator):
        Generator.setObjectName("Generator")
        Generator.resize(640, 480)
        self.centralwidget = QtWidgets.QWidget(Generator)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout.setObjectName("gridLayout")
        self.textEdit_3 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_3.setObjectName("textEdit_3")
        self.gridLayout.addWidget(self.textEdit_3, 2, 0, 1, 1)
        self.progressBar = QtWidgets.QProgressBar(self.centralwidget)
        self.progressBar.setProperty("value", 0)
        self.progressBar.setObjectName("progressBar")
        self.gridLayout.addWidget(self.progressBar, 4, 0, 1, 2)
        self.verticalLayout = QtWidgets.QVBoxLayout()
        self.verticalLayout.setObjectName("verticalLayout")
        self.pushButton = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton.setObjectName("pushButton")
        self.pushButton.clicked.connect(self.open_sales_report)
        self.verticalLayout.addWidget(self.pushButton)
        self.gridLayout.addLayout(self.verticalLayout, 0, 1, 1, 1)
        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.pushButton_3 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_3.clicked.connect(self.open_report)
        self.pushButton_3.setObjectName("pushButton_3")
        self.verticalLayout_3.addWidget(self.pushButton_3)
        self.gridLayout.addLayout(self.verticalLayout_3, 2, 1, 1, 1)
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        self.pushButton_2 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_2.clicked.connect(self.open_stock)
        self.pushButton_2.setObjectName("pushButton_2")
        self.verticalLayout_2.addWidget(self.pushButton_2)
        self.gridLayout.addLayout(self.verticalLayout_2, 1, 1, 1, 1)
        self.textEdit_2 = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit_2.setObjectName("textEdit_2")
        self.gridLayout.addWidget(self.textEdit_2, 1, 0, 1, 1)
        self.textEdit = QtWidgets.QTextEdit(self.centralwidget)
        self.textEdit.setObjectName("textEdit")
        self.gridLayout.addWidget(self.textEdit, 0, 0, 1, 1)
        self.pushButton_4 = QtWidgets.QPushButton(self.centralwidget)
        self.pushButton_4.setObjectName("pushButton_4")
        self.gridLayout.addWidget(self.pushButton_4, 3, 0, 1, 2)
        self.pushButton_4.clicked.connect(self.write_to_report)
        Generator.setCentralWidget(self.centralwidget)
        self.statusbar = QtWidgets.QStatusBar(Generator)
        self.statusbar.setObjectName("statusbar")
        Generator.setStatusBar(self.statusbar)

        self.retranslateUi(Generator)
        QtCore.QMetaObject.connectSlotsByName(Generator)

    def retranslateUi(self, Generator):
        _translate = QtCore.QCoreApplication.translate
        Generator.setWindowTitle(_translate("Generator", "Генератор отчетов"))
        self.pushButton.setText(_translate("Generator", "Отчет по продажам"))
        self.pushButton_3.setText(_translate("Generator", "Отчет для записи"))
        self.pushButton_2.setText(_translate("Generator", "Остатки"))
        self.pushButton_4.setText(_translate("Generator", "Записать в файл"))


class Main(QMainWindow, Ui_Generator):
    def __init__(self):
        QMainWindow.__init__(self)
        self.setupUi(self)
        self.textEdit.setReadOnly(True)
        self.textEdit_2.setReadOnly(True)
        self.textEdit_3.setReadOnly(True)
        self.progressBar.setFormat('Индикатор выполнения программы')
        self.centerOnScreen()

    def centerOnScreen(self):
        resolution = QDesktopWidget().screenGeometry()
        self.move((resolution.width() / 2) - (self.frameSize().width() / 2),
                  (resolution.height() / 2) - (self.frameSize().height() / 2))

    def open_sales_report(self):
        self.textEdit.setText(str(QFileDialog.getOpenFileName(self,'Вложите отчет с продажами', "","Excel (*.xls *.xlsx)")[0]))

    def open_stock(self):
        self.textEdit_2.setText(str(QFileDialog.getOpenFileName(self, 'Вложите отчет с остатками',"","Excel (*.xls *.xlsx)")[0]))

    def open_report(self):
        self.textEdit_3.setText(str(QFileDialog.getOpenFileName(self,'Вложите отчет для записи' ,"","Excel (*.xls *.xlsx)")[0]))

    def write_to_report(self):
        realization = self.textEdit.toPlainText()
        stock = self.textEdit_2.toPlainText()
        report = self.textEdit_3.toPlainText()

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
            #try:
            parser = TradeParser(path_realization=realization, path_stock=stock, sales_report=report)
            percent = 0
            self.progressBar.setFormat('Идет выполнение программы, ждите')
            for i in range(100):
                percent += 1
                self.progressBar.setValue(percent)
            parser.write_to_excel()
            self.progressBar.setValue(0)
            self.progressBar.setFormat('Индикатор выполнения программы')
            success = QMessageBox()
            success.setIcon(QMessageBox.Information)
            success.setInformativeText("Загрузка прошла успешно!")
            success.setWindowTitle("Успех")
            success.setStandardButtons(QMessageBox.Ok)
            retval = success.exec_()
            '''except Exception as e:
            self.progressBar.setValue(0)
            self.progressBar.setFormat('Индикатор выполнения программы')
            errmsg = QMessageBox()
            errmsg.setIcon(QMessageBox.Warning)
            errmsg.setInformativeText(str(e))
            errmsg.setWindowTitle("Ошибка")
            errmsg.setStandardButtons(QMessageBox.Ok)
            retval = errmsg.exec_()'''

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Main()
    window.show()
    sys.exit(app.exec_())

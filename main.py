from PyQt5 import QtCore, QtGui, QtWidgets
from PyQt5.QtWidgets import QMessageBox


import easygui as eg
import xlrd
import xlsxwriter as xlwt
import sys

class AlignDelegate(QtWidgets.QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super(AlignDelegate, self).initStyleOption(option, index)
        option.displayAlignment = QtCore.Qt.AlignCenter

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("MainWindow")
        MainWindow.resize(741, 692)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.Load = QtWidgets.QPushButton(self.centralwidget)
        self.Load.setGeometry(QtCore.QRect(520, 120, 181, 71))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.Load.setFont(font)
        self.Load.setObjectName("Load")
        self.Save = QtWidgets.QPushButton(self.centralwidget)
        self.Save.setGeometry(QtCore.QRect(520, 210, 181, 71))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.Save.setFont(font)
        self.Save.setObjectName("Save")
        self.Save_txt = QtWidgets.QPushButton(self.centralwidget)
        self.Save_txt.setGeometry(QtCore.QRect(520, 490, 181, 71))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.Save_txt.setFont(font)
        self.Save_txt.setObjectName("Save_txt")
        self.Search = QtWidgets.QPushButton(self.centralwidget)
        self.Search.setGeometry(QtCore.QRect(520, 300, 181, 71))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.Search.setFont(font)
        self.Search.setObjectName("Search")
        self.tableWidget = QtWidgets.QTableWidget(self.centralwidget)
        self.tableWidget.setGeometry(QtCore.QRect(20, 110, 451, 271))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.tableWidget.setFont(font)
        self.tableWidget.setRowCount(0)
        self.tableWidget.setObjectName("tableWidget")
        self.tableWidget.setColumnCount(0)
        self.tableWidget.horizontalHeader().setDefaultSectionSize(67)
        self.tableWidget.verticalHeader().setSortIndicatorShown(False)
        self.plainTextEdit = QtWidgets.QPlainTextEdit(self.centralwidget)
        self.plainTextEdit.setGeometry(QtCore.QRect(40, 420, 421, 251))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.plainTextEdit.setFont(font)
        self.plainTextEdit.setObjectName("plainTextEdit")
        self.Control_sum = QtWidgets.QLineEdit(self.centralwidget)
        self.Control_sum.setGeometry(QtCore.QRect(20, 40, 131, 31))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.Control_sum.setFont(font)
        self.Control_sum.setObjectName("Control_sum")
        self.label = QtWidgets.QLabel(self.centralwidget)
        self.label.setGeometry(QtCore.QRect(20, 10, 211, 31))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.label.setFont(font)
        self.label.setObjectName("label")
        self.label_2 = QtWidgets.QLabel(self.centralwidget)
        self.label_2.setGeometry(QtCore.QRect(150, 80, 211, 31))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.label_2.setFont(font)
        self.label_2.setObjectName("label_2")
        self.label_3 = QtWidgets.QLabel(self.centralwidget)
        self.label_3.setGeometry(QtCore.QRect(200, 390, 81, 31))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.label_3.setFont(font)
        self.label_3.setObjectName("label_3")
        self.spinBoxRow = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBoxRow.setGeometry(QtCore.QRect(230, 40, 111, 31))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.spinBoxRow.setFont(font)
        self.spinBoxRow.setObjectName("spinBoxRow")
        self.label_4 = QtWidgets.QLabel(self.centralwidget)
        self.label_4.setGeometry(QtCore.QRect(230, 10, 211, 31))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.label_4.setFont(font)
        self.label_4.setObjectName("label_4")
        self.label_5 = QtWidgets.QLabel(self.centralwidget)
        self.label_5.setGeometry(QtCore.QRect(420, 10, 211, 31))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.label_5.setFont(font)
        self.label_5.setObjectName("label_5")
        self.spinBoxColumn = QtWidgets.QSpinBox(self.centralwidget)
        self.spinBoxColumn.setGeometry(QtCore.QRect(420, 40, 111, 31))
        font = QtGui.QFont()
        font.setPointSize(13)
        self.spinBoxColumn.setFont(font)
        self.spinBoxColumn.setObjectName("spinBoxColumn")
        MainWindow.setCentralWidget(self.centralwidget)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

        self.add_functions()

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "Lab_7"))
        self.Load.setText(_translate("MainWindow", "Load Excel"))
        self.Save.setText(_translate("MainWindow", "Save Excel"))
        self.Save_txt.setText(_translate("MainWindow", "Save_txt"))
        self.Search.setText(_translate("MainWindow", "Search"))
        self.label.setText(_translate("MainWindow", "Control Sum"))
        self.label_2.setText(_translate("MainWindow", "Matrix - Labyrinth"))
        self.label_3.setText(_translate("MainWindow", "Result"))
        self.label_4.setText(_translate("MainWindow", "Row Count"))
        self.label_5.setText(_translate("MainWindow", "Column Count"))

    def add_functions(self):
        self.Load.clicked.connect(self.load)
        self.Search.clicked.connect(self.search)
        self.Save.clicked.connect(self.save_xlsx)
        self.Save_txt.clicked.connect(self.save_txt)
        self.spinBoxRow.valueChanged.connect(self.row_in_table)
        self.spinBoxColumn.valueChanged.connect(self.column_in_table)

    def load(self):
        file = File()
        a = eg.fileopenbox()
        if a!=None:
            file.set_workbook(a)
            file.set_sheet(0)
            mass_from_file = file.from_file_to_prog()[0]
            rez_from_file = file.from_file_to_prog()[1]
            self.Control_sum.setText(str(rez_from_file))
            self.tableWidget.setRowCount(len(mass_from_file))
            self.tableWidget.setColumnCount(len(mass_from_file[0]))
            for i in range(len(mass_from_file)):
                for j in range(len(mass_from_file[0])):
                    self.tableWidget.setItem(i, j,QtWidgets.QTableWidgetItem(str(mass_from_file[i][j])))

            for i in range(self.tableWidget.rowCount()):
                delegate = AlignDelegate(self.tableWidget)
                self.tableWidget.setItemDelegateForRow(i, delegate)

            self.spinBoxRow.setValue(self.tableWidget.rowCount())
            self.spinBoxColumn.setValue(self.tableWidget.columnCount())

    def search(self):
        mas = []
        self.plainTextEdit.clear()
        rez.clear()
        path.clear()
        for i in range(self.tableWidget.rowCount()):
            path.append([])
            for j in range(self.tableWidget.columnCount()):
                path[i].append(int(0))

        alg = Algoritm()
        if self.Control_sum.text()=="":
            msb = QMessageBox()
            msb.setIcon(QMessageBox.Warning)
            msb.setStyleSheet("QLabel {font-size:20px;}")
            msb.setStyleSheet("QPushButton  {font-size:15px;}")
            msb.setText("Control sum fiels is empty")
            msb.exec_()
        else:
            control_sum=int(self.Control_sum.text())
            if self.tableWidget.rowCount()==0 or self.tableWidget.columnCount()==0:
                msb = QMessageBox()
                msb.setIcon(QMessageBox.Warning)
                msb.setStyleSheet("QLabel {font-size:20px;}")
                msb.setStyleSheet("QPushButton  {font-size:15px;}")
                msb.setText("Matrix is empty")
                msb.exec_()
            else:
                for i in range(self.tableWidget.rowCount()):
                    mas.append([])
                    for j in range(self.tableWidget.columnCount()):
                        if self.tableWidget.item(i,j)==None:
                            self.tableWidget.setItem(i,j,QtWidgets.QTableWidgetItem(self.Control_sum.text()))
                        mas[i].append(int(self.tableWidget.item(i,j).text()))
                alg.set_mass(mas)
                alg.set_final_rez(control_sum)
                alg.step(0, 0, 0)
                for i in range(len(rez)):
                    self.plainTextEdit.appendPlainText(rez[i])
                if self.plainTextEdit.toPlainText()=="":
                    msb = QMessageBox()
                    msb.setIcon(QMessageBox.Warning)
                    msb.setStyleSheet("QLabel {font-size:20px;}")
                    msb.setStyleSheet("QPushButton  {font-size:15px;}")
                    msb.setText("No solutions found")
                    msb.exec_()

    def save_txt(self):
        if self.plainTextEdit.toPlainText()!="":
            file = File()
            mass1=[]
            mass1.append(self.plainTextEdit.toPlainText())
            a = eg.filesavebox()
            s = a + ".txt"
            if a!=None:
                file.prog_to_file(s, mass1)
                msb = QMessageBox()
                msb.setIcon(QMessageBox.Information)
                msb.setStyleSheet("QLabel {font-size:20px;}")
                msb.setStyleSheet("QPushButton  {font-size:15px;}")
                msb.setText("File Save")
                msb.exec_()
        else:
            msb = QMessageBox()
            msb.setIcon(QMessageBox.Warning)
            msb.setStyleSheet("QLabel {font-size:20px;}")
            msb.setStyleSheet("QPushButton  {font-size:15px;}")
            msb.setText("plainTextEdit is empty")
            msb.exec_()

    def save_xlsx(self):
        if self.tableWidget.rowCount() == 0 or self.tableWidget.columnCount() == 0:
            msb = QMessageBox()
            msb.setIcon(QMessageBox.Warning)
            msb.setStyleSheet("QLabel {font-size:20px;}")
            msb.setStyleSheet("QPushButton  {font-size:15px;}")
            msb.setText("Matrix is empty")
            msb.exec_()
        elif self.Control_sum.text()=="":
            msb = QMessageBox()
            msb.setIcon(QMessageBox.Warning)
            msb.setStyleSheet("QLabel {font-size:20px;}")
            msb.setStyleSheet("QPushButton  {font-size:15px;}")
            msb.setText("Control sum fiels is empty")
            msb.exec_()
        else:
            s = eg.filesavebox()
            if s != None:
                book = xlwt.Workbook(s + ".xlsx")
                sheet = book.add_worksheet()
                sheet.write(0, 0,int(self.Control_sum.text()))
                for i in range(self.tableWidget.rowCount()):
                    for j in range(self.tableWidget.columnCount()):
                        value=self.tableWidget.item(i, j).text()
                        if (value == self.Control_sum.text() or value==""):
                            value=""
                            sheet.write(i+1, j+1, value)
                        else:
                            sheet.write(i + 1, j + 1, int(value))
                book.close()

                msb = QMessageBox()
                msb.setIcon(QMessageBox.Information)
                msb.setStyleSheet("QLabel {font-size:20px;}")
                msb.setStyleSheet("QPushButton  {font-size:15px;}")
                msb.setText("File Save")
                msb.exec_()

    def row_in_table(self):
        self.tableWidget.setRowCount(self.spinBoxRow.value())
        for i in range(self.tableWidget.rowCount()):
            delegate = AlignDelegate(self.tableWidget)
            self.tableWidget.setItemDelegateForRow(i, delegate)
    def column_in_table(self):
        self.tableWidget.setColumnCount(self.spinBoxColumn.value())
        for i in range(self.tableWidget.columnCount()):
            delegate = AlignDelegate(self.tableWidget)
            self.tableWidget.setItemDelegateForColumn(i, delegate)


rez = []
path = []
class File:

    def __init__(self):
        self.workbook = ""
        self.sheet = ""

    def get_workbook(self):
        return self.workbook

    def get_sheet(self):
        return self.sheet

    def set_workbook(self, wb):
        self.workbook = wb

    def set_sheet(self, sh):
        self.sheet = sh

    def from_file_to_prog(self):
        book = xlrd.open_workbook(self.get_workbook())
        sheet = book.sheet_by_index(self.get_sheet())
        mass_from_file = [[sheet.cell_value(r, c) for c in range(1, sheet.ncols, 1)] for r in range(1, sheet.nrows, 1)]
        fin_rez = int(sheet.cell_value(0, 0))

        for i in range(len(mass_from_file)):
            for j in range(len(mass_from_file[0])):
                if mass_from_file[i][j] == '':
                    mass_from_file[i][j] = fin_rez
                mass_from_file[i][j] = int(mass_from_file[i][j])
        return (mass_from_file, fin_rez)

    def prog_to_file(self, s, ms):
        f = open(s, 'w')
        for i in range(len(ms)):
            f.write(ms[i])
            f.write("\n\n")
        f.close()
class Algoritm:

    def __init__(self):
        self.mass = []
        self.final_rez = 0

    def get_mass(self):
        return self.mass

    def get_final_rez(self):
        return self.final_rez

    def set_mass(self, ms):
        self.mass = ms

    def set_final_rez(self, fr):
        self.final_rez = fr

    def __str__(self):
        vivod = ''
        for i in range(len(self.mass)):
            for j in range(len(self.mass[0])):
                if self.mass[i][j] == self.get_final_rez():
                    vivod += "X" + " "
                else:
                    vivod += str(self.mass[i][j]) + " "
            vivod += "\n"
        return (vivod)

    def step(self, x, y, sum):
        vivod = ""
        sum += self.mass[x][y]
        if sum > self.get_final_rez():
            return False
        if x == len(self.mass) - 1 and y == len(self.mass[0]) - 1:
            if (sum != self.get_final_rez()):
                return False
            path[x][y] = True
            for i in range(len(self.mass)):
                for j in range(len(self.mass[0])):
                    if path[i][j] == True:
                        vivod += str(self.mass[i][j]) + " "
                    else:
                        vivod += "-" + " "
                vivod += "\n"
            # print(vivod)
            rez.append(vivod)
        path[x][y] = True

        # up
        if y > 0 and not (path[x][y - 1]):
            self.step(x, y - 1, sum)

        # down
        if y < len(self.mass[0])-1 and not (path[x][y + 1]):
            self.step(x, y + 1, sum)

        # left
        if x > 0 and not (path[x - 1][y]):
            self.step(x - 1, y, sum)

        # right
        if x < len(self.mass)-1 and not (path[x + 1][y]):
            self.step(x + 1, y, sum)

        path[x][y] = False
        return False


app = QtWidgets.QApplication(sys.argv)
MainWindow = QtWidgets.QMainWindow()
ui =Ui_MainWindow()
ui.setupUi(MainWindow)
MainWindow.show()
sys.exit(app.exec_())
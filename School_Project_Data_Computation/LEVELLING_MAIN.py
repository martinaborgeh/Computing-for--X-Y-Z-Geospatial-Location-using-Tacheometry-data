import pandas as pd
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QHeaderView,QFileDialog,QMainWindow
from LEVELLING import Levelling_Computation
from Tacheometry_computation import Import_data_from_excel
import sys
import csv
from levelling_ui import Ui_MainWindow


class Levelling(QMainWindow,Ui_MainWindow):
    def __init__(self,parent=None):
        super(Levelling, self).__init__(parent)
        self.setupUi(self)
        self.tableWidget1.horizontalHeader().setStretchLastSection(True)
        self.tableWidget1.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget_2.horizontalHeader().setStretchLastSection(True)
        self.tableWidget_2.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.pushButton_3.clicked.connect(self.method)
        self.GHANA_NATIONAL_POINT_EXPORT_TO_EXCEL_BUTTON.clicked.connect(self.export)
        self.pushButton.clicked.connect(self.clear_input_table)
        self.pushButton_2.clicked.connect(self.clear_output_table)
        self.pushButton_4.clicked.connect(self.Import_Input_data)
        self.EXPORT_WGS_POINT_TO_EXCEL_BUTTON.clicked.connect(self.Export_Input_Data)
    def Retrieving_data(self):
        POINT = []
        BS = []
        IS = []
        FS = []
        RMKS = []
        rowcount=self.tableWidget1.rowCount()
        for row in range(1,rowcount):
            try:
                a=self.tableWidget1.takeItem(row, 0).text()
                self.tableWidget1.setItem(row, 0, QtWidgets.QTableWidgetItem(a))
                BS.append(a)
            except Exception:
                BS.append('')
            try:
                b=self.tableWidget1.takeItem(row, 1).text()
                self.tableWidget1.setItem(row, 1,
                                         QtWidgets.QTableWidgetItem(b))
                IS.append(b)
            except Exception:
                IS.append('')
            try:
                c=self.tableWidget1.takeItem(row, 2).text()
                self.tableWidget1.setItem(row, 2,
                                         QtWidgets.QTableWidgetItem(c))
                FS.append(c)
            except Exception:
                FS.append('')
            try:
                d=self.tableWidget1.takeItem(row, 3).text()
                self.tableWidget1.setItem(row, 3,
                                         QtWidgets.QTableWidgetItem(d))
                RMKS.append(d)
            except Exception:
                RMKS.append('')


        for data in zip(BS,IS,FS,RMKS):
            if data[0]!='' or data[1]!='' or data[2]!='' or data[3]!='':
                POINT.append(data)
        # print(BS)
        # print(POINT)
        return POINT
    def Calcul(self):
        pointdata=self.Retrieving_data()
        initial_elevation=self.lineEdit.text()
        elevation = Levelling_Computation(pointdata, initial_elevation)
        i=1
        for result in zip(pointdata,elevation):
                if result[0][0]=='':
                    self.tableWidget_2.setItem(i, 0,
                                               QtWidgets.QTableWidgetItem(result[0][0]))

                else:
                    self.tableWidget_2.setItem(i, 0,
                                               QtWidgets.QTableWidgetItem(str(result[0][0])))

                if result[0][1]=='':
                    self.tableWidget_2.setItem(i, 1,
                                               QtWidgets.QTableWidgetItem(result[0][1]))
                else:
                    self.tableWidget_2.setItem(i, 1,
                                               QtWidgets.QTableWidgetItem(str(result[0][1])))

                if result[0][2]=='':
                    self.tableWidget_2.setItem(i, 2,
                                               QtWidgets.QTableWidgetItem(result[0][2]))
                else:
                    self.tableWidget_2.setItem(i, 2,
                                               QtWidgets.QTableWidgetItem(str(result[0][2])))

                if result[1]=='':
                    self.tableWidget_2.setItem(i, 3,
                                               QtWidgets.QTableWidgetItem(result[1]))
                else:
                    self.tableWidget_2.setItem(i, 3,
                                               QtWidgets.QTableWidgetItem(str(result[1])))

                if result[0][3]=='':
                    self.tableWidget_2.setItem(i, 4,
                                               QtWidgets.QTableWidgetItem(result[0][3]))
                else:
                    self.tableWidget_2.setItem(i, 4,
                                               QtWidgets.QTableWidgetItem(str(result[0][3])))

                i+=1

    def method(self):
        return self.Calcul()

    def export(self):
        POINT1 = []
        BS1 = []
        IS1 = []
        FS1 = []
        RMKS1 = []
        RL = []
        rowcount = self.tableWidget_2.rowCount()
        for row in range(1, rowcount):
            try:
                try:
                    a = self.tableWidget_2.takeItem(row, 0).text()
                    self.tableWidget_2.setItem(row, 0, QtWidgets.QTableWidgetItem(a))
                    BS1.append(a)
                except Exception:
                    BS1.append('')
                try:
                    b = self.tableWidget_2.takeItem(row, 1).text()
                    self.tableWidget_2.setItem(row, 1,
                                              QtWidgets.QTableWidgetItem(b))
                    IS1.append(b)
                except Exception:
                    IS1.append('')
                try:
                    c = self.tableWidget_2.takeItem(row, 2).text()
                    self.tableWidget_2.setItem(row, 2,
                                              QtWidgets.QTableWidgetItem(c))
                    FS1.append(c)
                except Exception:
                    FS1.append('')
                try:
                    d = self.tableWidget_2.takeItem(row, 3).text()
                    self.tableWidget_2.setItem(row, 3,
                                              QtWidgets.QTableWidgetItem(d))
                    RL.append(d)
                except Exception:
                    RL.append('')
                try:
                    d = self.tableWidget_2.takeItem(row, 4).text()
                    self.tableWidget_2.setItem(row, 4,
                                              QtWidgets.QTableWidgetItem(d))
                    RMKS1.append(d)
                except Exception:
                    RMKS1.append('')

            except Exception as e:
                print('the error is', e)
                continue

        for data in zip(BS1, IS1, FS1,RL, RMKS1):
            if data[0] != '' or data[1] != '' or data[2] != '' or data[3] != '' or data[4]!='':
                POINT1.append(data)
        print(POINT1)
        Save, check1 = QFileDialog.getSaveFileName(None, "QFileDialog.getOSaveFileNames()", "", "All Files(*)")
        link1 = None
        if check1:
            link1 = f"{Save}{'.csv'}"
            bs = [b[0] for b in POINT1]
            Is = [iS[1] for iS in POINT1]
            fs = [f[2] for f in POINT1]
            rl = [r[3] for r in POINT1]
            rmks = [rm[4] for rm in POINT1]
            pd.DataFrame({'BACKSIGHT': bs, 'INTERSIGHT': Is, 'FORESIGHT': fs, 'REDUCE\
            LEVEL': rl, 'REMARKS': rmks}).to_csv(link1, index=False)
    def Import_Input_data(self):
        Save, check1 = QFileDialog.getOpenFileName(None, "QFileDialog.getOpenFileName()", "", "All Files(*)")
        link1 = None
        if check1:
            link1 = Save
            from_excel = Import_data_from_excel(link1)
            row1 = 1
            for person in from_excel:
                self.tableWidget1.setItem(row1, 0, QtWidgets.QTableWidgetItem(str(person[0])))
                self.tableWidget1.setItem(row1, 1, QtWidgets.QTableWidgetItem(str(person[1])))
                self.tableWidget1.setItem(row1, 2, QtWidgets.QTableWidgetItem(str(person[2])))
                self.tableWidget1.setItem(row1, 3, QtWidgets.QTableWidgetItem(str(person[3])))
                row1+=1
    def Export_Input_Data(self):
        data=self.Retrieving_data()
        Save, check1 = QFileDialog.getSaveFileName(None, "QFileDialog.getOSaveFileNames()", "", "All Files(*)")
        link1 = None
        if check1:
            link1 = f"{Save}{'.csv'}"
            with open(link1,'w',newline='') as datafile:
                write=csv.writer(datafile)
                write.writerow(("BS","IS","FS","RMKS"))
                for item in data:
                    write.writerow(item)




    def clear_output_table(self):
        self.tableWidget_2.clearContents()

    def clear_input_table(self):
        self.tableWidget1.clearContents()

if __name__=='__main__':
    def my_exception(type,value,tback):
        print(tback,value,tback)
        sys.__excepthook__(type,value,tback)
    sys.excepthook=my_exception
    app = QtWidgets.QApplication(sys.argv)
    window = Levelling()
    app.exec_()

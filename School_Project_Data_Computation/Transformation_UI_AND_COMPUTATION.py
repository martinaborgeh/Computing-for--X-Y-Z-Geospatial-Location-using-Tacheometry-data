import numpy as np
import pandas as pd
from PyQt5.QtWidgets import QApplication,QFileDialog,QHeaderView
import sys
from PyQt5 import QtWidgets
from datum_ui import Ui_MainWindow
class Transformation(QtWidgets.QMainWindow,Ui_MainWindow):
    def __init__(self,parent=None):
        super(Transformation,self).__init__(parent)
        self.setupUi(self)
        self.tableWidget.horizontalHeader().setStretchLastSection(True)
        self.tableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget_2.horizontalHeader().setStretchLastSection(True)
        self.tableWidget_2.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.GHANA_NATIONAL_POINT_EXPORT_TO_EXCEL_BUTTON.clicked.connect(self.Grid_excel_output)
        self.EXPORT_WGS_POINT_TO_EXCEL_BUTTON.clicked.connect(self.WGS_excel_output)
        self.COORDINATE_CONVERSION_BUTTON.clicked.connect(self.insert_data)
        self.pushButton_2.clicked.connect(self.Clear_Grid_Table)
        self.pushButton.clicked.connect(self.Clear_WGS_Table)
        self.show()



    #retrieving Grid Data
    def Grid_retrieving_data(self):
        rowcount = self.ui_interface.tableWidget.rowCount()
        Lat = []
        Long = []
        CorN = []
        CorE = []
        EN = []
        EE = []
        for row in range(rowcount):
            try:

                CorNitem = self.tableWidget_2.takeItem(row, 0).text()
                self.tableWidget_2.setItem(row, 0, QtWidgets.QTableWidgetItem(CorNitem))
                CorN.append(float(CorNitem))

                CorEitem= self.tableWidget_2.takeItem(row, 1).text()
                self.tableWidget_2.setItem(row, 1, QtWidgets.QTableWidgetItem(CorEitem))
                CorE.append(float(CorEitem))

                Latitem = self.tableWidget_2.takeItem(row, 2).text()
                self.tableWidget_2.setItem(row, 2, QtWidgets.QTableWidgetItem(Latitem))
                Lat.append(float(Latitem))

                Longitem = self.tableWidget_2.takeItem(row, 3).text()
                self.tableWidget_2.setItem(row, 3, QtWidgets.QTableWidgetItem(Longitem))
                Long.append(float(Longitem))
            except Exception:
                continue

        for row in range(rowcount):
            try:
                ENitem = self.tableWidget_2.takeItem(row, 4).text()
                self.ui_interface.tableWidget_2.setItem(row, 4, QtWidgets.QTableWidgetItem(ENitem))
                EN.append(float(ENitem))

                EEitem = self.tableWidget_2.takeItem(row, 5).text()
                self.tableWidget_2.setItem(row, 5, QtWidgets.QTableWidgetItem(EEitem))
                EE.append(float(EEitem))
            except Exception:
                continue

        d = []
        for i in zip(Long, Lat):
            if self.GHANA_NATIONAL_UNIT_COMBOBOX.currentText() == 'GRID_METERS':
                d.append([float(i[0]), float(i[1]), 1, 0])
                d.append([float(i[1]), -float(i[0]), 0, 1])
            elif self.GHANA_NATIONAL_UNIT_COMBOBOX.currentText() == 'GRID_FEETS':
                d.append([float(i[0]) * 0.304799706846218, float(i[1]) / 0.304799706846218, 1, 0])
                d.append([float(i[1]) * 0.304799706846218, -float(i[0]) / 0.304799706846218, 0, 1])
        e = np.array(d)
        zipped_E_N = list(zip(CorE, CorN))
        grid = np.array([float(value) for value in list(np.concatenate(zipped_E_N).flat)])

        try:
            parm = np.linalg.inv(e).dot(grid)
            X = [(parm[0]*(float(j[0])-parm[2]))-(parm[1]*(float(j[1])-parm[3]))/(pow(parm[0],2)+pow(parm[1],2)) for j in zip(EE, EN)]
            Y = [(parm[1]*(float(j[0])-parm[2]))-(parm[0]*(float(j[1])-parm[3]))/(pow(parm[0],2)+pow(parm[1],2)) for j in zip(EE, EN)]
        except Exception as e:
            pass
        return list(zip(Y, X))

    def WGS_excel_output(self):
        Save,check1 =QFileDialog.getSaveFileName(None,"QFileDialog.getOSaveFileNames()","","All Files(*)")
        link1  = None
        if check1:
            link1=f"{Save}{'.csv'}"
        pd.DataFrame({'Latitude':[n[0] for n in self.Grid_retrieving_data()],'Long':[e[1] for e in self.Grid_retrieving_data()]}).to_csv(link1)
    # Retrieving WGS data
    def WGS_retrieving_data(self):
        rowcount=self.tableWidget.rowCount()
        Lat=[]
        Long=[]
        CorN=[]
        CorE=[]
        ELat=[]
        ELong=[]
        for row in range(rowcount):
            try:

                Latitem=self.tableWidget.takeItem(row, 0).text()
                self.tableWidget.setItem(row, 0, QtWidgets.QTableWidgetItem(Latitem))
                Lat.append(float(Latitem))

                Longitem=self.tableWidget.takeItem(row, 1).text()
                self.tableWidget.setItem(row, 1, QtWidgets.QTableWidgetItem(Longitem))
                Long.append(float(Longitem))

                CorNitem = self.tableWidget.takeItem(row, 2).text()
                self.tableWidget.setItem(row, 2, QtWidgets.QTableWidgetItem(CorNitem))
                CorN.append(float(CorNitem))

                CorEitem = self.tableWidget.takeItem(row, 3).text()
                self.tableWidget.setItem(row, 3, QtWidgets.QTableWidgetItem(CorEitem))
                CorE.append(float(CorEitem))
            except Exception:
                continue

        for row in range(rowcount):
            try:
                ELatitem = self.tableWidget.takeItem(row, 4).text()
                self.tableWidget.setItem(row, 4, QtWidgets.QTableWidgetItem(ELatitem))
                ELat.append(float(ELatitem))

                ELongitem = self.tableWidget.takeItem(row, 5).text()
                self.tableWidget.setItem(row, 5, QtWidgets.QTableWidgetItem(ELongitem))
                ELong.append(float(ELongitem))
            except Exception:
                continue

        d=[]
        for i in zip(Long,Lat):
            if self.WGS_COMBO_BOX.currentText()=='UTM_METERS':
                d.append([float(i[0]),float(i[1]),1,0])
                d.append([float(i[1]),-float(i[0]),0,1])
            elif self.WGS_COMBO_BOX.currentText()=='UTM_FEETS':
                d.append([float(i[0])/0.304799706846218, float(i[1])/0.304799706846218, 1, 0])
                d.append([float(i[1])/0.304799706846218, -float(i[0])/0.304799706846218, 0, 1])
        e=np.array(d)

        zipped_E_N=list(zip(CorE,CorN))
        grid=np.array([float(value) for value in list(np.concatenate(zipped_E_N).flat)])
        try:
            parm=np.linalg.inv(e).dot(grid)
            X=[(parm[0]*float(j[0]))+(parm[1]*float(j[1]))+parm[2] for j in zip(ELong,ELat)]
            Y=[(-parm[1]*float(j[0]))+(parm[0]*float(j[1]))+parm[3] for j in zip(ELong,ELat)]
            return list(zip(Y, X))
        except Exception as e:
            pass


    # Exporting Computed grid coordinates to excel
    def Grid_excel_output(self):
        retrive=self.WGS_retrieving_data()
        Save,check1 =QFileDialog.getSaveFileName(None,"QFileDialog.getOSaveFileNames()","","All Files(*)")
        link1  = None
        if check1:
            link1=f"{Save}{'.csv'}"
        pd.DataFrame({'N':[n[0] for n in retrive],'E':[e[1] for e in retrive]}).to_csv(link1)
    def Clear_Grid_Table(self):
        self.tableWidget_2.clearContents()
    def Clear_WGS_Table(self):
        self.tableWidget.clearContents()

    #inserting_data into column
    def insert_data(self):
            if self.COORDINATE_CONVERSION_COMMBO_BOX.currentText()=='GHANA NATIONAL  GRID':
                retreived=self.WGS_retrieving_data()
                row1=1
                try:
                    for person in retreived:
                        self.tableWidget_2.setItem(row1, 0, QtWidgets.QTableWidgetItem(str(person[0])))
                        self.tableWidget_2.setItem(row1, 1, QtWidgets.QTableWidgetItem(str(person[1])))
                        row1 += 1
                except Exception as e:
                    print(e)
            else:
                    retreived1 = self.Grid_retrieving_data()

                    row = 1
                    try:
                        for item in retreived1:
                            self.tableWidget.setItem(row, 0, QtWidgets.QTableWidgetItem(str(item[0])))
                            self.tableWidget.setItem(row, 1, QtWidgets.QTableWidgetItem(str(item[1])))
                            row += 1
                    except Exception as e:
                        print(e)
if __name__=='__main__':
    def my_exception(type,value,tback):
        print(tback,value,tback)
        sys.__excepthook__(type,value,tback)
    sys.excepthook=my_exception
    app=QApplication(sys.argv)
    window=Transformation()
    app.exec_()








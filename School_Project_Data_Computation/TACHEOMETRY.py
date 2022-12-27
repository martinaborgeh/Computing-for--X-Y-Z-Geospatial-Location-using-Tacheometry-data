from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMainWindow,QHeaderView,QFileDialog
import sys
from Tacheometry_computation import *
import csv
from tacheometry_ui import Ui_MainWindow

class Tacheometry(QMainWindow,Ui_MainWindow):
    def __init__(self,parent=None):
        super(Tacheometry, self).__init__(parent)
        self.setupUi(self)
        self.tableWidget1.horizontalHeader().setStretchLastSection(True)
        self.tableWidget1.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.tableWidget_2.horizontalHeader().setStretchLastSection(True)
        self.tableWidget_2.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.pushButton_3.clicked.connect(self.compute)
        self.GHANA_NATIONAL_POINT_EXPORT_TO_EXCEL_BUTTON.clicked.connect(self.Export_Output_Data)
        self.pushButton.clicked.connect(self.Export_Input_Table_data)
        self.pushButton_4.clicked.connect(self.Import_Input_Table)
        self.pushButton_5.clicked.connect(self.Clear_Input_data)
        self.pushButton_2.clicked.connect(self.Clear_Output_data)
        self.pushButton_6.clicked.connect(self.units_conversion)


        # self.centralwidget.setLayout(self.setWindowFlags(QtCore.Qt.WindowMaximizeButtonHint))
    def Retrieving_data(self):
        retrieved_data = []
        INST_STN = []
        BS=[]
        FS=[]
        IH = []
        HCR_LL = []
        HCR_RR = []
        VCR_LL=[]
        VCR_RR=[]
        HD_LL=[]
        HD_RR=[]
        VD_LL=[]
        VD_RR=[]
        CP=[]
        PH=[]

        rowcount = self.tableWidget1.rowCount()
        for row in range(1, rowcount):
            try:
                a = self.tableWidget1.takeItem(row, 0).text()
                self.tableWidget1.setItem(row, 0, QtWidgets.QTableWidgetItem(a))
                INST_STN.append(a)
            except Exception:
                INST_STN.append('')
            try:
                a = self.tableWidget1.takeItem(row, 1).text()
                self.tableWidget1.setItem(row, 1, QtWidgets.QTableWidgetItem(a))
                IH.append(a)
            except Exception:
                IH.append('')
            try:
                a = self.tableWidget1.takeItem(row, 2).text()
                self.tableWidget1.setItem(row, 2, QtWidgets.QTableWidgetItem(a))
                BS.append(a)
            except Exception:
                BS.append('')
            try:
                b = self.tableWidget1.takeItem(row, 3).text()
                self.tableWidget1.setItem(row, 3,
                                          QtWidgets.QTableWidgetItem(b))
                FS.append(b)
            except Exception:
                FS.append('')
            try:
                c = self.tableWidget1.takeItem(row, 4).text()
                self.tableWidget1.setItem(row, 4,
                                          QtWidgets.QTableWidgetItem(c))
                HCR_LL.append(c)
            except Exception:
                HCR_LL.append('')
            try:
                d = self.tableWidget1.takeItem(row, 5).text()
                self.tableWidget1.setItem(row, 5,
                                          QtWidgets.QTableWidgetItem(d))
                HCR_RR.append(d)
            except Exception:
                HCR_RR.append('')

            try:
                d = self.tableWidget1.takeItem(row, 6).text()
                self.tableWidget1.setItem(row, 6,
                                          QtWidgets.QTableWidgetItem(d))
                VCR_LL.append(d)
            except Exception:
                VCR_LL.append('')
            try:
                d = self.tableWidget1.takeItem(row, 7).text()
                self.tableWidget1.setItem(row, 7,
                                          QtWidgets.QTableWidgetItem(d))
                VCR_RR.append(d)
            except Exception:
                VCR_RR.append('')
            try:
                d = self.tableWidget1.takeItem(row, 8).text()
                self.tableWidget1.setItem(row, 8,
                                          QtWidgets.QTableWidgetItem(d))
                HD_LL.append(d)
            except Exception:
                HD_LL.append('')
            try:
                d = self.tableWidget1.takeItem(row, 9).text()
                self.tableWidget1.setItem(row, 9,
                                          QtWidgets.QTableWidgetItem(d))
                HD_RR.append(d)
            except Exception:
                HD_RR.append('')

            try:
                d = self.tableWidget1.takeItem(row, 10).text()
                self.tableWidget1.setItem(row, 10,
                                          QtWidgets.QTableWidgetItem(d))
                VD_LL.append(d)
            except Exception:
                VD_LL.append('')
            try:
                d = self.tableWidget1.takeItem(row, 11).text()
                self.tableWidget1.setItem(row, 11,
                                          QtWidgets.QTableWidgetItem(d))
                VD_RR.append(d)
            except Exception:
                VD_RR.append('')
            try:
                d = self.tableWidget1.takeItem(row, 12).text()
                self.tableWidget1.setItem(row, 12,
                                          QtWidgets.QTableWidgetItem(d))
                CP.append(d)
            except Exception:
                CP.append('')
            try:
                d = self.tableWidget1.takeItem(row, 13).text()
                self.tableWidget1.setItem(row, 13,
                                          QtWidgets.QTableWidgetItem(d))
                PH.append(d)
            except Exception:
                PH.append('')


        for data in zip(INST_STN,BS,FS,IH, HCR_LL, HCR_RR,VCR_LL,VCR_RR,HD_LL,HD_RR,VD_LL,VD_RR,CP,PH):
            if data[0] != '' or data[1] != '' or data[2] != '' or data[3] != '' or data[4] != '' or data[5] != '' or data[6] != ''\
                   or data[7] != '' or data[8] != '' or data[9] != '' or data[10] != '' or data[11] != '' or data[12] != '' or data[13] != '':
                retrieved_data.append(data)
        return retrieved_data
    def compute(self):
        try:
            Table_data=self.Retrieving_data()
            initial_Bearing=self.lineEdit_3.text()
            inst_stn_initial_Northings=self.lineEdit_4.text()
            inst_stn_initial_Eastings=self.lineEdit_2.text()
            inst_stn_initial_elevation=self.lineEdit.text()
            Backstion_initial_Northings=self.lineEdit_5.text()
            Backstion_initial_Eastings=self.lineEdit_7.text()
            Backstion_initial_Elevation=self.lineEdit_6.text()

            # print(Table_data)
            Computed_Points=Tacheo_Traver(Table_data,initial_Bearing,inst_stn_initial_Northings,inst_stn_initial_Eastings,Backstion_initial_Northings,Backstion_initial_Eastings)
            Elevation=Tacheo_Heightning(Table_data,inst_stn_initial_elevation,Backstion_initial_Elevation)
            #
            def Outputdata(xy_data, elevation):
                output = []
                for data in zip(xy_data, elevation):
                    output.append((str(data[0][0]), str(data[0][1]),str(data[1][2]), str(data[0][2]),str(data[1][0]), str(data[0][3]), str(data[0][4]), str(data[0][5]), str(data[0][6]),
                                           str(data[0][7]),str(data[1][1]), str(data[0][8])))

                return output




            computed_values = Outputdata(Computed_Points, Elevation)
            row1 = 1
            for person in computed_values:
                self.tableWidget_2.setItem(row1, 0, QtWidgets.QTableWidgetItem(str(person[0])))
                self.tableWidget_2.setItem(row1, 1, QtWidgets.QTableWidgetItem(str(person[1])))
                self.tableWidget_2.setItem(row1, 2, QtWidgets.QTableWidgetItem(str(person[2])))
                self.tableWidget_2.setItem(row1, 3, QtWidgets.QTableWidgetItem(str(person[3])))
                self.tableWidget_2.setItem(row1, 4, QtWidgets.QTableWidgetItem(str(person[4])))
                self.tableWidget_2.setItem(row1, 5, QtWidgets.QTableWidgetItem(str(person[5])))
                self.tableWidget_2.setItem(row1, 6, QtWidgets.QTableWidgetItem(str(person[6])))
                self.tableWidget_2.setItem(row1, 7, QtWidgets.QTableWidgetItem(str(person[7])))
                self.tableWidget_2.setItem(row1, 8, QtWidgets.QTableWidgetItem(str(person[8])))
                self.tableWidget_2.setItem(row1, 9, QtWidgets.QTableWidgetItem(str(person[9])))
                self.tableWidget_2.setItem(row1, 10, QtWidgets.QTableWidgetItem(str(person[10])))
                self.tableWidget_2.setItem(row1, 11, QtWidgets.QTableWidgetItem(str(person[11])))
                row1 += 1
            widget = self.label_13
            widget.setText('Computed Successfully')
            return Outputdata(Computed_Points, Elevation)
        except BaseException as e:
            self.label_13.setText(f'{e}')

    def Export_Output_Data(self):
        dataouput=self.compute()
        dataouput1=[]
        bearing_deg_min_seconds=Convert_decimal_Bearing_to_degree(dataouput)
        HIA_deg_min_seconds=Convert_decimal_HCR_ANGLE_to_degree(dataouput)
        VERT_ANGLE_deg_min_seconds=Convert_decimal_VCR_ANGLE_to_degree(dataouput)


        for data,deg_Bear,deg_HCR,deg_VCR in zip(dataouput, bearing_deg_min_seconds,HIA_deg_min_seconds,VERT_ANGLE_deg_min_seconds):
            to_list=list(data)
            to_list.pop(1)
            to_list.pop(1)
            to_list.pop(3)
            to_list.insert(3,deg_Bear[0])
            to_list.insert(4,deg_Bear[1])
            to_list.insert(5,deg_Bear[2])
            to_list.insert(6,deg_HCR[0])
            to_list.insert(7,deg_HCR[1])
            to_list.insert(8,deg_HCR[2])
            to_list.insert(9,deg_VCR[0])
            to_list.insert(10,deg_VCR[1])
            to_list.insert(11,deg_VCR[2])
            dataouput1.append(to_list)
        Save, check1 = QFileDialog.getSaveFileName(None, "QFileDialog.getOSaveFileNames()", "", "All Files(*)")
        link1 = None
        if check1:
            link1 = f"{Save}{'.csv'}"
        with open(link1, "w", newline='') as filename:
            write = csv.writer(filename)
            write.writerow(("STN", "HD", "VD", "B(DEG)", "B(MIN)", "B(SEC)", "HIA(DEG)", "HIA(MIN)", \
                            "HIA(SEC)", "VERT_ANGLE(DEG)", "VERT_ANGLE(MIN)", "VERT_ANGLE(SEC)", \
                            "LAT", "DEP", "N", "E", "ELEV", "PN"))
            for item in dataouput1:
                write.writerow(item)
        return dataouput1
    def Export_Input_Table_data(self):
        input_data=self.Retrieving_data()
        input_data1=self.Retrieving_data()
        Save, check1 = QFileDialog.getSaveFileName(None, "QFileDialog.getOSaveFileNames()", "", "All Files(*)")
        link1 = None
        if check1:
            link1 = f"{Save}{'.csv'}"
        with open(link1, "w", newline='') as filename:
            write = csv.writer(filename)
            write.writerow(("INST_STN","IH","BS","FS","HCR(LL)","HCR(RR)","VCR(LL)","VCR(RR)","HD(LL)","HD(RR)","VD(LL)","VD(RR)","CP","PH"))
            for data,data1 in zip(input_data,input_data1):
                to_list=list(data)
                to_list1=list(data1)
                to_list[1]=to_list[3]
                to_list[3]=to_list[2]
                to_list[2]=to_list1[1]
                write.writerow(to_list)
    def Import_Input_Table(self):
        Save, check1 = QFileDialog.getOpenFileName(None, "QFileDialog.getOpenFileName()", "", "All Files(*)")
        link1 = None
        if check1:
            link1 = Save
            data = Import_data_from_excel(link1)
            row1 = 1
            for person in data:
                self.tableWidget1.setItem(row1, 0, QtWidgets.QTableWidgetItem(str(person[0])))
                self.tableWidget1.setItem(row1, 1, QtWidgets.QTableWidgetItem(str(person[1])))
                self.tableWidget1.setItem(row1, 2, QtWidgets.QTableWidgetItem(str(person[2])))
                self.tableWidget1.setItem(row1, 3, QtWidgets.QTableWidgetItem(str(person[3])))
                self.tableWidget1.setItem(row1, 4, QtWidgets.QTableWidgetItem(str(person[4])))
                self.tableWidget1.setItem(row1, 5, QtWidgets.QTableWidgetItem(str(person[5])))
                self.tableWidget1.setItem(row1, 6, QtWidgets.QTableWidgetItem(str(person[6])))
                self.tableWidget1.setItem(row1, 7, QtWidgets.QTableWidgetItem(str(person[7])))
                self.tableWidget1.setItem(row1, 8, QtWidgets.QTableWidgetItem(str(person[8])))
                self.tableWidget1.setItem(row1, 9, QtWidgets.QTableWidgetItem(str(person[9])))
                self.tableWidget1.setItem(row1, 10, QtWidgets.QTableWidgetItem(str(person[10])))
                self.tableWidget1.setItem(row1, 11, QtWidgets.QTableWidgetItem(str(person[11])))
                self.tableWidget1.setItem(row1, 12, QtWidgets.QTableWidgetItem(str(person[12])))
                self.tableWidget1.setItem(row1, 13, QtWidgets.QTableWidgetItem(str(person[13])))
                row1 += 1
    def units_conversion(self):
        row1=1
        Table_Data=self.Retrieving_data()
        for data in Table_Data:
            if self.comboBox_3.currentText()=='FEETS':
                if self.comboBox.currentText()=='IH':
                    if data[3]=='':
                        self.tableWidget1.setItem(row1, 1, QtWidgets.QTableWidgetItem(''))
                    elif data[3]!='':
                        value=round(float(data[3])/0.3048,3)
                        self.tableWidget1.setItem(row1, 1, QtWidgets.QTableWidgetItem(str(value)))
                elif self.comboBox.currentText()=='HCR_LL':
                    if data[4]=='':
                        self.tableWidget1.setItem(row1, 4, QtWidgets.QTableWidgetItem(''))
                    elif data[4]!='':
                        value1=float(data[4])/0.3048
                        self.tableWidget1.setItem(row1, 4, QtWidgets.QTableWidgetItem(str(value1)))
                elif self.comboBox.currentText()=='HCR_RR':
                    if data[5]=='':
                        self.tableWidget1.setItem(row1, 5, QtWidgets.QTableWidgetItem(''))
                    elif data[5]!='':
                        value2=float(data[5])/0.3048
                        self.tableWidget1.setItem(row1, 5, QtWidgets.QTableWidgetItem(str(value2)))
                elif self.comboBox.currentText()=='VCR_LL':
                    if data[6]=='':
                        self.tableWidget1.setItem(row1, 6, QtWidgets.QTableWidgetItem(''))
                    elif data[6]!='':
                        value3=float(data[6])/0.3048
                        self.tableWidget1.setItem(row1, 6, QtWidgets.QTableWidgetItem(str(value3)))
                elif self.comboBox.currentText()=='VCR_RR':
                    if data[7] == '':
                        self.tableWidget1.setItem(row1, 7, QtWidgets.QTableWidgetItem(''))
                    elif data[7] != '':
                        value4 = float(data[7]) / 0.3048
                        self.tableWidget1.setItem(row1, 7, QtWidgets.QTableWidgetItem(str(value4)))
                elif self.comboBox.currentText()=='HD_LL':
                    if data[8]=='':
                        self.tableWidget1.setItem(row1, 8, QtWidgets.QTableWidgetItem(''))
                    elif data[8]!='':
                        value5=round(float(data[8])/0.3048,3)
                        self.tableWidget1.setItem(row1, 8, QtWidgets.QTableWidgetItem(str(value5)))
                elif self.comboBox.currentText()=='HD_RR':
                    if data[9]=='':
                        self.tableWidget1.setItem(row1, 9, QtWidgets.QTableWidgetItem(''))
                    elif data[9]!='':
                        value6=round(float(data[9])/0.3048,3)
                        self.tableWidget1.setItem(row1, 9, QtWidgets.QTableWidgetItem(str(value6)))
                elif self.comboBox.currentText()=='VD_LL':
                    if data[10] == '':
                        self.tableWidget1.setItem(row1, 10, QtWidgets.QTableWidgetItem(''))
                    elif data[10] != '':
                        value7 = round(float(data[10]) / 0.3048,3)
                        self.tableWidget1.setItem(row1, 10, QtWidgets.QTableWidgetItem(str(value7)))
                elif self.comboBox.currentText()=='VD_RR':
                    if data[11]=='':
                        self.tableWidget1.setItem(row1, 11, QtWidgets.QTableWidgetItem(''))
                    elif data[11]!='':
                        value8=round(float(data[11])/0.3048,)
                        self.tableWidget1.setItem(row1, 11, QtWidgets.QTableWidgetItem(str(value8)))
                elif self.comboBox.currentText() == 'PH':
                    if data[13] == '':
                        self.tableWidget1.setItem(row1, 13, QtWidgets.QTableWidgetItem(''))
                    elif data[13] != '':
                        value8 = round(float(data[13]) / 0.3048,3)
                        self.tableWidget1.setItem(row1, 13, QtWidgets.QTableWidgetItem(str(value8)))

            elif self.comboBox_3.currentText()== 'METRES':
                if self.comboBox.currentText()=='IH':
                    if data[3]=='':
                        self.tableWidget1.setItem(row1, 1, QtWidgets.QTableWidgetItem(''))
                    elif data[3]!='':
                        value9=round(float(data[3])*0.3048,3)
                        self.tableWidget1.setItem(row1, 1, QtWidgets.QTableWidgetItem(str(value9)))
                elif self.comboBox.currentText()=='HCR_LL':
                    if data[4]=='':
                        self.tableWidget1.setItem(row1, 4, QtWidgets.QTableWidgetItem(''))
                    elif data[4]!='':
                        value10=float(data[4])*0.3048
                        self.tableWidget1.setItem(row1, 4, QtWidgets.QTableWidgetItem(str(value10)))
                elif self.comboBox.currentText()=='HCR_RR':
                    if data[5]=='':
                        self.tableWidget1.setItem(row1, 5, QtWidgets.QTableWidgetItem(''))
                    elif data[5]!='':
                        value11=float(data[5])*0.3048
                        self.tableWidget1.setItem(row1, 5, QtWidgets.QTableWidgetItem(str(value11)))
                elif self.comboBox.currentText()=='VCR_LL':
                    if data[6]=='':
                        self.tableWidget1.setItem(row1, 6, QtWidgets.QTableWidgetItem(''))
                    elif data[6]!='':
                        value12=float(data[6])*0.3048
                        self.tableWidget1.setItem(row1, 6, QtWidgets.QTableWidgetItem(str(value12)))
                elif self.comboBox.currentText()=='VCR_RR':
                    if data[7] == '':
                        self.tableWidget1.setItem(row1, 7, QtWidgets.QTableWidgetItem(''))
                    elif data[7] != '':
                        value13 = float(data[7]) * 0.3048
                        self.tableWidget1.setItem(row1, 7, QtWidgets.QTableWidgetItem(str(value13)))
                elif self.comboBox.currentText()=='HD_LL':
                    if data[8]=='':
                        self.tableWidget1.setItem(row1, 8, QtWidgets.QTableWidgetItem(''))
                    elif data[8]!='':
                        value14=round(float(data[8])*0.3048,3)
                        self.tableWidget1.setItem(row1, 8, QtWidgets.QTableWidgetItem(str(value14)))
                elif self.comboBox.currentText()=='HD_RR':
                    if data[9]=='':
                        self.tableWidget1.setItem(row1, 9, QtWidgets.QTableWidgetItem(''))
                    elif data[9]!='':
                        value15=round(float(data[9])*0.3048,3)
                        self.tableWidget1.setItem(row1, 9, QtWidgets.QTableWidgetItem(str(value15)))
                elif self.comboBox.currentText()=='VD_LL':
                    if data[10] == '':
                        self.tableWidget1.setItem(row1, 10, QtWidgets.QTableWidgetItem(''))
                    elif data[10] != '':
                        value16 = round(float(data[10]) * 0.3048,3)
                        self.tableWidget1.setItem(row1, 10, QtWidgets.QTableWidgetItem(str(value16)))
                elif self.comboBox.currentText()=='VD_RR':
                    if data[11]=='':
                        self.tableWidget1.setItem(row1, 11, QtWidgets.QTableWidgetItem(''))
                    elif data[11]!='':
                        value17=round(float(data[11])*0.3048,3)
                        self.tableWidget1.setItem(row1, 11, QtWidgets.QTableWidgetItem(str(value17)))
                elif self.comboBox.currentText() == 'PH':
                    if data[13] == '':
                        self.tableWidget1.setItem(row1, 13, QtWidgets.QTableWidgetItem(''))
                    elif data[13] != '':
                        value18 = round(float(data[13]) * 0.3048,3)
                        self.tableWidget1.setItem(row1, 13, QtWidgets.QTableWidgetItem(str(value18)))
            row1+=1
    def Clear_Input_data(self):
        self.tableWidget1.clearContents()
    def Clear_Output_data(self):
        self.tableWidget_2.clearContents()


if __name__=='__main__':
    def my_exception(type,value,tback):
        print(tback,value,tback)
        sys.__excepthook__(type,value,tback)
    sys.excepthook=my_exception
    app = QtWidgets.QApplication(sys.argv)
    window = Tacheometry()
    app.exec_()

from PyQt5 import QtWidgets
import sys
from TACHEOMETRY import Tacheometry
from Transformation_UI_AND_COMPUTATION import Transformation
from LEVELLING_MAIN import Levelling
from homepage_ui import Ui_GEOMATIC_SUITE

class homepage(QtWidgets.QMainWindow,Ui_GEOMATIC_SUITE):
    def __init__(self,parent=None):
        super(homepage, self).__init__(parent)
        self.setupUi(self)
        self.pushButton.clicked.connect(self.Enter)
        self.show()
    def Enter(self):
        try:

            item = self.comboBox.currentText()
            if item=='SPIRIT_LEVELLING':
                self.data=Levelling()
                self.data.show()
            elif item=='TACHEOMETRIC DETAILING/LEVELLING/TRAVERSING':
                self.data1=Tacheometry()
                self.data1.show()
            elif item=='DATUM TRANSFORMATION':
                self.data2=Transformation()
                self.data2.show()

        except BaseException as e:
            print(e)


if __name__ == "__main__":
    def my_exception(type,value,tback):
        print(tback,value,tback)
        sys.__excepthook__(type,value,tback)
    sys.excepthook=my_exception
    app = QtWidgets.QApplication(sys.argv)
    window = homepage()
    app.exec_()

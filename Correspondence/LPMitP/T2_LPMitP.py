from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PyQt5.QtCore import pyqtSignal

# import xlsxwriter
# import xlrd
#from PSch.mysql_connector import update_sql_adm, retri_sql, insert_sql
#from PyQt5.QtGui import QColor
#from .MC_Computer import *
#from .Sch_code_connector import *
#from decimal import *
#import datetime


class T2_LPMitP(QDialog):

    toHome_s = pyqtSignal()
    toCorres_s = pyqtSignal()

    def __init__(self, path):
        QDialog.__init__(self)
        #loadUi(r"C:\Users\S T KWONG\AppData\Local\Programs\Python\Python39-32\Lib\site-packages\Gov_Code\Portal\Correspondence\LPMitP\T2_LPMitP.ui", self)
        loadUi(path + r"\Correspondence\LPMitP\T2_LPMitP.ui",self)

        self.T2_LPMitP_home_but.clicked.connect(self.toHome)
        self.T2_LPMitP_Return_but.clicked.connect(self.toCorres)

    def toHome(self):
        self.toHome_s.emit()
        self.close()

    def toCorres(self):
        self.toCorres_s.emit()
        self.close()



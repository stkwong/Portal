from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PyQt5.QtCore import pyqtSignal
from .Complaint.T2_Complaint import *
from .LPMitP.T2_LPMitP import *
from .Land.T2_Land import *
# import xlsxwriter
# import xlrd
#from PSch.mysql_connector import update_sql_adm, retri_sql, insert_sql
#from PyQt5.QtGui import QColor
#from .MC_Computer import *
#from .Sch_code_connector import *
#from decimal import *
#import datetime


class T1_Corres(QDialog):

    toHome_s = pyqtSignal()

    def __init__(self, path):
        QDialog.__init__(self)
        #loadUi(r"C:\Users\S T KWONG\AppData\Local\Programs\Python\Python39-32\Lib\site-packages\Gov_Code\Portal\Correspondence\T1_Corres.ui", self)
        loadUi(path + r"\Correspondence\T1_Corres.ui",self)

        #signal
        self.T1_Corres_Complaint_but.clicked.connect(self.toComplaint)
        self.T1_Corres_LPMitP_but.clicked.connect(self.toLPMitP)
        self.T1_Corres_Land_but.clicked.connect(self.toLand)

        #child and return
        self.T2_Complaint = T2_Complaint(path)
        self.T2_Complaint.toCorres_s.connect(self.retu_Corres)

        self.T2_LPMitP = T2_LPMitP(path)
        self.T2_LPMitP.toCorres_s.connect(self.retu_Corres)

        self.T2_Land = T2_Land(path)
        self.T2_Land.toCorres_s.connect(self.retu_Corres)

        #toHome
        self.T1_Corres_home_but.clicked.connect(self.toHome)

    def toHome(self):
        self.toHome_s.emit()
        self.close()

    def toComplaint(self):
        self.hide()
        self.T2_Complaint.show()

    def toLPMitP(self):
        self.hide()
        self.T2_LPMitP.show()

    def toLand(self):
        self.hide()
        self.T2_Land.show()

    def retu_Corres(self):
        self.show()

    # def closeEvent (self, event):
    #     if self.isDirectClose:
    #         event.accept()
    #     else:
    #         reply = QMessageBox.question(self, 'Close Without Saving?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
    #         if reply == QMessageBox.Yes:
    #             event.accept()
    #         else:
    #             event.ignore()


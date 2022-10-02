from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PyQt5.QtCore import pyqtSignal
from .CheckEI.T2_CheckEI import *
# import xlsxwriter
# import xlrd
#from PSch.mysql_connector import update_sql_adm, retri_sql, insert_sql
#from PyQt5.QtGui import QColor
#from .MC_Computer import *
#from .Sch_code_connector import *
#from decimal import *
#import datetime

class T1_Consult(QDialog):

    toHome_s = pyqtSignal()

    def __init__(self, path, p_to_database, p_to_documents, p_to_server):
        QDialog.__init__(self)
        #loadUi(r"C:\Users\S T KWONG\AppData\Local\Programs\Python\Python39-32\Lib\site-packages\Gov_Code\Portal\Consultancy\T1_Consult.ui", self)
        loadUi(path + r"\Consultancy\T1_Consult.ui", self)
        self.load_consult_data()

        # child and return
        self.T1_Consult_home_but.clicked.connect(self.home)
        self.T1_Consult_Update_Audit_Table_but.clicked.connect(self.load_consult_data)
        self.T2_CheckEI = T2_CheckEI(path, p_to_database, p_to_documents, p_to_server)
        self.T2_CheckEI.toConsult_s.connect(self.retu_Consult)
        self.T1_Consult_CheckEI_but.clicked.connect(self.toCheckEI)

    def home(self):
        self.toHome_s.emit()
        self.close()

    def load_consult_data(self):
        pass

    def toCheckEI(self):
        self.hide()
        self.T2_CheckEI.show()

    def retu_Consult(self):
        self.show()


from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PyQt5.QtCore import pyqtSignal
from .Content.T3_Content import *
from .Add_Upd_Caller.T3_Add_Upd_Caller import *
import datetime
from accdb_connector import retri_sql, insert_sql
# import xlsxwriter
# import xlrd
#from PSch.mysql_connector import update_sql_adm, retri_sql, insert_sql
#from PyQt5.QtGui import QColor
#from .MC_Computer import *
#from .Sch_code_connector import *
#from decimal import *


class T2_Complaint(QDialog):

    toHome_s = pyqtSignal()
    toCorres_s = pyqtSignal()

    def __init__(self, path):
        self.path = path
        QDialog.__init__(self)
        #loadUi(r"C:\Users\S T KWONG\AppData\Local\Programs\Python\Python39-32\Lib\site-packages\Gov_Code\Portal\Correspondence\Complaint\T2_Complaint.ui", self)
        loadUi(path + r"\Correspondence\Complaint\T2_Complaint.ui", self)
        self.isDirectClose = True

        # signal
        self.T2_Complaint_Add_Upd_Caller_but.clicked.connect(self.toAdd_Upd_Caller)

        # child and return
        self.T3_Content = T3_Content(path)
        self.T3_Content.toComplaint_s.connect(self.retu_Complaint)
        self.T3_Add_Upd_Caller = T3_Add_Upd_Caller(path)
        self.T3_Add_Upd_Caller.toComplaint_s.connect(self.retu_Complaint)

        #Attributes
        self.yr = str(datetime.date.today().year)
        self.to_win = "Close"

        #Method
        self.load_complaint_data()
        self.load_channel_data()

        # Att - but
        self.T2_Complaint_Update_Complaint_Table_but.clicked.connect(self.load_complaint_data)
        self.T2_Complaint_Return_but.clicked.connect(self.toCorres)

        # Att - le
        self.T2_Complaint_CaseNo_Yr_le.setText(self.yr)

        #toHome
        self.T2_Complaint_home_but.clicked.connect(self.toHome)

    def toHome(self):
        self.to_win = "home"
        self.close()

    def toCorres(self):
        self.to_win = "corres"
        self.close()

    def toAdd_Upd_Caller(self):
        self.hide()
        self.T3_Add_Upd_Caller.show()

    def load_complaint_data(self):
        pass

    def load_channel_data(self):
        channel_nlist = retri_sql(self.path, "Channel_log", "*")
        for row in range(len(channel_nlist)):
            self.T2_Complaint_Channel_com.addItem(channel_nlist[row][0])

    def save(self):
        self.isDirectClose = True

    def retu_Complaint(self):
        self.show()

    def closeEvent (self, event):
        if self.isDirectClose:
            if self.to_win == "home":
                self.toHome_s.emit()
            elif self.to_win == "corres":
                self.toCorres_s.emit()
            event.accept()
        else:
            reply = QMessageBox.question(self, 'Close Without Saving?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                if self.to_win == "home":
                    self.toHome_s.emit()
                elif self.to_win == "corres":
                    self.toCorres_s.emit()
                event.accept()
            else:
                event.ignore()


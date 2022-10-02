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


class T3_Content(QDialog):

    toComplaint_s = pyqtSignal()

    def __init__(self, path):
        QDialog.__init__(self)
        #loadUi(r"C:\Users\S T KWONG\AppData\Local\Programs\Python\Python39-32\Lib\site-packages\Gov_Code\Portal\Correspondence\Complaint\Content\T3_Content.ui", self)
        loadUi(path + r"\Correspondence\Complaint\Content\T3_Content.ui", self)
        self.isDirectClose = True
        self.T3_Content_Return_but.clicked.connect(self.toComplaint)

    def toComplaint(self):
        self.toComplaint_s.emit()
        self.close()

    def save(self):
        self.isDirectClose = True

    def closeEvent (self, event):
        if self.isDirectClose:
            event.accept()
        else:
            reply = QMessageBox.question(self, 'Close Without Saving?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                event.accept()
            else:
                event.ignore()


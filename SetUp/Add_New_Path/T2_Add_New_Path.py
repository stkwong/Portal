from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from accdb_connector import insert_path
from PyQt5.QtCore import pyqtSignal

class T2_Add_New_Path_Form(QDialog):

    Save_and_CheckEI_s = pyqtSignal()
    toCheckEI_s = pyqtSignal()

    def __init__(self, path):
        self.path = path
        QDialog.__init__(self)
        loadUi(path + r"\SetUp\Add_New_Path\T2_Add_New_Path.ui", self)
        self.T2_Add_New_Path_Add_but.clicked.connect(self.add_records)
        self.T2_Add_New_Path_clear_but.clicked.connect(self.clear_records)
        self.T2_Add_New_Path_Return_but.clicked.connect(self.return_toCheckEI)
        self.input_des = ""
        self.T2_Add_New_Path_Path_le.clear()
        # self.isDirectClose = True

    def add_records(self):
        if self.T2_Add_New_Path_Path_le.text() == "":
            print("please input Brief")
        elif self.input_des != "database" and self.input_des != "documents" and self.input_des != "server":
            print("incorrect input of input_des : ", self.input_des)
        else:
            self.New_Path = self.T2_Add_New_Path_Path_le.text()
            insert_path(self.path , "path_options", destination=self.input_des, saved_path=self.New_Path )
            self.Save_and_CheckEI_s.emit()
            self.close()

    def return_toCheckEI(self):
        self.toCheckEI_s.emit()
        self.T2_Add_New_Path_Path_le.clear()
        self.close()

    def clear_records(self):
        self.T2_Add_New_Path_Path_le.clear()

    # def closeEvent (self, event):
    #     if self.isDirectClose:
    #         event.accept()
    #     else:
    #         reply = QMessageBox.question(self, 'Close Without Saving?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
    #         if reply == QMessageBox.Yes:
    #             self.toCheckEI_s.emit()
    #             event.accept()
    #         else:
    #             event.ignore()
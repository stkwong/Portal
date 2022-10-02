from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from accdb_connector import update_sql_1key, retri_sql
from PyQt5.QtCore import pyqtSignal

class T3_Edit_Com_Form(QDialog):

    Save_and_CheckEI_s = pyqtSignal()
    toCheckEI_s = pyqtSignal()

    def __init__(self, path, p_to_database, p_to_documents, p_to_server):
        self.path = path
        self.p_to_database = p_to_database
        QDialog.__init__(self)
        loadUi(path + r"\Consultancy\CheckEI\Edit_Com\T3_Edit_Com.ui", self)
        self.T3_Edit_Com_Add_but.clicked.connect(self.add_records)
        self.isDirectClose = True
        self.T3_Edit_Com_Return_but.clicked.connect(self.return_toCheckEI)

    def retri_data(self, data_id):
        self.data_id = data_id
        CheckEI_nlist = retri_sql(self.p_to_database , "CheckEI_Data", "Sec", "Brief", "Details", "Comments", id=self.data_id )
        self.T3_Edit_Com_Sec_le.setText(CheckEI_nlist[0][0])
        self.T3_Edit_Com_Brief_le.setText(CheckEI_nlist[0][1])
        self.T3_Edit_Com_Details_le.setText(CheckEI_nlist[0][2])
        self.T3_Edit_Com_Comment_le.setText(CheckEI_nlist[0][3])

    def add_records(self):
        try:
            if self.T3_Edit_Com_Sec_le.text() == "":
                self.T3_Edit_Com_Error_le.setText("please input Section")
            elif self.T3_Edit_Com_Brief_le.text() == "":
                self.T3_Edit_Com_Error_le.setText("please input Brief")
            elif self.T3_Edit_Com_Details_le.text() == "":
                self.T3_Edit_Com_Error_le.setText("please input Details")
            elif self.T3_Edit_Com_Comment_le.text() == "":
                self.T3_Edit_Com_Error_le.setText("please input Comments")
            else:
                self.Section_text = self.T3_Edit_Com_Sec_le.text()
                Brief_text = self.T3_Edit_Com_Sec_le.text()
                Details_text = self.T3_Edit_Com_Details_le.text()
                self.Comments_text = self.T3_Edit_Com_Comment_le.text()
                update_sql_1key(self.p_to_database , "CheckEI_Data", "Sec", "id", self.data_id, self.Section_text )
                update_sql_1key(self.p_to_database, "CheckEI_Data", "Brief", "id", self.data_id, Brief_text)
                update_sql_1key(self.p_to_database, "CheckEI_Data", "Details", "id", self.data_id, Details_text)
                update_sql_1key(self.p_to_database, "CheckEI_Data", "Comments", "id", self.data_id, self.Comments_text)

                self.Save_and_CheckEI_s.emit()
                self.close()

        except ValueError:
            print( 'Please enter an integer' )

    def return_toCheckEI(self):
        self.toCheckEI_s.emit()
        self.close()

    def closeEvent (self, event):
        if self.isDirectClose:
            event.accept()
        else:
            reply = QMessageBox.question(self, 'Close Without Saving?', QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.toCheckEI_s.emit()
                event.accept()
            else:
                event.ignore()
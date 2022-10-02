from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from accdb_connector import insert_sql, retri_sql
from PyQt5.QtCore import pyqtSignal

class T3_Add_New_Form(QDialog):

    Save_and_CheckEI_s = pyqtSignal()
    toCheckEI_s = pyqtSignal()

    def __init__(self, path, p_to_database, p_to_documents, p_to_server):
        self.path = path
        self.p_to_database = p_to_database
        QDialog.__init__(self)
        loadUi(path + r"\Consultancy\CheckEI\Add_New\T3_Add_New.ui", self)
        self.T3_Add_New_Add_but.clicked.connect(self.add_records)
        self.T3_Add_New_clear_but.clicked.connect(self.clear_records)
        self.isDirectClose = True
        Sec_nlist = retri_sql(p_to_database, "CheckEI_Data", "Sec")
        Sec_lst = set([i for lst in Sec_nlist for i in lst])
        self.T3_Add_New_Sec_com.addItem("NA")
        self.T3_Add_New_Sec_com.addItems(Sec_lst)
        self.T3_Add_New_Sec_com.setCurrentText("NA")
        self.T3_Add_New_Sec_com.currentTextChanged.connect(self.choose_le_or_combo)
        self.T3_Add_New_Sec_le.textEdited.connect(self.choose_le_or_combo)
        self.T3_Add_New_Return_but.clicked.connect(self.return_toCheckEI)

    def choose_le_or_combo(self):
        print ("inside")
        if self.T3_Add_New_Sec_com.currentText() == "NA":
            print ("true")
            self.T3_Add_New_Sec_le.setVisible(True)
            self.Section_text = self.T3_Add_New_Sec_le.text()
        else:
            print ("false")
            self.T3_Add_New_Sec_le.setVisible(False)
            self.Section_text = self.T3_Add_New_Sec_com.currentText()

    def add_records(self):
        try:
            if not hasattr(self, "Section_text"):
                self.T3_Add_New_Error_le.setText("please input Section")
            elif self.T3_Add_New_Brief_le.text() == "":
                self.T3_Add_New_Error_le.setText("please input Brief")
            elif self.T3_Add_New_Details_le.text() == "":
                self.T3_Add_New_Error_le.setText("please input Details")
            elif self.T3_Add_New_Comment_le.text() == "":
                self.T3_Add_New_Error_le.setText("please input Comments")
            else:
                Brief_text = self.T3_Add_New_Sec_le.text()
                Details_text = self.T3_Add_New_Details_le.text()
                self.Comments_text = self.T3_Add_New_Comment_le.text()
                insert_sql(self.p_to_database , "CheckEI_Data", Sec=self.Section_text, Brief=Brief_text, Details=Details_text, Comments= self.Comments_text )
                self.Save_and_CheckEI_s.emit()
                self.close()

        except ValueError:
            print( 'Please enter an integer' )

    def return_toCheckEI(self):
        self.toCheckEI_s.emit()
        self.close()

    def clear_records(self):
        pass

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
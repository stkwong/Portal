from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from accdb_connector import insert_sql, retri_sql
from PyQt5.QtCore import pyqtSignal

class T2_Add_PMW_Slope_Form(QDialog):

    Save_and_toPMW_s = pyqtSignal()
    toPMW_s = pyqtSignal()

    def __init__(self, path, p_to_database):
        self.p_to_database = p_to_database
        QDialog.__init__(self)
        loadUi(path + r"\PMW\Add_PMW_Slope\T2_Add_PMW_Slope.ui", self)
        self.T2_Add_PMW_Add_but.clicked.connect(self.add_records)
        self.T2_Add_PMW_clear_but.clicked.connect(self.clear_records)
        self.isDirectClose = True
        self.T2_Add_PMW_Return_but.clicked.connect(self.return_toPMW)

        self.msgBoxClose = QMessageBox()
        self.msgBoxClose.setIcon(QMessageBox.Information)
        self.msgBoxClose.setText("Close without save?")
        self.msgBoxClose.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

        self.msgBoxClear = QMessageBox()
        self.msgBoxClear.setIcon(QMessageBox.Information)
        self.msgBoxClear.setText("Clear all inputs?")
        self.msgBoxClear.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

        self.T2_Add_PMW_Package_le.textChanged.connect(self.change_text)
        self.T2_Add_PMW_Feature_le.textChanged.connect(self.change_text)
        self.T2_Add_PMW_Location_le.textChanged.connect( self.change_text )

    def change_text(self):
        self.isDirectClose = False

    def add_records(self):
        try:
            self.Package = self.T2_Add_PMW_Package_le.text()
            self.Feature_No = self.T2_Add_PMW_Feature_le.text()
            self.Location = self.T2_Add_PMW_Location_le.text()
            self.Save_and_toPMW_s.emit()
            self.isDirectClose = True
            self.close()

        except ValueError:
            print( 'Error' )

    def return_toPMW(self):
        self.close()

    def clear_records(self):
        clearAns = self.msgBoxClear.exec()
        if clearAns == QMessageBox.Yes:
            self.T2_Add_PMW_Package_le.clear()
            self.T2_Add_PMW_Feature_le.clear()
            self.T2_Add_PMW_Location_le.clear()

    def closeEvent (self, event):
        if self.isDirectClose:
            self.toPMW_s.emit()
            event.accept()
        else:
            closeAns = self.msgBoxClose.exec()
            if closeAns == QMessageBox.Yes:
                self.clear_records()
                self.toPMW_s.emit()
                event.accept()
            else:
                event.ignore()
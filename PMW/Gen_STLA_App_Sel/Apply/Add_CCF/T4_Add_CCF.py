from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PyQt5.QtCore import pyqtSignal

class T4_Add_CCF_Form(QDialog):

    toApply_s = pyqtSignal()
    sel_and_toApply_s = pyqtSignal()

    def __init__(self, path, p_to_database, p_to_documents, p_to_server):
        QDialog.__init__(self)
        loadUi(path + r"\PMW\Gen_STLA_App_Sel\Apply\Add_CCF\T4_Add_CCF.ui",self)

        self.p_to_database = p_to_database
        self.p_to_documents = p_to_documents
        self.p_to_server = p_to_server

        self.msgBoxClose = QMessageBox()
        self.msgBoxClose.setIcon(QMessageBox.Information)
        self.msgBoxClose.setText("Close this page?")
        self.msgBoxClose.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

        self.T4_Add_CCF_return_but.clicked.connect(self.toApply)
        self.T4_Add_CCF_reload_but.clicked.connect(self.reset)
        self.T4_Add_CCF_execute_but.clicked.connect(self.ccf_execute)
        self.T4_Add_CCF_n_a_rad.setChecked(True)

        self.isDirectClose = True
        self.save = False

    def toApply(self):
        self.toApply_s.emit()
        self.close()

    def ccf_execute(self):
        if self.T4_Add_CCF_file_le.text() == "":
            print ("please input file")
        else:
            self.isDirectClose = True
            self.sel_and_toApply_s.emit()
            self.close()

    def reset(self):
        self.T4_Add_CCF_file_le.clear()
        self.T4_Add_CCF_n_a_rad.setChecked(True)

    def closeEvent (self, event):
        if self.isDirectClose:
            if not self.save:
                self.toApply_s.emit()
            event.accept()
        else:
            closeAns = self.msgBoxClose.exec()
            if closeAns == QMessageBox.Yes:
                self.reset()
                self.toApply_s.emit()
                event.accept()
            else:
                event.ignore()




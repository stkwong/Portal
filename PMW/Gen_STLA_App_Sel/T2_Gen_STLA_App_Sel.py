from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PyQt5.QtCore import pyqtSignal
from accdb_connector import retri_sql, insert_sql, update_sql_1key
from PMW.Gen_STLA_App_Sel.Apply.T3_Apply import *
import datetime

class T2_Gen_STLA_App_Sel_Form(QDialog):

    toHome_s = pyqtSignal()
    toPMW_s = pyqtSignal()

    def __init__(self, path, p_to_database, p_to_documents, p_to_server):

        QDialog.__init__(self)
        self.path = path
        loadUi(path + r"\PMW\Gen_STLA_App_Sel\T2_Gen_STLA_App_Sel.ui", self)

        self.ID = 0

        #child
        self.T3_Apply = T3_Apply_Form(path, p_to_database, p_to_documents, p_to_server)
        self.T3_Apply.toSTLA_Sel_s.connect(self.return_wo_save)

        self.T2_PMW_STLA_App_Sel_but.clicked.connect(self.to_STLA_App)
        self.T2_PMW_STLA_Reminder_Sel_but.clicked.connect(self.to_STLA_App_Reminder)
        self.T2_PMW_STLA_Return_Sel_but.clicked.connect(self.to_PMW)
        self.T2_PMW_STLA_home_Sel_but.clicked.connect(self.to_home)

        self.p_to_database = p_to_database
        self.p_to_server = p_to_server

        # msgBox
        self.msgBoxOverwrite = QMessageBox()
        self.msgBoxOverwrite.setIcon( QMessageBox.Information )
        self.msgBoxOverwrite.setText( "STLA Applied. Overwrite existing records?" )
        self.msgBoxOverwrite.setStandardButtons( QMessageBox.Yes | QMessageBox.No )

    def return_wo_save(self):
        self.show()

    def to_STLA_App(self):
        if self.STLA_applied():
            OverwriteAns = self.msgBoxOverwrite.exec()
            if OverwriteAns == QMessageBox.Yes:
                self.T3_Apply.create_New( self.ID )
                self.T3_Apply.show()
                self.close()
        else:
            self.T3_Apply.create_New(self.ID)
            self.T3_Apply.show()
            self.close()

    def to_STLA_App_Reminder(self):
        pass

    def to_PMW(self):
        self.hide()
        self.toPMW_s.emit()

    def to_home(self):
        self.hide()
        self.toHome_s.emit()

    def getID(self, id):
        self.ID = id

    def STLA_applied(self):
        if self.ID == 0:
            print ("ID is not correct")
            return True
        else:
            data_nlist = retri_sql(self.p_to_database, "PMW_records", "STLA_App_date", id=str(self.ID))
            print (data_nlist[0][0])
            if data_nlist[0][0] == "--":
                return False
            else:
                return True


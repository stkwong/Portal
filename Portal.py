import sys
import os
#from PSch.mysql_connector import update_sql_adm, retri_sql, insert_sql
#from PyQt5.QtGui import QColor
#from .MC_Computer import *
#from .Sch_code_connector import *
#from decimal import *
import datetime
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PMW.T1_PMW import *
from Correspondence.T1_Corres import *
from Consultancy.T1_Consult import *
from SetUp.T1_SetUp import *
from accdb_connector import retri_outgoing, update_sql_1key
import codecs

class Portal_Form(QDialog):

    def __init__(self):
        path = os.getcwd()
        self.path = path
        QDialog.__init__(self)
        loadUi(path + "\Portal_Form.ui", self)
        # Setup (sp_)
        self.SetUp = T1_SetUp_Form(self.path)
        self.setup_but.clicked.connect(self.toSetUp)
        self.SetUp.toHome_s.connect(self.retu_home_from_setup)
        self.start_but.clicked.connect(self.start)

        self.msgBoxUpdate = QMessageBox()
        self.msgBoxUpdate.setIcon(QMessageBox.Information)
        self.msgBoxUpdate.setText("Record modified?")
        self.msgBoxUpdate.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

        self.msgBoxClear = QMessageBox()
        self.msgBoxClear.setIcon(QMessageBox.Information)
        self.msgBoxClear.setText("Reset table?")
        self.msgBoxClear.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

    def start(self):
        # Attributes
        self.start_but.hide()

        self.p_to_database, self.p_to_documents, self.p_to_server = create_path(self.path)
        print ("at start: ",  self.p_to_database)

        self.yr = str(datetime.date.today().year)
        self.mth = str(datetime.date.today().month)
        self.d = str(datetime.date.today().day)

        self.refresh_but.clicked.connect(self.refresh_table)

        # Form_Attributes
        self.yr_le.setText(self.yr)
        self.mth_le.setText(self.mth)
        self.d_le.setText(self.d)

        # A. T1_PMW (pmw_)
        self.T1_PMW = T1_PMW(self.path, self.p_to_database, self.p_to_documents, self.p_to_server)
        self.pmw_but.clicked.connect(self.toPMW)
        self.T1_PMW.home_s.connect(self.retu_home)
        self.T1_PMW.T2_Gen_STLA_App_Sel.toHome_s.connect(self.retu_home)

        # B. T1_Consult (consult_)
        self.T1_Consult = T1_Consult(self.path, self.p_to_database, self.p_to_documents, self.p_to_server)
        self.consult_but.clicked.connect(self.toConsult)
        self.T1_Consult.toHome_s.connect(self.retu_home)

        # C. T1_Corres (corres_)
        self.T1_Corres = T1_Corres(self.path)
        self.corres_but.clicked.connect(self.toCorres)
        self.T1_Corres.toHome_s.connect(self.retu_home)

        # toHome
        self.T1_Corres.T2_Complaint.toHome_s.connect(self.retu_home)
        self.T1_Corres.T2_LPMitP.toHome_s.connect(self.retu_home)
        self.T1_Corres.T2_Land.toHome_s.connect(self.retu_home)
        self.T1_Consult.T2_CheckEI.toHome_s.connect(self.retu_home)

        self.retri_data()

    def toPMW(self):
        self.hide()
        self.T1_PMW.show()

    def toSetUp(self):
        self.hide()
        self.SetUp.show()

    def toConsult(self):
        self.hide()
        self.T1_Consult.show()

    def toCorres(self):
        self.hide()
        self.T1_Corres.show()

    def retu_home(self):
        self.show()

    def retu_home_from_setup(self):
        self.show()
        self.start_but.show()

    def refresh_table(self):
        self.reset_table()
        self.retri_data()

    def retri_data(self):

        self.outgoing_table.clearContents()

        for delta in range(14):
            concerned_date = datetime.date.today()  - datetime.timedelta(days=delta)
            concerned_date_str = concerned_date.strftime("%d/%m/%Y")
            print ("in retri_date :", self.p_to_database)
            outgoing_data = retri_outgoing(self.p_to_database, "Outgoing_records", "time_mod", "DESC", "*", date_mod= "'" + concerned_date_str + "'")
            row = 0
            for record_num in range(len(outgoing_data)):
                self.outgoing_table.insertRow(row)

                self.edit_records_but = QPushButton("Edit")
                self.edit_records_but.clicked.connect(self.edit_records)
                self.outgoing_table.setCellWidget(row, 0, self.edit_records_but)

                it = QTableWidgetItem()
                it.setText(outgoing_data[record_num][1])
                self.outgoing_table.setItem(row, 1, it)

                it = QTableWidgetItem()
                time_mod = outgoing_data[record_num][2].strftime("%H:%M:%S")
                it.setText(time_mod)
                self.outgoing_table.setItem(row, 2, it)

                for col in range(3, self.outgoing_table.columnCount()):
                    it = QTableWidgetItem()
                    it.setText(outgoing_data[record_num][col])
                    self.outgoing_table.setItem(row, col, it)
                row += 1

    def reset_table(self):
        clearAns = self.msgBoxClear.exec()
        if clearAns == QMessageBox.Yes:
            self.outgoing_table.clearContents()
            self.outgoing_table.setRowCount(0)

    def edit_records(self):
        edit_records_button = self.sender()
        index = self.outgoing_table.indexAt(edit_records_button.pos())
        edit_records_row = index.row()
        it = self.outgoing_table.item(edit_records_row, 12)
        if it and it.text():
            os.startfile(it.text())
            updateAns = self.msgBoxUpdate.exec()
            if updateAns == QMessageBox.Yes:
                now_time = datetime.datetime.now().time()
                now_date = datetime.date.today().strftime("%d/%m/%Y")
                update_sql_1key(self.p_to_database, "Outgoing_records", "time_mod", "path_to_file", "'" + it.text() + "'", now_time)
                update_sql_1key(self.p_to_database, "Outgoing_records", "date_mod", "path_to_file", "'" + it.text() + "'", now_date)
                self.refresh_table_wo_q()

    def reset_table_wo_q(self):
        self.outgoing_table.clearContents()
        self.outgoing_table.setRowCount(0)

if __name__ == '__main__':
    app = QApplication(sys.argv)
    Win = Portal_Form()
    Win.show()
    sys.exit(app.exec_())

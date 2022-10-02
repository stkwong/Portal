from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PyQt5.QtCore import pyqtSignal
from accdb_connector import retri_sql
import datetime

class T4_Add_Slope_Form(QDialog):

    toApply_s = pyqtSignal()
    sel_and_toApply_s = pyqtSignal()

    def __init__(self, path, p_to_database, p_to_documents, p_to_server):
        QDialog.__init__(self)
        loadUi(path + r"\PMW\Gen_STLA_App_Sel\Apply\Add_Slope\T4_Add_Slope.ui",self)

        self.p_to_database = p_to_database
        self.p_to_documents = p_to_documents
        self.p_to_server = p_to_server

        self.T4_Add_Slope_return_but.clicked.connect(self.toApply)
        #self.load_data()

    def toApply(self):
        self.toApply_s.emit()
        self.close()

    def load_data(self):
        self.T4_Add_Slope_table.clearContents()
        self.T4_Add_Slope_table.setRowCount( 1 )
        Selection_PMW_data = retri_sql(self.p_to_database, "PMW_records", "id", "Package", "Feature_No", "Location", "Status", "STLA_Status", "STLA_Exp_date", "NDC_date", "Future_Dev_date", "STLA_App_date", "STLA_Ext_App_date", status="'Selection'")
        row = 0
        for record_num in range(len(Selection_PMW_data)):
            self.T4_Add_Slope_table.insertRow(row)

            for col in range(0, 6):
                it = QTableWidgetItem()
                it.setText(Selection_PMW_data[record_num][col+1])
                self.T4_Add_Slope_table.setItem(row, col, it)

            self.SLTA_App_but = QPushButton("Add_Slope")
            self.T4_Add_Slope_table.setCellWidget(row, 8, self.SLTA_App_but)
            self.SLTA_App_but.clicked.connect(self.select_slope)

            it = QTableWidgetItem()
            it.setText(str(Selection_PMW_data[record_num][0]))
            self.T4_Add_Slope_table.setItem(row, 10, it)

            row += 1

    def select_slope(self):
        select_slope_button = self.sender()
        index = self.T4_Add_Slope_table.indexAt(select_slope_button.pos())
        it = self.T4_Add_Slope_table.item(index.row(), 10)
        data_id = it.text()
        feature_no_nlist = retri_sql(self.p_to_database, "PMW_records", "Feature_No", id=str(data_id))
        self.added_feature = feature_no_nlist[0][0]
        print ("self.added_feature : ", self.added_feature)
        self.sel_and_toApply_s.emit()
        self.close()




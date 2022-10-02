from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PyQt5.QtCore import pyqtSignal
from .Add_PMW_Slope.T2_Add_PMW_Slope import *
# import xlsxwriter
# import xlrd
#from PSch.mysql_connector import update_sql_adm, retri_sql, insert_sql
#from PyQt5.QtGui import QColor
#from .MC_Computer import *
#from .Sch_code_connector import *
#from decimal import *
#import datetime
from accdb_connector import retri_sql
from PMW.Gen_STLA_App_Sel.T2_Gen_STLA_App_Sel import *
import datetime

class T1_PMW(QDialog):

    home_s = pyqtSignal()

    def __init__(self, path, p_to_database, p_to_documents, p_to_server):
        QDialog.__init__(self)
        #loadUi(r"C:\Users\S T KWONG\AppData\Local\Programs\Python\Python39-32\Lib\site-packages\Gov_Code\Portal\PMW\T1_PMW.ui", self)
        loadUi(path + r"\PMW\T1_PMW.ui",self)

        # path

        self.p_to_database = p_to_database
        self.p_to_documents = p_to_documents
        self.p_to_server = p_to_server

        # buttons

        self.T1_PMW_home_but.clicked.connect(self.home)
        self.T1_PMW_Refresh_but.clicked.connect(self.load_data)
        self.T1_PMW_Insert_but.clicked.connect(self.add_pmw)

        # methods

        self.load_data()

        # child and return
        self.T2_Add_PMW_Slope = T2_Add_PMW_Slope_Form(path, p_to_database)
        self.T2_Add_PMW_Slope.Save_and_toPMW_s.connect(self.return_from_add_pmw_slope)
        self.T2_Add_PMW_Slope.toPMW_s.connect(self.return_wo_save)

        self.T2_Gen_STLA_App_Sel = T2_Gen_STLA_App_Sel_Form(path, p_to_database, p_to_documents, p_to_server)
        self.T2_Gen_STLA_App_Sel.toPMW_s.connect(self.return_wo_save)

        self.T2_Gen_STLA_App_Sel.T3_Apply.toPMW_s.connect(self.return_wo_save)

    def add_pmw(self):
        self.T2_Add_PMW_Slope.show()
        self.close()

    def return_wo_save(self):
        self.load_data()
        self.show()

    def home(self):
        self.home_s.emit()
        self.close()

    def load_data(self):
        self.T1_PMW_table.clearContents()
        self.T1_PMW_table.setRowCount(1)
        Selection_PMW_data = retri_sql(self.p_to_database, "PMW_records", "id", "Package", "Feature_No", "Location", "Status", "STLA_Status", "STLA_Exp_date", "NDC_date", "Future_Dev_date", "STLA_App_date", "STLA_Approve_date", "STLA_Ext_App_date", "SMS01_Part1_date", "Design_Com_date", "VO_date", "SMS_Part2_date", "TO_Issue_date", "Completion_date", status="'Selection'")
        row = 0
        for record_num in range(len(Selection_PMW_data)):
            self.T1_PMW_table.insertRow(row)

            for col in range(0, 6):
                it = QTableWidgetItem()
                it.setText(Selection_PMW_data[record_num][col+1])
                self.T1_PMW_table.setItem(row, col, it)

            NDC_date = Selection_PMW_data[record_num][7]
            self.NDC_but = QPushButton(NDC_date)
            self.T1_PMW_table.setCellWidget(row, 6, self.NDC_but)
            self.NDC_but.clicked.connect(self.create_NDC)

            Fut_Dev_date = Selection_PMW_data[record_num][8]
            self.Fut_Dev_but = QPushButton(Fut_Dev_date)
            self.T1_PMW_table.setCellWidget(row, 7, self.Fut_Dev_but)
            self.Fut_Dev_but.clicked.connect(self.create_Fut_Dev)

            SLTA_App_date = Selection_PMW_data[record_num][9]
            self.SLTA_App_but = QPushButton(SLTA_App_date)
            self.T1_PMW_table.setCellWidget(row, 8, self.SLTA_App_but)
            self.SLTA_App_but.clicked.connect(self.create_SLTA_App)

            it = QTableWidgetItem()
            it.setText( Selection_PMW_data[ record_num ][ 10 ] )
            self.T1_PMW_table.setItem( row , 9 , it )

            SLTA_Ext_App_date = Selection_PMW_data[record_num][11]
            self.SLTA_Ext_App_but = QPushButton(SLTA_Ext_App_date)
            self.T1_PMW_table.setCellWidget(row, 10, self.SLTA_Ext_App_but)
            self.SLTA_Ext_App_but.clicked.connect(self.create_SLTA_Ext_App)

            SMS01Part1_date = Selection_PMW_data[record_num][12]
            self.SMS01Part1_but = QPushButton(SMS01Part1_date)
            self.T1_PMW_table.setCellWidget(row, 11, self.SMS01Part1_but)
            self.SMS01Part1_but.clicked.connect(self.create_SMS01Part1)

            Design_Com_date = Selection_PMW_data[record_num][13]
            self.Design_Com_but = QPushButton(Design_Com_date)
            self.T1_PMW_table.setCellWidget(row, 12, self.Design_Com_but)
            self.Design_Com_but.clicked.connect(self.create_Design_Com)

            VO_date = Selection_PMW_data[record_num][14]
            self.VO_but = QPushButton(VO_date)
            self.T1_PMW_table.setCellWidget(row, 13, self.VO_but)
            self.VO_but.clicked.connect(self.create_VO)

            SMS01Part2_date = Selection_PMW_data[record_num][15]
            self.SMS01Part2_but = QPushButton(SMS01Part2_date)
            self.T1_PMW_table.setCellWidget(row, 14, self.SMS01Part2_but)
            self.SMS01Part2_but.clicked.connect(self.create_SMS01Part2)

            TO_Issue_date = Selection_PMW_data[record_num][16]
            self.TO_Issue_but = QPushButton(TO_Issue_date)
            self.T1_PMW_table.setCellWidget(row, 15, self.TO_Issue_but)
            self.TO_Issue_but.clicked.connect(self.create_TO_Issue)

            Completion_date = Selection_PMW_data[record_num][17]
            self.Completion_but = QPushButton(Completion_date)
            self.T1_PMW_table.setCellWidget(row, 16, self.Completion_but)
            self.Completion_but.clicked.connect(self.create_Completion)

            it = QTableWidgetItem()
            it.setText(str(Selection_PMW_data[record_num][0]))
            self.T1_PMW_table.setItem(row, 17, it)

            row += 1

    def create_NDC(self):
        pass

    def create_Fut_Dev(self):
        pass

    def create_SLTA_App(self):
        STLA_button = self.sender()
        index = self.T1_PMW_table.indexAt(STLA_button.pos())
        print ("index :" , index.row())
        it = self.T1_PMW_table.item(index.row(), 17)
        data_id = it.text()
        self.T2_Gen_STLA_App_Sel.getID(data_id)

        self.T2_Gen_STLA_App_Sel.show()
        self.close()

    def create_SLTA_Ext_App(self):
        pass

    def create_SMS01Part1(self):
        pass

    def create_Design_Com(self):
        pass

    def create_VO(self):
        pass

    def create_SMS01Part2(self):
        pass

    def create_TO_Issue(self):
        pass

    def create_Completion(self):
        pass

    def return_from_add_pmw_slope(self):
        insert_sql(self.p_to_database , "PMW_records", Package=self.T2_Add_PMW_Slope.Package, Feature_No=self.T2_Add_PMW_Slope.Feature_No, Location=self.T2_Add_PMW_Slope.Location, STLA_Status="Not Apply", STLA_Exp_date="--", STLA_Approve_date="--", NDC_date="--", Future_Dev_date="--", STLA_App_date="--", STLA_Ext_App_date="--", SMS01_Part1_date="--", Design_Com_date="--", VO_date="--", SMS_Part2_date="--", TO_Issue_date="--", Completion_date="--", Status="Selection")
        self.load_data()
        self.show()



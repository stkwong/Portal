from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PyQt5.QtCore import pyqtSignal
from accdb_connector import retri_sql
import datetime

class T4_Add_File_Thro_Form(QDialog):

    toApply_s = pyqtSignal()
    sel_and_toApply_s = pyqtSignal()

    def __init__(self, path, p_to_database, p_to_documents, p_to_server):
        QDialog.__init__(self)
        loadUi(path + r"\PMW\Gen_STLA_App_Sel\Apply\Add_File_Thro\T4_Add_File_Thro.ui",self)

        self.p_to_database = p_to_database
        self.p_to_documents = p_to_documents
        self.p_to_server = p_to_server

        self.msgBoxClose = QMessageBox()
        self.msgBoxClose.setIcon(QMessageBox.Information)
        self.msgBoxClose.setText("Close this page?")
        self.msgBoxClose.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

        self.T4_Add_File_Thro_return_but.clicked.connect(self.toApply)
        self.T4_Add_File_Thro_reload_but.clicked.connect(self.load_file_thro_recipient)
        self.T4_Add_File_Thro_execute_but.clicked.connect(self.cc_execute)

        self.isDirectClose = True
        self.save = False

        self.T4_Add_File_Thro_SGE_GE_com.currentTextChanged.connect(self.add_GE_SGE)
        self.T4_Add_File_Thro_STO_TO_com.currentTextChanged.connect( self.add_TO_STO )

    def add_TO_STO(self):
        self.T4_Add_File_Thro_SGE_GE_com.setCurrentText("")

    def add_GE_SGE(self):
        self.T4_Add_File_Thro_STO_TO_com.setCurrentText("")

    def load_file_thro_recipient(self):
        self.reset()
        self.T4_Add_File_Thro_SGE_GE_com.clear()
        self.T4_Add_File_Thro_STO_TO_com.clear()
        self.T4_Add_File_Thro_SGE_GE_com.addItem("")
        self.T4_Add_File_Thro_STO_TO_com.addItem("")

        cc_recipient_nlst = retri_sql(self.p_to_database, "Officer", "For_Post")
        cc_recipient_lst = [item for sublist in cc_recipient_nlst for item in sublist]
        cc_recipient_set = set(cc_recipient_lst)
        self.T4_Add_CC_recipient_com.addItems(cc_recipient_set)
        self.T4_Add_CC_recipient_com.setCurrentText("")

    def load_cc_attention(self):
        print ("in load attention")
        self.T4_Add_CC_attention_com.clear()
        self.T4_Add_CC_attention_com.show()
        self.T4_Add_CC_attention_com.addItem( "" )
        self.T4_Add_CC_attention_com.setCurrentText("")
        For_Post = self.T4_Add_CC_recipient_com.currentText()
        cc_attention_nlst = retri_sql(self.p_to_database, "Officer", "Officer_Name" , For_Post= "'" + For_Post + "'")
        cc_attention_lst = [item for sublist in cc_attention_nlst for item in sublist]
        self.T4_Add_CC_attention_com.addItem("Nil")
        self.T4_Add_CC_attention_com.addItems(cc_attention_lst)

    def load_cc_fax(self):
        print("in load fax")
        self.T4_Add_CC_fax_com.clear()
        self.T4_Add_CC_fax_com.show()
        self.T4_Add_CC_fax_com.addItem( "" )
        self.T4_Add_CC_fax_com.setCurrentText( "" )
        For_Post = self.T4_Add_CC_recipient_com.currentText()
        cc_fax_nlst = retri_sql(self.p_to_database, "Officer", "Fax", For_Post="'" + For_Post + "'")
        cc_fax_lst = [item for sublist in cc_fax_nlst for item in sublist]
        self.T4_Add_CC_fax_com.addItem("Nil")
        self.T4_Add_CC_fax_com.addItems(cc_fax_lst)

    def enable_rad(self):
        self.T4_Add_CC_rad_frame.show()

    # def select_slope(self):
    #     select_slope_button = self.sender()
    #     index = self.T4_Add_Slope_table.indexAt(select_slope_button.pos())
    #     it = self.T4_Add_Slope_table.item(index.row(), 10)
    #     data_id = it.text()
    #     feature_no_nlist = retri_sql(self.p_to_database, "PMW_records", "Feature_No", id=str(data_id))
    #     self.added_feature = feature_no_nlist[0][0]
    #     print ("self.added_feature : ", self.added_feature)
    #     self.sel_and_toApply_s.emit()
    #     self.save = True
    #     self.close()

    def cc_execute(self):
        print ("in_cc_execute")
        if self.T4_Add_CC_recipient_com.currentText() == "":
            print ("pleace input recipient")
        elif self.T4_Add_CC_attention_com.currentText() == "":
            print("pleace input attention")
        elif self.T4_Add_CC_fax_com.currentText() == "":
            print("pleace input fax")
        else:
            print ("#### to else")
            self.isDirectClose = True
            self.sel_and_toApply_s.emit()
            self.close()

    def reset(self):
        self.T4_Add_CC_recipient_com.clear()
        self.T4_Add_CC_attention_com.clear()
        self.T4_Add_CC_attention_com.hide()
        self.T4_Add_CC_fax_com.clear()
        self.T4_Add_CC_fax_com.hide()
        self.T4_Add_CC_rad_frame.hide()

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




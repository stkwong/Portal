from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PyQt5.QtCore import pyqtSignal
from accdb_connector import retri_sql, insert_sql, update_sql_1key

class T3_Add_Upd_Caller(QDialog):

    toComplaint_s = pyqtSignal()

    def __init__(self, path):
        self.path = path
        QDialog.__init__(self)
        #loadUi(r"C:\Users\S T KWONG\AppData\Local\Programs\Python\Python39-32\Lib\site-packages\Gov_Code\Portal\Correspondence\Complaint\Add_Upd_Caller\T3_Add_Upd_Caller.ui", self)
        loadUi(path + r"\Correspondence\Complaint\Add_Upd_Caller\T3_Add_Upd_Caller.ui", self)
        self.isDirectClose = True

        # Att - but
        self.T3_Add_Upd_Caller_Retri_but.clicked.connect(self.Retri)
        self.T3_Add_Upd_Caller_Save_but.clicked.connect(self.Save)
        self.T3_Add_Upd_Caller_Add_but.clicked.connect(self.Add)
        self.T3_Add_Upd_Caller_Return_but.clicked.connect(self.toComplaint)

        # Att - le
        self.T3_Add_Upd_Caller_inc_channel_le.textEdited.connect(self.setNotDirectClose)
        self.T3_Add_Upd_Caller_Name_le.textEdited.connect(self.setNotDirectClose)
        self.T3_Add_Upd_Caller_Email_le.textEdited.connect(self.setNotDirectClose)
        self.T3_Add_Upd_Caller_ContactNo_le.textEdited.connect(self.setNotDirectClose)
        self.T3_Add_Upd_Caller_Address_le.textEdited.connect(self.setNotDirectClose)

        # Att - check
        self.T3_Add_Upd_Caller_Def_Chinese_check.stateChanged.connect(self.setNotDirectClose)

        # Att - QMessageBox
        # self.msg = self.QMessageBox()
        # self.msg.setIcon(QMessageBox.Warning)
        # self.msg.setText("Close Without Saving?")
        # self.msg.setWindowTitle("Reminder")
        # self.msg.setStandardButtons(QMessageBox.Yes|QMessageBox.No)
        # self.msg.buttonClicked.connect(self.readYN)
        #self.msg.exec_()

    def toComplaint(self):
        self.close()

    def Retri(self):
        inc_channel_input = self.T3_Add_Upd_Caller_inc_channel_le.text()
        inc_channel_input = "'" + str(inc_channel_input) + "'"
        caller_info_nlist = retri_sql(self.path, "Channel_log", "*", inc_channel= inc_channel_input )
        self.T3_Add_Upd_Caller_Name_le.setText(caller_info_nlist[0][1])
        self.T3_Add_Upd_Caller_ContactNo_le.setText(caller_info_nlist[0][2])
        self.T3_Add_Upd_Caller_Address_le.setText(caller_info_nlist[0][3])
        self.T3_Add_Upd_Caller_Fax_le.setText(caller_info_nlist[0][4])
        self.T3_Add_Upd_Caller_Email_le.setText(caller_info_nlist[0][5])
        self.T3_Add_Upd_Caller_Def_Chinese_check.setChecked(caller_info_nlist[0][6])

    def Save(self):
        if self.T3_Add_Upd_Caller_inc_channel_le.text() == "":
            print ("please input inc_channel")
        else:
            inc_channel_input = self.T3_Add_Upd_Caller_inc_channel_le.text()
            name = self.T3_Add_Upd_Caller_Name_le.text()
            contactno = self.T3_Add_Upd_Caller_ContactNo_le.text()
            address = self.T3_Add_Upd_Caller_Address_le.text()
            fax = self.T3_Add_Upd_Caller_Fax_le.text()
            email = self.T3_Add_Upd_Caller_Email_le.text()
            if self.T3_Add_Upd_Caller_Def_Chinese_check.isChecked():
                chinese = -1
            else:
                chinese = 0
            update_sql_1key(self.path, "Channel_log", "Name", "inc_channel", inc_channel_input, name)
            update_sql_1key(self.path, "Channel_log", "ContactNo", "inc_channel", inc_channel_input, contactno)
            update_sql_1key(self.path, "Channel_log", "Address", "inc_channel", inc_channel_input, address)
            update_sql_1key(self.path, "Channel_log", "Fax", "inc_channel", inc_channel_input, fax)
            update_sql_1key(self.path, "Channel_log", "Email", "inc_channel", inc_channel_input, email)
            update_sql_1key(self.path, "Channel_log", "Def_Chinese", "inc_channel", inc_channel_input, chinese)
            self.isDirectClose = True

    def Add(self):
        if self.T3_Add_Upd_Caller_inc_channel_le.text() == "":
            print("please input inc_channel")
        elif self.inc_exist():
            print ("inc_channel exists")
        else:
            inc_channel_input = self.T3_Add_Upd_Caller_inc_channel_le.text()
            name = self.T3_Add_Upd_Caller_Name_le.text()
            contactno = self.T3_Add_Upd_Caller_ContactNo_le.text()
            address = self.T3_Add_Upd_Caller_Address_le.text()
            fax = self.T3_Add_Upd_Caller_Fax_le.text()
            email = self.T3_Add_Upd_Caller_Email_le.text()
            if self.T3_Add_Upd_Caller_Def_Chinese_check.isChecked():
                chinese = -1
            else:
                chinese = 0
            insert_sql(self.path, "Channel_log", inc_channel=inc_channel_input, Name=name, ContactNo=contactno, Address=address, Fax=fax, Email=email, Def_Chinese=chinese)

    def closeEvent (self, event):
        if self.isDirectClose:
            self.toComplaint_s.emit()
            event.accept()
        else:
            # self.msg = QMessageBox()
            # self.msg.setIcon(QMessageBox.Warning)
            # self.msg.setText("Close Without Saving?")
            # self.msg.setWindowTitle("Reminder")
            # self.msg.setStandardButtons(Yes|No)

            # if self.confirm_dec() == "Y":
            # returnValue = self.msg.exec()
            returnValue = self.QMessageBox.question(self, "1", "2", QMessageBox.Yes|QMessageBox.No)
            print (returnValue)
            if returnValue == self.msg.Yes:
                print ("yes")
                self.toComplaint_s.emit()
                event.accept()
            else:
                print("no")
                event.ignore()

    def inc_exist(self):
        inc_channel_input = self.T3_Add_Upd_Caller_inc_channel_le.text()
        inc_channel_input = "'" + str(inc_channel_input) + "'"
        caller_info_nlist = retri_sql(self.path, "Channel_log", "*", inc_channel=inc_channel_input)
        if len(caller_info_nlist) == 0:
            return False
        else:
            return True

    def setNotDirectClose(self):
        self.isDirectClose = False

    def readYN(self):
        print ("in function")

    # def confirm_dec(self):
    #     self.msg = QMessageBox()
    #     self.msg.setIcon(QMessageBox.Warning)
    #     self.msg.setText("Close Without Saving?")
    #     self.msg.setWindowTitle("Reminder")
    #     self.msg.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
    #     self.msg.buttonClicked.connect(self.readYN)
    #     self.msg.activateWindow()
    #     self.msg.show()
    #     returnValue = self.msg.exec_()
    #     print (returnValue)
    #     print (QMessageBox.Yes)
    #     if returnValue == QMessageBox.Yes:
    #         print ("Yes")
    #         return "Y"
    #     else:
    #         return "N"
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PyQt5.QtCore import pyqtSignal
from accdb_connector import retri_path, insert_path, update_path_1key
from SetUp.Add_New_Path.T2_Add_New_Path import *

class T1_SetUp_Form(QDialog):

    toHome_s = pyqtSignal()

    def __init__(self, path):
        QDialog.__init__(self)
        loadUi(path + r"\SetUp\T1_SetUp.ui", self)
        self.path = path

        self.T2_Add_New_Path = T2_Add_New_Path_Form(path)
        self.T2_Add_New_Path.Save_and_CheckEI_s.connect(self.return_from_add_new_path)
        self.T2_Add_New_Path.toCheckEI_s.connect(self.return_wo_save)

        # child and return
        self.setup_home_but.clicked.connect(self.home)

        path_to_database_nlist = retri_path(path, "path_options", "saved_path", destination="'" + 'database' + "'")
        self.setup_path_to_database_com.addItems([i for lst in path_to_database_nlist for i in lst])
        path_to_documents_nlist = retri_path(path, "path_options", "saved_path", destination="'" + 'documents' + "'")
        self.setup_path_to_documents_com.addItems([i for lst in path_to_documents_nlist for i in lst])
        path_to_server_nlist = retri_path(path, "path_options", "saved_path", destination="'" + 'server' + "'")
        self.setup_path_to_server_com.addItems([i for lst in path_to_server_nlist for i in lst])

        path_to_database_nlist = retri_path(self.path, "selected_path", "sel_path", destination="'" + 'database' + "'")
        self.setup_path_to_database_com.setCurrentText(path_to_database_nlist[0][0])
        path_to_documents_nlist = retri_path(self.path, "selected_path", "sel_path", destination="'" + 'documents' + "'")
        self.setup_path_to_documents_com.setCurrentText(path_to_documents_nlist[0][0])
        path_to_server_nlist = retri_path(self.path, "selected_path", "sel_path",destination="'" + 'server' + "'")
        self.setup_path_to_server_com.setCurrentText(path_to_server_nlist[0][0])

        self.setup_path_to_database_add_new_but.clicked.connect(self.add_new_database_path)
        self.setup_path_to_documents_add_new_but.clicked.connect(self.add_new_documents_path)
        self.setup_path_to_server_add_new_but.clicked.connect(self.add_new_server_path)

        self.setup_path_to_database_to_def_but.clicked.connect(self.database_to_def)
        self.setup_path_to_documents_to_def_but.clicked.connect(self.documents_to_def)
        self.setup_path_to_server_to_def_but.clicked.connect(self.server_to_def)

        self.setup_save_but.clicked.connect(self.save)

    def add_new_database_path(self):
        self.T2_Add_New_Path.input_des = "database"
        self.T2_Add_New_Path.show()
        self.close()

    def add_new_documents_path(self):
        self.T2_Add_New_Path.input_des = "documents"
        self.T2_Add_New_Path.show()
        self.close()

    def add_new_server_path(self):
        self.T2_Add_New_Path.input_des = "server"
        self.T2_Add_New_Path.show()
        self.close()

    def home(self):
        self.toHome_s.emit()
        self.close()

    def save(self):
        update_path_1key(self.path, "selected_path", "sel_path", "destination", "database", self.setup_path_to_database_com.currentText() )
        update_path_1key(self.path, "selected_path", "sel_path", "destination", "documents", self.setup_path_to_documents_com.currentText() )
        update_path_1key(self.path, "selected_path", "sel_path", "destination", "server", self.setup_path_to_server_com.currentText() )
        print(self.setup_path_to_database_com.currentText())



    def return_from_add_new_path(self):
        if self.T2_Add_New_Path.input_des == "database":
            self.setup_path_to_database_com.addItem(self.T2_Add_New_Path.New_Path)
            self.setup_path_to_database_com.setCurrentText(self.T2_Add_New_Path.New_Path)
        elif self.T2_Add_New_Path.input_des == "documents":
            self.setup_path_to_documents_com.addItem(self.T2_Add_New_Path.New_Path)
            self.setup_path_to_documents_com.setCurrentText(self.T2_Add_New_Path.New_Path)
        elif self.T2_Add_New_Path.input_des == "server":
            self.setup_path_to_server_com.addItem(self.T2_Add_New_Path.New_Path)
            self.setup_path_to_server_com.setCurrentText(self.T2_Add_New_Path.New_Path)
        else:
            print ("self.T2_Add_New_Path.input_des != either records")
        self.show()

    def return_wo_save(self):
        self.show()

    def database_to_def(self):
        self.setup_path_to_database_com.setCurrentText("DEF")

    def documents_to_def(self):
        self.setup_path_to_documents_com.setCurrentText("DEF")

    def server_to_def(self):
        self.setup_path_to_server_com.setCurrentText("DEF")

def create_path(path):
    path_to_database_nlist = retri_path(path, "selected_path", "sel_path", destination="'" + 'database' + "'")
    if path_to_database_nlist[0][0] == "DEF":
        p_to_database = path
    else:
        p_to_database = path_to_database_nlist[0][0]

    path_to_documents_nlist = retri_path(path, "selected_path", "sel_path", destination="'" + 'documents' + "'")
    if path_to_documents_nlist[0][0] == "DEF":
        p_to_documents = path
    else:
        p_to_documents = path_to_documents_nlist[0][0]

    path_to_server_nlist = retri_path(path, "selected_path", "sel_path", destination="'" + 'server' + "'")
    if path_to_server_nlist[0][0] == "DEF":
        p_to_server = path
    else:
        p_to_server = path_to_server_nlist[0][0]

    return p_to_database, p_to_documents, p_to_server


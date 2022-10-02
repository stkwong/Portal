from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PyQt5.QtCore import pyqtSignal
from .Add_New.T3_Add_New import *
from .Edit_Com.T3_Edit_Com import *
import datetime
from accdb_connector import retri_sql, insert_sql, update_sql_1key, retri_outgoing
from Consultancy.CheckEI.Gen_Letter.T3_CheckEI_Gen_Letter import *
from SetUp.T1_SetUp import *
import pandas as pd
from pandas_to_excel import save_cluster_xlsx
import numpy as np
import matplotlib.pyplot as plt
from math import floor


class T2_CheckEI(QDialog):

    toHome_s = pyqtSignal()
    toConsult_s = pyqtSignal()

    def __init__(self, path, p_to_database, p_to_documents, p_to_server):
        self.path = path
        QDialog.__init__(self)
        loadUi(path + r"\Consultancy\CheckEI\T2_CheckEI.ui", self)
        self.isDirectClose = True
        self.p_to_database = p_to_database
        self.p_to_documents = p_to_documents
        self.p_to_server = p_to_server

        # msgbox
        self.msgBoxClose = QMessageBox()
        self.msgBoxClose.setIcon(QMessageBox.Information)
        self.msgBoxClose.setText("Close this page?")
        self.msgBoxClose.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

        self.msgBoxClear = QMessageBox()
        self.msgBoxClear.setIcon(QMessageBox.Information)
        self.msgBoxClear.setText("Clear all inputs?")
        self.msgBoxClear.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

        self.msgBoxPrompt = QMessageBox()
        self.msgBoxPrompt.setIcon( QMessageBox.Information )
        self.msgBoxPrompt.setStandardButtons( QMessageBox.Ok )

        self.msgBoxChoose = QMessageBox()
        self.msgBoxChoose.setIcon( QMessageBox.Information )
        self.msgBoxChoose.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

        # signal
        consult_no_nlist = retri_sql(self.p_to_database, "Own_Consultancy", "Agreement_No")
        consult_no_list = [item for sublist in consult_no_nlist for item in sublist]
        self.T2_CheckEI_Consultancy_Com.addItems(consult_no_list)

        # child and return
        self.T3_Add_New = T3_Add_New_Form(path, p_to_database, p_to_documents, p_to_server)
        self.T3_Add_New.Save_and_CheckEI_s.connect(self.return_from_add_new)
        self.T3_Add_New.toCheckEI_s.connect(self.return_wo_save)

        self.T3_Edit_Com = T3_Edit_Com_Form(path, p_to_database, p_to_documents, p_to_server)
        self.T3_Edit_Com.Save_and_CheckEI_s.connect(self.return_from_edit_com)
        self.T3_Edit_Com.toCheckEI_s.connect(self.return_wo_save)

        self.Gen_Letter = T3_CheckEI_Gen_Letter(path, p_to_database, p_to_documents, p_to_server)
        self.Gen_Letter.toCheckEI_s.connect(self.return_wo_save)

        self.T2_CheckEI_Save_Gen_Status_but.clicked.connect(self.gen_status)

        # msgBox
        self.msgBoxOpen = QMessageBox()
        self.msgBoxOpen.setIcon( QMessageBox.Information )
        self.msgBoxOpen.setText( "Excel created ! Need to open the excel now?" )
        self.msgBoxOpen.setStandardButtons( QMessageBox.Yes | QMessageBox.No )

        #Attributes
        self.yr = str(datetime.date.today().year)
        self.to_win = "Close"
        self.T2_CheckEI_first_draft_rad.setChecked(True)
        self.batch_lst = []
        self.add_row_but = QPushButton("Add Row")
        self.T2_CheckEI_table.setCellWidget(0, 0, self.add_row_but)
        header = self.T2_CheckEI_table.horizontalHeader()
        header.setSectionResizeMode(QHeaderView.ResizeToContents)
        header.setSectionResizeMode(8, QHeaderView.Stretch)
        self.T2_CheckEI_First_Round_check.stateChanged.connect(self.change_first_round)

        # Att - but
        self.T2_CheckEI_Save_Table_but.clicked.connect(self.save)
        self.T2_CheckEI_Save_Gen_Report_but.clicked.connect(self.save_gen_report)
        self.T2_CheckEI_insert_batch_but.clicked.connect(self.insert_batch)
        self.add_row_but.clicked.connect(self.insert_row)
        self.T2_CheckEI_retri_but.clicked.connect(self.retri_data)
        self.T2_CheckEI_Clear_All_but.clicked.connect(self.reset_table)

        #toHome
        self.T2_CheckEI_home_but.clicked.connect(self.toHome)
        self.T2_CheckEI_Return_but.clicked.connect(self.toConsult)
        self.T2_CheckEI_final_rev_le.setText("0")
        self.T2_CheckEI_First_Round_check.setChecked(False)

    def change_first_round(self):
        if self.T2_CheckEI_First_Round_check.isChecked():
            self.T2_CheckEI_remarks_le.setText("First Round EI")
        else:
            self.T2_CheckEI_remarks_le.clear()

    def toHome(self):
        self.to_win = "home"
        self.close()

    def toConsult(self):
        self.to_win = "consult"
        self.close()

    def retri_data(self):
        self.isDirectClose = False
        if len(self.batch_lst) == 0:
            self.msgBoxPrompt.setText("Please Insert Batch No.")
            self.msgBoxPrompt.exec()
        else:
            agree_no = self.T2_CheckEI_Consultancy_Com.currentText()
            batch_lst_str = "$$".join(self.batch_lst)
            if self.T2_CheckEI_final_draft_rad.isChecked():
                rev_no = "1_" + self.T2_CheckEI_final_rev_le.text()
            else:
                rev_no = "0"
            con_batch_rev = agree_no + "!!!" + batch_lst_str + "!!!" + rev_no

            CheckEI_audit_nlist = retri_sql(self.p_to_database, "CheckEI_Audit_Records", "ind_batch_no", "id_in_batch", "feature_no", "sign_of_distress", "rec_pmw", "checkEI_data_id", "incoming_ref", "incoming_date", con_batch_rev="'" + con_batch_rev + "'")

            if len(CheckEI_audit_nlist) == 0:
                self.msgBoxPrompt.setText( "No Data Is Found" )
                self.msgBoxPrompt.exec()
            else:
                ind_batch_no = CheckEI_audit_nlist[0][0]
                self.ind_batch_no_lst = ind_batch_no.split("!!")
                id_in_batch = CheckEI_audit_nlist[0][1]
                id_in_batch_lst = id_in_batch.split("!!")
                feature_no = CheckEI_audit_nlist[0][2]
                feature_no_lst = feature_no.split("!!")
                sign_of_distress = CheckEI_audit_nlist[0][3]
                sign_of_distress_lst = sign_of_distress.split("!!")
                rec_pmw = CheckEI_audit_nlist[0][4]
                rec_pmw_lst = rec_pmw.split("!!")
                checkEI_data_id = CheckEI_audit_nlist[0][5]
                checkEI_data_id_lst = checkEI_data_id.split("!!")
                print ("CheckEI_audit_nlist[0][6] : ", CheckEI_audit_nlist[0][6])
                self.T2_CheckEI_incoming_le.setText(CheckEI_audit_nlist[0][6])
                if CheckEI_audit_nlist[0][7] != "":
                    incoming_date_datetime = datetime.datetime.strptime(CheckEI_audit_nlist[0][7], '%d %B %Y')
                    self.T2_CheckEI_incoming_day_le.setText(incoming_date_datetime.strftime("%d"))
                    self.T2_CheckEI_incoming_month_le.setText(incoming_date_datetime.strftime("%m"))
                    self.T2_CheckEI_incoming_year_le.setText(incoming_date_datetime.strftime("%y"))
                else:
                    self.T2_CheckEI_incoming_day_le.setText("")
                    self.T2_CheckEI_incoming_month_le.setText("")
                    self.T2_CheckEI_incoming_year_le.setText("")

                self.T2_CheckEI_table.clearContents()
                self.T2_CheckEI_table.setRowCount(1)
                self.add_row_but = QPushButton("Add Row")
                self.T2_CheckEI_table.setCellWidget(0, 0, self.add_row_but)
                self.add_row_but.clicked.connect(self.insert_row)

                for row in range(len(checkEI_data_id_lst)):
                    self.insert_row()

                    print("self.ind_batch_no_lst[row] : ", self.ind_batch_no_lst[row])
                    if self.ind_batch_no_lst[row] == "--":
                        self.T2_CheckEI_table.removeCellWidget(row, 2)
                        it = QTableWidgetItem()
                        it.setText(self.ind_batch_no_lst[row])
                        self.T2_CheckEI_table.setItem(row, 2, it)
                    elif self.ind_batch_no_lst[row] == "":
                        pass
                    else:
                        self.add_batch_combo.setCurrentText(self.ind_batch_no_lst[row])

                    it = QTableWidgetItem()
                    it.setText(id_in_batch_lst[row])
                    self.T2_CheckEI_table.setItem(row , 3, it)

                    it = QTableWidgetItem()
                    it.setText(feature_no_lst[row])
                    self.T2_CheckEI_table.setItem(row , 4, it)

                    if sign_of_distress_lst[row] == "T":
                        self.add_distress_check.setChecked(True)
                    elif sign_of_distress_lst[row] == "F":
                        self.add_distress_check.setChecked(False)
                    else:
                        it = QTableWidgetItem()
                        it.setText("--")
                        self.T2_CheckEI_table.setItem(row, 5, it)

                    if rec_pmw_lst[row] == "T":
                        self.add_recom_PMW_check.setChecked(True)
                    elif rec_pmw_lst[row] == "F":
                        self.add_recom_PMW_check.setChecked(False)
                    else:
                        it = QTableWidgetItem()
                        it.setText("--")
                        self.T2_CheckEI_table.setItem(row, 6, it)

                    if checkEI_data_id_lst[row] != "Nil" and checkEI_data_id_lst[row] != "" and checkEI_data_id_lst[row] != "EMPTY":
                        comment_data_nlist = retri_sql(self.p_to_database, "CheckEI_Data", "Sec", "Comments" , id=checkEI_data_id_lst[row])
                        self.sec_combo.setCurrentText(comment_data_nlist[0][0])
                        self.acti_com_combo.setCurrentText(comment_data_nlist[0][1])
                    elif checkEI_data_id_lst[row] == "Nil":
                        self.sec_combo.setCurrentText("--Nil--")

    def save(self, is_gen_letter=False):
        if len( self.batch_lst ) == 0:
            self.msgBoxPrompt.setText( "Please Insert Batch No." )
            self.msgBoxPrompt.exec()
        else:
            agree_no = self.T2_CheckEI_Consultancy_Com.currentText()
            batch_lst_str = "$$".join(self.batch_lst)
            if self.T2_CheckEI_final_draft_rad.isChecked():
                rev_no = "1_" + self.T2_CheckEI_final_rev_le.text()
            else:
                rev_no = "0"
            con_batch_rev = agree_no + "!!!" + batch_lst_str + "!!!" + rev_no
            if self.T2_CheckEI_incoming_le.text() == "":
                self.msgBoxChoose.setText( "Saving without input of incoming ref?" )
                ChooseAns = self.msgBoxChoose.exec()
                if ChooseAns == QMessageBox.Yes:
                    incoming_ref = ""
                else:
                    return
            else:
                incoming_ref = self.T2_CheckEI_incoming_le.text()
            in_day = self.T2_CheckEI_incoming_day_le.text()
            in_month = self.T2_CheckEI_incoming_month_le.text()
            in_year = self.T2_CheckEI_incoming_year_le.text()
            if in_day != "" and in_month != "" and in_year != "":
                if len(in_day) < 3 and len(in_month) < 3 and len(in_year) < 3 and in_day.isdigit() and in_month.isdigit() and in_year.isdigit():
                    incoming_date_str = in_day + ":" + in_month + ":" + in_year
                    incoming_date = datetime.datetime.strptime(incoming_date_str, '%d:%m:%y').strftime('%d %B %Y')
                else:
                    self.msgBoxPrompt.setText( "Input incoming date is incorrect" )
                    self.msgBoxPrompt.exec()
                    return
            else:
                self.msgBoxChoose.setText("Saving without input of incoming date?")
                ChooseAns = self.msgBoxChoose.exec()
                if ChooseAns == QMessageBox.Yes:
                    incoming_date = ""
                else:
                    return

            remarks = self.T2_CheckEI_remarks_le.text()

            self.isDirectClose = True

            date_mod = datetime.date.today()

            ind_batch_no_lst = []
            id_in_batch_lst = []
            feature_no_lst = []
            sign_of_distress_lst = []
            rec_pmw_lst = []
            checkEI_data_id_lst = []

            i = 0
            comment_dict = {}
            for row in range(self.T2_CheckEI_table.rowCount()-1):

                it = self.T2_CheckEI_table.item(row, 4)
                if it and it.text():
                    feature_no_lst.append(it.text())
                    if it.text() != "--":
                        i += 1
                        comment_dict["sub_comment_dict_%s" % i] = {}
                        comment_dict["sub_comment_dict_%s" % i]["comments"] = []
                        comment_dict["sub_comment_dict_%s" % i]["feature_no"] = it.text()
                else:
                    continue

                it = self.T2_CheckEI_table.cellWidget(row, 2)
                if isinstance(it, QComboBox):
                    if it.currentText():
                        ind_batch_no_lst.append(it.currentText())
                        comment_dict["sub_comment_dict_%s" % i]["ind_batch_no"] = it.currentText()
                    else:
                        ind_batch_no_lst.append("Nil")
                elif not it:
                    ind_batch_no_lst.append("--")
                else:
                    ind_batch_no_lst.append(it.text())

                it = self.T2_CheckEI_table.item(row, 3)
                if it and it.text():
                    id_in_batch_lst.append(it.text())
                    if it.text() != "--":
                        comment_dict["sub_comment_dict_%s" % i]["id_in_batch"] = it.text()
                else:
                    id_in_batch_lst.append("Nil")

                it = self.T2_CheckEI_table.cellWidget(row, 5)
                if isinstance(it, QCheckBox):
                    if it.isChecked():
                        sign_of_distress_lst.append("T")
                    else:
                        sign_of_distress_lst.append("F")
                elif not it:
                    sign_of_distress_lst.append("Nil")
                else:
                    sign_of_distress_lst.append(it.text())

                it = self.T2_CheckEI_table.cellWidget(row, 6)
                if isinstance(it, QCheckBox):
                    if it.isChecked():
                        rec_pmw_lst.append("T")
                    else:
                        rec_pmw_lst.append("F")
                elif not it:
                    rec_pmw_lst.append("Nil")
                else:
                    rec_pmw_lst.append(it.text())

                it_sec = self.T2_CheckEI_table.cellWidget(row, 7)
                it_com = self.T2_CheckEI_table.cellWidget(row, 8)
                if it_sec and it_sec.currentText() and it_com and it_com.currentText():
                    if it_sec.currentText() == "--Nil--":
                        comment_dict["sub_comment_dict_%s" % i]["comments"].append( " Nil Comment " )
                        checkEI_data_id_lst.append("Nil")
                    else:
                        if it_com.currentText() == "--Nil--":
                            checkEI_data_id_lst.append("Nil")
                        else:
                            id_nlist = retri_sql(self.p_to_database, "CheckEI_Data", "id", Sec="'" + it_sec.currentText() + "'", Comments= "'" + it_com.currentText() + "'")
                            checkEI_data_id_lst.append(str(id_nlist[0][0]))
                            num = len(comment_dict["sub_comment_dict_%s" % i]["comments"])
                            comment_dict["sub_comment_dict_%s" % i]["comments"].append( "(" + str(num + 1) + ") " + it_sec.currentText() + " - " + it_com.currentText())
                else:
                    checkEI_data_id_lst.append("EMPTY")

            date_comment = "--"

            if is_gen_letter:
                gen_dict = {}
                cur_my_post_nlist = retri_sql(self.p_to_database, "My_Post", "Department", "Section", "Post" , "Senior", "STO", "TO")
                cur_post = cur_my_post_nlist[0][2]
                cur_working_sec = cur_my_post_nlist[0][0] + "_"+ cur_my_post_nlist[0][1]
                my_senior_post = cur_my_post_nlist[0][3]
                my_STO_post = cur_my_post_nlist[0][4]
                my_TO_post = cur_my_post_nlist[0][5]

                posting_nlist = retri_sql(self.p_to_database, cur_working_sec, "P_Name" , "Tel", "Email", "Fax" ,  Post= "'" + cur_post + "'")
                my_name = posting_nlist[0][0]
                my_phone = posting_nlist[0][1]
                my_email = posting_nlist[0][2]
                my_fax = posting_nlist[0][3]

                my_initial = self.to_initial(my_name)
                my_initial_s = my_initial.lower()

                posting_nlist = retri_sql(self.p_to_database, cur_working_sec, "P_Name" , Post="'" + my_senior_post + "'")
                my_senior_name = posting_nlist[0][0]
                my_senior_initial = self.to_initial(my_senior_name)

                consult_nlist = retri_sql(self.p_to_database, "Own_Consultancy", "Fax" , "Address_1", "Address_2", "Address_3", "Address_4" , "Main_Contact_Name", "Main_Contact_Post", "Title_1", "Title_2", "EI_file", Agreement_No= "'" + agree_no + "'")
                consult_fax = consult_nlist[0][0]
                own_consult_address_1 = consult_nlist[0][1]
                own_consult_address_2 = consult_nlist[0][2]
                own_consult_address_3 = consult_nlist[0][3]
                own_consult_address_4 = consult_nlist[0][4]
                own_consult_main_contact_name = consult_nlist[0][5]
                own_consult_main_contact_post = consult_nlist[0][6]
                own_consult_agreement_title_1 = consult_nlist[0][7]
                own_consult_agreement_title_2 = consult_nlist[0][8]
                our_file_ref = consult_nlist[0][9]


                if len(self.batch_lst) > 1:
                    batch_lst_wo_last_lst = ["MMEI" + self.batch_lst[i] for i in range(len(self.batch_lst) - 1)]
                    last_batch = "MMEI" + self.batch_lst[-1]
                    batch_nos_show = " ,".join(batch_lst_wo_last_lst) + " and " + last_batch
                else:
                    batch_nos_show = "MMEI" + "".join(self.batch_lst)

                date_comment = date_mod.strftime( "%m/%d/%Y" )
                gen_dict["my_phone"] = my_phone
                gen_dict["my_fax"] = my_fax
                gen_dict["my_email"] = my_email
                gen_dict["our_file_ref"] = our_file_ref
                gen_dict["fax"] = consult_fax
                gen_dict["date"] = str(date_mod)
                gen_dict["own_consult_address_1"] = own_consult_address_1
                gen_dict["own_consult_address_2"] = own_consult_address_2
                gen_dict["own_consult_address_3"] = own_consult_address_3
                gen_dict["own_consult_address_4"] = own_consult_address_4
                gen_dict["own_consult_main_contact_name"] = own_consult_main_contact_name
                gen_dict["own_consult_main_contact_post"] = own_consult_main_contact_post
                gen_dict["own_consult_no"] = agree_no
                gen_dict["own_consult_agreement_title_1"] = own_consult_agreement_title_1
                gen_dict["own_consult_agreement_title_2"] = own_consult_agreement_title_2
                gen_dict["batch_nos_show"] = batch_nos_show
                gen_dict["rev_no"] = rev_no
                gen_dict["no_of_feature"] = str(i)
                gen_dict["my＿name"] = my＿name
                gen_dict["my_STO_post"] = my_STO_post
                gen_dict["my_TO_post"] = my_TO_post
                gen_dict["my_initial"] = my_initial
                gen_dict["my_senior_initial"] = my_senior_initial
                gen_dict["my_initial_s"] = my_initial_s
                gen_dict[ "incoming_ref" ] = incoming_ref
                gen_dict[ "is_first_round"] = self.T2_CheckEI_First_Round_check.isChecked()

                self.Gen_Letter.create_New(gen_dict, comment_dict)
                self.Gen_Letter.show()
                self.close()

            ind_batch_no = "!!".join(ind_batch_no_lst)
            id_in_batch = "!!".join(id_in_batch_lst)
            feature_no = "!!".join(feature_no_lst)
            sign_of_distress = "!!".join(sign_of_distress_lst)
            rec_pmw = "!!".join(rec_pmw_lst)
            checkEI_data_id = "!!".join(checkEI_data_id_lst)

            CheckEI_audit_nlist = retri_sql(self.p_to_database, "CheckEI_Audit_Records", "*", con_batch_rev= "'" + con_batch_rev + "'")
            if len(CheckEI_audit_nlist) == 0:
                insert_sql(self.p_to_database, "CheckEI_Audit_Records", con_batch_rev=con_batch_rev, date_mod=date_mod, ind_batch_no=ind_batch_no, id_in_batch=id_in_batch, feature_no=feature_no, sign_of_distress=sign_of_distress , rec_pmw=rec_pmw, checkEI_data_id=checkEI_data_id, incoming_ref=incoming_ref, incoming_date=incoming_date, remarks=remarks, date_comment=date_comment)
                self.msgBoxPrompt.setText( "Saved Successfully!" )
                self.msgBoxPrompt.exec()
                self.msgBoxChoose.setText( "Do you want to clear the existing input?" )
                ChooseAns = self.msgBoxChoose.exec()
                if ChooseAns == QMessageBox.Yes:
                    self.reset_table()
            else:
                self.msgBoxChoose.setText( "Existing data is found, Are you sure overwriting existing data?" )
                ChooseAns = self.msgBoxChoose.exec()
                if ChooseAns == QMessageBox.Yes:
                    update_sql_1key( self.p_to_database , "CheckEI_Audit_Records" , "date_mod" , "con_batch_rev" , con_batch_rev , date_mod )
                    update_sql_1key( self.p_to_database , "CheckEI_Audit_Records" , "ind_batch_no" , "con_batch_rev" , con_batch_rev , ind_batch_no )
                    update_sql_1key( self.p_to_database , "CheckEI_Audit_Records" , "id_in_batch" , "con_batch_rev" , con_batch_rev , id_in_batch )
                    update_sql_1key( self.p_to_database , "CheckEI_Audit_Records" , "feature_no" , "con_batch_rev" , con_batch_rev , feature_no )
                    update_sql_1key( self.p_to_database , "CheckEI_Audit_Records" , "sign_of_distress" , "con_batch_rev" , con_batch_rev , sign_of_distress )
                    update_sql_1key( self.p_to_database , "CheckEI_Audit_Records" , "rec_pmw" , "con_batch_rev" , con_batch_rev , rec_pmw )
                    update_sql_1key( self.p_to_database , "CheckEI_Audit_Records" , "checkEI_data_id" , "con_batch_rev" , con_batch_rev , checkEI_data_id )
                    update_sql_1key( self.p_to_database , "CheckEI_Audit_Records" , "incoming_ref" , "con_batch_rev" , con_batch_rev , incoming_ref )
                    update_sql_1key( self.p_to_database , "CheckEI_Audit_Records" , "incoming_date" , "con_batch_rev" , con_batch_rev , incoming_date )
                    update_sql_1key( self.p_to_database , "CheckEI_Audit_Records" , "date_comment" , "con_batch_rev" , con_batch_rev , date_comment )
                    self.msgBoxChoose.setText( "Do you want to clear the existing input?" )
                    ChooseAns = self.msgBoxChoose.exec()
                    if ChooseAns == QMessageBox.Yes:
                        self.reset_table()
                else:
                    return
        self.isDirectClose = True

    def gen_status(self):
        data_nlist = retri_outgoing( self.p_to_database , "CheckEI_Audit_Records" , "con_batch_rev" , "DESC" , "con_batch_rev", "incoming_ref", "incoming_date", "date_comment", "Remarks")

        batch_dict_draft = {}
        batch_dict_final = {}
        batch_lst = []

        for i in range(len(data_nlist)):
            con_batch_rev = data_nlist[ i ][ 0 ].split("!!!")
            consult_no = con_batch_rev[ 0 ]

            if consult_no == "CE 27/2019 (GE)":
                batch_group_lst = con_batch_rev[ 1 ].split("$$")

                for batch_no in batch_group_lst:
                    rev_no_str = con_batch_rev[ 2 ]
                    incoming_ref = data_nlist[ i ][ 1 ]
                    incoming_date = data_nlist[ i ][ 2 ]
                    date_comment = data_nlist[ i ][ 3 ]
                    remarks = data_nlist[ i ][ 4 ]

                    batch_lst.append(int(batch_no))

                    if "IEI" in remarks:
                        print ("remarks : ", remarks)
                        batch_dict_draft[ "MMEI%s" % batch_no ] = [ incoming_ref , incoming_date , "NA" , "NA" ]
                        batch_dict_final[ "MMEI%s" % batch_no ] = [ "NA" , "NA" , "NA" , remarks ]
                    else:
                        if rev_no_str == "0":
                            batch_dict_draft[ "MMEI%s" % batch_no ] = [incoming_ref, incoming_date, date_comment, remarks]
                        else:
                            batch_dict_final[ "MMEI%s" % batch_no ] = [incoming_ref, incoming_date, date_comment, remarks]

        batch_no_record_lst = [ ]
        draft_ref_no_record_lst = [ ]
        draft_ref_date_record_lst = [ ]
        draft_date_comment_record_lst = [ ]
        draft_remark_record_lst = [ ]
        final_ref_no_record_lst = [ ]
        final_ref_date_record_lst = [ ]
        final_date_comment_record_lst = [ ]
        final_remark_record_lst = [ ]

        for batch_num27 in range(max(batch_lst)):

            batch_lst_no = "MMEI" + str(batch_num27).zfill(4)
            if batch_lst_no in batch_dict_draft and batch_lst_no in batch_dict_final:
                batch_no_record_lst.append(batch_lst_no)
                draft_ref_no_record_lst.append(batch_dict_draft[batch_lst_no][0])
                draft_ref_date_record_lst.append(batch_dict_draft[batch_lst_no][1])
                draft_date_comment = batch_dict_draft[batch_lst_no][2]
                if draft_date_comment == "--":
                    draft_date_comment_record_lst.append("")
                else:
                    draft_date_comment_record_lst.append(draft_date_comment)
                draft_remark_record_lst.append(batch_dict_draft[batch_lst_no][3])
                final_ref_no_record_lst.append(batch_dict_final[batch_lst_no][0])
                final_ref_date_record_lst.append(batch_dict_final[batch_lst_no][1])
                final_date_comment = batch_dict_draft[ batch_lst_no ][ 2 ]
                if final_date_comment == "--":
                    final_date_comment_record_lst.append( "" )
                else:
                    final_date_comment_record_lst.append( final_date_comment )
                final_remark_record_lst.append(batch_dict_final[batch_lst_no][3])

            elif batch_lst_no in batch_dict_draft:
                batch_no_record_lst.append(batch_lst_no)
                draft_ref_no_record_lst.append(batch_dict_draft[ batch_lst_no ][ 0 ])
                draft_ref_date_record_lst.append(batch_dict_draft[ batch_lst_no ][ 1 ])
                draft_date_comment = batch_dict_draft[ batch_lst_no ][ 2 ]
                if draft_date_comment == "--":
                    draft_date_comment_record_lst.append( "" )
                else:
                    draft_date_comment_record_lst.append( draft_date_comment )
                draft_remark_record_lst.append(batch_dict_draft[ batch_lst_no ][ 3 ])
                final_ref_no_record_lst.append("")
                final_ref_date_record_lst.append("")
                final_date_comment_record_lst.append("")
                final_remark_record_lst.append("")

            elif batch_lst_no in batch_dict_final:
                batch_no_record_lst.append(batch_lst_no)
                draft_ref_no_record_lst.append("")
                draft_ref_date_record_lst.append("")
                draft_date_comment_record_lst.append("")
                draft_remark_record_lst.append("")
                final_ref_no_record_lst.append(batch_dict_final[ batch_lst_no ][ 0 ])
                final_ref_date_record_lst.append(batch_dict_final[ batch_lst_no ][ 1 ])
                final_date_comment = batch_dict_final[ batch_lst_no ][ 2 ]
                if final_date_comment == "--":
                    final_date_comment_record_lst.append( "" )
                else:
                    final_date_comment_record_lst.append( final_date_comment )
                final_remark_record_lst.append(batch_dict_final[ batch_lst_no ][ 3 ])

            else:
                batch_no_record_lst.append(batch_lst_no)
                draft_ref_no_record_lst.append("")
                draft_ref_date_record_lst.append("")
                draft_date_comment_record_lst.append("")
                draft_remark_record_lst.append("")
                final_ref_no_record_lst.append("")
                final_ref_date_record_lst.append("")
                final_date_comment_record_lst.append("")
                final_remark_record_lst.append("")


        data_dict = {"BatchNo." : batch_no_record_lst,
                     "Draft Ref No." : draft_ref_no_record_lst,
                     "Draft Sub Date" : draft_ref_date_record_lst,
                     "Draft Com Date" : draft_date_comment_record_lst,
                     "Draft Remarks" : draft_remark_record_lst,
                     "Final Ref No." : final_ref_no_record_lst,
                     "Final Sub Date": final_ref_date_record_lst ,
                     "Final Com Date" : final_date_comment_record_lst,
                     "Final Remarks" : final_remark_record_lst}

        all_data_df = pd.DataFrame (data_dict)

        path_to_file = self.p_to_server + "\\unit1\\CE 27_2019 (GE)\\EI_RMI\\EI (by GE)\\Comments on Draft EI\\Submission Status\\recorded on " + datetime.date.today().strftime("%d_%m_%Y") + ".xlsx"
        writer = pd.ExcelWriter( path_to_file , engine="xlsxwriter" )
        all_data_df.to_excel( writer , sheet_name= "on" + datetime.date.today().strftime("%d_%m_%Y") )
        writer.close()

        openAns = self.msgBoxOpen.exec()
        if openAns == QMessageBox.Yes:
            os.startfile( path_to_file )

    def to_initial(self, name ):
        name_lst = name.split(" ")
        initial_lst = [name_lst[i][0] for i in range(len(name_lst))]
        initial = "".join(initial_lst)
        return initial

    def save_gen_report(self):
        self.save(True)
        self.isDirectClose = True

    def insert_batch(self):
        if self.T2_CheckEI_batch_le.text() == "":
            self.msgBoxPrompt.setText("Please Input Batch No.")
            self.msgBoxPrompt.exec()
        elif len(self.T2_CheckEI_batch_le.text()) != 4:
            self.msgBoxPrompt.setText( "Input Batch No. is incorrect" )
            self.msgBoxPrompt.exec()
        elif self.T2_CheckEI_batch_le.text() in self.batch_lst:
            self.msgBoxPrompt.setText( "Inserted Batch No. is Duplicated" )
            self.msgBoxPrompt.exec()
        else:
            self.batch_lst.append(self.T2_CheckEI_batch_le.text())
        self.T2_CheckEI_Batch_lst_lab.setText(", ".join(self.batch_lst))

    def insert_row(self):
        if len(self.batch_lst) == 0:
            self.msgBoxPrompt.setText( "Please Input Batch No." )
            self.msgBoxPrompt.exec()
        else:
            rowPosition = self.T2_CheckEI_table.rowCount()
            self.T2_CheckEI_table.insertRow(rowPosition)
            self.add_row_but = QPushButton("Add Slope")
            self.add_row_but.clicked.connect(self.insert_row)
            self.T2_CheckEI_table.setCellWidget(rowPosition, 0, self.add_row_but)
            self.del_row_but = QPushButton("Delete Row")
            self.del_row_but.clicked.connect(self.del_row)
            self.T2_CheckEI_table.setCellWidget(rowPosition - 1, 0, self.del_row_but)
            self.add_com_but = QPushButton("Add Com.")
            self.add_com_but.clicked.connect(self.add_com)
            self.T2_CheckEI_table.setCellWidget(rowPosition - 1, 1, self.add_com_but)
            self.add_batch_combo = QComboBox(self)
            self.add_batch_combo.addItems(self.batch_lst)
            self.T2_CheckEI_table.setCellWidget(rowPosition - 1, 2, self.add_batch_combo)
            self.add_distress_check = QCheckBox(self)
            self.T2_CheckEI_table.setCellWidget(rowPosition - 1, 5, self.add_distress_check)
            self.add_recom_PMW_check = QCheckBox(self)
            self.T2_CheckEI_table.setCellWidget(rowPosition - 1, 6, self.add_recom_PMW_check)
            self.add_sec_combo(rowPosition - 1)
            self.T2_CheckEI_table.setCellWidget(rowPosition - 1, 7, self.sec_combo)
            self.add_new_but = QPushButton("Add New")
            self.add_new_but.clicked.connect(self.open_add_new)
            self.T2_CheckEI_table.setCellWidget(rowPosition - 1, 9, self.add_new_but)
            self.edit_com_but = QPushButton("Edit Com.")
            self.edit_com_but.clicked.connect(self.open_edit_com)
            self.T2_CheckEI_table.setCellWidget(rowPosition - 1, 10, self.edit_com_but)

    def activate_comment_combo(self):
        combobox = self.sender()
        acti_com_row = combobox.property("row")
        it = self.T2_CheckEI_table.cellWidget(acti_com_row, 7)
        if it:
            self.acti_com_combo = QComboBox(self)
            if it.currentText() != "--Nil--":
                self.acti_com_combo = QComboBox(self)
                Com_nlist = retri_sql(self.p_to_database, "CheckEI_Data", "Comments", Sec= "'" + it.currentText() + "'")
                Com_lst = set([i for lst in Com_nlist for i in lst])
                self.acti_com_combo.addItems(Com_lst)
            self.acti_com_combo.addItem("--Nil--")
            self.acti_com_combo.setCurrentText("--Nil--")
            self.T2_CheckEI_table.setCellWidget(acti_com_row, 8, self.acti_com_combo)

    def return_wo_save(self):
        self.show()

    def add_com(self):
        add_com_button = self.sender()
        index = self.T2_CheckEI_table.indexAt(add_com_button.pos())
        add_com_row = index.row() + 1
        it = self.T2_CheckEI_table.item(add_com_row-1, 4)
        if it and it.text():
            self.T2_CheckEI_table.insertRow(add_com_row)
            self.del_row_but = QPushButton("Delete Row")
            self.del_row_but.clicked.connect(self.del_row)
            self.T2_CheckEI_table.setCellWidget(add_com_row, 0, self.del_row_but)
            self.add_com_but = QPushButton("Add Com.")
            self.add_com_but.clicked.connect(self.add_com)
            self.T2_CheckEI_table.setCellWidget(add_com_row, 1, self.add_com_but)
            self.add_sec_combo(add_com_row)
            self.T2_CheckEI_table.setCellWidget(add_com_row, 7, self.sec_combo)
            self.add_new_but = QPushButton("Add New")
            self.add_new_but.clicked.connect(self.open_add_new)
            self.T2_CheckEI_table.setCellWidget(add_com_row, 9, self.add_new_but)
            self.edit_com_but = QPushButton("Edit Com")
            self.edit_com_but.clicked.connect(self.open_edit_com)
            self.T2_CheckEI_table.setCellWidget(add_com_row, 10, self.edit_com_but)
            for i in range(2 , 8):
                it = QTableWidgetItem()
                it.setText("--")
                self.T2_CheckEI_table.setItem(add_com_row, i, it)

    def add_sec_combo(self, row):
        self.sec_combo = QComboBox(self)
        Sec_nlist = retri_sql(self.p_to_database, "CheckEI_Data", "Sec")
        Sec_lst = set([i for lst in Sec_nlist for i in lst])
        self.sec_combo.addItem("")
        self.sec_combo.addItem("--Nil--")
        self.sec_combo.addItems(Sec_lst)
        self.sec_combo.setCurrentText("")
        self.sec_combo.setProperty("row", row)
        self.sec_combo.currentTextChanged.connect(self.activate_comment_combo)

    def open_add_new(self):
        add_new_button = self.sender()
        index = self.T2_CheckEI_table.indexAt(add_new_button.pos())
        self.add_new_row = index.row()
        self.T3_Add_New.show()
        self.isDirectClose = True
        self.close()

    def open_edit_com(self):
        edit_com_button = self.sender()
        index = self.T2_CheckEI_table.indexAt(edit_com_button.pos())
        self.edit_com_row = index.row()
        it_sec = self.T2_CheckEI_table.cellWidget(self.edit_com_row, 7)
        it_com = self.T2_CheckEI_table.cellWidget(self.edit_com_row, 8)
        if it_sec and it_sec.currentText() and it_sec.currentText() != "--Nil--" and it_com and it_com.currentText() and it_com.currentText() != "--Nil--":
            id_nlist = retri_sql(self.p_to_database, "CheckEI_Data", "id", Sec="'" + it_sec.currentText() + "'", Comments="'" + it_com.currentText() + "'")
            self.T3_Edit_Com.retri_data(str(id_nlist[0][0]))
            self.T3_Edit_Com.show()
            self.isDirectClose = True
            self.close()
        else:
            self.msgBoxPrompt.setText( "Problem encountered for retrieving comment data " )
            self.msgBoxPrompt.exec()

    def return_from_add_new(self):
        self.show()
        self.add_sec_combo(self.add_new_row)
        self.sec_combo.setCurrentText( self.T3_Add_New.Section_text )
        self.T2_CheckEI_table.setCellWidget(self.add_new_row, 7, self.sec_combo)
        self.acti_com_combo.addItem(self.T3_Add_New.Comments_text)
        self.acti_com_combo.setCurrentText(self.T3_Add_New.Comments_text)

    def return_from_edit_com(self):
        self.show()
        self.add_sec_combo(self.edit_com_row)
        self.T2_CheckEI_table.setCellWidget(self.edit_com_row, 7, self.sec_combo)
        self.sec_combo.setCurrentText(self.T3_Edit_Com.Section_text)
        self.acti_com_combo.addItem(self.T3_Edit_Com.Comments_text)
        self.acti_com_combo.setCurrentText(self.T3_Edit_Com.Comments_text)

    def del_row(self):
        del_button = self.sender()
        index = self.T2_CheckEI_table.indexAt(del_button.pos())
        del_row = index.row()
        self.T2_CheckEI_table.removeRow(del_row)

    def reset_table(self):
        clearAns = self.msgBoxClear.exec()
        if clearAns == QMessageBox.Yes:
            self.T2_CheckEI_table.clearContents()
            self.T2_CheckEI_table.setRowCount(1)
            self.add_row_but = QPushButton("Add Row")
            self.T2_CheckEI_table.setCellWidget(0, 0, self.add_row_but)
            self.add_row_but.clicked.connect(self.insert_row)
            self.batch_lst = []
            self.T2_CheckEI_incoming_le.clear()
            self.T2_CheckEI_incoming_day_le.clear()
            self.T2_CheckEI_incoming_month_le.clear()
            self.T2_CheckEI_incoming_year_le.clear()
            self.T2_CheckEI_final_rev_le.setText("0")
            self.T2_CheckEI_first_draft_rad.setChecked(True)
            self.T2_CheckEI_batch_le.clear()
            self.T2_CheckEI_Batch_lst_lab.clear()
            self.T2_CheckEI_remarks_le.clear()
            self.T2_CheckEI_First_Round_check.setChecked(False)

    def closeEvent (self, event):
        if self.isDirectClose:
            if self.to_win == "home":
                self.toHome_s.emit()
            elif self.to_win == "consult":
                self.toConsult_s.emit()
            event.accept()
        else:
            closeAns = self.msgBoxClose.exec()
            if closeAns == QMessageBox.Yes:
                self.reset_table()
                if self.to_win == "home":
                    self.toHome_s.emit()
                elif self.to_win == "consult":
                    self.toConsult_s.emit()
                event.accept()
            else:
                event.ignore()

    # def plot(self):
    #     left , width = 0.05 , 0.9
    #     bottom , height = 0.05 , 0.37
    #     spacing = 0.06
    #     rect_plotA = [ left , bottom , width , height ]
    #     rect_plotB = [ left , bottom + spacing + height , width , height ]
    #     fig = plt.figure( figsize=(10 , 10) )
    #     ax_plotA = fig.add_axes( rect_plotA )
    #     ax_plotB = fig.add_axes( rect_plotB , sharex=ax_plotA )
    #     fig.suptitle( 'Status of MM/EI Submission and Audit' , fontsize=20 )

from __future__ import print_function
from mailmerge import MailMerge
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PyQt5.QtCore import pyqtSignal
from accdb_connector import insert_sql, retri_sql, update_sql_1key
import datetime
import os
from docx2pdf import convert


class T3_CheckEI_Gen_Letter(QDialog):

    toCheckEI_s = pyqtSignal()

    def __init__(self, path, p_to_database, p_to_documents, p_to_server):

        QDialog.__init__(self)
        self.path = path
        loadUi(path + r"\Consultancy\CheckEI\Gen_Letter\T3_CheckEI_Gen_Letter.ui", self)
        self.T3_CheckEI_Gen_Letter_Execute_but.clicked.connect(self.execute)
        self.T3_CheckEI_Gen_Letter_Return_but.clicked.connect(self.toCheckEI)
        self.T3_CheckEI_Gen_Letter_First_Round_rad.toggled.connect(self.create_docu_title)
        self.T3_CheckEI_Gen_Letter_Rev_No_le.textChanged.connect(self.create_docu_title)

        self.p_to_database = p_to_database
        self.p_to_server = p_to_server

        # msgBox
        self.msgBoxOpen = QMessageBox()
        self.msgBoxOpen.setIcon(QMessageBox.Information)
        self.msgBoxOpen.setText("Letter created ! Need to open the letter now?")
        self.msgBoxOpen.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

        # msgBox
        self.msgBoxOverwrite = QMessageBox()
        self.msgBoxOverwrite.setIcon(QMessageBox.Information)
        self.msgBoxOverwrite.setText("Overwrite existing document?")
        self.msgBoxOverwrite.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

    def create_New(self, gen_dict, comment_dict):
        self.T3_CheckEI_Gen_Letter_My_Phone_le.setText(gen_dict["my_phone"])
        self.T3_CheckEI_Gen_Letter_My_Fax_le.setText(gen_dict["my_fax"])
        self.T3_CheckEI_Gen_Letter_My_Email_le.setText(gen_dict["my_email"])
        self.T3_CheckEI_Gen_Letter_Our_File_Ref.setText(gen_dict["our_file_ref"])
        self.T3_CheckEI_Gen_Letter_Date_le.setText(gen_dict["date"])
        self.T3_CheckEI_Gen_Letter_Fax_le.setText(gen_dict["fax"])
        self.T3_CheckEI_Gen_Letter_Own_Consult_Address_1_le.setText(gen_dict["own_consult_address_1"])
        self.T3_CheckEI_Gen_Letter_Own_Consult_Address_2_le.setText(gen_dict["own_consult_address_2"])
        self.T3_CheckEI_Gen_Letter_Own_Consult_Address_3_le.setText(gen_dict["own_consult_address_3"])
        self.T3_CheckEI_Gen_Letter_Own_Consult_Address_4_le.setText(gen_dict["own_consult_address_4"])
        self.T3_CheckEI_Gen_Letter_Own_Consult_Main_Contact_Name_le.setText(gen_dict["own_consult_main_contact_name"])
        self.T3_CheckEI_Gen_Letter_Own_Consult_Main_Contact_Post.setText(gen_dict["own_consult_main_contact_post"])
        self.T3_CheckEI_Gen_Letter_Own_Consult_Agreement_No.setText(gen_dict["own_consult_no"])
        self.T3_CheckEI_Gen_Letter_Consult_Agreement_Title_1_le.setText(gen_dict["own_consult_agreement_title_1"])
        self.T3_CheckEI_Gen_Letter_Consult_Agreement_Title_2_le.setText(gen_dict["own_consult_agreement_title_2"])
        self.T3_CheckEI_Gen_Letter_Batch_Nos.setText(gen_dict["batch_nos_show"])
        self.T3_CheckEI_Gen_Letter_Rev_No_le.setText(gen_dict["rev_no"])
        self.T3_CheckEI_Gen_Letter_No_of_Feature_le.setText(gen_dict["no_of_feature"])
        self.T3_CheckEI_Gen_Letter_My_Name_le.setText(gen_dict["my＿name"])
        self.T3_CheckEI_Gen_Letter_STO_Post_le.setText(gen_dict["my_STO_post"])
        self.T3_CheckEI_Gen_Letter_TO_Post_le.setText(gen_dict["my_TO_post"])
        self.T3_CheckEI_Gen_Letter_My_Initial_le.setText(gen_dict["my_initial"])
        self.T3_CheckEI_Gen_Letter_My_Senior_Initial_le.setText(gen_dict["my_senior_initial"])
        self.T3_CheckEI_Gen_Letter_My_Initial_s_le.setText(gen_dict["my_initial_s"])
        self.T3_CheckEI_Gen_Letter_Recurrent_rad.setChecked(True)
        self.T3_CheckEI_Gen_Letter_First_Round_rad.setChecked(gen_dict["is_first_round"])
        self.T3_CheckEI_Gen_Letter_Incoming_File_Ref_le.setText(gen_dict["incoming_ref"])

        #self.T3_CheckEI_Gen_Letter_Incoming_File_Ref_le.setStyleSheet("""QLineEdit {background-color: yellow;} """)
        print ("comment_dict : ", comment_dict)
        self.dict_list = []
        self.not_all_nil = False
        for key, value in comment_dict.items():
            value["comments"] = "\n".join(value["comments"])
            if "Nil Comment" not in value["comments"]:
                self.not_all_nil = True
            self.dict_list.append(value)
        self.create_docu_title()

    def create_docu_title(self):
        if self.T3_CheckEI_Gen_Letter_Recurrent_rad.isChecked():
            recur_firstround = "Recurrent"
        else:
            recur_firstround = "First Round"
        if self.T3_CheckEI_Gen_Letter_Rev_No_le.text() == "0" or self.T3_CheckEI_Gen_Letter_Rev_No_le.text() == "":
            rev_no = "(Draft)"
        else:
            rev_no_lst = self.T3_CheckEI_Gen_Letter_Rev_No_le.text().split("_")
            rev_no = "(Final Rev. " + rev_no_lst[1] + ")"
        docx_str = "Comments on " + recur_firstround + " EI ("  " " + self.T3_CheckEI_Gen_Letter_Batch_Nos.text() + ")" + rev_no
        self.T3_Letter_Docu_Title_le.setText(docx_str)

    def execute(self):
        print ("self.not_all_nil : ", self.not_all_nil)
        if self.not_all_nil:
            template = self.p_to_database + r"\Database\Consultancy\CheckEI\CheckEI_Letter.docx"
            document = MailMerge(template)
        else:
            template = self.p_to_database + r"\Database\Consultancy\CheckEI\CheckEI_Letter_no_comment.docx"
            document = MailMerge( template )

        if self.T3_CheckEI_Gen_Letter_Rev_No_le.text() == "0":
            rev_no = "(Draft)"
        else:
            rev_no_lst = self.T3_CheckEI_Gen_Letter_Rev_No_le.text().split( "_" )
            if rev_no_lst[1] == "0" or rev_no_lst[1] == "":
                rev_no = "(Final)"
            else:
                rev_no = "(Final Rev. " + rev_no_lst[1] + ")"

        if self.T3_CheckEI_Gen_Letter_Recurrent_rad.isChecked():
            recur_firstround = "Recurrent"
        else:
            recur_firstround = "First Round"

        if "and" in self.T3_CheckEI_Gen_Letter_Batch_Nos.text():
            batch_s = "s"
        else:
            batch_s = ""

        if self.T3_CheckEI_Gen_Letter_No_of_Feature_le.text() == "1":
            No_of_Feature = "a"
            feature_s = ""
        else:
            No_of_Feature = self.T3_CheckEI_Gen_Letter_No_of_Feature_le.text()
            feature_s = "s"

        date_str = datetime.datetime.strptime(self.T3_CheckEI_Gen_Letter_Date_le.text(), '%Y-%m-%d').strftime('%d %B %Y')

        try:
            document.merge(
                My_Phone = self.T3_CheckEI_Gen_Letter_My_Phone_le.text(),
                My_Fax = self.T3_CheckEI_Gen_Letter_My_Fax_le.text(),
                My_Email = self.T3_CheckEI_Gen_Letter_My_Email_le.text(),
                Our_File_Ref = self.T3_CheckEI_Gen_Letter_Our_File_Ref.text(),
                Incoming_File_Ref = self.T3_CheckEI_Gen_Letter_Incoming_File_Ref_le.text(),
                Consult_Fax = self.T3_CheckEI_Gen_Letter_Fax_le.text(),
                Own_Consult_Address_1 = self.T3_CheckEI_Gen_Letter_Own_Consult_Address_1_le.text(),
                Own_Consult_Address_2 = self.T3_CheckEI_Gen_Letter_Own_Consult_Address_2_le.text(),
                Own_Consult_Address_3 = self.T3_CheckEI_Gen_Letter_Own_Consult_Address_3_le.text(),
                Own_Consult_Address_4=self.T3_CheckEI_Gen_Letter_Own_Consult_Address_4_le.text(),
                Own_Consult_Main_Contact_Name = self.T3_CheckEI_Gen_Letter_Own_Consult_Main_Contact_Name_le.text(),
                Own_Consult_Main_Contact_Post = self.T3_CheckEI_Gen_Letter_Own_Consult_Main_Contact_Post.text(),
                Own_Consult_Agreement_No = self.T3_CheckEI_Gen_Letter_Own_Consult_Agreement_No.text(),
                Own_Consult_Agreement_Title_1 = self.T3_CheckEI_Gen_Letter_Consult_Agreement_Title_1_le.text(),
                Own_Consult_Agreement_Title_2 = self.T3_CheckEI_Gen_Letter_Consult_Agreement_Title_2_le.text(),
                date=date_str,
                batch_s = batch_s,
                Batch_Nos = self.T3_CheckEI_Gen_Letter_Batch_Nos.text(),
                Rev_No = rev_no,
                No_of_Feature = No_of_Feature,
                feature_s = feature_s,
                My_Name = self.T3_CheckEI_Gen_Letter_My_Name_le.text(),
                STO_Post = self.T3_CheckEI_Gen_Letter_STO_Post_le.text(),
                TO_Post=self.T3_CheckEI_Gen_Letter_TO_Post_le.text(),
                My_Initial=self.T3_CheckEI_Gen_Letter_My_Initial_le.text(),
                My_Senior_Initial=self.T3_CheckEI_Gen_Letter_My_Senior_Initial_le.text(),
                My_Initial_s=self.T3_CheckEI_Gen_Letter_My_Initial_s_le.text(),
                Recur_FirstRound = recur_firstround
            )

            comment_table = self.dict_list
            document.merge_rows('ind_batch_no', comment_table)

        except:
            print ("something went wrong")

        consultant_dir = self.T3_CheckEI_Gen_Letter_Own_Consult_Agreement_No.text().replace("/", "_")
        path_to_file = self.p_to_server + "\\unit1\\" + consultant_dir + "\\" + "EI_RMI\\EI (by GE)\\Comments on Draft EI" + "\\" + self.T3_Letter_Docu_Title_le.text() + ".docx"
        path_to_pdf = self.p_to_server + "\\unit1\\" + consultant_dir + "\\" + "EI_RMI\\EI (by GE)\\Comments on Draft EI" + "\\Comments in PDF\\" + self.T3_Letter_Docu_Title_le.text() + ".pdf"
        print ("###path to file = ", path_to_file)
        now_time = datetime.datetime.now().time()
        now_time_str = now_time.strftime("%H:%M:%S")
        now_date = datetime.date.today().strftime("%d/%m/%Y")

        try:
            with open(path_to_file) as f:
                OverwriteAns = self.msgBoxOverwrite.exec()
                if OverwriteAns == QMessageBox.Yes:
                    document.write(path_to_file)
                    id_nlst = retri_sql(self.p_to_database, "Outgoing_records", "id",
                                        path_to_file="'" + path_to_file + "'")

                    update_sql_1key(self.p_to_database, "Outgoing_records", "date_mod", "id", id_nlst[0][0], now_date)
                    update_sql_1key(self.p_to_database, "Outgoing_records", "time_mod", "id", id_nlst[0][0], now_time_str)

                    convert(path_to_file, path_to_pdf)

        except IOError:
            document.write(path_to_file)

            file_ref = self.T3_CheckEI_Gen_Letter_Our_File_Ref.text()
            recipient = self.T3_CheckEI_Gen_Letter_Own_Consult_Address_1_le.text()

            insert_sql(self.p_to_database, "Outgoing_records", date_mod=now_date, time_mod=now_time_str, type="Check_EI", identity1=self.T3_CheckEI_Gen_Letter_Batch_Nos.text() + "_" + rev_no, identity2="--", recipient=recipient, T1="Consult", T2="CheckEI", T3="--", T4="--", file_ref=file_ref , path_to_file=path_to_file)

            convert(path_to_file, path_to_pdf)

        openAns = self.msgBoxOpen.exec()
        if openAns == QMessageBox.Yes:
            os.startfile(path_to_file)

        self.toCheckEI()

    def toCheckEI(self):
        self.hide()
        self.toCheckEI_s.emit()


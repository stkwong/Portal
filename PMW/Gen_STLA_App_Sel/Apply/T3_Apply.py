from __future__ import print_function
from mailmerge import MailMerge
from PyQt5.QtWidgets import *
from PyQt5.uic import loadUi
from PyQt5.QtCore import pyqtSignal
from accdb_connector import retri_sql, insert_sql, update_sql_1key
from PMW.Gen_STLA_App_Sel.Apply.Add_Slope.T4_Add_Slope import *
from PMW.Gen_STLA_App_Sel.Apply.Add_CC.T4_Add_CC import *
from PMW.Gen_STLA_App_Sel.Apply.Add_CCF.T4_Add_CCF import *
import datetime
import os
import calendar


class T3_Apply_Form(QDialog):

    toPMW_s = pyqtSignal()
    toSTLA_Sel_s = pyqtSignal()

    def __init__(self, path, p_to_database, p_to_documents, p_to_server):

        QDialog.__init__(self)
        self.path = path
        loadUi(path + r"\PMW\Gen_STLA_App_Sel\Apply\T3_Apply.ui", self)

        self.T3_Apply_Execute_but.clicked.connect(self.execute)
        self.T3_Apply_Return_but.clicked.connect(self.toSTLA_Sel)
        self.T3_Apply_Add_Slope_but.clicked.connect(self.add_slope)
        self.T3_Apply_Add_Attention_but.clicked.connect(self.add_attention)
        self.T3_Apply_Reload_but.clicked.connect(self.load_feature)
        self.T3_Apply_Add_CC_but.clicked.connect(self.add_CC)
        self.T3_Apply_Add_CCF_but.clicked.connect( self.add_CCF )
        self.p_to_database = p_to_database
        self.p_to_server = p_to_server

        # msgBox
        self.msgBoxOpen = QMessageBox()
        self.msgBoxOpen.setIcon(QMessageBox.Information)
        self.msgBoxOpen.setText("Memo created ! Need to open the memo now?")
        self.msgBoxOpen.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

        # msgBox
        self.msgBoxOverwrite = QMessageBox()
        self.msgBoxOverwrite.setIcon(QMessageBox.Information)
        self.msgBoxOverwrite.setText("Overwrite existing document?")
        self.msgBoxOverwrite.setStandardButtons(QMessageBox.Yes | QMessageBox.No)

        #child
        self.T4_Add_Slope = T4_Add_Slope_Form( path , p_to_database , p_to_documents , p_to_server )
        self.T4_Add_Slope.toApply_s.connect( self.return_wo_save )
        self.T4_Add_Slope.sel_and_toApply_s.connect( self.insert_feature )

        self.T4_Add_CC = T4_Add_CC_Form(path, p_to_database, p_to_documents, p_to_server)
        self.T4_Add_CC.toApply_s.connect(self.return_wo_save)
        self.T4_Add_CC.sel_and_toApply_s.connect(self.insert_cc)

        self.T4_Add_CCF = T4_Add_CCF_Form( path , p_to_database , p_to_documents , p_to_server )
        self.T4_Add_CCF.toApply_s.connect( self.return_wo_save )
        self.T4_Add_CCF.sel_and_toApply_s.connect( self.insert_ccf )

    def add_CC(self):
        self.T4_Add_CC.show()
        self.T4_Add_CC.load_cc_recipient()
        self.close()

    def add_CCF(self):
        self.T4_Add_CCF.show()
        self.T4_Add_CCF.reset()
        self.close()

    def insert_cc(self):
        self.show()
        print ("self.T4_Add_CC.T4_Add_CC_recipient_com.currentText() : ", self.T4_Add_CC.T4_Add_CC_recipient_com.currentText())
        cc_recipient = self.T4_Add_CC.T4_Add_CC_recipient_com.currentText()
        cc_attention = self.T4_Add_CC.T4_Add_CC_attention_com.currentText()
        cc_fax = self.T4_Add_CC.T4_Add_CC_fax_com.currentText()
        if self.T4_Add_CC.T4_Add_CC_n_a_rad.isChecked():
            cc_encl = ""
        elif self.T4_Add_CC.T4_Add_CC_w_e_rad.isChecked():
            cc_encl = "(w/e)"
        else:
            cc_encl = "(w/o)"

        if self.T3_Apply_Officer_cc_le.text() != "":
            ori_cc_text = self.T3_Apply_Officer_cc_le.text()
            self.T3_Apply_Officer_cc_le.setText(ori_cc_text + "," + cc_recipient + ":" + cc_attention + ":" + cc_fax + ":" + cc_encl)
        else:
            self.T3_Apply_Officer_cc_le.setText(cc_recipient + ":" + cc_attention + ":" + cc_fax + ":" + cc_encl)

    def insert_ccf(self):
        self.show()
        ccf_recipient = self.T4_Add_CCF.T4_Add_CCF_file_le.text()
        if self.T4_Add_CCF.T4_Add_CCF_n_a_rad.isChecked():
            ccf_encl = ""
        elif self.T4_Add_CCF.T4_Add_CCF_w_e_rad.isChecked():
            ccf_encl = "(w/e)"
        else:
            ccf_encl = "(w/o)"

        if self.T3_Apply_Feature_cc_le.text() != "":
            ori_ccf_text = self.T3_Apply_Feature_cc_le.text()
            self.T3_Apply_Feature_cc_le.setText( ori_ccf_text + "," + ccf_recipient + ccf_encl )
        else:
            self.T3_Apply_Feature_cc_le.setText( ccf_recipient + ccf_encl )

    def create_New(self, id):
        self.id = id
        self.load_feature()

    def load_feature(self):
        self.clear_all_content()
        ##### read data
        feature_no_nlist = retri_sql( self.p_to_database , "PMW_records" , "Feature_No" , "Package" , "Location" , id=str(self.id) )
        feature_no = feature_no_nlist[ 0 ][ 0 ]
        package = feature_no_nlist[ 0 ][ 1 ]
        location = feature_no_nlist[ 0 ][ 2 ]

        cur_my_post_nlist = retri_sql( self.p_to_database , "My_Post" , "Department" , "Section" , "Post" , "Senior" ,
                                       "DLO_District" )
        cur_post = cur_my_post_nlist[ 0 ][ 2 ]
        cur_working_sec = cur_my_post_nlist[ 0 ][ 0 ] + "_" + cur_my_post_nlist[ 0 ][ 1 ]
        my_senior_post = cur_my_post_nlist[ 0 ][ 3 ]
        DLO_District_lst = cur_my_post_nlist[ 0 ][ 4 ].split( ";" )
        now_date = datetime.date.today()
        now_date_str = now_date.strftime( "%d %B %Y" )

        tentative_start_date = add_months( now_date , 3 )
        tentative_end_date = add_months( now_date , 9 )
        request_approve_date = add_months( now_date , 2 )

        posting_nlist = retri_sql( self.p_to_database , cur_working_sec , "P_Name" , "Tel" , "Email" , "Fax" ,
                                   Post="'" + cur_post + "'" )
        my_name = posting_nlist[ 0 ][ 0 ]
        my_phone = posting_nlist[ 0 ][ 1 ]
        my_email = posting_nlist[ 0 ][ 2 ]
        my_fax = posting_nlist[ 0 ][ 3 ]

        my_initial = to_initial( my_name )
        my_initial_s = my_initial.lower()

        posting_nlist = retri_sql( self.p_to_database , cur_working_sec , "P_Name" , Post="'" + my_senior_post + "'" )
        my_senior_name = posting_nlist[ 0 ][ 0 ]
        my_senior_initial = to_initial( my_senior_name )

        #### write data on form

        self.feature_no_lst = [ ]
        self.cc_feature_lst = [ ]
        self.T3_Apply_My_For_le.setText("CGE/SM")
        self.T3_Apply_My_Phone_le.setText(my_phone)
        self.T3_Apply_My_Fax_le.setText(my_fax)
        self.T3_Apply_My_Email_le.setText(my_email)
        self.T3_Apply_Our_File_Ref.setText("LD SMS/SLP/1/" + feature_no)
        self.T3_Apply_Date_le.setText(now_date_str)

        self.T3_Apply_Feature_No_le.setText(feature_no)
        self.T3_Apply_Package_No_le.setText(package)
        self.T3_Apply_Location_le.setText(location)
        self.T3_Apply_Tentative_Start_Date_le.setText(tentative_start_date)
        self.T3_Apply_Tentative_End_Date_le.setText(tentative_end_date)
        self.T3_Apply_Request_Approve_Date_le.setText(request_approve_date)

        self.T3_Apply_Recipient_com.addItems(DLO_District_lst)
        self.T3_Apply_Recipient_com.setCurrentText(DLO_District_lst[0])
        self.set_officer()
        self.T3_Apply_Total_Page_le.setText("1")
        self.T3_Apply_with_Encl_rad.setChecked(True)

        self.feature_no_lst.append(feature_no)
        self.T3_Apply_Tentative_Start_Date_le.setText(tentative_start_date)
        self.T3_Apply_Tentative_End_Date_le.setText(tentative_end_date)
        self.T3_Apply_Request_Approve_Date_le.setText(request_approve_date)

        self.T3_Apply_My_Name_le.setText(my＿name)
        self.T3_Apply_My_Initial_le.setText(my_initial)
        self.T3_Apply_My_Senior_Initial_le.setText(my_senior_initial)
        self.T3_Apply_My_Initial_s_le.setText(my_initial_s)

        self.set_feature()
        #self.T3_Apply_Incoming_File_Ref_le.setStyleSheet("""QLineEdit {background-color: yellow;} """)

    def set_feature(self):
        feature_lst_str = ", ".join(self.feature_no_lst)
        self.T3_Apply_Feature_No_le.setText(feature_lst_str)
        package_str = self.T3_Apply_Package_No_le.text().replace("/", "_")
        feature_lst_str_str = feature_lst_str.replace(",", "_")
        feature_lst_str_str = feature_lst_str_str.replace("/", "")
        Docu_Title = package_str + "_" + feature_lst_str_str + "(STLA_App)"
        self.T3_Apply_Docu_Title_le.setText(Docu_Title)
        if len(self.feature_no_lst) > 1:
            self.cc_feature_lst = self.feature_no_lst[1:]
            cc_feature_str = ",".join(self.cc_feature_lst)
            self.T3_Apply_Feature_cc_le.setText(cc_feature_str)

    def return_wo_save(self):
        self.show()

    def insert_feature(self):
        self.feature_no_lst.append( self.T4_Add_Slope.added_feature )
        print ("self.feature_no_lst : ", self.feature_no_lst)
        self.set_feature()
        self.show()

    def insert_file_thro(self):
        pass

    def set_officer(self):
        recipient = self.T3_Apply_Recipient_com.currentText()
        officer_nlist = retri_sql(self.p_to_database, "Officer", "Officer_Name", "Fax", For_Post= "'" + recipient + "'")
        Officer_Name_lst = []
        Fax_lst = []
        for i in range(len(officer_nlist)):
            Officer_Name_lst.append(officer_nlist[i][0])
            Fax_lst.append(officer_nlist[i][1])
        Officer_Name_set = set(Officer_Name_lst)
        Fax_set = set(Fax_lst)
        self.T3_Apply_Atention_com.addItems(Officer_Name_set)
        self.T3_Apply_Recipient_Fax_com.addItems(Fax_set)


    def execute(self):
        template = self.p_to_database + r"\Database\PMW\STLA App\STLA Application.docx"
        document = MailMerge(template)

        if self.T3_Apply_with_Encl_rad.isChecked():
            Title_Encl = " + Encl"
            End_Encl = "Encl."
        else:
            Title_Encl = ""
            End_Encl = ""

        if len(self.feature_no_lst)==1:
            with_s = ""
        else:
            with_s = "s"


        if self.T3_Apply_Officer_cc_le.text() != "":
            cc_dict_lst = [ ]
            cc_officer_lst = self.T3_Apply_Officer_cc_le.text().split( ", " )
            cc_officer_nlst = [ x.split(":") for x in cc_officer_lst]
            ###
            for i in range( len( cc_officer_lst ) ):
                if cc_officer_nlst[0][2] == "Nil":
                    if cc_officer_nlst[0][1] == "Nil":
                        attn_fax_text = ""
                    else:
                        attn_fax_text = " (attn: " + cc_officer_nlst[i][1] + ")"
                else:
                    if cc_officer_nlst[0][1] == "Nil":
                        attn_fax_text = " (fax: " + cc_officer_nlst[i][2] + ")"
                    else:
                        attn_fax_text = " (attn: " + cc_officer_nlst[i][1] + "     fax: " + cc_officer_nlst[i][2] + " ) "
                enclosure_text = cc_officer_nlst[ i ][ 3 ]
                if i == 0:
                    cc_dict_lst.append( {'cc': "c.c", 'cc_recipient': cc_officer_nlst[i][0] + attn_fax_text + enclosure_text} )
                else:
                    cc_dict_lst.append( {'cc_recipient': cc_officer_nlst[i][0] + attn_fax_text + enclosure_text})


            if self.T3_Apply_Feature_cc_le.text() != "":
                cc_dict_lst.append( {'cc_recipient': "LD SMS/SLP/1/" + self.T3_Apply_Feature_cc_le.text()} )

        elif self.T3_Apply_Feature_cc_le.text() != "":
            cc_dict_lst = [ ]
            cc_dict_lst.append( {'cc': "c.c" ,
                'cc_recipient': "LD SMS/SLP/1/" + self.T3_Apply_Feature_cc_le.text()} )

        else:
            cc_dict_lst = [ ]

        if self.T3_Apply_file_thro_le.text() != "":
            cc_dict_lst.append( {'cc': "file thro'" ,
                                 'cc_recipient': self.T3_Apply_file_thro_le.text()} )


        #date_str = datetime.datetime.strptime(self.T3_Apply_Date_le.text(), '%Y-%m-%d').strftime('%d %B %Y')
        date_str = self.T3_Apply_Date_le.text()
        try:
            document.merge(
                ###
                my_for = self.T3_Apply_My_For_le.text(),
                my_phone = self.T3_Apply_My_Phone_le.text(),
                my_fax = self.T3_Apply_My_Fax_le.text(),
                my_email = self.T3_Apply_My_Email_le.text(),
                our_file_ref = self.T3_Apply_Our_File_Ref.text(),
                date = date_str,
                ####
                feature_no = self.T3_Apply_Feature_No_le.text(),
                with_s = with_s,
                location = self.T3_Apply_Location_le.text(),
                package = self.T3_Apply_Package_No_le.text(),
                tentative_start_date = self.T3_Apply_Tentative_Start_Date_le.text(),
                tentative_end_date=self.T3_Apply_Tentative_End_Date_le.text(),
                request_approve_date=self.T3_Apply_Request_Approve_Date_le.text(),
                ####
                recipient = self.T3_Apply_Recipient_com.currentText(),
                attention = self.T3_Apply_Atention_com.currentText(),
                file_ref_A = self.T3_Apply_Recipient_File_Ref_A_le.text(),
                file_ref_B = self.T3_Apply_Recipient_File_Ref_B_le.text(),
                file_ref_date = self.T3_Apply_Recipient_File_Ref_Date_le.text(),
                recipient_fax=self.T3_Apply_Recipient_Fax_com.currentText(),
                total_page = self.T3_Apply_Total_Page_le.text(),
                Title_Encl = Title_Encl,
                ####
                my_name = self.T3_Apply_My_Name_le.text(),
                End_Encl = End_Encl,
                my_initial = self.T3_Apply_My_Initial_le.text(),
                my_senior_initial = self.T3_Apply_My_Senior_Initial_le.text(),
                my_initial_s = self.T3_Apply_My_Initial_s_le.text(),
            )

            document.merge_rows( 'cc' , cc_dict_lst )

        except:
            print ("something went wrong")


        package_str = self.T3_Apply_Package_No_le.text().replace("/", "_")

        now_time = datetime.datetime.now().time()
        now_time_str = now_time.strftime("%H:%M:%S")
        now_date = datetime.date.today().strftime("%d/%m/%Y")
        path_to_file = self.p_to_server + "\\unit1\\Kenneth KWONG\\PMW\\" + package_str + "\\" + self.T3_Apply_Docu_Title_le.text() + ".docx"

        try:
            with open(path_to_file) as f:
                OverwriteAns = self.msgBoxOverwrite.exec()
                if OverwriteAns == QMessageBox.Yes:
                    document.write(path_to_file)

                    id_nlst = retri_sql(self.p_to_database, "Outgoing_records", "id", path_to_file= "'" + path_to_file + "'")
                    print ("id_nlst : ", id_nlst[0][0])

                    update_sql_1key(self.p_to_database, "Outgoing_records", "date_mod", "id" , id_nlst[0][0], now_date)
                    update_sql_1key(self.p_to_database, "Outgoing_records", "time_mod", "id", id_nlst[0][0], now_time_str)

        except IOError:
            document.write(path_to_file)

            file_ref = self.T3_Apply_Our_File_Ref.text()
            recipient = self.T3_Apply_Recipient_com.currentText()

            insert_sql(self.p_to_database, "Outgoing_records", date_mod=now_date, time_mod=now_time_str, type="STLA_App", identity1=self.T3_Apply_Package_No_le.text(), identity2=self.T3_Apply_Feature_No_le.text(), recipient=recipient, T1="PMW", T2="STLA_App", T3="--", T4="--", file_ref=file_ref , path_to_file=path_to_file)

        for feature in self.feature_no_lst:
            feature_id_nlst = retri_sql(self.p_to_database, "PMW_records", "id", Feature_No="'" + feature + "'")
            update_sql_1key(self.p_to_database, "PMW_records", "STLA_Status", "id", feature_id_nlst[0][0], "Apply")
            update_sql_1key(self.p_to_database, "PMW_records", "STLA_App_date", "id", feature_id_nlst[0][0], now_date)

        openAns = self.msgBoxOpen.exec()
        if openAns == QMessageBox.Yes:
            os.startfile(path_to_file)

        self.toPMW()

    def clear_all_content(self):
        self.T3_Apply_Recipient_com.clear()
        self.T3_Apply_Atention_com.clear()
        self.T3_Apply_Recipient_Fax_com.clear()
        self.T3_Apply_Officer_cc_le.clear()
        self.T3_Apply_Feature_cc_le.clear()
        self.T3_Apply_file_thro_le.clear()

    def add_slope(self):
        self.T4_Add_Slope.load_data()
        self.T4_Add_Slope.show()
        self.close()

    def add_attention(self):
        pass

    def toPMW(self):
        self.hide()
        self.toPMW_s.emit()

    def toSTLA_Sel(self):
        self.hide()
        self.toSTLA_Sel_s.emit()

def add_months(sourcedate , months):
    month = sourcedate.month - 1 + months
    year = sourcedate.year + month // 12
    month = month % 12 + 1
    day = min( sourcedate.day , calendar.monthrange( year , month )[ 1 ] )
    full_date = datetime.date(year, month, day)
    return full_date.strftime("%d %B %Y")

def to_initial(name):
    name_lst = name.split(" ")
    initial_lst = [name_lst[i][0] for i in range(len(name_lst))]
    initial = "".join(initial_lst)
    return initial
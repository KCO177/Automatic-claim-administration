# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'gui_01.ui'
#
# Created by: PyQt5 UI code generator 5.15.4
#
# WARNING: Any manual changes made to this file will be lost when pyuic5 is
# run again.  Do not edit this file unless you know what you are doing.

import psycopg2 as pg

from PyQt5 import QtCore, QtGui, QtWidgets
from db_quality import SetGetData

'''
class Get_data:
    def connect_claim():
        # connecting to the db Qualita_test claim table

        try:

            conn = pg.connect(
                database='Qualita_test',
                user='postgres',
                password='kco177',
                host='localhost',
                port='5432')

            # create cursor object
            cur = conn.cursor()

        except (Exception, psycopg2.DatabaseError) as error:

            print('Error while creating PostgreSQL table', error)

        return conn, cur


    def fetch_data():
        # fetch all data from claim table
        conn, cur = Get_data.connect_claim()

        try:
            cur.execute('SELECT * FROM claims')

        except:
            print('error !')

        # store the results in data
        data = cur.fetchall()

        return data
'''

class Ui_MainWindow(object):
    def setupUi(self, MainWindow):
        MainWindow.setObjectName("Claim manager")
        MainWindow.resize(1034, 932)
        self.centralwidget = QtWidgets.QWidget(MainWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.gridLayout_8 = QtWidgets.QGridLayout(self.centralwidget)
        self.gridLayout_8.setObjectName("gridLayout_8")
        self.mail_menu_bar = QtWidgets.QFrame(self.centralwidget)
        self.mail_menu_bar.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.mail_menu_bar.setFrameShadow(QtWidgets.QFrame.Raised)
        self.mail_menu_bar.setObjectName("mail_menu_bar")
        self.verticalLayout_6 = QtWidgets.QVBoxLayout(self.mail_menu_bar)
        self.verticalLayout_6.setObjectName("verticalLayout_6")
        self.gridLayout_2 = QtWidgets.QGridLayout()
        self.gridLayout_2.setObjectName("gridLayout_2")

        self.select_claim = QtWidgets.QPushButton(self.mail_menu_bar)
        self.select_claim.setMinimumSize(QtCore.QSize(0, 75))
        self.select_claim.setMaximumSize(QtCore.QSize(16777215, 75))
        self.select_claim.setObjectName("select_claim")
        self.gridLayout_2.addWidget(self.select_claim, 0, 0, 1, 1)
        self.proces_claim = QtWidgets.QPushButton(self.mail_menu_bar)
        self.proces_claim.setMinimumSize(QtCore.QSize(0, 75))
        self.proces_claim.setMaximumSize(QtCore.QSize(16777215, 75))
        self.proces_claim.setObjectName("proces_claim")
        self.gridLayout_2.addWidget(self.proces_claim, 1, 0, 1, 1)
        self.confirm_QA = QtWidgets.QPushButton(self.mail_menu_bar)
        self.confirm_QA.setMinimumSize(QtCore.QSize(0, 75))
        self.confirm_QA.setMaximumSize(QtCore.QSize(16777215, 75))
        self.confirm_QA.setObjectName("confirm_QA")
        self.gridLayout_2.addWidget(self.confirm_QA, 2, 0, 1, 1)
        self.send_email = QtWidgets.QPushButton(self.mail_menu_bar)
        self.send_email.setMinimumSize(QtCore.QSize(0, 75))
        self.send_email.setMaximumSize(QtCore.QSize(16777215, 75))
        self.send_email.setObjectName("send_email")
        self.gridLayout_2.addWidget(self.send_email, 3, 0, 1, 1)
        self.verticalLayout_4 = QtWidgets.QVBoxLayout()
        self.verticalLayout_4.setObjectName("verticalLayout_4")
        self.datamatrix = QtWidgets.QCheckBox(self.mail_menu_bar)
        self.datamatrix.setObjectName("datamatrix")
        self.verticalLayout_4.addWidget(self.datamatrix)
        self.QR_BAR = QtWidgets.QCheckBox(self.mail_menu_bar)
        self.QR_BAR.setObjectName("QR_BAR")
        self.verticalLayout_4.addWidget(self.QR_BAR)
        self.OCR = QtWidgets.QCheckBox(self.mail_menu_bar)
        self.OCR.setObjectName("OCR")
        self.verticalLayout_4.addWidget(self.OCR)

        self.gridLayout_2.addLayout(self.verticalLayout_4, 4, 0, 1, 1)
        self.verticalLayout_6.addLayout(self.gridLayout_2)
        self.gridLayout_8.addWidget(self.mail_menu_bar, 0, 0, 1, 1)
        self.main_frame = QtWidgets.QVBoxLayout()
        self.main_frame.setObjectName("main_frame")
        self.horizontalLayout_2 = QtWidgets.QHBoxLayout()
        self.horizontalLayout_2.setObjectName("horizontalLayout_2")
        self.frame_mails = QtWidgets.QFrame(self.centralwidget)
        self.frame_mails.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_mails.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_mails.setObjectName("frame_mails")
        self.gridLayout_4 = QtWidgets.QGridLayout(self.frame_mails)
        self.gridLayout_4.setObjectName("gridLayout_4")
        self.mail_list = QtWidgets.QListWidget(self.frame_mails)
        self.mail_list.setStyleSheet("")
        self.mail_list.setObjectName("mail_list")
        self.gridLayout_4.addWidget(self.mail_list, 0, 0, 1, 1)
        self.horizontalLayout_2.addWidget(self.frame_mails)
        self.list_of_mail_contact = QtWidgets.QFrame(self.centralwidget)
        self.list_of_mail_contact.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.list_of_mail_contact.setFrameShadow(QtWidgets.QFrame.Raised)
        self.list_of_mail_contact.setObjectName("list_of_mail_contact")
        self.gridLayout_3 = QtWidgets.QGridLayout(self.list_of_mail_contact)
        self.gridLayout_3.setObjectName("gridLayout_3")
        self.verticalLayout_3 = QtWidgets.QVBoxLayout()
        self.verticalLayout_3.setObjectName("verticalLayout_3")
        self.label = QtWidgets.QLabel(self.list_of_mail_contact)
        self.label.setObjectName("label")
        self.verticalLayout_3.addWidget(self.label)
        self.contact_list = QtWidgets.QListWidget(self.list_of_mail_contact)
        self.contact_list.setObjectName("contact_list")
        item = QtWidgets.QListWidgetItem()
        self.contact_list.addItem(item)
        item = QtWidgets.QListWidgetItem()
        self.contact_list.addItem(item)
        self.verticalLayout_3.addWidget(self.contact_list)
        self.gridLayout_3.addLayout(self.verticalLayout_3, 0, 0, 1, 1)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.label_2 = QtWidgets.QLabel(self.list_of_mail_contact)
        self.label_2.setObjectName("label_2")
        self.horizontalLayout.addWidget(self.label_2)
        self.days_restrictor = QtWidgets.QSpinBox(self.list_of_mail_contact)
        self.days_restrictor.setProperty("value", 1)
        self.days_restrictor.setObjectName("days_restrictor")
        self.horizontalLayout.addWidget(self.days_restrictor)
        self.gridLayout_3.addLayout(self.horizontalLayout, 1, 0, 1, 1)
        self.horizontalLayout_2.addWidget(self.list_of_mail_contact)
        self.frame_4 = QtWidgets.QFrame(self.centralwidget)
        self.frame_4.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_4.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_4.setObjectName("frame_4")
        self.gridLayout = QtWidgets.QGridLayout(self.frame_4)
        self.gridLayout.setObjectName("gridLayout")
        self.Folder_tree = QtWidgets.QTreeWidget(self.frame_4)
        self.Folder_tree.setObjectName("Folder_tree")
        self.Folder_tree.headerItem().setText(0, "Claimfolder:")
        self.gridLayout.addWidget(self.Folder_tree, 0, 0, 1, 1)
        self.horizontalLayout_2.addWidget(self.frame_4)
        self.main_frame.addLayout(self.horizontalLayout_2)
        self.frame_5 = QtWidgets.QFrame(self.centralwidget)
        self.frame_5.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_5.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_5.setObjectName("frame_5")
        self.gridLayout_5 = QtWidgets.QGridLayout(self.frame_5)
        self.gridLayout_5.setObjectName("gridLayout_5")
        self.Console = QtWidgets.QLineEdit(self.frame_5)
        self.Console.setObjectName("Console")
        self.gridLayout_5.addWidget(self.Console, 0, 0, 1, 1)
        self.main_frame.addWidget(self.frame_5)
        self.frame_2 = QtWidgets.QFrame(self.centralwidget)
        self.frame_2.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.frame_2.setFrameShadow(QtWidgets.QFrame.Raised)
        self.frame_2.setObjectName("frame_2")
        self.gridLayout_6 = QtWidgets.QGridLayout(self.frame_2)
        self.gridLayout_6.setObjectName("gridLayout_6")
        self.table_of_claims = QtWidgets.QTableWidget(self.frame_2)
        self.table_of_claims.setObjectName("table_of_claims")

        #add cells regarding number of items in db
        data = SetGetData.fetch_data()
        sloupec = len(data[0])
        #print('pocet sloupcu: ', sloupec)
        radek = sum(isinstance(elem, tuple) for elem in data)
        #print('pocet radku: ', radek)
        self.table_of_claims.setColumnCount(sloupec)
        self.table_of_claims.setRowCount(radek)

        item = QtWidgets.QTableWidgetItem()
        self.table_of_claims.setVerticalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_of_claims.setVerticalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_of_claims.setHorizontalHeaderItem(0, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_of_claims.setHorizontalHeaderItem(1, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_of_claims.setHorizontalHeaderItem(2, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_of_claims.setHorizontalHeaderItem(3, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_of_claims.setHorizontalHeaderItem(4, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_of_claims.setHorizontalHeaderItem(5, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_of_claims.setHorizontalHeaderItem(6, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_of_claims.setHorizontalHeaderItem(7, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_of_claims.setHorizontalHeaderItem(8, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_of_claims.setHorizontalHeaderItem(9, item)
        item = QtWidgets.QTableWidgetItem()
        self.table_of_claims.setHorizontalHeaderItem(10, item)
        item = QtWidgets.QTableWidgetItem()

        #add cells regarding number of items in db
        for j in range(0, radek):
            for i in range(0, sloupec):
                self.table_of_claims.setItem(j, i, item)
                item = QtWidgets.QTableWidgetItem()

        self.gridLayout_6.addWidget(self.table_of_claims, 0, 0, 1, 1)
        self.main_frame.addWidget(self.frame_2)
        self.gridLayout_8.addLayout(self.main_frame, 0, 1, 2, 1)
        self.table_menu_bar = QtWidgets.QFrame(self.centralwidget)
        self.table_menu_bar.setFrameShape(QtWidgets.QFrame.StyledPanel)
        self.table_menu_bar.setFrameShadow(QtWidgets.QFrame.Raised)
        self.table_menu_bar.setObjectName("table_menu_bar")
        self.verticalLayout = QtWidgets.QVBoxLayout(self.table_menu_bar)
        self.verticalLayout.setObjectName("verticalLayout")
        self.verticalLayout_2 = QtWidgets.QVBoxLayout()
        self.verticalLayout_2.setObjectName("verticalLayout_2")
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout_2.addItem(spacerItem)
        self.update = QtWidgets.QPushButton(self.table_menu_bar)
        self.update.setMinimumSize(QtCore.QSize(0, 75))
        self.update.setObjectName("update")
        self.verticalLayout_2.addWidget(self.update)
        self.delete_2 = QtWidgets.QPushButton(self.table_menu_bar)
        self.delete_2.setMinimumSize(QtCore.QSize(0, 75))
        self.delete_2.setObjectName("delete_2")
        self.verticalLayout_2.addWidget(self.delete_2)
        self.exporttoexcell = QtWidgets.QPushButton(self.table_menu_bar)
        self.exporttoexcell.setMinimumSize(QtCore.QSize(0, 75))
        self.exporttoexcell.setObjectName("exporttoexcell")
        self.verticalLayout_2.addWidget(self.exporttoexcell)
        self.verticalLayout.addLayout(self.verticalLayout_2)
        self.gridLayout_8.addWidget(self.table_menu_bar, 1, 0, 1, 1)
        MainWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(MainWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 1034, 21))
        self.menubar.setObjectName("menubar")
        MainWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(MainWindow)
        self.statusbar.setObjectName("statusbar")
        MainWindow.setStatusBar(self.statusbar)

        self.retranslateUi(MainWindow)
        QtCore.QMetaObject.connectSlotsByName(MainWindow)

    def retranslateUi(self, MainWindow):
        _translate = QtCore.QCoreApplication.translate
        MainWindow.setWindowTitle(_translate("MainWindow", "KCO CLAIMBOT MANAGER"))
        self.select_claim.setText(_translate("MainWindow", "select claim"))
        self.proces_claim.setText(_translate("MainWindow", "process claim"))
        self.confirm_QA.setText(_translate("MainWindow", "confirm QA"))
        self.send_email.setText(_translate("MainWindow", "Send mail"))
        self.datamatrix.setText(_translate("MainWindow", "Data Matrix"))
        self.QR_BAR.setText(_translate("MainWindow", "QR BAR code"))
        self.OCR.setText(_translate("MainWindow", "OCR"))
        self.label.setText(_translate("MainWindow", "List of used mail contacts"))
        __sortingEnabled = self.contact_list.isSortingEnabled()
        self.contact_list.setSortingEnabled(False)
        item = self.contact_list.item(0)
        item.setText(_translate("MainWindow", "customer_mial@company.com"))
        item = self.contact_list.item(1)
        item.setText(_translate("MainWindow", "next mail@company.com"))
        self.contact_list.setSortingEnabled(__sortingEnabled)
        self.label_2.setText(_translate("MainWindow", "Days back"))
        self.Console.setText(_translate("MainWindow", "console processing"))

        #>>>



        item = self.table_of_claims.horizontalHeaderItem(0)
        item.setText(_translate("MainWindow", "ID"))
        item = self.table_of_claims.horizontalHeaderItem(1)
        item.setText(_translate("MainWindow", "date"))
        item = self.table_of_claims.horizontalHeaderItem(2)
        item.setText(_translate("MainWindow", "project"))
        item = self.table_of_claims.horizontalHeaderItem(3)
        item.setText(_translate("MainWindow", "part shortcut"))
        item = self.table_of_claims.horizontalHeaderItem(4)
        item.setText(_translate("MainWindow", "part_id"))
        item = self.table_of_claims.horizontalHeaderItem(5)
        item.setText(_translate("MainWindow", "customer"))
        item = self.table_of_claims.horizontalHeaderItem(6)
        item.setText(_translate("MainWindow", "type"))
        item = self.table_of_claims.horizontalHeaderItem(7)
        item.setText(_translate("MainWindow", "claim/qa"))
        item = self.table_of_claims.horizontalHeaderItem(8)
        item.setText(_translate("MainWindow", "failure"))
        item = self.table_of_claims.horizontalHeaderItem(9)
        item.setText(_translate("MainWindow", "NOKs"))
        item = self.table_of_claims.horizontalHeaderItem(10)
        item.setText(_translate("MainWindow", "Customer stock"))
        __sortingEnabled = self.table_of_claims.isSortingEnabled()
        self.table_of_claims.setSortingEnabled(False)

        #add data from database
        data = SetGetData.fetch_data() #Get_data.fetch_data()
        #print(data)
        sloupec = len(data[0])
        #print('pocet sloupcu: ', sloupec)
        radek = sum(isinstance(elem, tuple) for elem in data)
        #print('pocet radku: ', radek)

        for j in range (0, radek):

            for i in range(0, sloupec):
                item = self.table_of_claims.item(j, i)
                cell = str(data[j][i])
                item.setText(_translate("MainWindow", cell))
                #print(cell)

        self.table_of_claims.setSortingEnabled(__sortingEnabled)
        self.update.setText(_translate("MainWindow", "update"))
        self.delete_2.setText(_translate("MainWindow", "delete"))
        self.exporttoexcell.setText(_translate("MainWindow", "export to excell"))

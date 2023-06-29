import sys
import win32com.client
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QMainWindow, QApplication
from gui_01 import Ui_MainWindow
import os
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QTableView
from PyQt5.QtGui import QPalette, QColor
import datetime
from datetime import timedelta
from run_db import run_database

class Main(QMainWindow, Ui_MainWindow):
    app = QApplication([])
    # Force the style to be the same on all OSs:
    app.setStyle("Fusion")

    # Now use a palette to switch to dark colors:
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(53, 53, 53))
    palette.setColor(QPalette.WindowText, Qt.gray)
    palette.setColor(QPalette.Base, QColor(25, 25, 25))
    palette.setColor(QPalette.AlternateBase, QColor(53, 53, 53))
    palette.setColor(QPalette.ToolTipBase, Qt.black)
    palette.setColor(QPalette.ToolTipText, Qt.gray)
    palette.setColor(QPalette.Text, Qt.white)
    palette.setColor(QPalette.Button, QColor(53, 53, 53))
    palette.setColor(QPalette.ButtonText, Qt.white)
    palette.setColor(QPalette.BrightText, Qt.red)
    palette.setColor(QPalette.Link, QColor(42, 130, 218))
    palette.setColor(QPalette.Highlight, QColor(42, 130, 218))
    palette.setColor(QPalette.HighlightedText, Qt.black)

    app.setPalette(palette)



    def __init__(self):
        super().__init__()

        self.setupUi(self)


        # mail list from outlook
        for item in Mail.mail_list():
            self.mail_list.addItem(item)

        #self.select_claim = QtWidgets.QPushButton(self.mail_menu_bar)
        # get subject and body of mail
        self.mail_list.itemClicked.connect(Mail.on_item_clicked)


        # folder tree from folder path source
        claims = Files.get_list_of_file()
        documents = Files.get_list_of_documents()

        for claim in claims:
            claims_item = QtWidgets.QTreeWidgetItem(self.Folder_tree)
            claims_item.setText(0, claim)

            for document in documents[claim]:
                #print(document)
                document_item = QtWidgets.QTreeWidgetItem(claims_item)
                document_item.setText(0, document)
                claims_item.addChild(document_item)



        # show
        self.show()


class Mail:

    def mail_list():
        #>>> outsource to Outlook management
        outlook = win32com.client.Dispatch('outlook.application').GetNamespace('MAPI')

        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items  # .Folders
        current_time = datetime.datetime.now()
        received_dt = current_time - timedelta(days=7)
        received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
        messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
        messages = messages.Restrict("[SenderEmailAddress] = 'petr.koscelnik@ascorium.com'")  # 'info@novinky.knihobot.cz'

        messages.Sort('[ReceivedTime]', Descending=True)

        mail_str = ""
        mail_list = []
        mail_str_list = []
        for message in list(messages)[:100]:
            mail = (message.Subject, message.ReceivedTime, message.SenderEmailAddress)
            # print(mail)

            mail_str = (('{}/ {}/ {}').format(*mail))  # {} {}
            # print(mail_str)
            mail_list.append(mail_str)

        return mail_list

    def on_item_clicked(item):
        try:
            value = item.text()
            string_to_split = value

            print(string_to_split)
            subject = string_to_split.split('/')[0]
            #print(f"You clicked {value}")
            sub = str(subject)
            print(f'subject of mail: {sub}')

            #print('sub from on item clicked', sub)
            run_database.run_input_from_mail(sub)

            run_database.run()

        except Exception as e:
            print(e)


        # >>> output from pictures
        # >>> a iterace Mail_output.text_output s message y pictures_output pokud None ve vypisu

        #Mail.select_mail(sub) # >>> pro napojeni na funkce -> atribut subject
        # Mail_recieve.read_mail()

        #>>> vyresit update pro vsechny widgety po ukonceni funkce
        window = Main()
        window.show()


class Files:
    def get_list_of_file():
        files = os.listdir('C:/Users/START/PycharmProjects/databaze/claimfolder')
        #print(files)
        return files
    def get_list_of_documents():
        claims = Files.get_list_of_file()
        #print('claims:',claims)
        docs = []
        for claim in claims:
            path = 'C:/Users/START/PycharmProjects/databaze/claimfolder/' + claim
            documents = os.listdir(path)
            docustr = str(claim) + ':'+ str(documents)
            #print(docustr)
            docs.append(docustr)
        #print(docs)
        #docs = ["abc:['a','n','v']", "mnf = ['e','s','b']"]
        my_dict = {s.split(':')[0]: eval(s.split(':')[1]) for s in docs}
        #print(my_dict)
        return my_dict


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = Main()
    window.show()
    sys.exit(app.exec_())

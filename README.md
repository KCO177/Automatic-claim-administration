# Automatic-claim-administration
The application in incoming emails recognizes complaint e-mails and extracts important information from them, which it processes into a database and prepares pre-filled D3 documents and email messages with appropriate attachments. It also predicts the financial assumptions for settling the complaint. 

![schema](https://github.com/KCO177/Automatic-claim-administration/assets/28139409/114a6fee-04f1-4a23-af3c-02d21d7ee800)

## Requirements
Python 3.9

PyLib: psycopg2, pandas, openpyxl, win32com.client, os, re, nltk, thefuzz, PyQt5,cv2, pytesseract, pyzbar, decode, pylibdmtx , PIL

PostgreSQL

## Description
In this application, it was necessary to deal with the fragmentation of the BI system into individual Excel spreadsheets from which information is extracted into the database. Additionally, there was variation in product naming and numbering. The application is built and fine-tuned for specific projects, but its components can be used generally. After download will not run. If you wish to use the application, it will be necessary to align it with your BI system. This public version is without sensitive datas.

## Function

https://github.com/KCO177/Automatic-claim-administration/assets/28139409/03b9a7ab-26a8-42a1-a92a-93b3c96f45c1


## Run function
run_gui.py

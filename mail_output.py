from financial_assum import Expedition_df
from db_quality import SetGetData, Aggregation, Report
from qualita_db_dictionary import MyDict as Dic
import win32com.client
import datetime
import pandas as pd


class Mail_to_send:
    def prepare_qa_mail_message(claim_id):
        #prepare mail message to company
        claim_data = Aggregation.select(claim_id)
        data = claim_data
        d = data[0]

        #>>> pokud v textu nalezne vice chyb preved jednotlive do cestiny
        if ',' not in d[8]:
            vada = Dic.failures[d[8]]
        else:
            merged_failure = d[8]
            splited_failure = merged_failure.split(',')
            list_vada = []
            for failure in splited_failure:
                single_vada = Dic.failures[failure]
                list_vada.append(single_vada)
                delimiter=','
                vada = delimiter.join(list_vada)

        #NOK value
        if d[9] == 0:
            nok_value = '?'
        else:
            nok_value = d[9]

        #warehouse data:
        print('PART d[4]', d[4])
        part = d[3]

        if d[2] == 'P8':
            data_stock = SetGetData.warhouse_stock(part)

            if data_stock.empty:#shape[0]==0 and data_stock.shape[1]==0:
                stock = 0
            else:
                stock = data_stock.to_string(header=False, index= False)
            
        else: stock = 'No warehouse data for this project'

        stock = SetGetData.warhouse_stock(part)
        print(stock)

        ### production dates
        # get column names
        column_names = SetGetData.get_column_names_db('decoded')
        # get data
        data = (Aggregation.select_decoded(claim_id))
        if len(data) > 0:
            # make dataframe
            df = pd.DataFrame(data, columns=column_names)
            df = df.loc[:, ['prod_day', 'prod_time']]  # 'line'
            df['Datetime'] = pd.to_datetime(df['prod_day'] + ' ' + df['prod_time'])
            prod_date = df[['Datetime']].drop_duplicates()
            prod_date = prod_date.to_string(header= False, index=False)
        else:
            prod_date = '?'

        #next delivery
        next_delivery = Expedition_df.next_exped(claim_id)

        '''
        if delivery.empty:
            next_delivery = 0
        else:
            next_delivery = delivery
        '''
        mail_text = f'Dobrý den,\nzasílám informaci o reklamaci na projektu {(d[2])}-{d[3]}\n' \
                    f'REKLAMACE: {d[7]} {d[5]}\n' \
                    f'VADA: {vada}\n' \
                    f'PROJEKT: {d[2]}\n' \
                    f'DÍL: {part}\n' \
                    f'POČET KUSŮ: {nok_value}\n' \
                    f'DATUM VÝROBY:\n{prod_date}\n' \
                    f'SKLADOVÁ ZÁSOBA:\n{stock}\n' \
                    f'NÁSLEDUJÍCÍ VÝVOZ:\n{next_delivery}\n' \
                    f'\n' \
                    f'OKAMŽITÁ OPATŘENÍ:\n' \
                    f'Proškolení a seznámení s vadou:quality_technicican(@)industry.com\n' \
                    f'Třídění skladové zásoby:mistr(@)industry.com\n' \
                    f'\tPřetřízené boxy označte vývěskou v příloze. Nově vyrobené boxy značte závěskou (v příloze)'

        return mail_text
        #>>>improve to put in mail

#Mail_recieve.save_attachment()
#Mail_output.text_output('claim cupholder')


class Outlook_send:
    def send_email(claim_id):
        current_time = datetime.datetime.now()
        claim_data = Aggregation.select(claim_id)
        path_to_claim_folder = Report.claim_path(claim_id)
        name_qa = 'QA_' + str(current_time.year) + '_' + str(claim_id) + '_' + str(claim_data[0][8]) + '.xlsx'
        box_label_name = 'Box_label100% check ' + str(claim_data[0][6]) + ' ' + str(claim_data[0][8]) + '.xlsx'


        ol = win32com.client.Dispatch('Outlook.Application')  # connection reference
        olmailitem = 0x0  # dimension of message memory

        newmail = ol.CreateItem(olmailitem)

        subject = f'oznámení reklamace {claim_data[0][2]} - {claim_data[0][5]}'

        newmail.Subject = (subject)

        newmail.To = 'claim_victims@industry.com'
        newmail.Cc = 'someoneincopy@industry.com'

        path = path_to_claim_folder
        claim_folder_path = path.replace('/', '\\\\') #need to improve the path with '/' to '\\'


        qa_attachment = str(claim_folder_path) + '\\\\' + str(name_qa)
        box_label_attachment =  str(claim_folder_path) + '\\\\' + str(box_label_name)
        print(box_label_attachment)
        print(qa_attachment)

        newmail.Body= Mail_to_send.prepare_qa_mail_message(claim_id)

        attach = qa_attachment
        attach2 = box_label_attachment

        newmail.Attachments.Add(attach)
        newmail.Attachments.Add(attach2)
        newmail.Display()


#Outlook_send.send_email(3)
from db_quality import Report, SetGetData, Aggregation
from outlook_management import Mail_output
from mail_attachment import Mail_attachment, Read_from_attachment, Picture_processing, Picture_resizing_for_qa
from mail_output import Outlook_send
from financial_assum import Financial
from qualita_db_dictionary import MyDict as Dic
import threading
class run_database:

    def claim_update_switch(claim_id):
        data = Aggregation.select(claim_id)
        data_insp = data[0][2:5]
        print(data_insp)
        update_switch = False
        print(update_switch)
        #zjisti jestli je v listu None hodnota
        if None in data_insp:
            data_insp = [str(item).replace('None', ""'None'"") for item in data_insp]
            update_switch = True
        #zjisti jestli je znama LHD/RHD varianta
        for item in data_insp:
            if ',' in item:
                update_switch = True
        part_insp = data_insp[2]
        if part_insp in Dic.ROM_parts:
            if part_insp.endswith('LHD') or part_insp.endswith('RHD'):
                print('part variation known')
            else:
                print('part variation is not known')
                update_switch = True
        print('update_switch', update_switch)
        return update_switch
    def run_administration_output(claim_id):
        Report.make_complete_claim_folder(claim_id)
        direction = Report.claim_path(claim_id)
        Picture_resizing_for_qa.add_pictures_to_QA(path_to_the_claim_folder=direction)
        Outlook_send.send_email(claim_id)
        Financial.insert_new_item(claim_id)
    def run_input_from_mail(subject):
        sub = str(subject)
        print('subject from run_input_from_mail', sub)
        nc = (Mail_output.text_output(sub))

        #claim_date, project, part_name, part_id, customer, customer_claim_id, type_of_claim, claimed_failure, amount_of_claimed
        #>>> do db pridat sloupce stock, prod_time, prod_date

        SetGetData.insert_claim(nc[0],nc[1],nc[2],nc[3],nc[4],nc[5],nc[6],nc[7],nc[8],nc[9])
        claim_id = SetGetData.get_last_id() #get last Id

        # >>> make folder
        direction = Report.claim_path(claim_id)

        # >>> save mail and attachment
        Mail_attachment.save_attachment(direction, sub)
        #Picture_resizing_for_qa.add_pictures_to_QA(path_to_the_claim_folder=direction,path_to_the_template=Dic.quality_alert_template) ### >>> path

        # >>> output from pictures
        # >>> a iterace Mail_output.text_output s message y pictures_output pokud None ve vypisu

    def run():
        #try:
        #run_database.run_input_from_mail()
        claim_id = SetGetData.get_last_id()

        #decoding from attachment
        #claims UPDATE block
        if run_database.claim_update_switch(claim_id) == True:
            try:
                Read_from_attachment.update_from_img_to_db(claim_id)
            except Exception as e:
                print(e)

            input_list = Aggregation.select_decoded(claim_id)
            print(input_list)
            if len(input_list) > 0:
                print('data from images ready for updating claims table')
                Read_from_attachment.update_claim_table_from_db(claim_id)
            else:
                print('no data from images to updating claims table')

        #run administration processes
        run_database.run_administration_output(claim_id)


#run_database.run()
#run_database.run_input_from_mail('claim cupholder')

#print(Aggregation.select(2))
#print(SetGetData.get_last_id())

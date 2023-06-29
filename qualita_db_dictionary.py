class MyDict:
    #connect claim db
    conn_db = {'database':'Qualita_test','user':'postgres','password':'kco177','host':'localhost','port':'5432','table':'claims'}
    claim_table = ['id', 'claim_date', 'project', 'part_name', 'part_id', 'customer', 'customer_claim_id', 'type_of_claim', 'claimed_failure', 'amount_of_claimed']


    # paths
    path_warehosue_stock = 'Aktuální stav skladů.xlsx'
    path_claim_folder = 'C:/Users/START/PycharmProjects/databaze/claimfolder'
    path_production_folder = 'C:/Users/START/PycharmProjects/databaze/production_folder'

    #templates
    quality_alert_template = 'C:/Users/START/PycharmProjects/databaze/QA_template.xlsx' #''QA_template.xlsx'
    box_label_template = 'Label_template.xlsx'
    training_list_template = 'training_list_template.xlsx'

    #part and lines dictionaries and lists
    p8_parts_dictionary = {'AR': 'Antireflection', 'CV': 'Central Airvent', 'CH': 'Cupholder', 'GB': 'Glovebox','DP': 'DP upper skin', 'DF': 'Deco front'}
    project_dictionary = {'AR': 'P8', 'CA': 'P8', 'CH': 'P8', 'GB': 'P8', 'DF': 'P8', 'IP': 'P8', 'DP': 'P8'}
    projects = ['P8']
    customers = {'SMRC':'P8'}
    ROM_parts = ['AR', 'CA', 'CH', 'GB', 'DF']
    Spray_parts = ['IP', 'DP']
    failures = {'Deformation': 'Deformace;', 'Wrong position of label': 'Špatná pozice štítku;', 'NOK stitching': 'Chyba šití;', 'Foaming through ': 'Propěnění;', 'Bubble': 'Bublina;', 'Glossy spot': 'Lesklé fleky;', 'Damaged rib': 'Poškozené žebro;', 'Pinhole': 'Pinhole;', 'Matt spot': 'Matné fleky;', 'Delamination': 'Delaminace;', 'Dirt': 'Nečistoty;', 'NOK thickness': 'Špatná tloušťka;', 'Missing stiching': 'Chybí šití;', 'NOK Colour': 'Špatná barva;', 'incomplete box': 'Nekompletní box;', 'Fold': 'fold;', 'Missing PU': 'chybí PU;', 'Porosity': 'porozita;', 'Material excess':'Přetok'}
    part_id = {'B0004307': 'GB LHD', 'B0004309': 'GB RHD', 'D0323115': 'AR', 'D0323116': 'CH LHD', 'D0323516': 'CH RHD', 'D0323114': 'CA LHD', 'D0323514': 'CA RHD', 'B0004382': 'DF LH', 'B0004383': 'DF RH','B0004300': 'IP L2 LHD', 'B0004304': 'IP L2 RHD', 'B0004301': 'IP L3 LHD', 'B0004330': 'IP L3 RHD', 'B0004303': 'IP L4 LHD', 'B0004380': 'DP LHD', 'B0004381': 'DP RHD'}


    #reports
    df_overview = 'table_of_claims'
    quality_alert_excel_position = ['C5','R5','G6','R6','A8','D7','M45']
    index_position_in_claim_tuple = [5, 2, 3, 1, 8]
    box_label_excel_position = ['A1', 'A11', 'A21', 'A31', 'A41']

    #training lists
    training_list_position = ['C5', 'A8', 14, 2]

    #source:

    name_production_list_SPRAY = '/Průběh směny P8 SPRAY R011 KW'
    look_in_spray = ['Souhrny', 197, 18, 'B:Q', 218, 18]

    name_production_list_ROM = 'Průběh směny P8 B8 ROM R12 KW'
    look_in_ROM = ['OBSAZENÍ', 1, 150, 'A:G']



    #mail message

    mail_claim_recipients = 'claim_wictims@industry.com'
    mail_claim_recipients_Cc = 'someoneincopy@industry.com'

    mail_subject = 'oznámení reklamace'

    def invert_dictionary(dic):
        inv = {v: k for k, v in dic.items()}
        return inv

    def get_dictionary_from_csv(path):
        import pandas as pd

        df = pd.read_csv(path)
        dictionary = dict(zip(df.part_num, df.part_shortcut))
        return dictionary










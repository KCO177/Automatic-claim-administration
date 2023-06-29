import win32com.client
import datetime
import os
import re

#from db_quality import Report, Aggregation, Time
from qualita_db_dictionary import MyDict as Dic
#from run_gui import Mail
from datetime import timedelta
from collections import ChainMap
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
from thefuzz import fuzz
#from mail_output import Mail_to_send


class Mail_recieve:

    def list_of_incoming_mails():
        outlook = win32com.client.Dispatch('outlook.application').GetNamespace('MAPI')

        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items  # .Folders
        messages = messages.Restrict("[SenderEmailAddress] = 'petr.koscelnik@ascorium.com'")

        messages.Sort('[ReceivedTime]', Descending=True)
        for message in list(messages)[:30]:
            print(message.ReceivedTime, message.SenderEmailAddress, message.Subject)

    def read_mail():

        #>>> napojit vyber z GUI

        try:
            outlook = win32com.client.Dispatch('outlook.application').GetNamespace('MAPI')

            inbox = outlook.GetDefaultFolder(6)
            messages = inbox.Items  # .Folders
            current_time = datetime.datetime.now()
            received_dt = current_time - timedelta(days=31)
            received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
            messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
            messages = messages.Restrict("[SenderEmailAddress] = 'petr.koscelnik@ascorium.com'")
            message = messages.GetFirst()  # GetLast(), GetNext(), . Attachments

            # message.Subject #.Body, .To, .Recipients, .Sender, .Sender.Address
            if message.subject != None:
                message_text = message.subject + ' ' + message.body
            elif message.body != None:
                message_text = message.body
            else:
                message_text = 'no message'

            return message_text #message.subject, message.body
        except Exception as e:
            print(e)
    def select_mail(subj):
        print('subject from Mail_recieve.select_mail:', subj, (type(subj)))

        try:
            outlook = win32com.client.Dispatch('outlook.application').GetNamespace('MAPI')
            inbox = outlook.GetDefaultFolder(6)
            messages = inbox.Items  # .Folders
            #current_time = datetime.datetime.now()
            #received_dt = current_time - timedelta(days=31)
            #received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
            #messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
            #messages = messages.Restrict("[SenderEmailAddress] = 'petr.koscelnik@ascorium.com'")

            #for item in messages:
            #subject = item.Subject
            #subject = subject.encode('latin-1').decode('unicode_escape')  # Decode subject using appropriate encoding
            #subject = Mail_text.remove_specchar(subject)
            #print('encoded subject', subject)
            messages = messages.Restrict(f"[subject] = {subj}") #

            message = messages.GetFirst()  # GetLast(), GetNext(), . Attachments

            print('message from select_mail:', message)

            if message != None:
                # message.Subject #.Body, .To, .Recipients, .Sender, .Sender.Address
                print('message subject', message.subject)
                print('message.body', message.body)

                if message.subject != None:
                    message_text = message.subject + ' ' + message.body
                elif message.body != None:
                    message_text = message.body
                else:
                    message_text = 'no message'

                return message_text #message.subject, message.body
            else:
                print ('no message with subject', subj)
        except Exception as e:
            print(e)

    def save_attachment(subject):
        sub = str(subject)    
        outlook = win32com.client.Dispatch('outlook.application').GetNamespace('MAPI')

        inbox = outlook.GetDefaultFolder(6)
        messages = inbox.Items  # .Folders
        #current_time = datetime.datetime.now()
        #received_dt = current_time - timedelta(days=7)
        #received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
        #messages = messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")
        #messages = messages.Restrict("[SenderEmailAddress] = 'petr.koscelnik@ascorium.com'")
        messages = messages.Restrict(f"[subject] = {sub}")  #

        message = messages.GetFirst()  # GetLast(), GetNext(), . Attachments

        # message.Subject #.Body, .To, .Recipients, .Sender, .Sender.Address
        message_text = message.subject + ' ' + message.body

        outputDir = r'C:\Users\START\PycharmProjects\databaze\mail_test'

        for message in messages:
            attachments = message.attachments
            s = message.sender
            for attachment in attachments:
                print(f'attachment {attachment.FileName}')
                attachment.SaveASFile(os.path.join(outputDir, attachment.FileName))
                print(f'attachment {attachment.FileName} from {s} saved')
class Mail_text:
    def remove_specchar(string):
        cleaned_string = re.sub(r'[^a-zA-Z0-9]', '', string)
        return cleaned_string

    def collect_book(list_of_dic):
        # list of dic in format : (list_02, inv_list_02, list_03)
        from collections import ChainMap
        book = ChainMap(list_of_dic)
        return book
    def fuzzy_text(message, book, threshold):

        try:
            message_str = str(message)
            # transform string to list
            message_list = message_str.split()
            # input dictionaries and lists
            if threshold == 0:
                threshold_set = 80
            else:
                threshold_set = threshold
            match = []
            for i in message_list:
                for j in book:
                    threshold = fuzz.token_set_ratio(i, j)
                    if threshold >= threshold_set:
                        #check your matchings
                        #print(f'Token set ratio similarity :{i} : {j} score {threshold}')
                        match.append(j)

        except:
            match = None
        return match
    def recognise_claim_in_mail(message):
        #print(message)
        word = 'claim'
        for item in message:
            if word in item:
                print(f'claim in message')
            else:
                print('no claim in mail inbox for last 24hours')

        return message
    def message_from_mail(message_t):
        # make string for analazying
        separator = ''
        message = separator.join(message_t)
        return message
    def preprocess_text(message):

        #make all letter low A -> a

        # remove all special characters
        example_sent = re.sub(r'[^\w\s]', '', message)
        stop_words = set(stopwords.words('english'))
        word_tokens = word_tokenize(example_sent)
        # converts the words in word_tokens to lower case and then checks whether
        # they are present in stop_words or not
        #filtered_sentence = [w for w in word_tokens if not w.lower() in stop_words]
        # with no lower case conversion
        filtered_sentence = []

        for w in word_tokens:
            if w not in stop_words:
                filtered_sentence.append(w)


        return filtered_sentence
    # find part name
    def find_part(message):
        #print('find part method message', message)
        inv_project_part_dictionary = Dic.invert_dictionary(Dic.p8_parts_dictionary)
        #print('inv_project_part_dictionary ', inv_project_part_dictionary)
        for part in inv_project_part_dictionary:
            if part in message:
                part_name = (inv_project_part_dictionary[part])

                return part_name
            else:
                for part in Dic.ROM_parts:
                    if part in message:

                        return part
                    else:
                        for part in Dic.Spray_parts:
                            if part in message:

                                return part
                            else:
                                for part in Dic.part_id:
                                    if part in message:
                                        part_name = (Dic.part_id[part])

                                        return part_name

                                    else: None
    # find project name
    def find_project(message):
        for project in Dic.projects:
            if project in message:
                return project
            elif Mail_text.find_part(message) != None:
                part = Mail_text.find_part(message)
                if part in Dic.project_dictionary:
                    project = (Dic.project_dictionary[part])
                    return project
                else:
                    None

    def find_customer(message):
        for customer in Dic.customers:
            if customer in message:
                return customer
            elif Mail_text.find_project(message) != None:
                inv_customer_dic = {v: k for k, v in Dic.customers.items()}
                project = Mail_text.find_project(message)
                if project in inv_customer_dic:
                    customer = (inv_customer_dic[project])
                    return customer
                else:
                    None

    def find_part_id(message, part):
        book = Dic.part_id
        #treshold set up on 100
        match = Mail_text.fuzzy_text(message, book, 100)
        #print('part ID match', match)

        # pokud neni part id explicitne v textu zkus matchnout shortcut z metody find part
        a = len(match)
        if a == 0:
            book_reverse = Dic.invert_dictionary(book)
            #print(book_reverse, part)
            part_match = Mail_text.fuzzy_text(part, book_reverse, 90)
            #print(part_match)

            # preved na part id
            match = []
            for part in part_match:
                part_id = book_reverse[part]
                match.append(part_id)
                #print(part_id)

        #print ('part_id_match', match)




        # pokud je zadan dil bez L/R varianty vypis obe varianty oddelene carkou
        if len(match) > 1:
            if match[0] != match[1]:
                part_id = match[0] + ',' + match[1]
        else:
            part_id = match[0]

        print('find part ID', part_id)
        return part_id

    def find_failure(meSSage):
        #remove 'NOK' from message because of frequently use for failures
        # and make false match with failures in failure dictionary
        message = [word.lower() for word in meSSage]
        nok = 0
        while 'nok' in message:
            message.remove('nok')
            nok = +1
        print('nok removed from message', nok)
        book = Dic.failures
        match = Mail_text.fuzzy_text(message, book, 75)
        #check of match
        #print('failure match', match)

        for i in match:
            if i != None:
                #  pokud se vrati vice failure musi se spojit do stringu jinak hodi v nasledujici metode out of index bug
                delimiter = ','
                matched_failure = []
                for item in match:
                    if item not in matched_failure:
                        matched_failure.append(item)
                failure = delimiter.join(matched_failure)

            else:
                failure = None
            return failure

    def find_numbers_in_text(message, match_word):
        # In given lists(message) look for integers and look for matches with fuzz list (match word)
        # If there is int in message look for 3 indexes in front of and behind the int for match word
        # For match word 'claim' is set up only 1 index in front of and behind the int. Mostly presented in subject.
        # Threshold set as default = 80

        #check inputs
        #print(message)
        #print(match_word)

        #try:
        number_indexes = [i for i, token in enumerate(message) if token.isdigit()]
        #check for index with int output
        #print(number_indexes)
        number = []
        for index in number_indexes:
             number.append(message[index])
        #print(number)

        threshold_set = 80
        match = []

        ra = 3
        rb = 4

        if 'claim' in match_word:
            ra = 1
            rb = 2

        for i in number_indexes:

            # out of range case condition >>> nutne doresit else blok
            if i - 2 >= 0 and i + 2 < len(message) != None:


                # print (message [i-3:i+4])
                # nastaveni vyhledavaciho rozsahu okolo shodneho slova
                words_around = (message[i - ra:i + rb])



                #check for inputs
                #print('words around', words_around)
                #print(message [i])

                for w in words_around:
                    for j in match_word:
                        threshold = fuzz.token_set_ratio(w, j)
                        if threshold >= threshold_set:
                            #check for fuzzy match value
                            #print(f'Token set ratio similarity :{w} : {j} score {threshold}')
                            match.append(j)
                            match.append(message[i])

            else:  # i -1 >=0 and i + 2 <len(message):
                print('out of range')
        return match
        #except Exception as e:
        #    print(e)

    def find_nok_stock_numbers(match):
        #count sum of 'noks' and 'stock' findings
        nok_total = 0
        stock_total = 0

        for i in range(0, len(match), 2):
            if match[i] == 'nok':
                nok_total += int(match[i + 1])
            elif match[i] == 'stock':
                stock_total += int(match[i + 1])
        #check for outputs
        #print("stock =", stock_total)
        #print("nok =", nok_total)

        return nok_total, stock_total
        # >>>stock_amount do def pro assumptions
        # stock_amount = amount = def, amount[1]

        return(match)

    def find_claim_numbers(match):

        #method prepare for solving more dificult claim IDs current ID is only int in format 'nnnnnn'
        #print(match)
        return(match)
class Mail_output:

    def text_output(mail_subject):
        #try:

        sub = str(mail_subject)
        print(type(sub))

        print('sub from Mailoutput.text_output', sub, (type(sub)))

        #priprava var pro analyzu textu:
        inv = Dic.invert_dictionary(Dic.p8_parts_dictionary)
        book = ChainMap(Dic.p8_parts_dictionary, inv, Dic.claim_table, Dic.project_dictionary, Dic.projects,Dic.customers,Dic.failures)

        #>>> select claim
        mes = Mail_recieve.select_mail(sub)
        #print('Mail_recieve.select_mail_method', mes)
        message = Mail_text.preprocess_text(Mail_text.message_from_mail(mes))
        print('massage: ', message)

        #message = Mail_text.preprocess_text(Mail_text.message_from_mail(Mail_recieve.read_mail()))

        fuzzy_tuple = Mail_text.fuzzy_text(message,book,75)
        #print('fuzzy_tuple: ', fuzzy_tuple)

        #hledani part
        part = (Mail_text.find_part(message))
        if part == None:
            part = (Mail_text.find_part(fuzzy_tuple))
        elif part == 'CA':
            pass #>>>Debug vypis match set threshold to 95 'CA' se casto vyskytuje v 'Can'
        #print(part)

        #hledani projektu
        project = (Mail_text.find_project(fuzzy_tuple))
        #print(project)

        #hledani customer
        customer = (Mail_text.find_customer(fuzzy_tuple))
        #print(customer)

        #hledani int v textu
        match_word = ['nok', 'stock']
        amount = Mail_text.find_nok_stock_numbers(Mail_text.find_numbers_in_text(message, match_word))
        #print(amount)
        #print(message)

        #nok z int v textu
        #print('nok, stock', amount)
        nok_amount = amount[0]
        stock_amount = amount[1]

        #hledani claim number
        #claim je v pripade tohoto zakaznika ve formatu int v pripade jineho formatu je nutne upravit

        match_word_claim = ['claim']
        claim_number_match = Mail_text.find_numbers_in_text(message, match_word_claim)
        a = len(claim_number_match)
        if a != 0:
            for i in claim_number_match:
                if i != None:
                    claim_number = claim[1]
                    print('claim number', claim_number)
        else:
            claim_number = None
        #print('claim_number', claim_number_match)

        # hledani chyby
        # musi byt po hledani NOK number protoze odstranuje NOK z message kvuli
        # matchi s failure dictionary

        failure = Mail_text.find_failure(message)
        #print(failure)

        # hledani part id
        part_id = Mail_text.find_part_id(message, part)
        #print('part_id', part_id)

        # find date >>> today
        today = datetime.date.today()
        #print(today)

        datum = today.strftime('%Y - %m - %d') #Time.date_format(today)
        #print(datum)
        # datetime.date(2023, 3, 22) to format 2023 - 03 - 09


        ##### kolekce output listu #####
        new_claim = [datum,project,part,part_id,customer, claim_number,'claim',failure, nok_amount, stock_amount]
        print('new_claim:', new_claim)
        return new_claim

        #except Exception as e:
        #    print(e)


        # kontrola
        #print(amount)
        #print('new claim', new_claim)
        #return new_claim

        #except Exception as e:
        #print('exception in outlook management', e)

#print(Mail_recieve.select_mail('claim cupholder 218479'))
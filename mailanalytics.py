## import

from __future__ import print_function

import os.path
import json
import numpy as np
from datetime import datetime
import base64

from matplotlib import pyplot as plt

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


## Classe
class Mailytics():

    def __init__(self):
        print("Thanks for using Mailytics")

        # If modifying these scopes, delete the file token.json.
        self.SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

        self.token_file  = os.path.join('h:', 'Downloads', 'mail_analytics', 'saved_gmail_token.json')
        self.cred_file = os.path.join('h:', 'Downloads', 'mail_analytics', 'gmail_token.json')
        self.email_saved = os.path.join('h:', 'Downloads', 'mail_analytics')

        self.service = None
        self.messages = []

        self.auth()

    def auth(self):
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists(self.token_file):
            creds = Credentials.from_authorized_user_file(self.token_file, self.SCOPES)
        # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    self.cred_file, self.SCOPES)
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open(self.token_file, 'w') as token:
                token.write(creds.to_json())

        try:
            # Call the Gmail API
            self.service = build('gmail', 'v1', credentials=creds)

        except HttpError as error:
            # TODO(developer) - Handle errors from gmail API.
            print(f'An error occurred: {error}')


    ## Collector function
    '''
    Display all the labels of the mail account
    return: None
    '''
    def print_label(self):
        results = self.service.users().labels().list(userId='me').execute()
        labels = results.get('labels', [])

        if not labels:
            print('No labels found.')
            return
        print('Labels:')
        for label in labels:
            print(label['name'])


    '''
    Getting all last emails ids (only inbox mail)
    param: bool with_spam to include spam
    param: int count number of mail to retreive
    return: a list of Dict(Messages)
    '''
    def __get_email_ids_count(self, with_spam, count):
        list = []
        rest = count
        next_page = ''
        while rest > 0:
            # 100(message) is the default length page
            max = 100
            if rest <= max:
                max = rest
            if next_page != '':
                query = self.service.users().messages().list(userId="me", includeSpamTrash=with_spam, maxResults=max, pageToken=next_page, q="in:inbox")
            else:
                query = self.service.users().messages().list(userId="me", includeSpamTrash=with_spam, maxResults=max, q="in:inbox")
            results = query.execute()
            list += results["messages"]
            if "nextPageToken" in results:
                rest -= len(results["messages"])
                next_page = results["nextPageToken"]
            else:
                rest = 0
        return list

    '''
    Getting all last emails ids (only inbox mail)
    param: bool with_spam to include spam
    param: int count number of mail to retreive
    return: a list of Dict(Messages)
    '''
    def __get_email_ids_date(self, with_spam, mdate):
        list = []
        next_page = ''
        end = False
        while not end:
            if next_page != '':
                query = self.service.users().messages().list(userId="me", includeSpamTrash=with_spam, pageToken=next_page, q=f"in:inbox after:{mdate.strftime('%m/%d/%Y')}")
            else:
                query = self.service.users().messages().list(userId="me", includeSpamTrash=with_spam, q=f"in:inbox after:{mdate.strftime('%m/%d/%Y')}")
            results = query.execute()
            list += results["messages"]
            if "nextPageToken" in results:
                next_page = results["nextPageToken"]
            else:
                end = True
        return list

    '''
    Get the content of an email
    param: int id the id of the email you want to get
    return a Message
    '''
    def __get_email_content(self, id):
        query = self.service.users().messages().get(userId="me", id=id)
        results = query.execute()
        return results

    '''
    Get the body content of a mail
    param: Message message
    return: the text body
    '''
    def __get_email_body(self, message):
        # give a partId, return a text
        def search_body(partId):
            if partId.get("mimeType", "andouille") == "text/plain":
                return base64.urlsafe_b64decode(partId["body"]["data"].encode("ASCII")).decode("utf-8")
            else:
                queue = partId.get("parts", [])
                while len(queue) != 0:
                    text_body = search_body(queue.pop(0))
                    if text_body != "":
                        return text_body
                return ""
        return search_body(message["payload"])

    '''
    Get the list of filename of attached document of a mail
    param: Message message
    return: list of String
    '''
    def __get_att_doc(self, message):
        # give a partId, return a list of filename
        def search_body(partId):
            if partId.get("filename", "") != "":
                return [partId["filename"]]
            else:
                queue = partId.get("parts", [])
                filenames = []
                while len(queue) != 0:
                    filenames_temp = search_body(queue.pop(0))
                    if filenames_temp != []:
                        filenames += filenames_temp
                return filenames
        return search_body(message["payload"])

    '''
    Get all the data from the save or the Gmail API
    param: bool refresh force to use the Gmail API and delete old save
    param: int sample number of mail to make the study on
    return: (List of Messages, the number of mail recovered)
    '''
    def loading_messages_count(self, refresh=False, sample=1):
        self.messages = []
        file_email = os.path.join(self.email_saved, 'backup_email_count.txt')
        if refresh:
            print(f"[Collecting {sample} emails]")

            email_list = self.__get_email_ids_count(False, sample)
            print("[Collecting: getting content]")
            for email in email_list:
                self.append(self.__get_email_content(email["id"]))

            print("[Collecting: done]")
            print("[Saving emails]")

            f = open(file_email, 'w', encoding='utf-8')
            f.write(str(self.messages))
            f.close()

            print("[Saving: done]")

        else:
            print("[Loading the emails]")

            f = open(file_email, 'r', encoding='utf-8')
            self.messages = list(eval(f.read()))
            sample = len(self.messages)
            f.close()

            print(f"[Loading: done ({sample} emails loaded)]")

    '''
    Get all the data from the save or the Gmail API
    param: bool refresh force to use the Gmail API and delete old save
    param: int sample number of mail to make the study on
    return: (List of Messages, the number of mail recovered)
    '''
    def loading_messages_date(self, refresh=False, mdate=datetime.today()):
        self.messages = []
        file_email = os.path.join(self.email_saved, 'backup_email_date.txt')
        if refresh:
            print(f"[Collecting emails up from {mdate.strftime('%d/%m/%Y')}]")

            email_list = self.__get_email_ids_date(False, mdate)
            print("[Collecting: getting content]")
            for email in email_list:
                self.messages.append(self.__get_email_content(email["id"]))

            print(f"[Collecting: done ({len(self.messages)} mails)]")
            print("[Saving emails]")

            f = open(file_email, 'w', encoding='utf-8')
            f.write(str(self.messages))
            f.close()

            print("[Saving: done]")

        else:
            print("[Loading the emails]")

            f = open(file_email, 'r', encoding='utf-8')
            self.messages = list(eval(f.read()))
            sample = len(self.messages)
            f.close()

            print(f"[Loading: done ({sample} emails loaded)]")

    ## Stat functions controller
    '''
    Get the most used "from" email
    return: an ordered (by count desc) list of {item, count}
    '''
    def __rank_most_active_from(self):
        sender = {}
        for msg in self.messages:
            for item in msg["payload"]["headers"]:
                if item["name"] == "From":
                    sender[item["value"]] = sender.get(item["value"], 0) + 1

        list = sorted([{'item': key, 'count': value} for key, value in sender.items()], key=lambda d:d["count"], reverse=True)
        return list

    '''
    Get the most used mailing list
    return: an ordered list of {item, count}
    '''
    def __rank_mailinglist_from(self):
        sender = {}
        for msg in self.messages:
            is_listed = False
            for item in msg["payload"]["headers"]:
                if item["name"] == "List-ID":
                    is_listed = True
                    sender[item["value"]] = sender.get(item["value"], 0) + 1
            if not is_listed:
                key = "Sans liste"
                sender[key] = sender.get(key, 0) + 1

        list = sorted([{'item': key, 'count': value} for key, value in sender.items()], key=lambda d:d["count"], reverse=True)
        return list

    '''
    Get the count of mail with the world "newsletter"
    return: list of {item, count}
    '''
    def __rank_count_newsletter(self):
        sender = {}
        label_newsletter = "Newsletter"
        lebel_reste = "Reste"
        searching_key = "newsletter"
        for msg in self.messages:
            is_newslettered = False
            for item in msg["payload"]["headers"]:
                # if "newsletter" in the subject
                if item["name"] == "Subject":
                    if searching_key in item["value"].lower():
                        sender[label_newsletter] = sender.get(label_newsletter, 0) + 1
                        is_newslettered = True
            # if the keyword is in the body document
            if not is_newslettered:
                body_text = self.__get_email_body(msg)
                if searching_key in body_text.lower():
                    is_newslettered = True
                    sender[label_newsletter] = sender.get(label_newsletter, 0) + 1
            # not key word detected
            if not is_newslettered:
                sender[lebel_reste] = sender.get(lebel_reste, 0) + 1

        list = sorted([{'item': key, 'count': value} for key, value in sender.items()], key=lambda d:d["count"], reverse=True)
        return list

    '''
    Get the top mail with attached document
    return: list of {item, count}
    '''
    def __rank_most_att_doc(self):
        sender = {}
        for msg in self.messages:
            filenames = self.__get_att_doc(msg)
            nb_att_doc = len(filenames)
            sender[nb_att_doc] = sender.get(nb_att_doc, 0) + 1

        list = sorted([{'item': key, 'count': value} for key, value in sender.items()], key=lambda d:d["count"], reverse=True)
        return list

    ## Stat function displayer
    def display_piechart_aux(self, ranking_list, details_pie_chart=12,labels={}):
        # labels
        name_ranking = labels.get("name_ranking", "Mails Ranking")
        name_item = labels.get("name_item", "mails")

        print(f"[{name_ranking}] {len(ranking_list)} éléments différents")
        top_ranking_list= [ranking_list[0]["item"]]
        ranking_list_i = 0
        while ranking_list_i < len(ranking_list)-1 and ranking_list[ranking_list_i+1]["count"] == ranking_list[0]["count"]:
            top_ranking_list.append(ranking_list[ranking_list_i+1]["item"])
            rmf_i += 1

        print(f"[{name_ranking}] Le(ou les) gagnant est : {top_ranking_list}, avec un score de {ranking_list[0]['count']} mails touchés (chacun des {len(top_ranking_list)} gagnants)")

        # displaying a graph
        ranking_labels = []
        ranking_values = []
        show_max = details_pie_chart if len(top_ranking_list) <= details_pie_chart or len(top_ranking_list) >= 20 else len(top_ranking_list)
        counter_pie_chart = 1
        for item in ranking_list:
            if counter_pie_chart <= show_max:
                ranking_labels.append(item["item"])
                ranking_values.append(item["count"])
            elif counter_pie_chart == show_max + 1:
                ranking_labels.append("Reste")
                ranking_values.append(item["count"])
            else:
                ranking_values[show_max] += item["count"]
            counter_pie_chart += 1

        # wedges, text = plt.pie(ranking_values, autopct=make_autopct(ranking_values), labels=ranking_labels, startangle=90)
        wedges, text, per_txt = plt.pie(ranking_values, startangle=0, wedgeprops=dict(width=0.5), autopct=self.make_autopct(ranking_values), pctdistance=0.7)

        bbox_props = dict(boxstyle="square,pad=0.3", fc="w", ec="k", lw=0.72)
        kw = dict(arrowprops=dict(arrowstyle="-"),
                bbox=bbox_props, zorder=0, va="center")

        for i, p in enumerate(wedges):
            ang = (p.theta2 - p.theta1)/2. + p.theta1
            y = np.sin(np.deg2rad(ang))
            x = np.cos(np.deg2rad(ang))
            horizontalalignment = {-1: "right", 1: "left"}[int(np.sign(x))]
            connectionstyle = "angle,angleA=0,angleB={}".format(ang)
            kw["arrowprops"].update({"connectionstyle": connectionstyle})
            plt.annotate(ranking_labels[i], xy=(x, y), xytext=(1.35*np.sign(x), 1.4*y),
                        horizontalalignment=horizontalalignment, **kw)

        plt.title(f"[{name_ranking}] Les {details_pie_chart} {name_item} les plus actifs parmi les {len(self.messages)} mails collectés", pad=50)
        plt.show()

    def make_autopct(self, values):
        def my_autopct(pct):
            total = sum(values)
            val = int(round(pct*total/100.0))
            return '{p:.2f}%  ({v:d})'.format(p=pct,v=val)
        return my_autopct

    def __display_stat_headers(self):
        # on va rassembler tous les name de headers des mails pour voir s'il y a des mailing list
        # # stocker un exemple à chaque fois
        # print("tous les headers")
        # headers = {}
        # for msg in self.messages:
        #     for head in msg["payload"]["headers"]:
        #         if not head["name"] in headers:
        #             headers[head["name"]] = head["value"]
        #
        # # displaying
        # for key, value in headers.items():
        #     print(f"[{key}] : {value}")

        print("resultat de l'analyse")
        res = []
        for msg in self.messages:
            List_ID = False
            List_Id = False
            Mailing_list = False
            for head in msg["payload"]["headers"]:
                List_ID = List_ID or "List-ID" == head["name"]
                List_Id = List_Id or "List-Id" == head["name"]
                Mailing_list = Mailing_list or "Mailing-list" == head["name"]
            res.append({"List_ID": List_ID, "List_Id": List_Id, "Mailing_list": Mailing_list})

        nb_total = len(res)
        nb_trois = 0
        nb_mail_Id = 0
        nb_mail_ID = 0
        nb_Id_ID = 0
        nb_mail = 0
        nb_Id = 0
        nb_ID = 0
        nb_zero = 0
        for item in res:
            mail = 1 if item["Mailing_list"] else 0
            Id = 1 if item["List_Id"] else 0
            ID = 1 if item["List_ID"] else 0
            tot = mail + Id + ID
            if tot == 3:
                nb_trois += 1
            elif tot == 2:
                if not item["List_Id"]:
                    nb_mail_ID += 1
                elif not item["List_ID"]:
                    nb_mail_Id += 1
                else:
                    nb_Id_ID += 1
            elif tot == 1:
                if item["List_Id"]:
                    nb_Id += 1
                elif item["List_ID"]:
                    nb_ID += 1
                else:
                    nb_mail += 1
            else:
                nb_zero += 1
        print("total :", nb_total)
        print("mail Id :", nb_mail_Id)
        print("mail ID :", nb_mail_ID)
        print("Id ID :", nb_Id_ID)
        print("mail :", nb_mail)
        print("Id :", nb_Id)
        print("ID :", nb_ID)
        print("zero :", nb_zero)

    def display_rank_most_active_from(self):
        rmf_list = self.__rank_most_active_from()
        self.display_piechart_aux(rmf_list, details_pie_chart=12,labels={"name_ranking": "rank_most_active_from", "name_item": "envoyeurs"})

    def display_rank_mailinglist_from(self):
        rmf_list = self.__rank_mailinglist_from()
        self.display_piechart_aux(rmf_list, details_pie_chart=12,labels={"name_ranking": "rank_mailinglist_from", "name_item": "mailinglist"})

    def display_count_newsletter(self):
        cnl_list = self.__rank_count_newsletter()
        self.display_piechart_aux(cnl_list, details_pie_chart=2,labels={"name_ranking": "count_newletter", "name_item": "types"})

    def display_most_att_doc(self):
        mad_list = self.__rank_most_att_doc()
        print(mad_list)
        self.display_piechart_aux(mad_list, details_pie_chart=5,labels={"name_ranking": "most_att_doc", "name_item": "pièces jointes"})

## Main sequence
# parameters
sample = 600
mdate = datetime.strptime("01/08/2022", "%d/%m/%Y")

# create a new mailytics object
mails = Mailytics()

# getting messages
# mails.loading_messages_count(refresh=False, sample=sample)
mails.loading_messages_date(refresh=False, mdate=mdate)

# displaying stats
# mails.display_rank_most_active_from()
# mails.display_rank_mailinglist_from()
# mails.display_count_newsletter()
mails.display_most_att_doc()



"""
Nos paramètres à tester :
- l'envoyeur
- les destinataires
- s'il contient des pièces jointes (combien ont des pièces jointes)
- s'il contient certaint mot (ex: newsletter)

Fonctions:
    - classement des envoyeurs les plus actifs
    - le top des mailing lists
    - proportion des mails avec le mot clé "newsletter"
    - proportion des mails avec pièces jointes sur le total
- proportion des mails avec pièces jointes sur le total des mails envoyé par le top 1 envoyeur


Sender, Subject, From, To, Cc, Date, (Mailing-list)
"""

# TODO
"""
Ajouter une option sur le loader de date pour donner une date d'arrivé en plus de la date de départ
"""

















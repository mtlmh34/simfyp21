import os
import re

from flask import Flask, render_template, url_for, request, flash, redirect
import imaplib
# For connection
import easyimap as e
import smtplib
import logging
import xlsxwriter
from openpyxl import load_workbook
# For Machine Learning
import pandas as pd
from joblib import dump, load
import numpy as np
# import tkinter
# from tkinter import *
# import tkinter.messagebox

import nltk

nltk.download('punkt')
nltk.download('stopwords')
from nltk import PorterStemmer, word_tokenize
from nltk.corpus import stopwords
from functions import mainFunctions

# Import function from another python file
from sklearn.feature_extraction.text import CountVectorizer

app = Flask(__name__)


@app.route('/', methods=['GET', 'POST'])
def login():
    loading_gif = url_for('static', filename='imgload.gif')
    image_file = url_for('static', filename='login.png')  # image
    # declare variables
    global server
    global username
    global password
    global imap_url
    global platform
    try:
        if request.method == 'POST':
            username = request.form['email']  # id='email' from html form
            password = request.form['password']
            platform = request.form['platform']

            # Authenticates and retrieves email
            if platform == 'Gmail':
                imap_url = 'imap.gmail.com'
            else:
                imap_url = 'outlook.office365.com'
            server = e.connect(imap_url, username, password)

            return redirect('/inbox')
        else:
            return render_template('index.html', image_file=image_file, loading_gif=loading_gif)

    except imaplib.IMAP4.error:
        '''window = Tk()
        window.attributes('-topmost', True)
        window.wm_withdraw()
        window.geometry(f"1x1+{round(window.winfo_screenwidth() / 2)}+{round(window.winfo_screenheight() / 2)}")
        tkinter.messagebox.showerror(title="Invalid credentials", message="Please re-enter user credentials",
                                     parent=window)'''
        return render_template('index.html', image_file=image_file)


##############################################################################################

def cleaning(string):
    string = re.sub("[^0-9a-zA-Z\ ]", "", str(string))
    string = string.lower()
    string = string.strip()

    return string


def stem(string):
    tokenized = word_tokenize(string)
    stemmed = []
    stemmer = PorterStemmer()

    for word in tokenized:
        stemmed.append(stemmer.stem(word))

    return ' '.join(stemmed)


def remove_stopwords(string):
    STOP_WORDS = set(stopwords.words('english'))

    tokenized = word_tokenize(string)
    filtered = []

    for word in tokenized:
        if word not in STOP_WORDS:
            filtered.append(word)

    return " ".join(filtered)


######################################################################################

@app.route('/inbox')
# retrieve and log emails to csv and perform ML prediction
def email():
    global email
    global subject_list
    global body_list
    global email_address_list
    global percentage_list
    global result_list

    percentage_list = []
    subject_list = []
    body_list = []
    email_address_list = []
    date_list = []
    result_list = []
    # Authenticates and retrieves email

    # logging emails
    logger = logging.getLogger('logger1')
    handler = logging.FileHandler('log1.txt')
    logger.addHandler(handler)
    logger.setLevel(logging.INFO)

    server = e.connect(imap_url, username, password)
    # inbox = server.listup()
    inbox = server.listids()
    if os.path.isfile('logs.xlsx'):
        # Check if email sheet exist
        wb = load_workbook("logs.xlsx")  # open an Excel file and return a workbook
        # check if email sheet exist
        if username in wb.sheetnames:  # user have login before
            print(username + ' exists')
            # server = e.connect(imap_url, username, password)
            email = server.mail(server.listids()[0])
            ws = wb[username]
            excelRow = 2
            '''test = ws["A2"].value
            test = test.strip()
            test1 = email.title
            test1 = test1.strip(test1)
            for a, b in zip(test, test1):
                if a != b:
                    print(a + b)'''
            if ws["A2"].value == email.title and ws["B2"].value == email.from_addr:
                print("This is working")
                for x in range(1, ws.max_row):
                    # read from excel file
                    excelColumn = 'A'
                    excelPosition = excelColumn + str(excelRow)
                    subject_list.append(ws[excelPosition].value)
                    print(excelPosition)

                    excelColumn = chr(ord(excelColumn) + 1)
                    excelPosition = excelColumn + str(excelRow)
                    email_address_list.append(ws[excelPosition].value)
                    print(excelPosition)

                    excelColumn = chr(ord(excelColumn) + 1)
                    excelPosition = excelColumn + str(excelRow)
                    body_list.append(ws[excelPosition].value)
                    print(excelPosition)

                    excelColumn = chr(ord(excelColumn) + 1)
                    excelPosition = excelColumn + str(excelRow)
                    print(excelPosition)

                    excelColumn = chr(ord(excelColumn) + 1)
                    excelPosition = excelColumn + str(excelRow)
                    result_list.append(ws[excelPosition].value)
                    print(excelPosition)

                    excelColumn = chr(ord(excelColumn) + 1)
                    excelPosition = excelColumn + str(excelRow)
                    print(excelPosition)

                    excelColumn = chr(ord(excelColumn) + 1)
                    excelPosition = excelColumn + str(excelRow)
                    print(excelPosition)

                    excelColumn = chr(ord(excelColumn) + 1)
                    excelPosition = excelColumn + str(excelRow)
                    print(excelPosition)

                    excelColumn = chr(ord(excelColumn) + 1)
                    excelPosition = excelColumn + str(excelRow)
                    percentage_list.append(ws[excelPosition].value)
                    print(excelPosition)

                    excelRow += 1
                    wb.save("logs.xlsx")

            '''else:  # user have no new email
                print("Finding new email")
                checkTitle = ws["A2"].value
                checkAddr = ws["B2"].value
                # server = e.connect(imap_url, username, password)
                # inbox = server.listids()
                for x in range(0, len(inbox)):  # change 10 to len(inbox) to get 100 mails
                    email = server.mail(server.listids()[x])
                    if checkTitle == email.title and checkAddr == email.from_addr:
                        break
                    else:
                        # log.txt file
                        logger.info("----------------------------------------------------------------")
                        logger.info("Email Title:")
                        logger.info(email.title)
                        logger.info("Email from:")
                        logger.info(email.from_addr)
                        logger.info("Message: ")
                        logger.info(email.body)

                        string = email.body
                        string = cleaning(string)
                        string = stem(string)
                        string = remove_stopwords(string)

                        # ML
                        n_df = pd.DataFrame({'text': string}, index=[0])
                        n_df.head()
                        vectorizer = load(r'naivebayesVectorizer.joblib')  # load vectorizer
                        nbclf = load(r'naivebayes.joblib')  # load the naivebayes ml model

                        x_matrix = vectorizer.transform(n_df['text'])
                        my_prediction = nbclf.predict(x_matrix)
                        percentage = nbclf.predict_proba(x_matrix)
                        # percentage = np.array(percentage)
                        # percentage = ['{:f}'.format(item) for item in percentage]
                        np.set_printoptions(formatter={'float_kind': '{:f}'.format})

                        if my_prediction == 1:
                            ml_result = 'Phishing'
                            percentage = format(percentage[0][1], '.12f')  # to 12decimal place
                            percentage = float(percentage) * 100  # convert to percent
                            percentage = str(percentage) + '%'
                        elif my_prediction == 0:
                            ml_result = 'Non-Phishing'
                            percentage = format(percentage[0][0], '.12f')  # to 12decimal place
                            percentage = float(percentage) * 100  # convert to percent
                            percentage = str(percentage) + '%'

                        logger.info("Email attachment: ")
                        logger.info(email.attachments)
                        emailAttachment = email.attachments
                        if not emailAttachment:
                            emailAttach = "-"
                        else:
                            attachment = emailAttachment[0]
                            attachment = str(attachment)
                            attach = attachment.split(',')
                            emailAttach = str(attach[0])
                            emailAttach = emailAttach[1:]
                        logger.info("----------------------------------------------------------------")

                        # Run function and counter check with ML result
                        functionResult = 100
                        emailConFormat = mainFunctions.content_formatting(email.body)
                        # spellingResult = mainFunctions.spelling_check(emailConFormat)
                        spellingResult = mainFunctions.spelling_check(str(email.title))
                        emailValidResult = mainFunctions.email_valid(email.from_addr)
                        attachmentResult = mainFunctions.attachment_check(emailAttach)
                        linkResult = mainFunctions.check_link(email.body)

                        function_indi = " || "
                        # compile result
                        if spellingResult:
                            functionResult -= 25
                            function_indi += " A"
                        else:
                            function_indi += " 0"

                        if emailValidResult:
                            functionResult -= 25
                            function_indi += " B"
                        else:
                            function_indi += " 0"

                        if attachmentResult:
                            functionResult -= 25
                            function_indi += " C"
                        else:
                            function_indi += " 0"

                        if linkResult:
                            functionResult -= 25
                            function_indi += " D"
                        else:
                            function_indi += " 0"

                        if functionResult > 50:
                            function_result = 'Phishing'
                        else:
                            function_result = 'Non-Phishing'

                        # counter check
                        if function_result == ml_result:
                            result = ml_result
                        else:
                            result = 'Suspicious'

                        # Extra just to check
                        function_result = function_result + " || " + str(functionResult) + function_indi
                        # insert row in excel
                        ws.insert_rows(2)
                        # excel file
                        excelColumn = 'A'
                        excelPosition = excelColumn + str(excelRow)
                        ws[excelPosition] = email.title
                        print(excelPosition)

                        excelColumn = chr(ord(excelColumn) + 1)
                        excelPosition = excelColumn + str(excelRow)
                        ws[excelPosition] = email.from_addr
                        print(excelPosition)

                        excelColumn = chr(ord(excelColumn) + 1)
                        excelPosition = excelColumn + str(excelRow)
                        ws[excelPosition] = emailConFormat
                        print(excelPosition)

                        excelColumn = chr(ord(excelColumn) + 1)
                        excelPosition = excelColumn + str(excelRow)
                        ws[excelPosition] = "-"
                        print(excelPosition)

                        excelColumn = chr(ord(excelColumn) + 1)
                        excelPosition = excelColumn + str(excelRow)
                        ws[excelPosition] = result
                        print(excelPosition)

                        excelColumn = chr(ord(excelColumn) + 1)
                        excelPosition = excelColumn + str(excelRow)
                        ws[excelPosition] = emailAttach
                        print(excelPosition)

                        excelColumn = chr(ord(excelColumn) + 1)
                        excelPosition = excelColumn + str(excelRow)
                        ws[excelPosition] = ml_result
                        print(excelPosition)

                        excelColumn = chr(ord(excelColumn) + 1)
                        excelPosition = excelColumn + str(excelRow)
                        ws[excelPosition] = function_result
                        print(excelPosition)

                        excelColumn = chr(ord(excelColumn) + 1)
                        excelPosition = excelColumn + str(excelRow)
                        ws[excelPosition] = percentage
                        print(excelPosition)

                        excelRow += 1
                        # wb.save("logs.xlsx")'''
            # wb.save("logs.xlsx")
            # load into list for web
            wb = load_workbook("logs.xlsx")
            wsNew = wb[username]
            print(wsNew["A2"].value)
            print(excelRow)
            excelRow = 2
            for x in range(1, wsNew.max_row):
                # read from excel file
                excelColumn = 'A'
                excelPosition = excelColumn + str(excelRow)
                subject_list.append(wsNew[excelPosition].value)
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                email_address_list.append(wsNew[excelPosition].value)
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                body_list.append(wsNew[excelPosition].value)
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                result_list.append(wsNew[excelPosition].value)
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                percentage_list.append(wsNew[excelPosition].value)
                print(excelPosition)

                excelRow += 1

        else:  # create new sheet for new email
            print(username + ' does not exists')
            wb.create_sheet(username)
            print(username + ' created')
            ws = wb[username]
            excelRow = 2
            # server = e.connect(imap_url, username, password)
            # inbox = server.listids()
            for x in range(0, len(inbox)):  # change 10 to len(inbox) to get 100 mails
                email = server.mail(server.listids()[x])
                # log.txt file
                logger.info("----------------------------------------------------------------")
                logger.info("Email Title:")
                logger.info(email.title)
                logger.info("Email from:")
                logger.info(email.from_addr)
                logger.info("Message: ")
                logger.info(email.body)

                # store email subject, body in list
                email_address_list.append(email.from_addr)
                subject_list.append(email.title)
                body = mainFunctions.content_formatting(email.body)
                body_list.append(body)

                # ML
                string = body
                string = cleaning(string)
                string = stem(string)
                string = remove_stopwords(string)

                n_df = pd.DataFrame({'text': string}, index=[0])
                n_df.head()
                vectorizer = load(r'naivebayesVectorizer.joblib')  # load vectorizer
                nbclf = load(r'naivebayes.joblib')  # load the naivebayes ml model
                # nbclf = load(r'mlp.joblib')

                x_matrix = vectorizer.transform(n_df['text'])
                my_prediction = nbclf.predict(x_matrix)
                percentage = nbclf.predict_proba(x_matrix)
                # percentage = np.array(percentage)
                # percentage = ['{:f}'.format(item) for item in percentage]
                np.set_printoptions(formatter={'float_kind': '{:f}'.format})

                if my_prediction == 1:
                    ml_result = 'Phishing'
                    percentage = format(percentage[0][1], '.12f')  # to 12decimal place
                    percentage = float(percentage) * 100  # convert to percent
                    percentage = str(percentage) + '%'
                    percentage_list.append(percentage)
                elif my_prediction == 0:
                    ml_result = 'Non-Phishing'
                    percentage = format(percentage[0][0], '.12f')  # to 12decimal place
                    percentage = float(percentage) * 100  # convert to percent
                    percentage = str(percentage) + '%'
                    percentage_list.append(percentage)

                logger.info("Email attachment: ")
                logger.info(email.attachments)
                emailAttachment = email.attachments
                if not emailAttachment:
                    emailAttach = "-"
                else:
                    attachment = emailAttachment[0]
                    attachment = str(attachment)
                    attach = attachment.split(',')
                    emailAttach = str(attach[0])
                    emailAttach = emailAttach[1:]
                logger.info("----------------------------------------------------------------")

                # Run function and counter check with ML result
                functionResult = 100
                emailConFormat = mainFunctions.content_formatting(email.body)
                # spellingResult = mainFunctions.spelling_check(emailConFormat)
                spellingResult = mainFunctions.spelling_check(str(email.title))
                emailValidResult = mainFunctions.email_valid(email.from_addr)
                attachmentResult = mainFunctions.attachment_check(emailAttach)
                linkResult = mainFunctions.check_link(email.body)

                # compile result
                if spellingResult:
                    functionResult -= 25
                if emailValidResult:
                    functionResult -= 25
                if attachmentResult:
                    functionResult -= 25
                if linkResult:
                    functionResult -= 25

                if functionResult > 50:
                    function_result = 'Phishing'
                else:
                    function_result = 'Non-Phishing'

                # counter check
                if function_result == ml_result:
                    result = ml_result
                else:
                    result = 'Suspicious'

                result_list.append(result)

                # excel file titles
                excelColumn1 = 'A1'
                excelPosition1 = excelColumn1
                ws[excelPosition1] = "Email Subject"
                # ws.write(excelPosition1, "Email Subject")
                print(excelPosition1)

                excelColumn2 = 'B1'
                excelPosition2 = excelColumn2
                ws[excelPosition2] = "Name and Email Address"
                # ws.write(excelPosition2, "Name and Email Address")
                print(excelPosition2)

                excelColumn3 = 'C1'
                excelPosition3 = excelColumn3
                ws[excelPosition3] = "Email Content"
                # ws.write(excelPosition3, "Email Content")
                print(excelPosition3)

                excelColumn4 = 'D1'
                excelPosition4 = excelColumn4
                ws[excelPosition4] = "Listing"
                # ws.write(excelPosition4, "Listing")
                print(excelPosition4)

                excelColumn5 = 'E1'
                excelPosition5 = excelColumn5
                ws[excelPosition5] = "Classification"
                # ws.write(excelPosition5, "Classification")
                print(excelPosition5)

                excelColumn6 = 'F1'
                excelPosition6 = excelColumn6
                ws[excelPosition6] = "Attachment"
                # ws.write(excelPosition6, "Attachment")
                print(excelPosition6)

                excelColumn7 = 'G1'
                excelPosition7 = excelColumn7
                ws[excelPosition7] = "Ml Result"
                # ws.write(excelPosition6, "ML Result")
                print(excelPosition7)

                excelColumn8 = 'H1'
                excelPosition8 = excelColumn8
                ws[excelPosition8] = "Function Result"
                # ws.write(excelPosition8, "Function Result")
                print(excelPosition8)

                excelColumn9 = 'I1'
                excelPosition9 = excelColumn9
                ws[excelPosition9] = "Percentage"
                # ws.write(excelPosition9, "Percentage")
                print(excelPosition9)

                # excel file
                excelColumn = 'A'
                excelPosition = excelColumn + str(excelRow)
                ws[excelPosition] = email.title
                # ws.write(excelPosition, email.title)
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                ws[excelPosition] = email.from_addr
                # ws.write(excelPosition, email.from_addr)
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                ws[excelPosition] = emailConFormat
                # ws.write(excelPosition, emailConFormat)
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                ws[excelPosition] = "-"
                # ws.write(excelPosition, "-")
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                ws[excelPosition] = result
                # ws.write(excelPosition, result)
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                ws[excelPosition] = emailAttach
                # ws.write(excelPosition, emailAttach)
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                ws[excelPosition] = ml_result
                # ws.write(excelPosition, ml_result)
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                ws[excelPosition] = function_result
                # ws.write(excelPosition, function_result)
                print(excelPosition)

                excelColumn = chr(ord(excelColumn) + 1)
                excelPosition = excelColumn + str(excelRow)
                ws[excelPosition] = percentage
                # ws.write(excelPosition, percentage)
                print(excelPosition)

                excelRow += 1
            # ws.close()
            wb.save("logs.xlsx")

    else:
        # excel file
        excel = xlsxwriter.Workbook('logs.xlsx')
        bold = excel.add_format({'bold': True})
        worksheet = excel.add_worksheet(username)
        excelRow = 2

        # server = e.connect(imap_url, username, password)
        # inbox = server.listup()
        # inbox = server.listids()
        email = server.mail(server.listids()[0])

        for x in range(0, len(inbox)):  # change 10 to len(inbox) to get 100 mails
            email = server.mail(server.listids()[x])
            # log.txt file
            logger.info("----------------------------------------------------------------")
            logger.info("Email Title:")
            logger.info(email.title)
            logger.info("Email from:")
            logger.info(email.from_addr)
            logger.info("Message: ")
            logger.info(email.body)

            # store email subject, body in list
            email_address_list.append(email.from_addr)
            subject_list.append(email.title)
            body = mainFunctions.content_formatting(email.body)
            body_list.append(body)

            # ML
            string = body
            string = cleaning(string)
            string = stem(string)
            string = remove_stopwords(string)

            n_df = pd.DataFrame({'text': string}, index=[0])
            n_df.head()
            vectorizer = load(r'naivebayesVectorizer.joblib')  # load vectorizer
            nbclf = load(r'naivebayes.joblib')  # load the naivebayes ml model
            # nbclf = load(r'mlp.joblib')

            x_matrix = vectorizer.transform(n_df['text'])
            my_prediction = nbclf.predict(x_matrix)
            percentage = nbclf.predict_proba(x_matrix)
            # percentage = np.array(percentage)
            # percentage = ['{:f}'.format(item) for item in percentage]
            np.set_printoptions(formatter={'float_kind': '{:f}'.format})

            if my_prediction == 1:
                ml_result = 'Phishing'
                percentage = format(percentage[0][1], '.12f')  # to 12decimal place
                percentage = float(percentage) * 100  # convert to percent
                percentage = str(percentage) + '%'
                percentage_list.append(percentage)
            elif my_prediction == 0:
                ml_result = 'Non-Phishing'
                percentage = format(percentage[0][0], '.12f')  # to 12decimal place
                percentage = float(percentage) * 100  # convert to percent
                percentage = str(percentage) + '%'
                percentage_list.append(percentage)

            logger.info("Email attachment: ")
            logger.info(email.attachments)
            emailAttachment = email.attachments
            if not emailAttachment:
                emailAttach = "-"
            else:
                attachment = emailAttachment[0]
                attachment = str(attachment)
                attach = attachment.split(',')
                emailAttach = str(attach[0])
                emailAttach = emailAttach[1:]
            logger.info("----------------------------------------------------------------")

            # Run function and counter check with ML result
            functionResult = 100
            emailConFormat = mainFunctions.content_formatting(email.body)
            # spellingResult = mainFunctions.spelling_check(emailConFormat)
            spellingResult = mainFunctions.spelling_check(str(email.title))
            emailValidResult = mainFunctions.email_valid(email.from_addr)
            attachmentResult = mainFunctions.attachment_check(emailAttach)
            linkResult = mainFunctions.check_link(email.body)

            # compile result
            if spellingResult:
                functionResult -= 25
            if emailValidResult:
                functionResult -= 25
            if attachmentResult:
                functionResult -= 25
            if linkResult:
                functionResult -= 25

            if functionResult > 50:
                function_result = 'Phishing'
            else:
                function_result = 'Non-Phishing'

            # counter check
            if function_result == ml_result:
                result = ml_result
            else:
                result = 'Suspicious'

            result_list.append(result)

            # excel file titles
            excelColumn1 = 'A1'
            excelPosition1 = excelColumn1
            worksheet.write(excelPosition1, "Email Subject", bold)
            print(excelPosition1)

            excelColumn2 = 'B1'
            excelPosition2 = excelColumn2
            worksheet.write(excelPosition2, "Name and Email Address", bold)
            print(excelPosition2)

            excelColumn3 = 'C1'
            excelPosition3 = excelColumn3
            worksheet.write(excelPosition3, "Email Content", bold)
            print(excelPosition3)

            excelColumn4 = 'D1'
            excelPosition4 = excelColumn4
            worksheet.write(excelPosition4, "Listing", bold)
            print(excelPosition4)

            excelColumn5 = 'E1'
            excelPosition5 = excelColumn5
            worksheet.write(excelPosition5, "Classification", bold)
            print(excelPosition5)

            excelColumn6 = 'F1'
            excelPosition6 = excelColumn6
            worksheet.write(excelPosition6, "Attachment", bold)
            print(excelPosition6)

            excelColumn6 = 'G1'
            excelPosition6 = excelColumn6
            worksheet.write(excelPosition6, "ML Result", bold)
            print(excelPosition6)

            excelColumn6 = 'H1'
            excelPosition6 = excelColumn6
            worksheet.write(excelPosition6, "Function Result", bold)
            print(excelPosition6)

            excelColumn7 = 'I1'
            excelPosition7 = excelColumn7
            worksheet.write(excelPosition7, "Percentage", bold)
            print(excelPosition7)

            # excel file
            excelColumn = 'A'
            excelPosition = excelColumn + str(excelRow)
            worksheet.write(excelPosition, email.title)
            print(excelPosition)

            excelColumn = chr(ord(excelColumn) + 1)
            excelPosition = excelColumn + str(excelRow)
            worksheet.write(excelPosition, email.from_addr)
            print(excelPosition)

            excelColumn = chr(ord(excelColumn) + 1)
            excelPosition = excelColumn + str(excelRow)
            worksheet.write(excelPosition, emailConFormat)
            print(excelPosition)

            excelColumn = chr(ord(excelColumn) + 1)
            excelPosition = excelColumn + str(excelRow)
            worksheet.write(excelPosition, "-")
            print(excelPosition)

            excelColumn = chr(ord(excelColumn) + 1)
            excelPosition = excelColumn + str(excelRow)
            worksheet.write(excelPosition, result)
            print(excelPosition)

            excelColumn = chr(ord(excelColumn) + 1)
            excelPosition = excelColumn + str(excelRow)
            worksheet.write(excelPosition, emailAttach)
            print(excelPosition)

            excelColumn = chr(ord(excelColumn) + 1)
            excelPosition = excelColumn + str(excelRow)
            worksheet.write(excelPosition, ml_result)
            print(excelPosition)

            excelColumn = chr(ord(excelColumn) + 1)
            excelPosition = excelColumn + str(excelRow)
            worksheet.write(excelPosition, function_result)
            print(excelPosition)

            excelColumn = chr(ord(excelColumn) + 1)
            excelPosition = excelColumn + str(excelRow)
            worksheet.write(excelPosition, percentage)
            print(excelPosition)

            excelRow += 1
        excel.close()
    server = smtplib.SMTP('smtp.gmail.com', 587)  # smtp settings, change accordingly.
    server.ehlo()
    server.starttls()  # secure connection
    return render_template("inbox.html", len=len(subject_list), subject=subject_list,
                           address=email_address_list, body=body_list,
                           result_list=result_list, percentage_list=percentage_list)


@app.route('/inbox/<num>')
def showEmail(num):
    getresult = result_list[int(num)]
    getpercentage = percentage_list[int(num)]

    return render_template("inbox1.html", len=len(subject_list), subject=subject_list,
                           address=email_address_list, body=body_list, num=num,
                           result=getresult, percentage=getpercentage)


@app.route('/inbox/blacklist')
def blacklist():
    wb = load_workbook('blacklist.xlsx')
    sheet = wb["blacklist"]
    row_count = sheet.max_row
    list = []  # nested list of email address and status
    # nested list will look like this in the end [[email,status], [email1,status1], [email2,status2]]
    for i in range(1, row_count + 1):
        for k in range(1, 3):
            if k == 1:
                emailadd = sheet.cell(row=i, column=k).value
                list1 = []
                list1.append(emailadd)
            if k == 2:
                status = sheet.cell(row=i, column=k).value
                list1.append(status)
                list.append(list1)

    return render_template("blacklist.html", list=list)


@app.route('/inbox/blacklist/new', methods=['GET', 'POST'])
def blacklistnew():
    wb = load_workbook('blacklist.xlsx')
    sheet = wb["blacklist"]  # values will be saved to excel sheet"blacklist"
    col2 = 'blacklisted'  # value of 2nd column in excel
    if request.method == 'POST':
        email1 = request.form['email1']  # id='email1' from html form
        if email1.strip() != "":
            col1 = email1.strip()
            sheet.append([col1, col2])

        email2 = request.form['email2']
        if email2.strip() != "":
            col1 = email2.strip()
            sheet.append([col1, col2])

        email3 = request.form['email3']
        if email3.strip() != "":
            col1 = email3.strip()
            sheet.append([col1, col2])

        email4 = request.form['email4']
        if email4.strip() != "":
            col1 = email4.strip()
            sheet.append([col1, col2])
        wb.save('blacklist.xlsx')
        wb.close()
        return redirect('/inbox/blacklist')
    else:
        return render_template("blacklistnew.html")


@app.route('/inbox/whitelist')
def whitelist():
    wb = load_workbook('whitelist.xlsx')
    sheet = wb["whitelist"]
    row_count = sheet.max_row
    list = []  # nested list of email address and status
    # nested list will look like this in the end [[email,status], [email1,status1], [email2,status2]]
    for i in range(1, row_count + 1):
        for k in range(1, 3):
            if k == 1:
                emailadd = sheet.cell(row=i, column=k).value
                list1 = []
                list1.append(emailadd)
            if k == 2:
                status = sheet.cell(row=i, column=k).value
                list1.append(status)
                list.append(list1)

    return render_template("whitelist.html", list=list)


@app.route('/inbox/whitelist/new', methods=['GET', 'POST'])
def whitelistnew():
    wb = load_workbook('whitelist.xlsx')
    sheet = wb["whitelist"]  # values will be saved to excel sheet"whitelist"
    col2 = 'whitelisted'  # value of 2nd column in excel
    if request.method == 'POST':
        email1 = request.form['email1']  # id='email1' from html form
        if email1.strip() != "":
            col1 = email1.strip()
            sheet.append([col1, col2])

        email2 = request.form['email2']
        if email2.strip() != "":
            col1 = email2.strip()
            sheet.append([col1, col2])

        email3 = request.form['email3']
        if email3.strip() != "":
            col1 = email3.strip()
            sheet.append([col1, col2])

        email4 = request.form['email4']
        if email4.strip() != "":
            col1 = email4.strip()
            sheet.append([col1, col2])
        wb.save('whitelist.xlsx')
        wb.close()
        return redirect('/inbox/whitelist')
    else:
        return render_template("whitelistnew.html")


@app.route('/inbox/quarantine')
def showQuarantine():
    from openpyxl import load_workbook
    wb = load_workbook('logs.xlsx')
    ws = wb["Sheet1"]
    percentageList = []
    subjectList = []
    bodyList = []
    emailAddressList = []
    resultList = []

    for row in ws.rows:
        if row[4].value == "Suspicious":
            subjectList.append(row[0].value)
            emailAddressList.append(row[1].value)
            bodyList.append(row[2].value)
            resultList.append(row[4].value)
            percentageList.append(row[8].value)

    return render_template("quarantine.html", len=len(subjectList), subject=subjectList,
                           address=emailAddressList, body=bodyList,
                           result_list=resultList, percentage_list=percentageList)


# to run application
if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=False)
    # app.run()

import re
import easyimap as e
import pandas as pd
import numpy as np
from flask import request

from joblib import dump, load
from nltk import word_tokenize, PorterStemmer
from nltk.corpus import stopwords

#
# def cleaning(string):
#     string = re.sub("[^0-9a-zA-Z\ ]", "", str(string))
#     string = string.lower()
#     string = string.strip()
#
#     return string
#
#
# def stem(string):
#     tokenized = word_tokenize(string)
#     stemmed = []
#     stemmer = PorterStemmer()
#
#     for word in tokenized:
#         stemmed.append(stemmer.stem(word))
#
#     return ' '.join(stemmed)
#
#
# def remove_stopwords(string):
#     STOP_WORDS = set(stopwords.words('english'))
#
#     tokenized = word_tokenize(string)
#     filtered = []
#
#     for word in tokenized:
#         if word not in STOP_WORDS:
#             filtered.append(word)
#
#     return " ".join(filtered)
#
#
# global email
# global subject_list
# global body_list
# global email_address_list
# global percentage_list
#
# percentage_list = []
# subject_list = []
# body_list = []
# email_address_list = []
# date_list = []
# # Authenticates and retrieves email
#
# imap_url = 'imap.gmail.com'
# username = 'jerrettfg@gmail.com'
# password = 'Yongyong97'
# server = e.connect(imap_url, username, password)
# #inbox = server.listup()
# inbox = server.listids()
# email = server.mail(server.listids()[0])
#
# for x in range(5, 10): # len(inbox)
#         email = server.mail(server.listids()[x])
#
#         string = email.body
#         string = cleaning(string)
#         string = stem(string)
#         string = remove_stopwords(string)
#
#         # store email subject, body in list
#         email_address_list.append(email.from_addr)
#         subject_list.append(email.title)
#         body_list.append(email.body)
#
#         # ML
#         n_df = pd.DataFrame({'text': string}, index=[0])
#         n_df.head()
#         vectorizer = load(r'naivebayesVectorizer.joblib')  # load vectorizer
#         nbclf = load(r'naivebayes.joblib')  # load the naivebayes ml model
#
#         x_matrix = vectorizer.transform(n_df['text'])
#         my_prediction = nbclf.predict(x_matrix)
#         percentage = nbclf.predict_proba(x_matrix)
#         #percentage = np.array(percentage)
#         #percentage = ['{:f}'.format(item) for item in percentage]
#         np.set_printoptions(formatter={'float_kind':'{:f}'.format})
#         if my_prediction == 1:
#             result = 'Phishing'
#             percentage = format(percentage[0][1], '.12f') # to 12decimal place
#             percentage = float(percentage) * 100 # convert to percent
#             percentage = str(percentage) + '%'
#             percentage_list.append(percentage)
#         elif my_prediction == 0:
#             result = 'Non-Phishing'
#             percentage = format(percentage[0][0], '.12f') # to 12decimal place
#             percentage = float(percentage) * 100 # convert to percent
#             percentage = str(percentage) + '%'
#             percentage_list.append(percentage)
#
#         print('PERCENTAGE', percentage)
#         print(result)
#         for a in percentage_list:
#             print(a)
#
#         date_list.append(email.date)
#         for i in date_list:
#             print(i)
#
#         #percentage_l = list(percentage)    #convert numpy array to list
#         #print(percentage_l)
#
#         if my_prediction == 1:
#             result = 'Phishing'
#             #percentage = percentage_l[1]
#             #percentage_list.append(percentage)
#         elif my_prediction == 0:
#             result = 'Non-Phishing'
#             #percentage = percentage_l[0]
#            # percentage_list.append(percentage)

# import time
# 
# start = time.time()
# print("hello")
# end = time.time()
# print(end - start)
# from openpyxl import load_workbook
#
#
# wb = load_workbook('logs.xlsx')
# sheet = wb["blacklist"]  # values will be saved to excel sheet"blacklist"
# col2 = 'blacklisted'  # value of 2nd column in excel
#
# email1 = '  '  # id='email1' from html form
# if email1.strip() != "":
#     col1 = email1.strip()
#     sheet.append([col1, col2])
#
# email2 = 'jerrett@gmail.com'
# if email2.strip() != "":
#     col1 = email2.strip()
#     sheet.append([col1, col2])
#
# email3 = 'jerrett@gmail.com'
# if email3.strip() != "":
#     col1 = email3.strip()
#     sheet.append([col1, col2])
#
# email4 = 'jerrett@gmail.com'
# if email4.strip() != "":
#     col1 = email4.strip()
#     sheet.append([col1, col2])
#
# wb.save('logs.xlsx')
# wb.close()
#

# from openpyxl import load_workbook
# wb = load_workbook('logs.xlsx')
# ws = wb["Sheet1"]
# percentage_list = []
# subject_list = []
# body_list = []
# email_address_list = []
# result_list = []
#
# for row in ws.rows:
#     if row[4].value == "Suspicious":
#         subject_list.append(row[0].value)
#         email_address_list.append(row[1].value)
#         body_list.append(row[2].value)
#         result_list.append(row[3].value)
#
# for a in subject_list:
#     print(a)
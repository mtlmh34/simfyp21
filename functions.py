# Functions

# For connection
import openpyxl  # import openpyxl package
import sys
import os
import hashlib
import bs4
import re
import requests
import xlsxwriter
import pefile  # import pefile package
import textblob  # import textblob package
from is_safe_url import is_safe_url
from pysafebrowsing import SafeBrowsing
from nested_lookup import nested_lookup
from bs4 import BeautifulSoup


class mainFunctions:

    def email_valid(self):
        # Function to check if email address is legitimate
        # Import py3-validate-email package
        from validate_email import validate_email
        from openpyxl import load_workbook
        emailAd = self
        # To only check string that contain @
        if "@" in emailAd:
            # To get only the email address
            if "<" in emailAd:
                emailAddress = emailAd.split('<')
                emailAd = emailAddress[1]
                emailAd = emailAd[:-1]
                is_valid = validate_email(emailAd, smtp_timeout=10)
                # print(is_valid)
                if is_valid:
                    return True
                else:
                    return False
            else:
                is_valid = validate_email(emailAd, smtp_timeout=10)
                # print(is_valid)
                if is_valid:
                    return True
                else:
                    return False
        else:
            return False

    def attachment_check(self):
        extensionsToCheck = ['.zip', '.exe', '.scr', '.rar', '.7z', '.iso', '.r09']
        if any(ext in self for ext in extensionsToCheck):
            # print("Unsafe")
            return False
        else:
            # print("Safe")
            return True

    def content_formatting(self):
        def remove_tags(html):
            # parse html content
            soup = BeautifulSoup(html, "html.parser")

            for data in soup(['style', 'script']):
                # Remove tags
                data.decompose()

            # return data by retrieving the tag content
            return ' '.join(soup.stripped_strings)

        try:
            HTML_DOC = self
            if not HTML_DOC:
                print("Empty")
            else:
                display = remove_tags(HTML_DOC)
                mystring = display.replace("_x000D_", "\n")
                mstring = mystring.replace("\n", "\x0A")
                return mstring
        except:
            return self

    def check_link(self):

        def extract_link(body):
            regex = r"\b((?:https?://)?(?:(?:www\.)?(?:[\da-z\.-]+)\.(?:[a-z]{2,6})|(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)|(?:(?:[0-9a-fA-F]{1,4}:){7,7}[0-9a-fA-F]{1,4}|(?:[0-9a-fA-F]{1,4}:){1,7}:|(?:[0-9a-fA-F]{1,4}:){1,6}:[0-9a-fA-F]{1,4}|(?:[0-9a-fA-F]{1,4}:){1,5}(?::[0-9a-fA-F]{1,4}){1,2}|(?:[0-9a-fA-F]{1,4}:){1,4}(?::[0-9a-fA-F]{1,4}){1,3}|(?:[0-9a-fA-F]{1,4}:){1,3}(?::[0-9a-fA-F]{1,4}){1,4}|(?:[0-9a-fA-F]{1,4}:){1,2}(?::[0-9a-fA-F]{1,4}){1,5}|[0-9a-fA-F]{1,4}:(?:(?::[0-9a-fA-F]{1,4}){1,6})|:(?:(?::[0-9a-fA-F]{1,4}){1,7}|:)|fe80:(?::[0-9a-fA-F]{0,4}){0,4}%[0-9a-zA-Z]{1,}|::(?:ffff(?::0{1,4}){0,1}:){0,1}(?:(?:25[0-5]|(?:2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(?:25[0-5]|(?:2[0-4]|1{0,1}[0-9]){0,1}[0-9])|(?:[0-9a-fA-F]{1,4}:){1,4}:(?:(?:25[0-5]|(?:2[0-4]|1{0,1}[0-9]){0,1}[0-9])\.){3,3}(?:25[0-5]|(?:2[0-4]|1{0,1}[0-9]){0,1}[0-9])))(?::[0-9]{1,4}|[1-5][0-9]{4}|6[0-4][0-9]{3}|65[0-4][0-9]{2}|655[0-2][0-9]|6553[0-5])?(?:/[\w\.-]*)*/?)\b"

            if body.startswith('<'):
                links = re.findall(regex, body)
                return links

        def check_urls(url_list):
            KEY = "AIzaSyABO6DPGmHpCs8U5ii1Efkp1dUPJHQfGpo"
            s = SafeBrowsing(KEY)
            safe = 0
            malicious = 0

            for url in url_list:
                if is_safe_url(url, {"example.com", "www.example.com", "https://www.example.com"}):
                    # boo whether contains malicious link or not
                    r = s.lookup_urls([url])
                    if False in nested_lookup('malicious', r):
                        # print('not malicious')
                        safe += 1
                    else:
                        # print('malicious')
                        malicious += 1

            if malicious > 0:
                # print("{} links in this emails is/are malicious! ".format(malicious))
                return False
            else:
                # print("All links in this emails are safe. ")
                return True

        # TEST
        try:
            if not self:
                return True
            linkList = extract_link(self)
            print(linkList)
            if not linkList:
                return True
            result = check_urls(linkList)
            return result
        except:
            return False


'''
    def check_function(emailSub, emailAdr, emailCon, emailAtt):

        emailConFormat = content_formatting(emailCon)
        spelling_check(formatted_content)
        email_valid(emailAdr)
        attachment_check(emailAtt)

        result = 'Non-Phishing'
        return result
'''

# mainFunctions.malware_check()
# mainFunctions.spelling_check()
# mainFunctions.email_valid()

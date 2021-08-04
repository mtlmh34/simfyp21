import re

# find for html tag
def html_exists(text):
    if re.search("html", str(text), re.IGNORECASE):
        return 1
    else:
        return 0

# regex to get domains count
def count_domain(text):
    myregex = r'(?:[a-zA-Z0-9](?:[a-zA-Z0-9\-]{,61}[a-zA-Z0-9])?\.)+[a-zA-Z]{2,6}'
    domain_list = re.findall(myregex, str(text))
    return len(domain_list)


# dots count in an email, need to change to max dot count in urls
def count_dots(text):
    dots_list = re.findall(".", str(text))
    return len(dots_list)

# email contains account term
def account_exists(text):
    if re.search ("account", str(text), re.IGNORECASE):
        return 1
    else:
        return 0

# paypal
def paypal_exists(text):
    if re.search("paypal", str(text), re.IGNORECASE):
        return 1
    else:
        return 0

# login
def login_exists(text):
    if re.search("login", str(text), re.IGNORECASE):
        return 1
    else:
        return 0

# bank
def bank_exists(text):
    if(re.search("bank", str(text), re.IGNORECASE)):
        return 1
    else:
        return 0
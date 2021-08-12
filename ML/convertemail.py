import re
import pandas as pd
from IPython.display import display

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

#Create new df to be saved into csv file later
new_df = pd.DataFrame(columns=['Class','HtmlExists','DomainCount','DotCount',
                           'AccountExists','PaypalExists','LoginExists','BankExists'])

df = pd.read_csv("email.csv")
print(df.shape, "before clean")
display(df)
df = df.dropna()
print(df.shape, "after clean")

new_df["Class"] = df["Class"]
new_df['HtmlExists'] = df['Text'].apply(html_exists)
new_df['DomainCount'] = df['Text'].apply(count_domain)
new_df['DotCount'] = df['Text'].apply(count_dots)
new_df['AccountExists'] = df['Text'].apply(account_exists)
new_df['PaypalExists'] = df['Text'].apply(paypal_exists)
new_df['LoginExists'] = df['Text'].apply(login_exists)
new_df['BankExists'] = df['Text'].apply(bank_exists)
new_df.to_csv("email2.csv", index=False)


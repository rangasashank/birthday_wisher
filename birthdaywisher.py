import pandas as pd
import datetime
import smtplib
import os

# Enter your authentication details
GMAIL_ID = ''
GMAIL_PSWD = ''

def send_email(to, sub, msg):
    print(f"Email to {to} sent with subject: {sub} and message {msg}" )
    s = smtplib.SMTP('smtp.gmail.com', 587)
    s.starttls()
    s.login(GMAIL_ID, GMAIL_PSWD)
    s.sendmail(GMAIL_ID, to, f"Subject: {sub}\n\n{msg}")
    s.quit()


df = pd.read_excel("/Users/sashi/Desktop/Projects/birthday wisher/Birthdayfile.xlsx")
today = datetime.datetime.now().strftime("%d-%m")
year_now = datetime.datetime.now().strftime("%Y")


index_list = []

for index, item in df.iterrows():

    bday = item['Birthday'].strftime("%d-%m")
    year = item['Year']
    email = item['Email']
    dialogue = item['Dialogue']

    if (today==bday) and year_now not in str(year):
        send_email(email, "Happy Birthday", dialogue)
        index_list.append(index)

for item in index_list:
    yr = df.loc[item, 'Year']
    df.loc[item, 'Year'] = str(yr)+','+ str(year_now)

df.to_excel('/Users/sashi/Desktop/Projects/birthday wisher/Birthdayfile.xlsx', index=False)
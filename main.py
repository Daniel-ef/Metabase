import os

import smtplib
from email import encoders
from email.mime.multipart import MIMEMultipart, MIMEBase
from email.mime.text import MIMEText
from urllib.parse import quote
from urllib.request import urlretrieve

import pandas as pd

from settings import CHILDREN, MESSAGE_TEMPLATE, MESSAGE_SUBJECT,\
    CURATOR_EMAIL, CURATOR_PASS, START_DATE, END_DATE, SEND_TO_PARENTS


def download_file(start_date, end_date, user_id, file_name):

    if file_name in os.listdir("."):
        print("File exists")
        return False

    print("\nDownloading {} ...".format(file_name))
    request_string = 'https://metabase.foxford.ru/public/question/782ef688-5f95-41ba-8cab-c285650cca86.xlsx?parameters='
    params = """
        [{"type":"date/single","target":["variable",["template-tag","startDate"]],"value":"%s"},
        {"type":"date/single","target":["variable",["template-tag","endDate"]],"value":"%s"},
        {"type":"category","target":["variable",["template-tag","user_id"]],"value":"%s"},
        {"type":"category","target":["variable",["template-tag","school_year_id"]],"value":"6"}]
        """ % (start_date, end_date, user_id)

    url = request_string + quote(params)
    urlretrieve(url, file_name)
    print("File downloaded")
    return True


def hide_columns(file_name):
    print('Trying to hide columns')
    columns_to_delete = [0, 1, 3, 4]

    df = pd.read_excel(file_name)
    df.drop(df.columns[columns_to_delete], axis=1, inplace=True)
    writer = pd.ExcelWriter(file_name)
    df.to_excel(writer)
    writer.save()
    print("Columns hidden")


def send_email(filename, email):
    with open(filename, "rb") as fp:
        attachment = fp.read()
    body = MIMEText(MESSAGE_TEMPLATE)

    # Compose attachment
    part = MIMEBase('application', "octet-stream")
    part.set_payload(attachment)
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', 'attachment; filename="%s"' % filename)

    # Compose message
    msg = MIMEMultipart()
    msg["Subject"] = MESSAGE_SUBJECT
    msg["From"] = CURATOR_EMAIL
    msg["To"] = email

    msg.attach(part)
    msg.attach(body)

    # Establish connection
    server = smtplib.SMTP_SSL()
    server_addres = CURATOR_EMAIL.split("@")[1]
    server.connect('smtp.%s' % server_addres)
    server.login(CURATOR_EMAIL, CURATOR_PASS)

    # Send mail
    server.sendmail(msg["From"], msg["To"], msg.as_string())
    server.quit()
    print("Mail to {} sent".format(email))


if __name__ == "__main__":
    for child in CHILDREN:
        file_name = "{}-{}-{}.xlsx".format(
            START_DATE.replace('-', ''),
            END_DATE.replace('-', ''),
            child['name']
        )

        if download_file(START_DATE, END_DATE, child['user_id'], file_name):
            hide_columns(file_name)

        if SEND_TO_PARENTS:
            send_email(file_name, child['parent_email'])


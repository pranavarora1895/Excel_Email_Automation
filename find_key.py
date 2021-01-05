from tkinter import Tk
from tkinter.filedialog import askopenfilename
from openpyxl import load_workbook
import os
import smtplib

EMAIL_ADDRESS = os.environ.get('EMAIL_ADDRESS')
EMAIL_PASSWORD = os.environ.get('EMAIL_PASS')


def mail_sending_to_approved(receiver_email_id):
    """Function to send emails"""

    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        smtp.ehlo()
        smtp.starttls()
        smtp.ehlo()
        smtp.login(EMAIL_ADDRESS, EMAIL_PASSWORD)

        subject = 'Congrats!! You got approved!'
        body = open('mailbody.txt', 'r').read()

        msg = f'Subject: {subject}\n\n{body}'

        smtp.sendmail(EMAIL_ADDRESS, receiver_email_id, msg)


def get_excel_path():
    """Getting the file (GUI)"""
    Tk().withdraw()
    return askopenfilename(filetypes=[('Excel Files', '*.csv;*.xlsx;*.xlsm;*.xltx;*.xltm')],
                           title='Select your Excel File and Click Open')


path = get_excel_path()
wb = load_workbook(path)
ws = wb.active
unique_id = []

# crawling through all rows, and if the keyword matches then it will send the email. One email per email-id
# irrespective of number of occurrences
for row in ws.iter_rows(min_row=2, values_only=True):
    if 'Approved' in row:
        if row[0] in unique_id:
            continue
        else:
            mail_sending_to_approved(row[0])
            unique_id.append(row[0])
print(f'Email sent to {unique_id}')
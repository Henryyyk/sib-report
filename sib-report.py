from __future__ import print_function

import os.path
import pandas as pd
import smtplib
import email.message

from IPython.display import display
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1wOJWAb6CJjrh-kt9VTB9POKChyw9LnUW96MuTTUTRrs'
SAMPLE_RANGE_NAME = 'company1!B6:J1000'

def data_collect():  # OBTEM DADOS DO SHEETS
    global report
    report = []
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)

        # Chamada de API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                    range=SAMPLE_RANGE_NAME).execute()
        # Obtem os valores em RANGE_NAME
        report.append(result['values'])
        print('Dados coletados.')
        return report
        # Print dos valores armazenados em result
    except HttpError as err:
        print(err)


def enviar_email():  # ENVIO DE EMAIL
    corpo_email = mailAuto

    msg = email.message.Message()
    msg['Subject'] = f"{std} Support Report"
    msg['From'] = 'adhenrique.santos@gmail.com'
    msg['To'] = mail
    password = 'vbtpdgnnnegfgmql'
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email)

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print(f'Email para {std} enviado!')


if __name__ == '__main__':
    data_collect()  # Coleta os dados da planilha

tabela = []  # ORGANIZANDO TABELA
for x in report:
    for y in x:
        tabela.append(y)
tb = pd.DataFrame.from_records(tabela, columns=[[
    'Ticket ID',
    'Type',
    'Status',
    'Requester',
    'Description',
    'Open On',
    'Days Open',
    'Due Date',
    'Work Log'
]])

# Lista de estudios SIB
studios = [
    {
        'studio': 'Company 1',
        'reponsaveis': 'João Teixeira',
        'e-mail': 'ad_henrique@live.com'
    },
    {
        'studio': 'Company 2',
        'reponsaveis': 'Henrique Santos',
        'e-mail': 'adhenrique.santos@gmail.com'
    }
]

# Envia e-mails para os SIB
for e in studios:
    std = e.get('studio')
    resp = e.get('reponsaveis')
    mail = e.get('e-mail')
    mailAuto = f"""
        <p>Prezado {resp},</p>
        <p>Segue o report de suporte.</p>
        <p>{tb.to_html(index=False)}</p>
        <p>Henrique</p>
        """
    enviar_email()

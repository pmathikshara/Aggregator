from __future__ import print_function
import google.auth

import os.path
import re
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
try:
    from BeautifulSoup import BeautifulSoup
except ImportError:
    from bs4 import BeautifulSoup

SCOPES = ['https://www.googleapis.com/auth/presentations.readonly',
          'https://www.googleapis.com/auth/spreadsheets']

EMAIL_DUMP_DB = '1a2m9rZXviGS2c3i3_zSWvf1tDIKgUOsCFpZ8-Z24GKI'
RANGE_NAME = 'RawData!A1:F'


def main():
    """
    Driver code - look at block diagram for flow information. 
    If a new entry is added runs dumpToPresentation.py 
    In a while True, polls every 10 minutes. 
    """
    creds = None
    entry_count = 0
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('sheets', 'v4', credentials=creds)
        # TODO - Add a trigger signal and remove while True.
        while(True):
            sheet = service.spreadsheets()
            result = sheet.values().get(spreadsheetId=EMAIL_DUMP_DB,
                                        range=RANGE_NAME).execute()
            values = result.get('values', [])
            if(entry_count != len(values)):
                entry_count = len(values)
                exec(open("dumpToPresentation.py").read())
                time.sleep(600)

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    main()

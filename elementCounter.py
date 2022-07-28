from __future__ import print_function
import google.auth

import os.path
import re
import ParsePresentation
import time

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
try:
    from BeautifulSoup import BeautifulSoup
except ImportError:
    from bs4 import BeautifulSoup

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/presentations.readonly',
          'https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
ELEMENT_COUNT_DB = '1QAuFviz167FutIgv7UkkYAgylB9YinGWDwIjPZWPf4U-Z24GKI'
RANGE_NAME = 'RawData!A1:B'
values = list()


def elementCount():
    """Shows basic usage of the Slides API.
    Prints the number of slides and elments in a sample presentation.
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

        # Call the Sheets API
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=ELEMENT_COUNT_DB,
                                    range=RANGE_NAME).execute()
        if(values != result.get('values', [])):

        if not values:
            print('No data found.')
            return

        for row in values:
            parsedHTML = BeautifulSoup(row[4], 'html.parser')
            for link in parsedHTML.find_all('a'):
                if(link.get('href').find("docs.google.com") >= 0):
                    start_index = link.get('href').find("/d/") + 3
                    end_index = (link.get('href')[start_index:]).find("/")
                    ID = link.get('href')[start_index:][:end_index]
                    values = [
                        [
                            row[5], ID
                        ]
                    ]
                    body = {
                        'values': values
                    }
                    result = service.spreadsheets().values().append(
                        spreadsheetId=Slides_DB, range="A1:B1",
                        valueInputOption="USER_ENTERED", body=body).execute()
                    print(
                        f"{(result.get('updates').get('updatedCells'))} cells appended.")
                    ParsePresentation.parsePresentation(ID)

    except HttpError as err:
        print(err)


if __name__ == '__main__':
    while(True):
        elementCount()
        time.sleep(600)

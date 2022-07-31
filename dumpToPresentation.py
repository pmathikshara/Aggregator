from __future__ import print_function
import google.auth
import os
import os.path
import re
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


def dumpToPresentation():
    """
    Parses information from the email dump and identifies the google doc links. 
    Checks if the document already exisits if not adds it to presDB.
    Also calls parsePresentation.py to initiate the parse process and add elements to the elementCount DB. 
    """
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/presentations.readonly',
              'https://www.googleapis.com/auth/spreadsheets']

    EMAIL_DUMP_DB = '1a2m9rZXviGS2c3i3_zSWvf1tDIKgUOsCFpZ8-Z24GKI'
    Slides_DB = '1hND2jYigcJibPWKjcBkb4ktzFs6If9msnqYYik9ohrs'
    RANGE_NAME = 'RawData!A1:F'

    creds = None
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
        sheet = service.spreadsheets()
        result = sheet.values().get(spreadsheetId=EMAIL_DUMP_DB,
                                    range=RANGE_NAME).execute()
        values = result.get('values', [])

        if not values:
            print('No data found.')
            return

        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            "/Users/mathikshara/Documents/Work/Scraping/the-sandbox-357506-457186fbf5da.json", scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(
            Slides_DB).sheet1

        try:
            for row in values:
                parsedHTML = BeautifulSoup(row[4], 'html.parser')
                for link in parsedHTML.find_all('a'):
                    if(link.get('href').find("docs.google.com") >= 0):
                        start_index = link.get('href').find("/d/") + 3
                        end_index = (link.get('href')[start_index:]).find("/")
                        ID = link.get('href')[start_index:][:end_index]

                        if(sheet.find(ID) == None):
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
                            os.system("python3 ParsePresentation.py " + ID)
        except:
            print("No HTML or No Link")
    except HttpError as err:
        print(err)


if __name__ == '__main__':
    dumpToPresentation()

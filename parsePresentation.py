from __future__ import print_function
import google.auth

import os.path
import re
import time
import sys

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

ELEMENT_COUNT_DB = "1QAuFviz167FutIgv7UkkYAgylB9YinGWDwIjPZWPf4U"
RANGE_NAME = "ElementCount!A1:C"
LISTING_DB = "1BH9teyYj8oGYR-0ZzBktPo86a14a5z2G13XFnZdU9Mw"


def extractEmails(txt, _email):
    _result = re.search(r'[\w.+-]+@[\w-]+\.[\w.-]+', txt)
    if(_result != None):
        _result = _result.group(0)
    else:
        _result = _email
    return _result


def extractPrice(txt):
    _result = re.search("[\$\£\€](\d+(?:\.\d{1,3})?)", txt)
    if(_result != None):
        _result = _result.group(0)
    return _result


def extractURLs(txt):
    _result = re.search("(?P<url>https?://[^\s]+)", txt)
    if(_result != None):
        _result = _result.group("url")
    return _result


def parsePresentation(PRESENTATION_ID):
    """ 
    Parses out the presentation using Slides API. Looks at each sentence (parse out the textRun content) and does the following,
    - Append content count per slide and over all slide count to element count DB. To be used in identifying if an item is sold. 
    - Check if the content has an email, price or URL. If so uses accordingly. 
    - If none - considers it to be a general sentence. 
    - If a slide does not have link or price - it is considered to be a logstics slide. 

    Assumptions:
    - One listing per page 
    - Price is listing price - which is not the case in 90% of the listing
    - One person per presentation

    # TODO
    1. Parse out original price and new price 
    2. Parse out condition 
    3. Link code not working for google slides - look into this. 
    4. Look for phone numbers in the email 
    """
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
        # See if slide is in element count DB - if not add it to it
        sheets_service = build('sheets', 'v4', credentials=creds)
        sheet = sheets_service.spreadsheets()
        result = sheet.values().get(spreadsheetId=ELEMENT_COUNT_DB,
                                    range=RANGE_NAME).execute()
        values = result.get('values', [])
        slides_service = build('slides', 'v1', credentials=creds)
        presentation = slides_service.presentations().get(
            presentationId=PRESENTATION_ID).execute()
        slides = presentation.get('slides')

        slide_count = ""
        for i, slide in enumerate(slides):
            slide_count = slide_count + " ," + \
                str(len(slide.get('pageElements')))

        if(PRESENTATION_ID not in values[0][:][0]):
            values = [
                [
                    PRESENTATION_ID, len(slides), slide_count
                ]
            ]
            body = {
                'values': values
            }
            result = sheets_service.spreadsheets().values().append(
                spreadsheetId=ELEMENT_COUNT_DB, range="A1:C1",
                valueInputOption="USER_ENTERED", body=body).execute()
            print(
                f"{(result.get('updates').get('updatedCells'))} cells appended.")

        # Parse out info from slide and add to listing db.
        isListing = False
        slideContent = ""
        price = ""
        URL = ""
        email = ""
        for j, object in enumerate(slides):
            content = ""
            slideContent = ""
            for ele in range(len(object['pageElements'])):
                res = list(fun(object['pageElements'][ele], 'shape'))
                if(len(res) > 0):
                    res_1 = list(fun(res[0], 'text'))
                    if(len(res_1) > 0):
                        for elements in res_1[0]['textElements']:
                            if('textRun' in elements.keys()):
                                content = elements['textRun']['content']
                                content = content.strip()

                                # Get email
                                email = extractEmails(content, email)
                                if(email != None):
                                    print(" ")

                                # Get price
                                price = extractPrice(content)
                                if(price != None):
                                    print(price)
                                    isListing = True

                                # Get URLs
                                URL = extractURLs(content)
                                if(URL != None):
                                    # print(URL)
                                    isListing = True

                                # Remove URLs from content
                                content = re.sub(
                                    r'''(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'".,<>?«»“”‘’]))''', " ", content)
                                slideContent = slideContent + " " + content
            print(slideContent)

            # See if the content is a listing or if it is a logsitics page.
            if(isListing):
                # Sleep to keep withing google sheets API limit
                time.sleep(1)
                values = [
                    [slideContent, price, URL, email]
                ]
                body = {
                    'values': values
                }
                result = sheets_service.spreadsheets().values().append(
                    spreadsheetId=LISTING_DB, range="A:D",
                    valueInputOption="USER_ENTERED", body=body).execute()
                print(
                    f"{(result.get('updates').get('updatedCells'))} cells appended.")

    except HttpError as err:
        print(err)


def fun(d, srch_key):
    if srch_key in d:
        yield d[srch_key]
    for k in d:
        if isinstance(d[k], list):
            for i in d[k]:
                for j in fun(i):
                    yield j


if __name__ == '__main__':
    parsePresentation(sys.argv[1])

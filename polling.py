from __future__ import print_function
import google.auth

import os.path
import re
import time
import sys
import airtable

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.cloud import vision
import gspread
from oauth2client.service_account import ServiceAccountCredentials
try:
    from BeautifulSoup import BeautifulSoup
except ImportError:
    from bs4 import BeautifulSoup

SCOPES = ['https://www.googleapis.com/auth/presentations.readonly',
          'https://www.googleapis.com/auth/spreadsheets']

ELEMENT_COUNT_DB = "1QAuFviz167FutIgv7UkkYAgylB9YinGWDwIjPZWPf4U"
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
        _result = _result.groups()
    return _result


def extractURLs(txt):
    _result = re.search("(?P<url>https?://[^\s]+)", txt)
    if(_result != None):
        _result = _result.group("url")
    return _result


def listingToAirtable(data):
    at = airtable.Airtable('appP28g4PMQOHJjtY', 'keyidKxie1VMA9kDn')
    at.get('Sales')
    print(at.create('Sales', data))


def detect_text_uri(uri):
    """Detects text in the file located in Google Cloud Storage or on the Web.
    """
    from google.cloud import vision
    client = vision.ImageAnnotatorClient()
    image = vision.Image()
    image.source.image_uri = uri

    response = client.text_detection(image=image)
    texts = list()
    texts = response.text_annotations

    if response.error.message:
        raise Exception(
            '{}\nFor more info on error messages, check: '
            'https://cloud.google.com/apis/design/errors'.format(
                response.error.message))
    return texts


def polling(PRESENTATION_ID):
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
    RANGE_NAME = "ElementCount!A1:I"

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
        slides_service = build('slides', 'v1', credentials=creds)
        presentation = slides_service.presentations().get(
            presentationId=PRESENTATION_ID).execute()
        slides = presentation.get('slides')

        # Parse out info from slide and add to listing db.
        isListing = False
        email = ""
        location = ""
        for j, object in enumerate(slides):
            slideContent = ""
            isTitle = True
            item = ""
            condition = ""
            price = ""
            price_list = list()
            URL = ""
            imageURL = ""

            for ele in range(len(object['pageElements'])):
                content = ""
                if('image' in object['pageElements'][ele].keys()):
                    imageURL = object['pageElements'][ele]["image"]["contentUrl"]
                    try:
                        texts = detect_text_uri(imageURL)
                        for text in texts:
                            if(any(substring in (text.description).lower() for substring in ["sold", "taken"])):
                                continue
                    except:
                        print("Image screwed up")

                res = list(fun(object['pageElements'][ele], 'shape'))
                if(len(res) > 0):
                    res_1 = list(fun(res[0], 'text'))
                    if(len(res_1) > 0):
                        for elements in res_1[0]['textElements']:
                            if('textRun' in elements.keys()):
                                content = elements['textRun']['content']
                                content = content.strip()

                                if(content.lower() in ["sold"]):
                                    continue

                                if(any(substring in content.lower() for substring in ['pickup', 'pick up'])):
                                    location = content

                                if(isTitle and len(content) > 1):
                                    item = content
                                    isTitle = False

                                # Get email
                                email = extractEmails(content, email)
                                if(email != None):
                                    print(" ")

                                # Get price
                                price_return = (extractPrice(content))

                                if(price_return != None):
                                    price_list.append(price_return[0])
                                    isListing = True

                                if(content.lower() in ['amazon', 'link', 'links', 'product', 'amazon link', 'photo']):
                                    continue

                                slideContent = slideContent + " " + content
                                if('link' in elements['textRun']['style'].keys()):
                                    URL = elements['textRun']['style']['link']['url']
                                    if("mailto" in URL):
                                        continue
                                    else:
                                        isListing = True

                                if('used' in content.lower() or 'condition' in content.lower()):
                                    condition = content
                                    continue

            # See if the content is a listing or if it is a logsitics page.
            if(isListing and len(slideContent) > 1):
                # Sleep to keep withing google sheets API limit
                try:
                    price = "$" + str(min(price_list))
                except:
                    price = ""
                time.sleep(1)
                values = [
                    [item, slideContent, price, imageURL,
                        condition, URL, location, email, "https://drive.google.com/file/d/"+PRESENTATION_ID]
                ]
                body = {
                    'values': values
                }
                result = sheets_service.spreadsheets().values().append(
                    spreadsheetId=LISTING_DB, range="A:I",
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
    # If modifying these scopes, delete the file token.json.
    SCOPES = ['https://www.googleapis.com/auth/presentations.readonly',
              'https://www.googleapis.com/auth/spreadsheets']

    Slides_DB = '1hND2jYigcJibPWKjcBkb4ktzFs6If9msnqYYik9ohrs'
    LISTING_DB = "1BH9teyYj8oGYR-0ZzBktPo86a14a5z2G13XFnZdU9Mw"
    RANGE_NAME = 'Sheet1!A1:B'

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

        # Delete Listing DB
        request = service.spreadsheets().values().clear(spreadsheetId=LISTING_DB,
                                                        range="A:I", body={})
        response = request.execute()

        # Add header
        values = [
            ["Item", "Description", "Price", "Photo", "Condition",
                "Link", "Location", "Email", "Source"]
        ]
        body = {
            'values': values
        }
        result = service.spreadsheets().values().append(
            spreadsheetId=LISTING_DB, range="A1:I1",
            valueInputOption="USER_ENTERED", body=body).execute()

        if not values:
            print('No data found.')

        # Get pres ids
        result = sheet.values().get(spreadsheetId=Slides_DB,
                                    range="A1:B").execute()
        values = result.get('values', [])
        for row in values:
            polling(row[1])

        at = airtable.Airtable('appP28g4PMQOHJjtY', 'keyidKxie1VMA9kDn')
        scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
                 "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]
        creds = ServiceAccountCredentials.from_json_keyfile_name(
            "/Users/mathikshara/Documents/Work/Scraping/the-sandbox-357506-457186fbf5da.json", scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_key(
            "1BH9teyYj8oGYR-0ZzBktPo86a14a5z2G13XFnZdU9Mw").sheet1

        # Iterate records on airtable
        checked_rows = []
        for r in at.iterate("Sales", fields=["Item", "Description", "Price", "Photo", "Condition", "Link", "Location", "Email", "Source"]):
            print(".")
            ID = r['id']
            desc = r['fields']['Description']
            item = r['fields']['Item']
            try:
                checked_rows.append(sheet.find(item).row)
                time.sleep(1)
            except:
                print("Item missing or changed")
                print(at.delete("Sales", ID))
        print(checked_rows)

        for i in range(1, len(sheet.col_values(1))):
            print("*")
            if(i in checked_rows):
                continue
            else:
                vals = sheet.row_values(i + 1)
                time.sleep(1)
                data = {
                    "Price": vals[2],
                    "Condition": vals[4],
                    "Link": vals[5],
                    "Location": vals[6],
                    "Email": vals[7],
                    "Description": vals[1],
                    "Item": vals[0],
                    "Photo":
                        [{
                            "url": vals[3]
                        }],
                    "Source": vals[8]
                }
                listingToAirtable(data)
        values_list = sheet.row_values(1)

    except HttpError as err:
        print(err)

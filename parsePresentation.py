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

                                if(any(substring in (content).lower() for substring in ["sold", "taken"])):
                                    continue

                                if(any(substring in content.lower() for substring in ['pickup', 'pick up'])):
                                    location = content

                                if(isTitle and len(content) > 1 and not any(substring in (content).lower() for substring in ["available"])):
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
                                    # price.append(price_return[0])
                                    isListing = True

                                # Remove URLs from content
                                # content = re.sub(
                                #     r'''(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'".,<>?«»“”‘’]))''', " ", content)

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
                    peice = ""
                time.sleep(5)
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
                print(result)

                print(
                    f"{(result.get('updates').get('updatedCells'))} cells appended.")
                data = {
                    "Price": price,
                    "Condition": condition,
                    "Link": URL,
                    "Location": location,
                    "Email": email,
                    "Description": slideContent,
                    "Item": item,
                    "Photo":
                        [{
                            "url": imageURL
                        }],
                    "Source": "https://drive.google.com/file/d/"+PRESENTATION_ID
                }
                listingToAirtable(data)

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
    PRESENTATION_ID = sys.argv[1]
    # PRESENTATION_ID = "1qOuhUYXYdQlwFJnOGP06dRs7inZ8IRItdWYfXmZ6SZc"
    parsePresentation(PRESENTATION_ID)

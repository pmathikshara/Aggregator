from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from docx import Document
import os
from pathlib import Path
import glob2
import re
import csv
f = open('log.csv', 'w')
writer = csv.writer(f, delimiter=',')
price = ""
row = ""


def extractEmails(txt, _email):
    _result = re.search(r'[\w.+-]+@[\w-]+\.[\w.-]+', txt)
    if(_result != None):
        _result = _result.group(0)
    else:
        _result = _email
    return _result


def extractPrice(txt):
    _result = re.search("[\$\£\€](\d+(?:\.\d{1,3})?)", txt)
    return _result


def extractURLs(txt):
    _result = re.search("(?P<url>https?://[^\s]+)", txt)
    if(_result != None):
        _result = _result.group("url")
    return _result


for eachfile in glob2.glob("/Users/mathikshara/Documents/Work/Scraping/presentations/*.pptx"):
    prs = Presentation(eachfile)
    print(eachfile)
    print("----------------------")
    email = ""
    for slide in prs.slides:
        isListing = False
        slideContent = ""
        price = ""
        URL = ""
        print("********************")
        for shape in slide.shapes:
            content = ""

            if hasattr(shape, "text"):
                # Print all content on the slide
                content = shape.text

                # Get email
                email = extractEmails(content, email)
                if(email != None):
                    print(" ")

                # Get price
                price = extractPrice(content)
                if(price != None):
                    # for i in price:
                    print(price)
                    isListing = True
                # if("original" in content.lower()):
                #     if(content.count("$") >= 2):
                #         price = extractPrice(content)
                #         if(price != None):
                #             # for i in price:
                #             print(price)
                #             isListing = True

                # Get URLs
                URL = extractURLs(content)
                if(URL != None):
                    # print(URL)
                    isListing = True
                else:
                    if hasattr(shape, "text_frame"):
                        for paragraph in shape.text_frame.paragraphs:
                            for run in paragraph.runs:
                                URL = run.hyperlink.address
                                if URL is not None:
                                    content = content.replace(run.text, '')
                                    isListing = True
                                    # print(URL)

                # Remove URLs from content
                content = re.sub(
                    r'''(?i)\b((?:https?://|www\d{0,3}[.]|[a-z0-9.\-]+[.][a-z]{2,4}/)(?:[^\s()<>]+|\(([^\s()<>]+|(\([^\s()<>]+\)))*\))+(?:\(([^\s()<>]+|(\([^\s()<>]+\)))*\)|[^\s`!()\[\]{};:'".,<>?«»“”‘’]))''', " ", content)
                slideContent = slideContent + " " + content

        # Write to CSV
        if(isListing):
            row = [slideContent, price, URL, email]
            writer.writerow(row)

        # TODO
        # Multiple prices - original, used, old, new... Should factor in these things :/
        # Does not identify free stuff
        # Assumes one item per slide
        # If slide has
        # Should get email of the dude and add it if it's not in the presentation


f.close()

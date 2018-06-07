# -*- coding: utf-8 -*-

# Import dependencies
import cchardet as chardet
import cookielib
import httplib2
import httplib
import urllib2
import socket
import time
import re
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles.fills import PatternFill
from bs4 import BeautifulSoup, UnicodeDammit

# Configure opener to accept cookies and handle HTTP/HTTPS
cookies = cookielib.LWPCookieJar()
handlers = [
    urllib2.HTTPHandler(),
    urllib2.HTTPSHandler(),
    urllib2.HTTPCookieProcessor(cookies)
    ]
opener = urllib2.build_opener(*handlers)
opener.addheaders = [('User-Agent', 'Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/44.0.2403.107 Safari/537.36'),('Accept', '*/*')]
urllib2.install_opener(opener)

# # Compile RegEx for string matching
# re.compile

# Define helper functions
def meta_redirect(content):
    soup  = BeautifulSoup(content, 'html.parser' )
    result=soup.find("meta",attrs={"http-equiv":"refresh"})
    if result:
        wait,text=result["content"].split(";")
        if text.strip().lower().startswith("url="):
            url=text[4:]
            return url
    return None

def get_content(url):

    h=httplib2.Http(".cache")
    # h.add_certificate(key,cert,"")

    resp, content = h.request(url,"GET")

    # follow the chain of redirects
    while meta_redirect(content):
        resp, content = h.request(meta_redirect(content),"GET")

    return content

def clean_wb(wb,ws,ws2):
#     wb.remove(ws2)
#     wb.create_sheet("Output")

    for row in ws['B2:D{0}'.format(ws.max_row)]:
      for cell in row:
        cell.value = None

    for row in ws2['A1:{0}{1}'.format(get_column_letter(ws2.max_column),ws2.max_row)]:
      for cell in row:
        cell.value = None

    wb.save("Ads_Crawler.xlsx")

# EXCEL
wb = load_workbook('Ads_Crawler.xlsx')
ws = wb['Input']
ws2 = wb['Output']
search_list = []
url_list = []
src = ""
html = ""

clean_wb(wb,ws,ws2)

# Construct search_list
for col, cell in enumerate(ws['F'], start=1):
   if cell.value and cell.value != u'Search for:':
       cell_str = (cell.value).encode('utf-8').strip()
       search_list.append(cell_str)

       # Populate 'Output' tab headers
       ws2.cell(row=1,column=col).value = cell_str

# /EXCEL

# EXCEL
# Construct url_list
for row in ws['A{}:A{}'.format(ws.min_row + 1, ws.max_row)]:
    for cell in row:
        if cell.value:
            url_list.append("http://{0}/ads.txt".format( (cell.value).encode('utf-8').strip()) )

# Add one to list length to accound for start=2 below
max_row = len(url_list) + 1
# /EXCEL

# Loop through url_list with index starting at 2 to match row index in spreadsheet
for row, url in enumerate(url_list, start=2):
    print "____________________________"
    print url
    print ""

    ws2['{0}{1}'.format("A",row)] = url

    # Attempt to open URL
    try:
        html = urllib2.urlopen(url, timeout = 60)#"http://cronista.com/ads.txt"/"http://mundodrama.com/ads.txt"
        html = html.read()

        print chardet.detect(html)

        # Check for refresh chain and handle accordingly w/ urllib2 or httplib2
        if "refresh" in html.lower():
            src = get_content(url)
        else:
            soup = BeautifulSoup( html, 'html.parser', from_encoding="windows-1255")
            src = soup.get_text()

    # Except for errors and print to terminal
    except httplib.BadStatusLine as e:
        print e.reason #http error
        src = ""
        html = ""
    except urllib2.URLError as e:
        print e.reason    #not catch
        src = ""
        html = ""
    except socket.timeout as e:
        print e    #catched
        src = ""
        html = ""

    # Log failed URLs in Notes section
    if chardet.detect(html)['encoding'] == None and chardet.detect(html)['confidence'] == None:
        print "Confidence and Encoding for {0} are None".format(url)
        ws['{0}{1}'.format("D",row)] = "Verify URL manually."

    # If successful, continue with searches
    else:
        # Loop through search_set from wb
        for col, search_str in enumerate(search_list, start=2):
        # for search_str in search_set:
            # Check for desired text
            text_found = re.search(r'(\<|\,| ){0}(\>|\,| )'.format(search_str), src)

            # Confirm if found and choose column accordingly
            if text_found:
                # Input tab
                column = "B"

                # Output tab
                response = "Pass"
                colorFill = PatternFill(start_color='00FF00',
                   end_color='00FF00',
                   fill_type='solid')

                # Log response to terminal
                print "{0} found!".format(search_str)
            else:
                # Input tab
                column = "C"

                # Output tab
                response = "Fail"
                colorFill = PatternFill(start_color='FF0000',
                   end_color='FF0000',
                   fill_type='solid')

                # Log response to terminal
                print "{0} not found!".format(search_str)

            # Populate 'Output' tab body
            ws2.cell(row=row,column=col).value = response
            ws2.cell(row=row,column=col).fill= colorFill

            # Populate spreadsheet using column above
            if ws['{0}{1}'.format(column,row)].value:
                current_val = ws['{0}{1}'.format(column,row)].value
                ws['{0}{1}'.format(column,row)] = "{0}, {1}".format(current_val, search_str)
            else:
                ws['{0}{1}'.format(column,row)] = search_str

    print "____________________________"

    # Write confirmation to file
    wb.save("Ads_Crawler.xlsx")

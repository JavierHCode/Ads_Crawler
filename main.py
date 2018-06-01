# AUTOBROWSER (AB)
import splinter
import selenium
import time
import re
# /AB

# EXCEL
from openpyxl import load_workbook
# /EXCEL

# AB
    # Set driver variables
webdriver = selenium.webdriver
#ActionChains = webdriver.common.action_chains.ActionChains # import action_chains to access invisible html elements
Keys = webdriver.common.keys
Browser = splinter.Browser

    # Set chrome options
prefs = {
    'credentials_enable_service': False,
    'profile': {
        'password_manager_enabled': False,
        'default_content_setting_values' : {
            'automatic_downloads': True,
        },

    },
}
chrome_options = webdriver.ChromeOptions()
chrome_options.add_experimental_option("prefs",prefs)
chrome_options.add_argument('--window-size=1200,1000')
# /AB

# EXCEL
wb = load_workbook('Ads_Crawler.xlsx')
ws = wb['Sheet1']
search_str = ws["f1"].value
url_list = []
# /EXCEL

# EXCEL
for row in ws['A{}:A{}'.format(ws.min_row + 1, ws.max_row)]:
    for cell in row:
        url_list.append("http://{0}/ads.txt".format(cell.value))

#print url_list
# /EXCEL
for row, url in enumerate(url_list, start=2):
    url_match = ""

    # Open chrome
    browser = Browser('chrome', options=chrome_options)
    browser.visit(url)

    # Create driver variable for selenium
    selenium_driver = browser.driver

    # Check for desired "text" (UPDATE THIS VARIABLE NAME)
    src = selenium_driver.page_source
    text_found = re.search(r'{0}'.format(search_str), src)

    if text_found:
        url_match = "Yes"

    else:
        url_match = "No"

    # Write to file
    ws['C{0}'.format(row)] = url_match
    wb.save("Ads_Crawler.xlsx")

    # Close browser
    browser.quit()

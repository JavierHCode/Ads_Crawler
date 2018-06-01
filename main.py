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
from selenium.common.exceptions import TimeoutException
    # Set driver variables
webdriver = selenium.webdriver
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

# Add one to list length to accound for start=2 below
max_row = len(url_list) + 1

# /EXCEL
for row, url in enumerate(url_list, start=2):
    url_match = ""

    print url

    if row == 2:
        # Set browser variable
        browser = Browser('chrome', options=chrome_options)

        # Create driver variable for selenium
        selenium_driver = browser.driver
        selenium_driver.set_page_load_timeout(6)

        # Open chrome
        try:
            browser.visit(url)

            # Test for "pre" element
            while browser.is_element_not_present_by_tag("pre",wait_time=0.05):
                continue

        except TimeoutException:
            pass

    else:
        # Close old tab
        browser.windows[0].close()

        # Switch current window to new tab
        browser.windows.current = browser.windows[0]

        # Visit URL
        try:
            browser.visit(url)

            # Test for "pre" element
            while browser.is_element_not_present_by_tag("pre",wait_time=0.05):
                continue

        except TimeoutException:
            pass

    # Check for desired text
    src = selenium_driver.page_source
    text_found = re.search(r'{0}'.format(search_str), src)

    # Confirm if found
    if text_found:
        url_match = "Yes"
    else:
        url_match = "No"

    print url_match

    # Write confirmation to file
    ws['C{0}'.format(row)] = url_match
    wb.save("Ads_Crawler.xlsx")

    # Open new tab
    selenium_driver.execute_script("window.open('');")

    if row == max_row:
        # Close browser
        browser.quit()

#! python3
# getNewDesignCost.py
# Enter an ECO# and this will open a Selenium browser, user enters login info and then
# proceeds to the ECO and generates the affected item cost difference
# Will need to download chrome driver

import sys
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException

# Launch the Selenium browser and navigate to Omnify
chrome = webdriver.Chrome()
chrome.maximize_window()
eco_url = 'https://plm.mdsaero.com/Omnify7/Apps/Desktop/Change/Reports/BOMFieldCheck.aspx?id=3040'
chrome.get(eco_url)

# Set the login type to LDAP
loginSelect = Select(chrome.find_element_by_id("dlLogin_AuthenticationMethod"))
loginSelect.select_by_value("1")

# Wait for the page to load the proper fields for LDAP
WebDriverWait(chrome, 5).until(EC.text_to_be_present_in_element((By.ID, "lblLogin_AuthenticationMethod_Description"), 'Specify network'))

# Defines and fills in the user credentials
# Fill in credentials :)
username = ''
password = ''
usernameElem = chrome.find_element_by_id("tbxUserName")
usernameElem.send_keys(username)
passwordElem = chrome.find_element_by_id("tbxPassword")
passwordElem.send_keys(password)
chrome.implicitly_wait(1)

# Logs the user in after the credentials are filled in
loginButton = chrome.find_element_by_id("btnLogin")
loginButton.click()

# Waits for the user to sign into their account and for the webpage to load, set with a 60 sec timeout
WebDriverWait(chrome, 60).until(EC.url_contains(eco_url))

# Navigates to the webpage of the desired ECO to the BOM Item Field Check report

# Change the 'Check Field' to 'DESIGN COST PER UNIT (CAD)'
costSelect = Select(chrome.find_element_by_id("dlToolbar_FieldName"))
costAtt = 'Attribute: DESIGN COST PER UNIT (CAD)'
costSelect.select_by_visible_text(costAtt)

# Waits for webpage to switch fields to the design cost attribute 
WebDriverWait(chrome, 5).until(EC.text_to_be_present_in_element((By.ID, "lblParentItem_Field"), costAtt))

# Check the Find # box to include it as a field in the tables
findBox = chrome.find_element_by_id("cblBOMFields_DisplayOpt_3")
if not findBox.is_selected():
    findBox.click()
    print('Adding Find # column to tables')
    WebDriverWait(chrome, 5).until(EC.text_to_be_present_in_element((By.CSS_SELECTOR, '#dgOriginalBOM > tbody > tr.Header > td:nth-child(4)'), 'Find'))

# Check to see how many affected items there are
affItemSelect = Select(chrome.find_element_by_id("dlOptions_AffectedItems"))
affItemsOptions = affItemSelect.options
affItemList = []                       
for option in affItemsOptions:
    affItemList.append(option.text)
                
print("Affected items are: {}".format(affItemList))

# Loop through each affected item to gather the necessary data
for item in affItemList:

    # Select the affected item
    affItemSelect = Select(chrome.find_element_by_id("dlOptions_AffectedItems"))
    affItemSelect.select_by_visible_text(item)
    WebDriverWait(chrome, 5).until(EC.text_to_be_present_in_element((By.ID, "lblParentItem_PN"), item))

    # Gather the Original BOM into a list of lists
    cells = []
    originalBOMTable = []
    originalBOMTableElem = chrome.find_element_by_id("dgOriginalBOM")
    cellElemList = originalBOMTableElem.find_elements_by_tag_name('td')    

    # Convert the WebElement items into a list of table cell texts
    for cell in cellElemList:
        cells.append(cell.text)

    # Parse the list of table cell texts and convert into a proper row
    for i in range(0, len(cells), 8):
        c0 = cells[i+0]
        c1 = cells[i+1]
        c2 = cells[i+2]
        c3 = cells[i+3]
        c4 = cells[i+4]
        c5 = cells[i+5]
        c6 = cells[i+6]
        c7 = cells[i+7]
        originalBOMTable.append([c0, c1, c2, c3, c4, c5, c6, c7])

    for i in originalBOMTable:
        print(i)
    

#! python3
# getNewDesignCost.py
# Enter an ECO# and this will open a Selenium browser, user enters login info and then
# proceeds to the ECO and generates the affected item cost difference

import sys, math
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

# Get the ecoURL from the user when they run the script
if len(sys.argv) > 1:
    # Get url from the command line.
    ecoURL = str(sys.argv[1])
    print('\nProcessing url: {}'.format(ecoURL))

# Launch the Selenium browser and navigate to the ECO in Omnify
options = Options()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome = webdriver.Chrome(options=options)
chrome.maximize_window()
chrome.get(ecoURL)

# Set the login type to LDAP
loginSelect = Select(chrome.find_element_by_id("dlLogin_AuthenticationMethod"))
loginSelect.select_by_value("1")

# Wait for the page to load the proper fields for LDAP
WebDriverWait(chrome, 5).until(EC.text_to_be_present_in_element((By.ID, "lblLogin_AuthenticationMethod_Description"), 'Specify network'))

# Defines and fills in the user credentials
username = '' # Enter usename here in Windows login format ex. 'michael chaplin'
password = '' # Enter password here
usernameElem = chrome.find_element_by_id("tbxUserName")
usernameElem.send_keys(username)
passwordElem = chrome.find_element_by_id("tbxPassword")
passwordElem.send_keys(password)
chrome.implicitly_wait(1)

# Logs the user in after the credentials are filled in
loginButton = chrome.find_element_by_id("btnLogin")
loginButton.click()

# Waits for the user to sign into their account and for the webpage to load, set with a 60 sec timeout
WebDriverWait(chrome, 5).until(EC.url_contains(ecoURL))

# Click on the 'Affected Items' tab and waits for webpage to switch to the affected item tab of the ECO
newUrl = ecoURL.replace('Object', 'Reports/BOMFieldCheck')
chrome.get(newUrl)
WebDriverWait(chrome, 5).until(EC.text_to_be_present_in_element((By.ID, "lblPageOptionsAITitle"), 'Affected Items'))

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
    if not option.text[0] == 'D':
            affItemList.append(option.text)
    else:
        continue

print("\nAffected items are: {}\n".format(affItemList))
costData = []

# Loop through each affected item to gather the necessary data
for item in affItemList:

    # TODO: Make the table parsing happen on multiple threads to speed it up fast
    print('Selecting affected item: {}...'.format(item))

    # Select the affected item
    affItemSelect = Select(chrome.find_element_by_id("dlOptions_AffectedItems"))
    affItemSelect.select_by_visible_text(item)
    WebDriverWait(chrome, 5).until(EC.text_to_be_present_in_element((By.ID, "lblParentItem_PN"), item))

    # Get the original design costo
    originalDesignCost = round(float(chrome.find_element_by_id("lblParentItem_OldValue").text), 2)
    
    # Gather the Original BOM into a list of lists
    cells = []
    originalBOMCost = []

    # Finds the original BOM and skips over items like E and M parts that don't have a BOM
    try:
        originalBOMTableElem = chrome.find_element_by_id("dgOriginalBOM")
    except NoSuchElementException:
        print('Skipping over {} as it does not have a BOM\n'.format(item))
        continue        
    
    originalBOMCellElemList = originalBOMTableElem.find_elements_by_tag_name('td')
    print('Processing original BOM of {}...'.format(item))

    # Convert the WebElement items into a list of table cell texts
    for cell in originalBOMCellElemList:
        cells.append(cell.text)

   #print(cells)

    # Parse the list of table cell texts and take the necessary data for the original BOM
    for i in range(0, len(cells), 8):
        c2 = cells[i+2] # Quantity
        c7 = cells[i+7] # Design Cost
        if i < 1:
            c8 = 'Extended Cost'
        elif c7 == ' ':
            # Blank value is a string with 1 space in here ' '
            continue
        else:
            c8 = float(c2) * float(c7)    # Extended Cost
            originalBOMCost.append(c8)

    # Calculate the total cost of the originalBOM
    originalBOMCostTotal = round(sum(originalBOMCost, 1), 2)
    print('Finished processing original BOM of {}\n'.format(item))

    # Gather the Proposed BOM into a list of lists
    cells = []
    proposedBOMCost = []
    proposedBOMTableElem = chrome.find_element_by_id("dgProposedBOM")
    proposedBOMCellElemList = proposedBOMTableElem.find_elements_by_tag_name('td')
    print('Processing proposed BOM of {}...'.format(item))

    # Convert the WebElement items into a list of table cell texts
    for cell in proposedBOMCellElemList:
        cells.append(cell.text)

    # Parse the list of table cell texts and take the necessary data for the proposed BOM
    for i in range(0, len(cells), 9):
        c2 = cells[i+2] # Quantity
        c8 = cells[i+8] # Design Cost
        if i < 1:
            c9 = 'Extended Cost'
        elif c8 == '':
            # Blank value is a an empty string here ''
            continue
        else:
            c9 = float(c2) * float(c8)    # Extended Cost
            proposedBOMCost.append(c9)

    # Calculate the total cost of the proposedBOM
    proposedBOMCostTotal = round(sum(proposedBOMCost, 1), 2)
    print('Finished processing proposed BOM of {}\n'.format(item))

    # Store the data to iterate back through later
    costData.append([item, originalDesignCost, originalBOMCostTotal, proposedBOMCostTotal])

# Close the browser
print('Preparing summary...\n')
chrome.quit()

# Process the costData and display to the user
for i in range(len(costData)):
    affItemData = costData[i]
    print('Affected Item {}:\n \
          Original Design Cost = $ {}\n \
          Original BOM Design Cost Total = $ {}\n \
          Proposed BOM Design Cost Total = $ {}\n \
          '.format(affItemData[0], affItemData[1], affItemData[2], affItemData[3]))

#! python3
# getNewDesignCost.py
# Enter an ECO# and this will open a Selenium browser, user enters login info and then
# proceeds to the ECO and generates the affected item cost difference

import sys, math, time
from tabulate import tabulate
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException

# Calculates the total cost of a BOM
def calcTotalCost(item, nameOfBOM, itemBOM, *colNames):
	
	# Check how many rows are in the BOM
    print(f'Processing {nameOfBOM} BOM of {item}...')
    itemBOMRows = itemBOM.find_elements_by_tag_name('tr')
    numRows = len(itemBOMRows)
	
	# Find the column numbers of the quantity and design cost columns
    headerRow = itemBOMRows[0]    
    headerCols = headerRow.find_elements_by_tag_name('td')
    colNums = getTableColNums(headerCols, colNames)
    qtyCol = colNums[colNames[0]]
    costCol = colNums[colNames[1]]
        
	# Calculate and increment the extended cost of each BOM line item
    totalBOMCost = 0
    for row in range(1, numRows):
        qty = chrome.find_element_by_css_selector(f'#dg{nameOfBOM}BOM > tbody > tr:nth-child({row + 1}) > td:nth-child({qtyCol + 1})').text
        cost = chrome.find_element_by_css_selector(f'#dg{nameOfBOM}BOM > tbody > tr:nth-child({row + 1}) > td:nth-child({costCol + 1})').text         
        if not cost == '' and not cost == ' ':
            totalBOMCost += (float(qty) * float(cost))
	
	# Format the final cost of the BOM
    totalBOMCost = round(totalBOMCost, 2)    
    print(f'Finished processing {nameOfBOM} BOM of {item}\n')
    return totalBOMCost

# Returns a dictionary for column indices for a provided list of <td> elements and column names Strings
def getTableColNums(headerCols, colNames):

	# Initialize some helper variables
	colNum = 0
	colIndexDict = {}
	
	# Add all the requested columns to the dictionary
	# colIndexDict = {'Qty': '', 'Value': ''} for the normal use
	for colName in colNames:
		colIndexDict[colName] = ''

	# Finds the column numbers for each requested column
	for col in headerCols:		
		# Checks if the colName was requested and assigns the colNum to the dict as a value
		if not colIndexDict.get(col.text, -1) == -1:
			colIndexDict[col.text] = colNum
		
		# Increment counter		
		colNum += 1
	
	return colIndexDict

# Start the timer
startTime = time.time()

# Get the ECO Number from the user when they run the script
if len(sys.argv) > 1:
    # Get url from the command line.
    ecoNum = str(sys.argv[1])
    print(f'\nProcessing ECO: {ecoNum}')

ecoSearchURL = 'https://plm.mdsaero.com/Omnify7/Apps/Desktop/Search/Default.aspx?obj=1&fld0=Change%20Number&val0=' + ecoNum

# Launch the Selenium browser and navigate to the ECO in Omnify
options = Options()
options.add_experimental_option('excludeSwitches', ['enable-logging'])
chrome = webdriver.Chrome(options=options)
chrome.maximize_window()
chrome.get(ecoSearchURL)

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
WebDriverWait(chrome, 5).until(EC.text_to_be_present_in_element((By.ID, "dg_Results_lnkObjectForm_0"), ecoNum))
ecoID = chrome.find_element_by_id("dg_Results_lnkObjectForm_0").get_attribute('href').split('=')[1]

# Click on the 'Affected Items' tab and waits for webpage to switch to the affected item tab of the ECO
ecoURL = 'https://plm.mdsaero.com/Omnify7/Apps/Desktop/Change/Reports/BOMFieldCheck.aspx?id=XXXX'.replace('XXXX', ecoID)
chrome.get(ecoURL)
WebDriverWait(chrome, 5).until(EC.text_to_be_present_in_element((By.ID, "lblPageOptionsAITitle"), 'Affected Items'))

# Change the 'Check Field' to 'DESIGN COST PER UNIT (CAD)'
costSelect = Select(chrome.find_element_by_id("dlToolbar_FieldName"))
costAtt = 'Attribute: DESIGN COST PER UNIT (CAD)'
costSelect.select_by_visible_text(costAtt)

# Waits for webpage to switch fields to the design cost attribute
WebDriverWait(chrome, 5).until(EC.text_to_be_present_in_element((By.ID, "lblParentItem_Field"), costAtt))

# Check to see how many affected items there are
affItemSelect = Select(chrome.find_element_by_id("dlOptions_AffectedItems"))
affItemsOptions = affItemSelect.options
affItemList = []
for option in affItemsOptions:
    if not option.text[0] == 'D':
        # Ignores 'D' type items 
        affItemList.append(option.text)

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
    
    # Finds the original BOM and skips over items like E and M parts that don't have a BOM
    try:
        originalBOMTableElem = chrome.find_element_by_id("dgOriginalBOM")
    except NoSuchElementException:
        print('Skipping over {} as it does not have an original BOM\n'.format(item))
        continue         
    
    # Finds the proposed BOM and skips over items like E and M parts that don't have a BOM
    try:
        proposedBOMTableElem = chrome.find_element_by_id("dgProposedBOM")
    except NoSuchElementException:
        print('Skipping over {} as it does not have a proposed BOM\n'.format(item))
        continue 
    
    # Process the BOMs to caluclate the total costs
    originalBOMCostTotal = calcTotalCost(item, 'Original', originalBOMTableElem, 'Qty', 'Value')
    proposedBOMCostTotal = calcTotalCost(item, 'Proposed', proposedBOMTableElem, 'Qty', 'New Value')

    # Store the data to iterate back through later
    costData.append([item, originalDesignCost, originalBOMCostTotal, proposedBOMCostTotal])

# Stop the timer
stopTime = time.time()
duration = round(stopTime - startTime, 2)

# Close the browser
print(f'Finished processing in {duration} seconds. Now preparing summary...\n')
chrome.quit()

# Process the costData and display to the user
headers = ['Affected Item', 'Original Item\n Cost ($)', 'Total Original\n BOM Cost ($)', 'Total Proposed\n BOM Cost ($)']
print(tabulate(costData, headers=headers, tablefmt="presto"))

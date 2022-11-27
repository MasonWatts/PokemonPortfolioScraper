# -------- Includes --------

import os, requests, csv, time, xlsxwriter
from urllib.parse import urlparse, urljoin
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from datetime import datetime

# -------- Global Variables --------

# Web driver instance used for data scraping
driver = 0
# Excel spreadsheet for output
workbook = 0
# Spreadsheet page for summary info
summaryWorksheet = 0
# Index of the next open row on the summary spreadsheet page
summaryIndex = 0
# Cache of info that overrides default card URLs (some links on Pokellector to TCGPlayer are not correct)
cardCache = [['1','2','3']]
# Cache of info indicating the average condition of cards in each set
conditionCache = [['1','2']]

slabList = [['1','2','3']]

themeDeckList = [['1','2']]

# -------- Helper Functions: General --------

# Initializes the web browser and logs into Pokellector
def initBrowser():
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", os.getcwd())
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/plain, text/csv")
    driver = webdriver.Firefox(executable_path=r'geckodriver.exe', firefox_profile=profile)
    driver.maximize_window()
    driver.get("https://www.pokellector.com/signin/")
    userName = driver.find_element_by_xpath('//*[@id="columnLeft"]/form/label[1]/input')
    userName.send_keys("mason3195")
    password = driver.find_element_by_xpath('//*[@id="columnLeft"]/form/label[2]/input')
    password.send_keys("vU827rsKxNmk9F$@")
    loginButton = driver.find_element_by_xpath('//*[@id="columnLeft"]/form/div/button[2]')
    loginButton.click()
    return driver

# -------- Helper Functions: Individual Cards --------

# Indexes all set-specific pages in your Pokellector collection
def getAllSetLinks():
    urls = []
    driver.get("https://www.pokellector.com/my-collection/")
    collectionSeries = driver.find_elements_by_class_name("collection-series")
    for collection in collectionSeries:
        aTags = collection.find_elements_by_tag_name("a")
        for tag in aTags:
            href = tag.get_attribute("href")
            if "/sets/" in href:
                urls.append(href)
    return urls

# Indexes card-specific pages in a given set, dividing into collected and not collected
def getAllCardLinks(url):
    urls = []
    urlsMissing = []
    driver.get(url)
    cards = driver.find_elements_by_class_name('card.checked')
    for card in cards:
        aTags = card.find_elements_by_tag_name("a")
        for tag in aTags:
            href = tag.get_attribute("href")
            if "/card/" in href:
                urls.append(href)
    cardsMissing = driver.find_elements_by_class_name('card')
    
    for card in cardsMissing:
        if card in cards:
            continue
        aTags = card.find_elements_by_tag_name("a")
        for tag in aTags:
            href = tag.get_attribute("href")
            if "/card/" in href:
                urlsMissing.append(href)
    return urls, urlsMissing

# Checks the marketplace link override cache for the given card, returning the correct link if applicable
def searchCache(setName, cardName):
    for x in range(len(cardCache)):
        if len(cardCache[x]) < 3:
            continue
        if setName == cardCache[x][0] and cardName == cardCache[x][1]:
            return cardCache[x][2]
    return ""

def getCardPrice(vendorUrl, setCondition):
    if setCondition == "NM":
        conditionedUrl = vendorUrl + "?Language=English&Condition=Near%20Mint&page=1"
    if setCondition == "LP":
        conditionedUrl = vendorUrl + "?Language=English&Condition=Lightly%20Played&page=1"
    if setCondition == "MP":
        conditionedUrl = vendorUrl + "?Language=English&Condition=Moderately%20Played&page=1"
    driver.get(conditionedUrl)
    print(conditionedUrl)
    time.sleep(2)
    try:
        priceElement = driver.find_element_by_xpath('/html/body/div[2]/div/div/section[2]/section/section[2]/section/section/section/section[1]/div[2]/div[1]')
    except:
        time.sleep(4)
        priceElement = driver.find_element_by_xpath('/html/body/div[2]/div/div/section[2]/section/section[2]/section/section/section/section[1]/div[2]/div[1]')
    priceStr = priceElement.get_attribute('innerText')
    return priceStr

# Accesses marketplace URL and gets current market price for the card at the given condition
def processCard(url, setName, cardName, setCondition):
    soup = BeautifulSoup(requests.get(url).content, "html.parser")
    priceStr = ""
    suspectStr = ""
    vendorUrl = ""
    # Check cache: if no hit, process normally; if hit, go to that page
    cacheLink = searchCache(setName, cardName)
    if cacheLink == "" or cacheLink is None:
        # Find "priceblurb" area which contains vendor link and price
        for div in soup.findAll("div", "priceblurb"):
            vendorFound = False
            for cite in div.findAll("cite"):
                if "tcgplayer" in str(cite):
                    for aTag in cite.findAll("a"):
                        href = aTag.attrs.get("href")
                        if href == "" or href is None:
                            continue
                        vendorUrl = href
                    vendorFound = True
            if not vendorFound:
                continue
    else:
        vendorUrl = cacheLink
        
    if not (setName and cardName) in vendorUrl:
        suspectStr = "!"

    if vendorUrl == "" or vendorUrl is None:
        priceStr = "?"
        suspectStr = "?"
        vendorUrl = "?"
    else:
        priceStr = getCardPrice(vendorUrl, setCondition)

    return priceStr, suspectStr, vendorUrl

# Gets the condition for the given set from the conditions input cache
def searchConditionCache(setName):
    for x in range(len(conditionCache)):
        if len(conditionCache[x]) < 2:
            continue
        if setName == conditionCache[x][0]:
            return conditionCache[x][1]
    return ""

# Creates a worksheet page for the given set and fills it out with market price info
def processSet(setLink, summaryIndex):
    setMoniker = setLink.rsplit('/', 1)[-1]
    setId = setMoniker.split("-",1)[1]
    setName = (" ".join(setMoniker.split("-",1)[1:]))
    setCondition = searchConditionCache(setName)
    if setCondition == "" or setCondition is None:
        setCondition = "NM"
    setWorksheet = workbook.add_worksheet(setName.replace("-"," "))
    setIndex = 0
    setWorksheet.write(setIndex, 0, "Card Name")
    setWorksheet.write(setIndex, 1, "Unit Price")
    setWorksheet.write(setIndex, 2, "Quantity")
    setWorksheet.write(setIndex, 3, "Sum Price")
    setWorksheet.write(setIndex, 4, "Suspect Link")
    setWorksheet.write(setIndex, 5, "Vendor Link")
    setIndex += 1
    ownedCardLinks, missingCardLinks = getAllCardLinks(setLink)
    print(f"Found {len(ownedCardLinks)} cards in set {setLink}")
    allCardLinks = ownedCardLinks + missingCardLinks
    priceSumMissing = 0
    for cardLink in allCardLinks:
        try:
            cardMoniker = cardLink.rsplit('/', 1)[-1]
            cardName = cardMoniker.split(setName,1)[0].replace("-", " ").rstrip()
            if cardLink in ownedCardLinks:
                priceStr, suspectStr, linkStr = processCard(cardLink, setName, cardName, setCondition)
                if priceStr == "" or priceStr is None:
                    continue
                priceNum = float(priceStr[1:])
                setWorksheet.write(setIndex, 0, cardName)
                setWorksheet.write(setIndex, 1, priceNum)
                setWorksheet.write(setIndex, 2, '1')
                setWorksheet.write(setIndex, 3, f"=B{setIndex+1}*C{setIndex+1}")
                setWorksheet.write(setIndex, 4, suspectStr)
                setWorksheet.write(setIndex, 5, linkStr)
                setIndex += 1

        except:
            print("Error on ")
            print(cardName)
            print(cardLink)
    setName = setName.replace("-"," ")
    setWorksheet.write(setIndex+2, 0, "Set Price: ")
    setWorksheet.write(setIndex+2, 1, f"=SUM(D2:D{setIndex})")
    summaryWorksheet.write(summaryIndex, 0, setName)
    summaryWorksheet.write(summaryIndex, 1, setCondition)
    sumReference = "='"
    sumReference += setName
    sumReference += f"'!B{setIndex+3}"
    summaryWorksheet.write(summaryIndex, 2, sumReference)
    summaryWorksheet.write(summaryIndex, 3, len(ownedCardLinks))
    summaryWorksheet.write(summaryIndex, 4, len(missingCardLinks))
    summaryWorksheet.write(summaryIndex, 5, priceSumMissing)
    summaryIndex += 1

# -------- Helper Functions: Graded Cards --------

def getSlabPrice(setName, cardName, cardGrade):
    driver.get("https://www.pokemonprice.com/")
    searchBox = driver.find_element_by_xpath('/html/body/nav/div/form/input[3]')
    searchBox.send_keys(setName + ' ' + cardName)
    searchBox.send_keys(Keys.ENTER)
    time.sleep(2)
    searchText = setName + ' ' + cardName
    searchResults = driver.find_elements_by_tag_name('a')
    for x in range(len(searchResults)):
        if(searchText not in searchResults[x].text):
            continue
        linkText = searchResults[x].get_attribute('href')
    driver.get(linkText)
    print('HERE1')
    time.sleep(5)
    print('HERE')
    driver.execute_script("window.scrollTo(0, 1000)") 
    filterBox = driver.find_element_by_xpath('/html/body/div[3]/div[1]/div/div[1]/div[4]/div/div[2]/label/input')
    filterBox.send_keys(cardGrade)
    select = Select(driver.find_element_by_name('prices_length'))
    select.select_by_visible_text('100')
    print('HERE3')
    time.sleep(0.5)
    print('HERE4')
    salesTable = driver.find_element_by_xpath('/html/body/div[3]/div[1]/div/div[1]/div[4]/div/table/tbody')
    prices = []
    print('HERE5')
    for row in salesTable.find_elements_by_xpath(".//tr"):
        columns = row.find_elements_by_tag_name('td')
        if len(columns) < 7:
            continue
        if cardGrade not in columns[1].text:
            continue
        print(columns[2].text)
        prices.append(float(columns[2].text.replace('$', '')))
    sumPrice = 0.0
    print('HERE6')
    for x in range(len(prices)):
        sumPrice += prices[x]
    averagePrice = sumPrice / len(prices)
    print('HERE7')
    print(averagePrice)
    print('------------')
    return averagePrice

def processGradedCards():
    worksheet = workbook.add_worksheet("Graded Cards")
    index = 0
    worksheet.write(index, 0, "Card Name")
    worksheet.write(index, 1, "Set")
    worksheet.write(index, 2, "Grade")
    worksheet.write(index, 3, "Price")
    index += 1

    try:
        driver.get("https://www.pokemonprice.com/")
        time.sleep(3)
        buttonAgree = driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[2]/div/button[3]')
        buttonAgree.click()
    except:
        print('Agreement not there')

    for x in range(len(slabList)):
        setName = slabList[x][0]
        cardName = slabList[x][1]
        cardGrade = slabList[x][2]
        price = getSlabPrice(setName, cardName, cardGrade)
        worksheet.write(index, 0, cardName)
        worksheet.write(index, 1, setName)
        worksheet.write(index, 2, cardGrade)
        worksheet.write(index, 3, price)
        index += 1
    
    worksheet.write(index+2, 0, "Total Price ")
    worksheet.write(index+2, 3, f"=SUM(D2:D{index})")
    summaryWorksheet.write(summaryIndex, 0, "Graded Cards")
    summaryWorksheet.write(summaryIndex, 2, f"='Graded Cards'!D{index+3}")

# -------- Helper Functions: Theme Decks --------

def getDeckList(deckUrl):
    decklist = []
    driver.get(deckUrl)
    
    setBox = driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[1]/div[6]/div[4]/div/table[1]/tbody/tr[2]/td/table[2]/tbody/tr/td/a')
    setName = setBox.get_attribute('innerText')
    setName = setName.lower().replace(' ', '-')
    print(setName)
    listTable = driver.find_element_by_xpath('/html/body/div[2]/div[1]/div[1]/div[6]/div[4]/div/table[2]/tbody')
    for row in listTable.find_elements_by_xpath(".//tr"):
        columns = row.find_elements_by_tag_name('td')
        if len(columns) < 3:
            continue
        try:
            rarityElement = columns[2].find_element_by_tag_name('a')
            rarityImage = rarityElement.get_attribute('title')
            holo = ('holo' in rarityImage.lower())
        except:
            holo = False
        cardName = columns[1].text.lower().replace(' ', '-').replace('é', 'e').replace('\'', "")
        cardName = cardName.replace('♂', '-m').replace('♀', '-f')
        innerElement = columns[1].find_element_by_tag_name('a')
        cardSet = innerElement.get_attribute('title')
        cardSet = cardSet[cardSet.find("(")+1:cardSet.find(")")]
        cardNumber = cardSet
        cardSetParts = cardSet.split(' ')
        cardSet = ' '.join(cardSetParts[0:-1])
        cardSet = cardSet.lower().rstrip().replace(' ', '-')
        cardNumber = ''.join(c for c in cardNumber if c.isnumeric())
        if cardSet != 'base-set' and cardSet != 'base-set-2' and holo:
            cardName = cardName + "-" + cardNumber
        cardQuantity = columns[0].text.replace('×', '')
        decklist.append([cardSet, cardName, cardQuantity])
        print(cardName + ", " + cardQuantity + ", " + cardSet)
    return decklist

def processThemeDeck(deckName, deckUrl, summaryIndex):
    worksheet = workbook.add_worksheet(deckName)
    index = 0
    worksheet.write(index, 0, "Card Name")
    worksheet.write(index, 1, "Unit Price")
    worksheet.write(index, 2, "Quantity")
    worksheet.write(index, 3, "Sum Price")
    worksheet.write(index, 4, "Suspect Link")
    worksheet.write(index, 5, "Vendor Link")
    index += 1

    deckList = getDeckList(deckUrl)
    for x in range(len(deckList)):
        if len(deckList[x]) < 3:
            continue
        setName = deckList[x][0]
        cardName = deckList[x][1]
        cardQuantity = deckList[x][2]
        setCondition = 'NM'
        cardLink = "https://shop.tcgplayer.com/pokemon/" + setName + "/" + cardName
        priceStr = getCardPrice(cardLink, setCondition)
        if priceStr == "" or priceStr is None:
            continue
        priceNum = float(priceStr[1:])
        worksheet.write(index, 0, cardName)
        worksheet.write(index, 1, priceNum)
        worksheet.write(index, 2, cardQuantity)
        worksheet.write(index, 3, f"=B{index+1}*C{index+1}")
        index += 1
    
    worksheet.write(index+2, 0, "Set Price: ")
    worksheet.write(index+2, 1, f"=SUM(D2:D{index})")
    summaryWorksheet.write(summaryIndex, 0, deckName)
    sumReference = "='"
    sumReference += deckName
    sumReference += f"'!B{index+3}"
    summaryWorksheet.write(summaryIndex, 2, sumReference)
    summaryIndex += 1

# -------- Main --------

# Announce process start
now = datetime.now()
current_time = now.strftime("%H:%M:%S")
print("Current Time =", current_time)  

# Read in marketplace link correction cache
with open('Inputs\MarketplaceLinkCorrection.csv', newline='') as f:
    reader = csv.reader(f)
    cardCache = list(reader)

# Read in set condition cache
with open('Inputs\SetCondition.csv', newline='') as f:
    reader = csv.reader(f)
    conditionCache = list(reader)

# Read in theme deck list
with open('Inputs\ThemeDecks.csv', newline='') as f:
    reader = csv.reader(f)
    themeDeckList = list(reader)

# Read in graded cards list
with open('Inputs\GradedCards.csv', newline='') as f:
    reader = csv.reader(f)
    slabList = list(reader)

# Create output spreadsheet and set up summary page
workbook = xlsxwriter.Workbook('Collection.xlsx')
summaryWorksheet = workbook.add_worksheet('Summary')
summaryWorksheet.write(summaryIndex, 0, "Set")
summaryWorksheet.write(summaryIndex, 1, "Condition")
summaryWorksheet.write(summaryIndex, 2, "Total Price")
summaryWorksheet.write(summaryIndex, 3, "Cards Collected")
summaryWorksheet.write(summaryIndex, 4, "Cards Not Collected")
summaryWorksheet.write(summaryIndex, 5, "Cost to Complete")
summaryIndex += 1

# Initialize web browser and catalogue all available sets
driver = initBrowser()
availableSetLinks = getAllSetLinks()
print(f"Found {len(availableSetLinks)} sets")

# Process each detected set, accessing market price info and adding each set to its own page
#for setLink in availableSetLinks:
    #processSet(setLink, summaryIndex)
    #summaryIndex += 1

# Process theme deck list, accessing market price info and adding each set to its own page
for x in range(len(themeDeckList)):
    processThemeDeck(themeDeckList[x][0], themeDeckList[x][1], summaryIndex)
    summaryIndex += 1

# Process graded cards list, accessing market price info and summing all on one page
#processGradedCards()
#summaryIndex += 1

# After all sets are added, create collection value sum cell
summaryWorksheet.write(summaryIndex + 2, 0, "Collection Price:")
summaryWorksheet.write(summaryIndex + 2, 2, f"=SUM(C2:C{summaryIndex})")

# Close up workbook
workbook.close()
driver.close()

# Announce process completion
now = datetime.now()
current_time = now.strftime("%H:%M:%S")
print("Current Time =", current_time)    
print("Done")
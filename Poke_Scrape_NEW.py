import os, requests, csv, time, xlsxwriter
from urllib.parse import urlparse, urljoin
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import datetime

preferredVendor = "tcgplayer"
backupVendor = "trollandtoad"
workbook = 0
summaryWorksheet = 0
driver = 0
summaryIndex = 0
cardCache = [['1','2','3']]
conditionCache = [['1','2']]

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
    userName.send_keys("USERNAME")
    password = driver.find_element_by_xpath('//*[@id="columnLeft"]/form/label[2]/input')
    password.send_keys("PASSWORD")
    loginButton = driver.find_element_by_xpath('//*[@id="columnLeft"]/form/div/button[2]')
    loginButton.click()
    return driver

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

def searchCache(setName, cardName):
    for x in range(len(cardCache)):
        if len(cardCache[x]) < 3:
            continue
        if setName == cardCache[x][0] and cardName == cardCache[x][1]:
            return cardCache[x][2]
    return ""

def getCardPrice(url, setName, cardName, setCondition):
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
        if setCondition == "NM":
            conditionedUrl = vendorUrl + "&Condition=Near%20Mint&page=1"
        if setCondition == "LP":
            conditionedUrl = vendorUrl + "&Condition=Lightly%20Played&page=1"
        if setCondition == "MP":
            conditionedUrl = vendorUrl + "&Condition=Moderately%20Played&page=1"
        driver.get(conditionedUrl)
        time.sleep(2)
        try:
            priceElement = driver.find_element_by_xpath('//*[@id="app"]/div/section[2]/section/div[2]/section/section/section/section[1]/div[2]/div[1]')
        except:
            time.sleep(4)
            priceElement = driver.find_element_by_xpath('//*[@id="app"]/div/section[2]/section/div[2]/section/section/section/section[1]/div[2]/div[1]')
        priceStr = priceElement.get_attribute('innerText')

    return priceStr, suspectStr, vendorUrl

def searchConditionCache(setName):
    for x in range(len(conditionCache)):
        if len(conditionCache[x]) < 2:
            continue
        if setName == conditionCache[x][0]:
            return conditionCache[x][1]
    return ""

def processSet(setLink, sumIndex):
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
                priceStr, suspectStr, linkStr = getCardPrice(cardLink, setName, cardName, setCondition)
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
            # else:
            #     priceStr, suspectStr, linkStr = getCardPrice(cardLink, setName, cardName, setCondition)
            #     priceNum = float(priceStr[1:])
            #     priceSumMissing += priceNum

        except:
            print("Error on ")
            print(cardName)
            print(cardLink)
    setName = setName.replace("-"," ")
    setWorksheet.write(setIndex+2, 0, "Set Price: ")
    setWorksheet.write(setIndex+2, 1, f"=SUM(D2:D{setIndex})")
    summaryWorksheet.write(sumIndex, 0, setName)
    summaryWorksheet.write(sumIndex, 1, setCondition)
    sumReference = "='"
    sumReference += setName
    sumReference += f"'!B{setIndex+3}"
    summaryWorksheet.write(sumIndex, 2, sumReference)
    summaryWorksheet.write(sumIndex, 3, len(ownedCardLinks))
    summaryWorksheet.write(sumIndex, 4, len(missingCardLinks))
    summaryWorksheet.write(sumIndex, 5, priceSumMissing)

#Main--------------------

now = datetime.now()

current_time = now.strftime("%H:%M:%S")
print("Current Time =", current_time)  

with open('cache.csv', newline='') as f:
        reader = csv.reader(f)
        cardCache = list(reader)

with open('set_condition.csv', newline='') as f:
        reader = csv.reader(f)
        conditionCache = list(reader)

workbook = xlsxwriter.Workbook('Collection.xlsx')
summaryWorksheet = workbook.add_worksheet('Summary')

summaryWorksheet.write(summaryIndex, 0, "Set")
summaryWorksheet.write(summaryIndex, 1, "Condition")
summaryWorksheet.write(summaryIndex, 2, "Total Price")
summaryWorksheet.write(summaryIndex, 3, "Cards Collected")
summaryWorksheet.write(summaryIndex, 4, "Cards Not Collected")
summaryWorksheet.write(summaryIndex, 5, "Cost to Complete")
summaryIndex += 1

driver = initBrowser()

availableSetLinks = getAllSetLinks()
print(f"Found {len(availableSetLinks)} sets")

for setLink in availableSetLinks:
    processSet(setLink, summaryIndex)
    summaryIndex += 1

summaryWorksheet.write(summaryIndex+2, 0, "Collection Price:")
summaryWorksheet.write(summaryIndex+2, 2, f"=SUM(C2:C{summaryIndex})")

workbook.close()
driver.close()

now = datetime.now()

current_time = now.strftime("%H:%M:%S")
print("Current Time =", current_time)    
print("Done")
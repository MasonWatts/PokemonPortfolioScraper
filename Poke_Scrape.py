import os, requests, csv, xlsxwriter
from urllib.parse import urlparse, urljoin
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import datetime

preferredVendor = "trollandtoad"
backupVendor = "collectorscache"
workbook = 0
summaryWorksheet = 0
driver = 0
summaryIndex = 0
cardCache = [['1','2','3']] 

def initBrowser():
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", os.getcwd())
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/plain, text/csv")
    driver = webdriver.Firefox(executable_path=r'C:/Users/Mason/Desktop/Pokemon Scraper/geckodriver', firefox_profile=profile)
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
    cardsMissing = driver.find_elements_by_class_name('card.checked')
    for card in cardsMissing:
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

def getCardPrice(url, setName, cardName):
    soup = BeautifulSoup(requests.get(url).content, "html.parser")
    priceStr = ""
    vendorUrl = ""
    # Check cache: if no hit, process normally; if hit, go to that page
    cacheLink = searchCache(setName, cardName)
    if cacheLink == "" or cacheLink is None:
        # Find "priceblurb" area which contains vendor link and price
        for div in soup.findAll("div", "priceblurb"):
            vendorFound = False
            for cite in div.findAll("cite"):
                if preferredVendor in str(cite) or backupVendor in str(cite):
                    for aTag in cite.findAll("a"):
                        href = aTag.attrs.get("href")
                        if href == "" or href is None:
                            continue
                        vendorUrl = href
                    vendorFound = True
            if not vendorFound:
                continue
            for subdiv in div.findAll("div"):
                tag = subdiv.attrs.get("class")
                if tag == "" or tag is None:
                    continue
                if "price" in tag:
                    price = subdiv.getText()
                    if "$" in price:
                        priceStr = price
        # Check if vendor link is suspect (doesn't seem to match the card)
        suspectStr = "!"
        if not vendorUrl == "":
            if "trollandtoad" in vendorUrl:
                soup = BeautifulSoup(requests.get(vendorUrl).content, "html.parser")
                if setName == "Sun Moon":
                    setName = "Sun & Moon"
                for title in soup.findAll("img"):
                    text = title.attrs.get("alt")
                    if cardName in text and setName in text:
                        suspectStr = ""
                        break
            elif "collectorscache" in vendorUrl:
                soup = BeautifulSoup(requests.get(vendorUrl).content, "html.parser")
                if " EX" in cardName:
                    cardName = cardName.replace(" EX", "-EX")
                if "M " in cardName:
                    cardName = cardName.replace("M ", "Mega-")
                if "BREAK" in setName:
                    setName = setName.replace("BREAK", "Break")
                for title in soup.findAll("title"):
                    text = title.getText()
                    if "1st" in text:
                        continue
                    if cardName in text and setName in text:
                        suspectStr = ""
                        break
    else:
        vendorUrl = cacheLink
        suspectStr = ""
        if "trollandtoad" in vendorUrl:
            soup = BeautifulSoup(requests.get(vendorUrl).content, "html.parser")
            for element in soup.findAll("span"):
                text = element.getText()
                if "$" in text:
                    priceStr = text
                    break
        elif "collectorscache" in vendorUrl:
            soup = BeautifulSoup(requests.get(vendorUrl).content, "html.parser")
            for element in soup.findAll("script"):
                if "price" in str(element):
                    sections = str(element).split(",")
                    for text in sections:
                        if "price" in text:
                            values = text.split("'")
                            for value in values:
                                if "$" in value:
                                    priceStr = value
                                    break
    return priceStr, suspectStr, vendorUrl

def processSet(setLink, sumIndex):
    setMoniker = setLink.rsplit('/', 1)[-1]
    setId = setMoniker.split("-",1)[1]
    setName = (" ".join(setMoniker.split("-",1)[1:]))
    setWorksheet = workbook.add_worksheet(setName.replace("-"," "))
    setIndex = 0
    setWorksheet.write(setIndex, 0, "Card Name")
    setWorksheet.write(setIndex, 1, "Unit Price")
    setWorksheet.write(setIndex, 2, "Quantity")
    setWorksheet.write(setIndex, 3, "Sum Price")
    setWorksheet.write(setIndex, 4, "Suspect Link")
    setWorksheet.write(setIndex, 5, "Vendor Link")
    setIndex += 1
    availableCardLinks, numMissing = getAllCardLinks(setLink)
    print(f"Found {len(availableCardLinks)} cards in set {setLink}")
    for cardLink in availableCardLinks:
        #try:
            cardMoniker = cardLink.rsplit('/', 1)[-1]
            cardName = cardMoniker.split(setName,1)[0].replace("-", " ").rstrip()
            priceStr, suspectStr, linkStr = getCardPrice(cardLink, setName.replace("-"," "), cardName)
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
        #except:
            #print("Error on ")
            #print(cardLink)
    setName = setName.replace("-"," ")
    setWorksheet.write(setIndex+2, 0, "Set Price: ")
    setWorksheet.write(setIndex+2, 1, f"=SUM(D2:D{setIndex})")
    summaryWorksheet.write(sumIndex, 0, setName)
    sumReference = "='"
    sumReference += setName
    sumReference += f"'!B{setIndex+3}"
    summaryWorksheet.write(sumIndex, 1, sumReference)
    summaryWorksheet.write(sumIndex, 2, len(availableCardLinks))
    summaryWorksheet.write(sumIndex, 3, numMissing)

#Main--------------------

now = datetime.now()

current_time = now.strftime("%H:%M:%S")
print("Current Time =", current_time)  

with open('cache.csv', newline='') as f:
        reader = csv.reader(f)
        cardCache = list(reader)

workbook = xlsxwriter.Workbook('Collection.xlsx')
summaryWorksheet = workbook.add_worksheet('Summary')

summaryWorksheet.write(summaryIndex, 0, "Set")
summaryWorksheet.write(summaryIndex, 1, "Total Price")
summaryWorksheet.write(summaryIndex, 2, "Cards Collected")
summaryWorksheet.write(summaryIndex, 3, "Cards Not Collected")
summaryIndex += 1

driver = initBrowser()

availableSetLinks = getAllSetLinks() #Main request, see above Def.
numResponses = len(availableSetLinks)
print(f"Found {len(availableSetLinks)} sets")

for setLink in availableSetLinks:
    processSet(setLink, summaryIndex)
    summaryIndex += 1
# processSet("https://www.pokellector.com/sets/DET-Detective-Pikachu", summaryIndex)
# summaryIndex += 1

summaryWorksheet.write(summaryIndex+2, 0, "Collection Price:")
summaryWorksheet.write(summaryIndex+2, 1, f"=SUM(B2:B{summaryIndex})")

workbook.close()
driver.close()

now = datetime.now()

current_time = now.strftime("%H:%M:%S")
print("Current Time =", current_time)    
print("Done")
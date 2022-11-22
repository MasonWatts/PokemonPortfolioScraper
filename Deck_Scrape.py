import os, requests, csv, xlsxwriter, time
from urllib.parse import urlparse, urljoin
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from datetime import datetime

workbook = 0
worksheet = 0
driver = 0
summaryIndex = 0

def initBrowser():
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", os.getcwd())
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/plain, text/csv")
    driver = webdriver.Firefox(executable_path=r'C:/Users/Mason/Desktop/Pokemon Scraper/geckodriver', firefox_profile=profile)
    driver.get('https://shop.tcgplayer.com/pokemon/base-set/mewtwo')
    conditionFilter = driver.find_element_by_xpath('/html/body/div[4]/section[3]/div[1]/div[2]/div/div/ul[4]/li[2]/a')
    conditionFilter.click()
    time.sleep(2)
    return driver

def getDeckList(deckUrl):
    decklist = []
    driver.get(deckUrl)
    setBox = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[7]/div/div[1]/div[4]/table[1]/tbody/tr[2]/td/table[2]/tbody/tr/td/a')
    setName = setBox.get_attribute('innerText')
    setName = setName.lower().replace(' ', '-')
    print(setName)
    listTable = driver.find_element_by_xpath('/html/body/div[1]/div[2]/div/div[7]/div/div[1]/div[4]/table[2]/tbody')
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

def getCardPrice(setName, cardName):
    url = "https://shop.tcgplayer.com/pokemon/" + setName + "/" + cardName
    driver.get(url)
    time.sleep(2)
    try:
        priceBox = driver.find_element_by_xpath('/html/body/div[4]/section[1]/div/section/div[4]/div[2]/div/div/div[2]/span[1]')
    except:
        time.sleep(2)
        priceBox = driver.find_element_by_xpath('/html/body/div[4]/section[1]/div/section/div[4]/div[2]/div/div/div[2]/span[1]')
    price = priceBox.get_attribute('innerText')
    price = price.replace('$', '')
    return float(price)

#Main--------------------

now = datetime.now()

current_time = now.strftime("%H:%M:%S")
print("Current Time =", current_time)  

workbook = xlsxwriter.Workbook('Deck.xlsx')
worksheet = workbook.add_worksheet('Output')

driver = initBrowser()

decklist = getDeckList('https://bulbapedia.bulbagarden.net/wiki/Overgrowth_(TCG)')

rowIndex = 0
total = 0.0
for x in range(len(decklist)):
    if len(decklist[x]) < 3:
        continue
    cardPrice = getCardPrice(decklist[x][0], decklist[x][1])
    worksheet.write(rowIndex, 0, decklist[x][1])
    worksheet.write(rowIndex, 1, decklist[x][2])
    worksheet.write(rowIndex, 2, cardPrice)
    worksheet.write(rowIndex, 3, cardPrice * float(decklist[x][2]))
    rowIndex += 1
    total += cardPrice * float(decklist[x][2])

rowIndex += 1
worksheet.write(rowIndex, 3, total)

workbook.close()
driver.close()

now = datetime.now()

current_time = now.strftime("%H:%M:%S")
print("Current Time =", current_time)    
print("Done")
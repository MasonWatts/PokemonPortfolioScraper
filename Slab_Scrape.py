import os, requests, csv, xlsxwriter, time
from urllib.parse import urlparse, urljoin
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from datetime import datetime

workbook = 0
worksheet = 0
driver = 0

def initBrowser():
    profile = webdriver.FirefoxProfile()
    profile.set_preference("browser.download.folderList", 2)
    profile.set_preference("browser.download.manager.showWhenStarting", False)
    profile.set_preference("browser.download.dir", os.getcwd())
    profile.set_preference("browser.helperApps.neverAsk.saveToDisk", "text/plain, text/csv")
    driver = webdriver.Firefox(executable_path=r'C:/Users/Mason/Desktop/Pokemon Scraper/geckodriver', firefox_profile=profile)
    return driver

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
    time.sleep(5)
    filterBox = driver.find_element_by_xpath('/html/body/div[3]/div[1]/div/div[1]/div[3]/div/div[2]/label/input')
    filterBox.send_keys(cardGrade)
    select = Select(driver.find_element_by_name('prices_length'))
    select.select_by_visible_text('100')
    time.sleep(0.5)
    salesTable = driver.find_element_by_xpath('/html/body/div[3]/div[1]/div/div[1]/div[3]/div/table/tbody')
    prices = []
    for row in salesTable.find_elements_by_xpath(".//tr"):
        columns = row.find_elements_by_tag_name('td')
        if len(columns) < 7:
            continue
        if cardGrade not in columns[1].text:
            continue
        print(columns[2].text)
        prices.append(float(columns[2].text.replace('$', '')))
    sumPrice = 0.0
    for x in range(len(prices)):
        sumPrice += prices[x]
    averagePrice = sumPrice / len(prices)
    print(averagePrice)
    print('------------')
    return averagePrice

#Main--------------------

now = datetime.now()

current_time = now.strftime("%H:%M:%S")
print("Current Time =", current_time)  

slabList = [['1','2','3']] 
with open('slabs.csv', newline='') as f:
    reader = csv.reader(f)
    slabList = list(reader)

workbook = xlsxwriter.Workbook('Deck.xlsx')
worksheet = workbook.add_worksheet('Output')

driver = initBrowser()

try:
    driver.get("https://www.pokemonprice.com/")
    time.sleep(3)
    buttonAgree = driver.find_element_by_xpath('/html/body/div[1]/div/div/div/div[2]/div/button[3]')
    buttonAgree.click()
except:
    print('Agreement not there')

for x in range(len(slabList)):
    price = getSlabPrice(slabList[x][0], slabList[x][1], slabList[x][2])

workbook.close()
#driver.close()

now = datetime.now()

current_time = now.strftime("%H:%M:%S")
print("Current Time =", current_time)    
print("Done")
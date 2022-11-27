from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import os, csv
profile = 0
driver = 0
def initializeDriver():
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

    driver.get("https://www.pokellector.com/my-collection/")

    collection_series = driver.find_elements_by_class_name("collection-series")

    for collection in collection_series:
        a_tags = collection.find_elements_by_tag_name("a")
        for tag in a_tags:
            href = tag.get_attribute("href")
            if "/sets/" in href:
                print(href)

    driver.close()
    

initializeDriver()
print("Done")

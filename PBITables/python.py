from selenium import webdriver
from selenium.webdriver.common.keys import Keys

driver = webdriver.Chrome("D:\drivers\chromedriver.exe")
driver.get("https://www.google.com")

driver.implicitlyWait(20)
driver.close()
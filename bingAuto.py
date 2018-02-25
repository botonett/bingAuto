import requests #get the source page with this
import os #use this to start edge browers
from bs4 import BeautifulSoup #make minified html readable for debug
import fileinput
import time
import pyautogui #auto control of mouse and key board
    #dependencies - pip install image
from selenium import webdriver
from random import randint
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from selenium.common.exceptions import NoAlertPresentException
from selenium.webdriver.chrome.options import Options
import smtplib
import re
def get_credits():
    try:
        chrome_options1 = Options()
        chrome_options1.add_argument("user-data-dir=C:/Users/bing/AppData/Local/Google/Chrome/User Data")
        driver1 = webdriver.Chrome(chrome_options = chrome_options1)
        driver1.get('https://account.microsoft.com/rewards/')
        current_credits = (driver1.find_elements_by_xpath("""//*[@id="userStatus"]/div/mee-rewards-user-status-counter/div[1]/div/div/div/div/p[1]""")[0].text).replace(',','')
        driver1.quit()
        return int(current_credits)
    except:
        print("failed")


print(get_credits())

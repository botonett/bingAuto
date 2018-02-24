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
#get credits

def get_credits():
    try:
        chrome_options1 = Options()
        chrome_options1.add_argument("user-data-dir=C:/Users/bing/AppData/Local/Google/Chrome/User Data")
        driver1 = webdriver.Chrome(chrome_options = chrome_options1)
        driver1.get('https://account.microsoft.com/rewards/dashboard')
        current_credits = driver1.find_element_by_class_name("info-title").text
    except(KeyboardInterrupt, SystemExit):
        raise
    except:
        os.system("startChrome.bat")
        
        
        time.sleep(15)
        os.system("KillChrome.bat")
        #driver1.quit()
        chrome_options1 = Options()
        chrome_options1.add_argument("user-data-dir=C:/Users/bing/AppData/Local/Google/Chrome/User Data")
        driver1 = webdriver.Chrome(chrome_options = chrome_options1)
        driver1.get('https://account.microsoft.com/rewards/dashboard')
        current_credits = driver1.find_element_by_class_name("info-title").text
    driver1.quit()
    return current_credits
def hello():
    print("Hello World")
    print("testing update!")

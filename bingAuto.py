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
        driver1.get('https://www.bing.com/')
        time.sleep(4)
        current_credits = driver1.find_element_by_id("id_rc").text
        counter = 5
        while(counter > 0):
            counter = counter - 1
            try:
                int(current_credits)
                break
            except:
                time.sleep(4)
                current_credits = driver1.find_element_by_id("id_rc").text
        driver1.quit()
        print("Current Credits: "+ str(current_credits))
        return int(current_credits)
    except Exception as E:
        print("failed to get credit after 20 seconds with error: " + str(E))

def get_profile():
    profile = []
    with open("C:\\Users\\bing\\Desktop\\Bing2.0\\data\\profile.dat", 'r') as pf:
            for line in pf:
                 profile.append(line.strip())
    with open("C:\\Users\\bing\\Desktop\\Bing2.0\\data\\shutdown.dat", 'r') as pf:
            for line in pf:
                 profile.append(line.strip())
    return profile

def getAccount():
    profile = []
    with open("C:\\Users\\bing\\Desktop\\Bing2.0\\data\\reportEngine.dat", 'r') as pf:
            for line in pf:
                 profile.append(line.strip())
    return profile[0],profile[1]

#send email function
def send_email(user, pwd, recipient, subject, body):
    import smtplib

    gmail_user = user
    gmail_pwd = pwd
    FROM = 'Report Engine'
    TO = recipient if type(recipient) is list else [recipient]
    SUBJECT = subject
    TEXT = body

    # Prepare actual message
    message = """From: %s\nTo: %s\nSubject: %s\n\n%s
    """ % (FROM, ", ".join(TO), SUBJECT, TEXT)
    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.ehlo()
        server.starttls()
        server.login(gmail_user, gmail_pwd)
        server.sendmail(FROM, TO, message)
        server.close()
        print('successfully sent the report mail')
    except:
        print("failed to send mail")
def keywords():
    #make a list ready for search
    keyWords = []
    
    with open('C:\\Users\\bing\\Desktop\\Bing2.0\\data\\keyWords_Backup.dat', 'r') as f:
        for line in f:
            keyWords.append(line.strip())
    print('Library Preparation Successful: ' + str(len(keyWords))+ ' Keywords')
    return keyWords

#start PC Search on Edge
def search(range1):
    counter = 0
    temp2 = 0
    timeSpent = 16.5
    minutes = 0
    seconds = 0
    range1 = range1

    
    for x in range(0, range1):
        try:
            driver = webdriver.Edge()
            driver.get('http://bing.com')
        except(KeyboardInterrupt, SystemExit):
            raise
        except:
            print("edge start failed, killing application")
            
            #driver.quit()
           
            print("restarting search process")
            driver = webdriver.Edge()
            driver.get('http://bing.com')
        time.sleep(3)
        seed = randomNum(size)
        pyautogui.typewrite(keyWords[seed] + '\n', interval=0.1)
        counter += 1
        
        time.sleep(3)
        driver.execute_script("window.scrollTo(0, 200)")
        time.sleep(1.5)
        driver.execute_script("window.scrollTo(0, 150)")
        time.sleep(2)
        driver.execute_script("window.scrollTo(0, 300)")
        time.sleep(2)
        driver.execute_script("window.scrollTo(0, 150)")
        pyautogui.mouseDown(x=339, y=314, button='left')
        pyautogui.mouseUp(x=339, y=314, button='left')
        time.sleep(2)
        pyautogui.mouseDown(x=339, y=314, button='left')
        pyautogui.mouseUp(x=312, y=336, button='left')
        time.sleep(3)
        driver.quit()
        temp2 = randint(10,25)
        timeSpent = timeSpent + temp2
        print('Searched '+ str(counter)+ ' out of ' + str(range1) + ': ' + str(keyWords[seed]))
        print('Waiting for: '+ str(temp2) +' Seconds' )
        time.sleep(temp2)

        print('Total Words Seached in This Session:' + str(counter))
       
        if(counter == 0):
            timeSpent = 0
        minutes = int(timeSpent/60)
        seconds = timeSpent % 60
        print('Total Time Spent is ' + str(minutes) + ' minute(s) ' + str(seconds) + " second(s).")
        
    return counter

#generate a random number that is not repeated
def randomNum(size):
    seed = randint(0,size-1)
    calledNum = []
    for item in calledNum:
        if(item == seed):
            seed = randint(0,size-1)
        else:
            calledNum.append(seed)
    return seed
def mobile_search(range1):
    counter = 0
    temp2 = 0
    timeSpent = 16.5
    minutes = 0
    seconds = 0
    range1 = range1
    for x in range(0, range1):
        try:
            mobile_emulation = {
                "deviceMetrics": { "width": 360, "height": 640, "pixelRatio": 3.0 },
                "userAgent": "Mozilla/5.0 (Linux; Android 4.2.1; en-us; Nexus 5 Build/JOP40D) AppleWebKit/535.19 (KHTML, like Gecko) Chrome/18.0.1025.166 Mobile Safari/535.19" }
            chrome_options = Options()
            chrome_options.add_experimental_option("mobileEmulation", mobile_emulation)
            chrome_options.add_argument("user-data-dir=C:/Users/bing/AppData/Local/Google/Chrome/User Data")
            driver = webdriver.Chrome(chrome_options = chrome_options)
            driver.get('http://bing.com')
        except(KeyboardInterrupt, SystemExit):
            raise
        except:
            print("open chrome failed, kiling application")
            os.system("KillChrome.bat")
            #driver.quit()
           
            os.system("startChrome.bat")
            time.sleep(5)
            
            os.system("KillChrome.bat")
            
            print("restarting search process")
            mobile_emulation = {
                "deviceMetrics": { "width": 360, "height": 640, "pixelRatio": 3.0 },
                "userAgent": "Mozilla/5.0 (Linux; Android 4.2.1; en-us; Nexus 5 Build/JOP40D) AppleWebKit/535.19 (KHTML, like Gecko) Chrome/18.0.1025.166 Mobile Safari/535.19" }
            chrome_options = Options()
            chrome_options.add_experimental_option("mobileEmulation", mobile_emulation)
            chrome_options.add_argument("user-data-dir=C:/Users/bing/AppData/Local/Google/Chrome/User Data")
            driver = webdriver.Chrome(chrome_options = chrome_options)
            driver.get('http://bing.com')
        time.sleep(3)
        seed = randomNum(size)
        search_box = driver.find_element_by_name('q')
        search_box.send_keys(keyWords[seed])
        search_box.submit()
        #pyautogui.typewrite(keyWords[seed] + '\n', interval=0.1)
        counter += 1
        
        time.sleep(3)
        driver.execute_script("window.scrollTo(0, 200)")
        time.sleep(1.5)
        driver.execute_script("window.scrollTo(0, 150)")
        time.sleep(2)
        driver.execute_script("window.scrollTo(0, 300)")
        time.sleep(2)
        driver.execute_script("window.scrollTo(0, 150)")
        pyautogui.mouseDown(x=339, y=314, button='left')
        pyautogui.mouseUp(x=339, y=314, button='left')
        time.sleep(2)
        pyautogui.mouseDown(x=339, y=314, button='left')
        pyautogui.mouseUp(x=312, y=336, button='left')
        time.sleep(3)
        driver.quit()
        temp2 = randint(10,25)
        timeSpent = timeSpent + temp2
        print('Mobile Searched '+ str(counter)+ ' out of ' + str(range1) + ': ' + str(keyWords[seed]))
        print('Waiting for: '+ str(temp2) +' Seconds' )
        time.sleep(temp2)
        
        print('Total Words Seached in This Mobile Session:' + str(counter))
        
        if(counter == 0):
            timeSpent = 0
        minutes = int(timeSpent/60)
        seconds = timeSpent % 60
        print('Total Time Spent is ' + str(minutes) + ' minute(s) ' + str(seconds) + " second(s).")       
    return counter

#shutdown the machine
def shutdown(shutdown):
    if(shutdown == "True"):
        print('No more task, shutting down in 5 seconds!')
        time.sleep(5)
        os.system("shutdown.bat")
    else:
        print('No more task, testing has been completed, wait 5s!')
        time.sleep(5)

if __name__ == "__main__":
    profile = get_profile()
    Account = profile[0]
    VM = profile[1].split("=")[1]
    Host = profile[2].split("=")[1]
    Report = profile[3].split("=")[1]
    PCSeach = int(profile[4].split("=")[1])
    MobileSearch = int(profile[5].split("=")[1])
    with open("C:\\Users\\bing\\Desktop\\Bing2.0\\data\\shutdown.dat", 'r') as pf:
            for line in pf:
                 profile.append(line.strip())
    Shutdown = profile[7]
    keyWords = keywords()
    size = len(keyWords)
    global keyWords
    global size
    print('Initialize Human Like Search Sequence:')
    presearch_credits = get_credits()
    postsearch_credits = 0
    gain = 0
    get_credit_failed = False
    if(isinstance(presearch_credits,int) == False):
        get_credit_failed = True
    search(PCSeach)
    mobile_search(MobileSearch)
    if(get_credit_failed == False):
        pass
        postsearch_credits = get_credits()
        gain = postsearch_credits - presearch_credits
    else:
        postsearch_credits = "Failed to get credits"
        gain = "Failed to get credits"
    user, pwd = getAccount()
    subject = Account + ' on '+ Host + ' ' + VM +' gained: ' + str(gain) + ' credits.' 
    body = (Account +' currently has: ' + str(postsearch_credits)) + ' credits!'
    send_email(user, pwd, Report, subject, body)
    shutdown(Shutdown)

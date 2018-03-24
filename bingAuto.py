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
from selenium.webdriver.common.alert import Alert
import smtplib
import re
import datetime
from unipath import Path
import sys

current_user = os.getlogin()
current_working_dir, filename = os.path.split(os.path.abspath(__file__))
home = Path(current_working_dir).parent
sys.path.append(home)
os.chdir(current_working_dir)
def get_credits():
    try:
        chrome_options1 = Options()
        chrome_options1.add_argument("user-data-dir=C:/Users/"+current_user+"/AppData/Local/Google/Chrome/User Data")
        chrome_options1.add_argument(current_working_dir)
        chrome_options1.add_argument("--log-level=3")
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
        notify("failed to get credit after 20 seconds with error: " + str(E))

def get_profile():
    profile = []
    with open(home + "\\data\\profile.dat", 'r') as pf:
            for line in pf:
                 profile.append(line.strip())
    with open(home+"\\data\\shutdown.dat", 'r') as pf:
            for line in pf:
                 profile.append(line.strip())
    return profile

def getAccount():
    profile = []
    with open(home+"\\data\\reportEngine.dat", 'r') as pf:
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
    
    with open(home + '\\data\\keyWords_Backup.dat', 'r') as f:
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
    edge_path = current_working_dir+ "\\MicrosoftWebDriver.exe"
    
    for x in range(0, range1):
        try:
            driver = webdriver.Edge(edge_path)
            driver.get('http://bing.com')
        except(KeyboardInterrupt, SystemExit):
            raise
        except Exception as E:
            print("edge start failed, killing application " + str(E))
            notify("PC search failed with error "+ str(E))
            #driver.quit()
            print("restarting search process")
            driver = webdriver.Edge(edge_path)
            driver.get('http://bing.com')
        time.sleep(3)
        seed = randomNum(size)
        pyautogui.typewrite(keyWords[seed] + '\n', interval=0.1)
        counter += 1
        time.sleep(3)
        try:
            driver.switchTo().alert().dismiss();
        except:
            pass
        driver.execute_script("window.scrollTo(0, 200)")
        time.sleep(1.5)
        try:
            driver.switchTo().alert().dismiss();
        except:
            pass
        driver.execute_script("window.scrollTo(0, 150)")
        time.sleep(2)
        try:
            driver.switchTo().alert().dismiss();
        except:
            pass
        driver.execute_script("window.scrollTo(0, 300)")
        time.sleep(2)
        try:
            driver.switchTo().alert().dismiss();
        except:
            pass
        driver.execute_script("window.scrollTo(0, 150)")
        pyautogui.mouseDown(x=339, y=314, button='left')
        pyautogui.mouseUp(x=339, y=314, button='left')
        time.sleep(2)
        pyautogui.mouseDown(x=339, y=314, button='left')
        pyautogui.mouseUp(x=312, y=336, button='left')
        time.sleep(3)
        driver.quit()
        temp2 = randint(3,5)
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
            chrome_options.add_argument("user-data-dir=C:/Users/"+current_user+"/AppData/Local/Google/Chrome/User Data")
            chrome_options.add_argument(current_working_dir)
            chrome_options.add_argument("--log-level=3")
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
            try:
                driver.switchTo().alert().dismiss();
            except:
                pass
            driver.execute_script("window.scrollTo(0, 200)")
            time.sleep(1.5)
            try:
                driver.switchTo().alert().dismiss();
            except:
                pass
            driver.execute_script("window.scrollTo(0, 150)")
            time.sleep(2)
            try:
                driver.switchTo().alert().dismiss();
            except:
                pass
            driver.execute_script("window.scrollTo(0, 300)")
            time.sleep(2)
            try:
                driver.switchTo().alert().dismiss();
            except:
                pass
            driver.execute_script("window.scrollTo(0, 150)")
            try:
                driver.switchTo().alert().dismiss();
            except:
                pass
            pyautogui.mouseDown(x=339, y=314, button='left')
            pyautogui.mouseUp(x=339, y=314, button='left')
            time.sleep(2)
            pyautogui.mouseDown(x=339, y=314, button='left')
            pyautogui.mouseUp(x=312, y=336, button='left')
            time.sleep(3)
            driver.quit()
            temp2 = randint(3,5)
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
        #except(KeyboardInterrupt, SystemExit):
        #    raise
        except Exception as E:
            notify("Mobile search failed with error "+ str(E))
            driver.quit()    
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
        
def get_progress_legacy():
    try:
        chrome_options1 = Options()
        chrome_options1.add_argument("user-data-dir=C:/Users/"+current_user+"/AppData/Local/Google/Chrome/User Data")
        chrome_options1.add_argument(current_working_dir)
        driver1 = webdriver.Chrome(chrome_options = chrome_options1)
        driver1.get('https://www.bing.com/')
        time.sleep(4)
        driver1.find_element_by_id("id_rh").click()
        time.sleep(4)
        driver1.switch_to_frame("bepfm")
        data =  driver1.find_elements_by_class_name("breakdown")
        timeout = 1
        while(data == None):
            print("Failed to get progress - retrying up to 20 times!")
            print("Try: " + str(timeout))
            time.sleep(1)
            timeout = timeout + 1 
            data =  driver1.find_elements_by_class_name("breakdown")
            if((timeout == 20) and (data == None)):
                driver1.quit()
                return "failed","failed"
                break
        counter = 5
        while(counter > 0):
            if(len(data) > 0):
                
                try:
                    PC_SEARCH = (data[0].text).split("\n")[1][12:]
                    MAX_PC = int(PC_SEARCH.split("/")[1])
                    CUR_PC = int(PC_SEARCH.split("/")[0])
                except Exception as E:
                    notify(str(E))
                    PC_SEARCH = "failed"
                
                try:
                    MOBILE_SEARCH = (data[0].text).split("\n")[2][9:]
                    MAX_MOBILE = int(MOBILE_SEARCH.split("/")[1])
                    CUR_MOBILE = int(MOBILE_SEARCH.split("/")[0])
                except Exception as E:
                    notify(str(E))
                    MOBILE_SEARCH = "failed"
                driver1.quit()
                return PC_SEARCH,MOBILE_SEARCH
                break
            else:
                counter = counter - 1
                time.sleep(1)
                data =  driver1.find_elements_by_class_name("breakdown")
        driver1.quit()
        return "failed","failed"
    except Exception as E:
        print("failed to get current progress: " + str(E))
        notify("Failed to get current progress: " + str(E))
        driver1.quit()
        return "failed","failed"

def fortune():
    quotes = []
    with open("fortune_database2.txt", 'r') as pf:
        for line in pf:
            quotes.append(line.strip())
    seed = randomNum(len(quotes)-1)
    return quotes[seed]    

def notify(error):
    current_working_dir, filename = os.path.split(os.path.abspath(__file__))
    profile = get_profile()
    user, pwd = getAccount()
    Account = profile[0]
    VM = profile[1].split("=")[1]
    Host = profile[2].split("=")[1]
    Report = profile[3].split("=")[1]
    PCSeach = int(profile[4].split("=")[1])
    MobileSearch = int(profile[5].split("=")[1])
    SMSemail = Report.split("@")[1]
    subject = "Critical error on: " + Host + " " + VM + " " + filename
    body = "The following error has occured: " + error
    send_email(user, pwd, Report, subject, body)

def isInt(var):
    try:
        int(var)
        return True
    except:
        return False
    
def advanced_progress():
    try:
        report = []
        chrome_options1 = Options()
        chrome_options1.add_argument("user-data-dir=C:/Users/"+current_user+"/AppData/Local/Google/Chrome/User Data")
        chrome_options1.add_argument(current_working_dir)
        driver1 = webdriver.Chrome(chrome_options = chrome_options1)
        driver1.get('https://account.microsoft.com/rewards/pointsbreakdown')
        time.sleep(4)
        data =  driver1.find_elements_by_class_name("title-detail")
        timeout = 1
        while(data == None):
            print("Failed to get progress - retrying up to 20 times!")
            print("Try: " + str(timeout))
            time.sleep(1)
            timeout = timeout + 1 
            data =  driver1.find_elements_by_class_name("title-detail")
            if((timeout == 20) and (data == None)):
                driver1.quit()
                return "failed"
                break
        time.sleep(1)
        for item in data:
            report.append((item.text).split("\n"))
        driver1.quit()
        return report
    except Exception as E:
        print("failed to get current progress: " + str(E))
        notify("Advanced progress retrival failed with error: " + str(E))
        driver1.quit()
        return "failed"
    
def processReport(var,report):
    for item in report:
        for leaf in item:
            if(var == leaf):
                return item
    return None

def processLeaf(var,leaf):
    if(var == "Available points"):
        return leaf[0]
    elif(var == "Streak count"):
        return leaf[0]
    elif(var == "Microsoft Edge bonus"):
        return leaf[1].split(" / ")
    elif(var == "PC search"):
        return leaf[1].split(" / ")
    elif(var == "Mobile search"):
        return leaf[1].split(" / ")
    elif(var == "Shop & earn"):
        return leaf[1]
    elif(var == "Other activities"):
        return leaf[1].split(" / ")
    else:
        return None
def progressCheck(progress):
    for item in progress:
        if(isInt(item) == False):
            return False

if __name__ == "__main__":
    #begins advanced report setup
    PC_search = []
    Mobile_search = []
    Microsoft_Edge_bonus = []
    Other_activities = []
    report = advanced_progress()
    advanced_available_points = ""
    if(report != "failed"):
        advanced_available_points = ""
        temp = processReport("PC search",report)
        if(temp != None):
            PC_search = processLeaf("PC search",temp)
        else:
            PC_search = None
        if(progressCheck(PC_search) == False):
            PC_search = None
           
        temp = processReport("Mobile search",report)
        if(temp != None):
            Mobile_search = processLeaf("Mobile search",temp)
        else:
            Mobile_search = None  
        if(progressCheck(Mobile_search) == False):
            Mobile_search = None
            
        temp = processReport("Microsoft Edge bonus",report)
        if(temp != None):
            Microsoft_Edge_bonus = processLeaf("Microsoft Edge bonus",temp)
        else:
            Microsoft_Edge_bonus = None
        if(progressCheck(Microsoft_Edge_bonus) == False):
            Microsoft_Edge_bonus = None
            
        temp = processReport("Available points",report)
        if(temp != None):
            advanced_available_points = processLeaf("Available points",temp)
        else:
            advanced_available_points = None
        if(isInt(advanced_available_points) == False):
            advanced_available_points = None
        else:
            advanced_available_points = int(advanced_available_points)

        temp = processReport("Other activities",report)
        if(temp != None):
            Other_activities = processLeaf("Other activities",temp)
        else:
            Other_activities = None
        if(progressCheck(Other_activities) == False):
            Other_activities = None
        
    else:
        PC_search = None
        Mobile_search = None
        Microsoft_Edge_bonus = None
        advanced_available_points = None
        Other_activities = None
    
    total_legacy_search = 0
    total_adaptive_search = 0
    time1 = datetime.datetime.now()
    profile = get_profile()
    Account = profile[0]
    VM = profile[1].split("=")[1]
    Host = profile[2].split("=")[1]
    Report = profile[3].split("=")[1]
    PCSeach = int(profile[4].split("=")[1])
    MobileSearch = int(profile[5].split("=")[1])
    SMSemail = Report.split("@")[1]
    with open(home+"\\data\\shutdown.dat", 'r') as pf:
            for line in pf:
                 profile.append(line.strip())
    Shutdown = profile[7]
    global keyWords
    global size
    keyWords = keywords()
    size = len(keyWords)
    print('Initialize Human Like Search Sequence:')
    #get presearch credits
    if(advanced_available_points == None):
        presearch_credits = get_credits()
    else:
        presearch_credits = advanced_available_points
    postsearch_credits = 0
    gain = 0
    get_credit_failed = False
    if(isInt(presearch_credits) == False):
        get_credit_failed = True
    
    if(PC_search == None):
        #check for pc search stat
        PC_SEARCH,MOBILE_SEARCH = get_progress_legacy()
        if(PC_SEARCH != "failed"):
            MAX_PC = int(PC_SEARCH.split("/")[1])
            CUR_PC = int(PC_SEARCH.split("/")[0])
            while(CUR_PC < MAX_PC):
                print("Current PC Search Progress: "+ str(PC_SEARCH))
                diff = int((MAX_PC - CUR_PC)/5)
                total_adaptive_search = total_adaptive_search + diff + 1
                print("Making " + str(diff+1) + " additional searches!")
                search(diff+1)
                PC_SEARCH,MOBILE_SEARCH = get_progress_legacy()
                MAX_PC = int(PC_SEARCH.split("/")[1])
                CUR_PC = int(PC_SEARCH.split("/")[0])
        else:
            #legacy search if adaptive search failed
            notify("PC adaptive search failed, begins legacy search with " + str(PCSeach) + " searches.")
            total_legacy_search = PCSeach + total_legacy_search
            search(PCSeach)
    else:
        try:
            #do adaptive search with advanced progress
            MAX_PC = int(PC_search[1])
            CUR_PC = int(PC_search[0])
            while(CUR_PC < MAX_PC):
                    print("Current PC Search Progress: "+ str(CUR_PC))
                    diff = int((MAX_PC - CUR_PC)/5)
                    total_adaptive_search = total_adaptive_search + diff + 1
                    print("Making " + str(diff+1) + " additional searches!")
                    search(diff+1)
                    report = advanced_progress()
                    temp = processReport("PC search",report)
                    PC_search = processLeaf("PC search",temp)
                    MAX_PC = int(PC_search[1])
                    CUR_PC = int(PC_search[0])
        except Exception as E:
            print("Adaptive PC search with advanced report failed with error: " + str(E))
            notify("PC adaptive search failed, begins legacy search with " + str(PCSeach) + " searches.")
            total_legacy_search = PCSeach + total_legacy_search
            search(PCSeach)

    if(Mobile_search == None):
        #check for mobile search stat
        PC_SEARCH,MOBILE_SEARCH = get_progress_legacy()
        if(MOBILE_SEARCH != "failed"):
            MAX_MOBILE = int(MOBILE_SEARCH.split("/")[1])
            CUR_MOBILE = int(MOBILE_SEARCH.split("/")[0])

            while(CUR_MOBILE < MAX_MOBILE):
                print("Current Mobile Search Progress: "+ str(MOBILE_SEARCH))
                diff = int((MAX_MOBILE - CUR_MOBILE)/5)
                total_adaptive_search = total_adaptive_search + diff + 1
                print("Making " + str(diff+1) + " additional searches!")
                mobile_search(diff+1)
                PC_SEARCH,MOBILE_SEARCH = get_progress_legacy()
                MAX_MOBILE = int(MOBILE_SEARCH.split("/")[1])
                CUR_MOBILE = int(MOBILE_SEARCH.split("/")[0])
        else:
            #legacy search if adaptive search failed
            notify("Mobile adaptive search failed, begins legacy search with " + str(MobileSearch) + " searches.")
            total_legacy_search = MobileSearch + total_legacy_search
            mobile_search(MobileSearch)
    else:
        #do adaptive search with advanced progress
        try:
            #do adaptive search with advanced progress
            MAX_MOBILE = int(Mobile_search[1])
            CUR_MOBILE = int(Mobile_search[0])
            while(CUR_MOBILE < MAX_MOBILE):
                    print("Current Mobile Search Progress: "+ str(CUR_MOBILE))
                    diff = int((MAX_MOBILE - CUR_MOBILE)/5)
                    total_adaptive_search = total_adaptive_search + diff + 1
                    print("Making " + str(diff+1) + " additional searches!")
                    mobile_search(diff+1)
                    report = advanced_progress()
                    temp = processReport("Mobile search",report)
                    Mobile_search = processLeaf("Mobile search",temp)
                    MAX_MOBILE = int(Mobile_search[1])
                    CUR_MOBILE = int(Mobile_search[0])
        except Exception as E:
            print("Adaptive Mobile search with advanced report failed with error: " + str(E))
            notify("Mobile adaptive search failed, begins legacy search with " + str(MobileSearch) + " searches.")
            total_legacy_search = MobileSearch + total_legacy_search
            mobile_search(MobileSearch)
    
    if(get_credit_failed == False):
        postsearch_credits = get_credits()
        gain = postsearch_credits - presearch_credits
    else:
        postsearch_credits = "Failed to get credits"
        gain = "Failed to get credits"
    if(PC_search != None):
        PC_SEARCH = PC_search[0] + "/"+ PC_search[1]
    else:
        PC_SEARCH = "Failed"

    if(Mobile_search != None):
        MOBILE_SEARCH = Mobile_search[0] + "/" + Mobile_search[1]
    else:
        MOBILE_SEARCH = "Failed"

    if(Microsoft_Edge_bonus != None):
        EdgeStat = Microsoft_Edge_bonus[0] + "/" + Microsoft_Edge_bonus[1]
    else:
        EdgeStat = "Failed"
    if(Other_activities != None):
        others = Other_activities[0] + "/" + Other_activities[1]
    else:
        others = "Failed"
    time2 = datetime.datetime.now()
    timeDiff = time2 - time1
    user, pwd = getAccount()
    if(SMSemail == 'tmomail.net'):
        subject = 'Gained: ' + str(gain) + ' credits.'
        body = (Account +' currently has: ' + str(postsearch_credits)) + ' credits!'
        send_email(user, pwd, Report, subject, body)
        subject = "Total time spent: " + str(timeDiff)[:10]
        body = ' '
        send_email(user, pwd, Report, subject, body)
        subject ="Mobile Progress: " + MOBILE_SEARCH
        body = "PC Progress: " + PC_SEARCH
        send_email(user, pwd, Report, subject, body)
    else:
        subject = Account + ' on '+ Host + ' ' + VM +' gained: ' + str(gain) + ' credits.'
        body = ((Account +' currently has: ' + str(postsearch_credits)) + ' credits!' + "\n" +
                "Total time spent: " + str(timeDiff)[:10]
                + "\n" + "PC Progress: " + PC_SEARCH + "\n" + "Mobile Progress: " +
                MOBILE_SEARCH + "\n" + "Edge Bonus: "+ str(EdgeStat)+ "\n"+ "Other activities: "+ others + "\n" +
                "Total adaptive searches done: " + str(total_adaptive_search)+ "\n"+
                "Total legacy seaches done: "+ str(total_legacy_search)+ "\n" +fortune())
        send_email(user, pwd, Report, subject, body)

    shutdown(Shutdown)

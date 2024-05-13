import os
import time
import datetime
import shutil
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service


# --------------------------------------User Info---------------------------------------- #

user = 'MUNEEB'


def CreateFolder(path, name):

    os.chdir(path)
    os.mkdir(name)


def MoveFiles(dPath, rPath, aPath, currFolder, files):

    # Supplemental File

    while True:

        path1 = dPath + files[0]
        path2 = aPath + currFolder + '\\' + files[0]

        if os.path.exists(path1):

            shutil.move(path1, path2)

            break
        
        else:

            time.sleep(15)
    
    # PA2 File
    
    while True:

        path1 = rPath + files[1]
        path2 = aPath + currFolder + '\\' + files[1]

        if os.path.exists(path1):

            shutil.copy(path1, path2)

            break
        
        else:

            time.sleep(15)

    # Recap File
    
    while True:

        path1 = rPath + files[2]
        path2 = aPath + currFolder + '\\' + files[2]

        if os.path.exists(path1):

            shutil.copy(path1, path2)

            break
        
        else:

            time.sleep(15)


# ------------------------------------Current date--------------------------------------- #


currDate = datetime.date.today()
currFolder = currDate.strftime("%m-%d-%Y")
yyyymmdd = currDate.strftime("%Y%m%d")


# -------------------------------------File Names---------------------------------------- #


suppFile = 'ResurgentSupplement.csv'
PA2File = 'PA2_8357_' + yyyymmdd + '.txt'
recapFile = 'RECAP_8357_' + yyyymmdd + '.xlsx'

# List of File Names

fileNames = [suppFile, PA2File, recapFile]


# ----------------------------------Directories Path------------------------------------- #


# Path to downloads folder

downPath = f'C:\\Users\\{user}\\Downloads\\'

# Path to reports folder

reportsPath = 'Y:\\'

# Path to Resurgent's remit folder

adminPath = 'N:\\!Daily Operations\\Client Related\\Resurgent\\Remits\\'


# ----------------Creates instance of ChromeDriver-------------------- #


# Path for Chrome

chromePath = f'{os.getcwd()}\\chromedriver.exe'

chromeService = Service(chromePath)

# Creates a driver instance using Service

driver = webdriver.Chrome(service=chromeService)


# ----------------------Custom Dates & Monday------------------------- #


currDate = datetime.date.today()

if currDate.strftime("%w") == '1': 
    
    monday = True
    
    friday = currDate - datetime.timedelta(days=3)
    sunday = currDate - datetime.timedelta(days=1)
    
    # Converts datetime objects in strings to send as input keys
    
    inpFriday = friday.strftime("%Y-%m-%d")
    inpSunday = sunday.strftime("%Y-%m-%d")
    
else: monday = False


# -----------------------USAePay Login Page-------------------------- #


driver.get('https://secure.usaepay.com/login')

time.sleep(10)

# xPaths for email input
# Password input box &
# Login button

xpathEmail1 = "/html/body/div[3]/div[2]/div/form/div/div[1]/div/input"
xpathPass1 = "/html/body/div[3]/div[2]/div/form/div/div[2]/div/input"
xpathLogin1 = "/html/body/div[3]/div[2]/div/form/div/button"

# Access to Email and password division

inpEmail1 = driver.find_element(By.XPATH, xpathEmail1)
inpPassword1 = driver.find_element(By.XPATH, xpathPass1)

# Sends email and password credentials.

inpEmail1.send_keys("arehman")
inpPassword1.send_keys("Unifin:12345")

# Login button

login1 = driver.find_element(By.XPATH, xpathLogin1)
login1.click()


# ----------------Prompts User for custom dates input---------------- #


while True:
    
    inp = input('\n' + 'Do you want to run report on custom dates (Y/N): ')
    
    if inp == 'Y' or inp == 'N': 
        
        if inp == 'Y': custom = True
        else: custom = False
            
        break
        
    else: print('\n' + 'Invalid Input')


# ------------------------ePay Reports Page-------------------------- #


driver.get('https://secure.usaepay.com/console/reports/')

time.sleep(5)

# XPath to Resurgent Supplement class

xpathResSupplement = "/html/body/center[1]/table/tbody/tr/td/center[2]/table[4]" \
                     "/tbody/tr[2]/td[1]/table/tbody/tr[1]/td[2]/a[1]"

resSupplement = driver.find_element(By.XPATH, xpathResSupplement)

# Clicks on Resurgent Supplement link

resSupplement.click()

time.sleep(5)

# XPath to export button

xpathExport = "/html/body/center[1]/div/form/div/table/tbody/tr[3]/td/input[4]"

exportReport = driver.find_element(By.XPATH, xpathExport)

# Clicks on export button

exportReport.click()

time.sleep(5)


# -----------------------Uportal Login Page-------------------------- #


driver.get('http://uportal.unifininc.com/admin/login')

time.sleep(5)

# xPaths for email input
# Password input box &
# Login button

xpathEmail2 = "/html/body/div/div[2]/div/div/div/div/form/div[1]/input"
xpathPass2 = "/html/body/div/div[2]/div/div/div/div/form/div[2]/input"
xpathLogin2 = "/html/body/div/div[2]/div/div/div/div/form/div[4]/div[1]/button"

# Access to Email and password division

inpEmail2 = driver.find_element(By.XPATH, xpathEmail2)
inpPassword2 = driver.find_element(By.XPATH, xpathPass2)

# Sends email and password credentials.

inpEmail2.send_keys("muneeb1")
inpPassword2.send_keys("Unifin@112233##")

# Login button

login2 = driver.find_element(By.XPATH, xpathLogin2)
login2.click()

time.sleep(10)


# ----------------Gets custom dates input from user----------------- #


if custom:

    startingDate = input('\n' + 'Enter the beginning date in format yyyy-mm-dd: ')
    endingDate = input('\n' + 'Enter the closing date in format yyyy-mm-dd: ')


# ---------------------Client Reporting Page------------------------ #


driver.get('http://uportal.unifininc.com/admin/clientreports')

time.sleep(10)

# xPaths for Daily Report button, Resurgent Remit Checkbox
# & Resurgent Recap check box & Run report button

xpathDaily = "/html/body/div/div/main/div/form/div[3]/div/button[4]"
xpathResRemit = "/html/body/div/div/main/div/form/div[2]/table/tbody/tr[49]/td[6]/input"
xpathResRecap = "/html/body/div/div/main/div/form/div[2]/table/tbody/tr[50]/td[6]/input"
xpathRunReport = "/html/body/div/div/main/div/form/div[3]/button"
xpathStartDate = "/html/body/div/div/main/div/form/div[1]/input[3]"
xpathEndDate = "/html/body/div/div/main/div/form/div[1]/input[4]"

if not custom and not monday:

    # Daily Report button

    dailyRep = driver.find_element(By.XPATH, xpathDaily)
    dailyRep.click()

if not custom and monday:
    
    # Dates to be entered are friday to sunday
    # Finds input elements of start and end date
    # send friday and sunday as input respectively
    
    startDate = driver.find_element(By.XPATH, xpathStartDate)
    endDate = driver.find_element(By.XPATH, xpathEndDate)
    
    startDate.send_keys(inpFriday)
    endDate.send_keys(inpSunday)

if custom:

    # Finds input elements of start and end date
    # inputs custom date given by user

    startDate = driver.find_element(By.XPATH, xpathStartDate)
    endDate = driver.find_element(By.XPATH, xpathEndDate)

    startDate.send_keys(startingDate)
    endDate.send_keys(endingDate)
    
# Inputs Resurgent Remit & Resurgent Recap checkbox
# Clicks the Run button after 2 seconds pause

checkRemit = driver.find_element(By.XPATH, xpathResRemit)
checkRemit.click()

checkRecap = driver.find_element(By.XPATH, xpathResRecap)
checkRecap.click()

time.sleep(2)

runReport = driver.find_element(By.XPATH, xpathRunReport)
runReport.click()


# ----------------------------Creates Folder & Moves Files------------------------------- #


# CreateFolder(adminPath, currFolder)
# MoveFiles(downPath, reportsPath, adminPath, currFolder, fileNames)


# ------------------------------------Closing Edge--------------------------------------- #


print('\n' + '-----------------------------------------------------------' + '\n')

while True:

    inp = input('Want to continue? ')

    if inp == 'Y' or inp == 'y':
        break

    else:
        print('\n' + 'What do you mean bro?' + '\n')

# Close the browser

driver.quit()
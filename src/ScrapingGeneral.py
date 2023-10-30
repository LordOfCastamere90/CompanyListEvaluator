from selenium import webdriver
from time import sleep
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.remote.webelement import WebElement
from selenium_recaptcha_solver import RecaptchaSolver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.actions.wheel_input import ScrollOrigin
from bs4 import BeautifulSoup
import os
import openai
from dotenv import load_dotenv
from idlelib.query import Goto
from openpyxl import load_workbook

excelSheetName = "IBA.xlsx"
wb = load_workbook(excelSheetName)
sheet = wb.active



service = Service()
options = webdriver.ChromeOptions()
options.add_experimental_option('detach', True)
options.add_argument("--start-maximized")
driver = webdriver.Chrome(service=service, options=options)
actions = ActionChains(driver)

driver.get('chrome://settings/appearance')
sleep(2)
driver.execute_script('chrome.settingsPrivate.setDefaultZoom(0.70);')
sleep(2)

#Variabeln:
website = "https://clusterix.io/companies"
emailadresse = "schnittger@innoscripta.com"
passwort = "ehz6afn"
# Begin with row (Normaly 1):
rowNumber = 297
ZeilenZahl = rowNumber + 1



#Variabeln:
website = "https://clusterix.io/companies"
driver.get(website)
sleep(10)

'''
#Cookieeinstellungen:
shadow_parent = driver.find_element(By.CSS_SELECTOR, '#usercentrics-root')
outer = driver.execute_script('return arguments[0].shadowRoot', shadow_parent)
outer.find_element(By.CSS_SELECTOR, "button[data-testid='uc-accept-all-button']").click()
'''
email = driver.find_element(By.NAME,"email")
email.send_keys(emailadresse)

password = driver.find_element(By.NAME,"password")
password.send_keys(passwort)

anmelden = driver.find_element(By.XPATH, '//*[@id="login-form"]/div[3]/div[3]/button')
anmelden.click()

sleep(10)

#reading specific column 
columns = sheet["A"]

for count, data in enumerate(columns[rowNumber:],ZeilenZahl):
    print(count)
    print(data.value)
    companyField = driver.find_element(By.XPATH,'//*[@id="main_content"]/div/div[2]/div[1]/div/div/div[1]/div[1]/div[2]/div/div/input')
    companyField.send_keys(Keys.CONTROL, "a")
    companyField.send_keys(Keys.DELETE)
    companyField.send_keys(data.value)
    searchField = driver.find_element(By.XPATH, '//*[@id="main_content"]/div/div[2]/div[1]/div/div/div[1]/div[2]/div')
    searchField.click()
    sleep(2)
    companyName = driver.find_elements(By.CLASS_NAME,'CompanyBox_company-name__dUWw0')
    companySize = driver.find_elements(By.CLASS_NAME,'CompanyBox_number-of-employee__7umhd')
    kommentarElements = driver.find_elements(By.CSS_SELECTOR,'div.CompanyBox_comments__NJXr\+.CompanyBox_part__RF7W2.CompanyBox_button__9y6iA')
    
    # Name from List in 2 versions (AG and Aktiengesellschaft etc.)
    nameFromList = (data.value).lower()
    nameIntoList = nameFromList.split()
    nameFromListV2 = (data.value).lower()
    
    if "gmbh" in nameFromList:
        nameFromListV2 = nameFromList.replace("gmbh", "gesellschaft mit beschränkter haftung")

    for x in range(len(nameIntoList)):    
        if nameIntoList[x] == "ag":
            nameIntoList[x] = "aktiengesellschaft"
            nameFromListV2 = ' '.join(nameIntoList)

    for x in range(len(nameIntoList)):    
        if nameIntoList[x] == "kg":
            nameIntoList[x] = "kommanditgesellschaft"
            nameFromListV2 = ' '.join(nameIntoList)

                      
    kommNumberTemp = 0
    companySizeClean = 0
    get_url = ''
    
    for count2, comp in enumerate(companyName):
            
        nameFromClx = (comp.text).lower()
                   
        if nameFromList in nameFromClx or nameFromListV2 in nameFromClx:
            
            #driver.execute_script("arguments[0].scrollIntoView();", comp)
            #actions.move_to_element(comp).perform()
            
            scroll_origin = ScrollOrigin.from_viewport(10, 10)
            ActionChains(driver)\
            .scroll_from_origin(scroll_origin, 0, -500)\
            .perform()
            
            '''
            scroll_origin = ScrollOrigin.from_element(comp)
            ActionChains(driver)\
            .scroll_from_origin(scroll_origin, 0, 300)\
            .perform()
            '''
            # Kommentaranzahl entscheidet über relevantesten Firmeneintrag in CLX            
            kommElement = kommentarElements[count2]
            html = kommElement.get_attribute('innerHTML')
            soup = BeautifulSoup(html, 'html.parser')
            span_text = soup.find('span').text
            
            if span_text == '-':
                continue
            kommNumber = span_text.replace('+','')
            kommNumber = kommNumber.replace('-','')
            
            if kommNumber == '':
                kommNumber = 0
            if '@' in str(kommNumber):
                kommNumber = 10

            kommNumber = int(kommNumber)
            
            
            if kommNumber >= kommNumberTemp:
                
                WebDriverWait(driver, 3).until(EC.element_to_be_clickable(kommentarElements[count2])).click()
                sleep(2)
                #Relevante Daten:
                get_url = driver.current_url
                companySizeClean = str(companySize[count2].text).replace('(','')
                companySizeClean = companySizeClean.replace(')','')
                if companySizeClean == '':
                    companySizeClean = 0
                companySizeClean = str(companySizeClean).replace('.','')
                sizeInt = int(companySizeClean)
                kommNumberTemp = kommNumber
            
                driver.find_element(By.XPATH, '//*[@id="main_content"]/div/div[4]/div[2]/div/div[1]/div[1]/div[2]/div[2]/div[1]').click()
                sleep(1) 
            
            
    columns = sheet["F"+ str(count)].value=companySizeClean
    columns = sheet["G"+ str(count)].value=str(get_url)
    wb.save(excelSheetName)


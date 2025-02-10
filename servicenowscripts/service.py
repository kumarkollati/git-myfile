import os

from openpyxl.workbook import Workbook
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support import expected_conditions as EC
import time
from selenium.webdriver.support.ui import Select
import smtplib
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.keys import Keys
import re
from selenium import webdriver
import pandas as pd

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
def read_values_from_file(filename):

    current_dir = os.path.dirname(os.path.abspath(__file__))

    filepath = os.path.join(current_dir, filename)

    values_dict = {}
    values_list = []

    with open(filepath, 'r') as file:

        for line in file:

            variable_name, value = line.strip().split('=')  # Assuming variables are separated by "="

            variable_name = variable_name.strip()

            value = value.strip()

            if variable_name == "eprids":

                value = value.split(',')  # Assuming search_ids are comma-separated

            values_dict[variable_name] = value
            values_list = value


    return values_list

# Read values from the file

eprids = read_values_from_file("input_file.txt")
print(eprids)
output = []
userid = "kumar.kollati-wipro@hpe.com"
passWord = ""
print("Driver Path intialised")
# driver_path = ChromeDriverManager().install()
service = Service(executable_path='D:/ChromeDriver/chromedriver.exe')
# service = Service(executable_path='C:/chromedriver/chromedriver_win32/chromedriver.exe')
chrome_options = Options()
print("Driver Path intialised")
chrome_options.add_argument('--ignore-certificate-errors')
print("Driver Path intialised")
chrome_options.add_argument('--ignore-ssl-errors')
print("Driver Path intialised")
driver = webdriver.Chrome(service=service, options=chrome_options)

# driver = webdriver.Chrome(service=Service(driver_path))
print("Driver Path intialised")
driver.maximize_window()
driver.get(r'https://hpe.service-now.com/now/nav/ui/classic/params/target/incident_list.do%3Fsysparm_query%3Dopened_at%253E%253Djavascript%3Ags.beginningOfLast6Months()%255Eu_eprid_alertSTARTSWITH%26sysparm_first_row%3D1%26sysparm_view%3Dtext_search')
print("url got hit")
time.sleep(4)
email = driver.find_element(by=By.ID, value='input27')
print("username entered")
email.send_keys(userid)
nextbutton = driver.find_element(by=By.CLASS_NAME, value='button.button-primary')
nextbutton.click()
time.sleep(5)
print("Button Clicked")
print("Sign in with Password or Okta Verify")

element1 = driver.find_element(By.CSS_SELECTOR, "div.authenticator-verify-list.authenticator-list > div > div:nth-child(2) > div.authenticator-description > div.authenticator-button > a")
element1.click()
time.sleep(20)
driver.implicitly_wait(15)
time.sleep(80)

macro_component_container = driver.find_element(by=By.XPATH,value='/html/body/macroponent-f51912f4c700201072b211d4d8c26010')
shadow_root = driver.execute_script('return arguments[0].shadowRoot', macro_component_container)
time.sleep(3)
frame = shadow_root.find_element(by=By.CSS_SELECTOR,value='#gsft_main')
driver.switch_to.frame(frame)

for eprid in eprids:
    element1 = driver.find_element(By.CLASS_NAME, value='list_filter_toggle.icon-filter.btn.btn-icon')
    element1.click()
    print("Filter button clicked")
    time.sleep(5)
    divcontainer = driver.find_element(By.CLASS_NAME, value='filterContainer')
    tableid = driver.find_element(by=By.ID, value='incidentfilters_table')
    element1 = driver.find_element(By.XPATH,"/html/body/div[1]/div[1]/span/div/div[4]/div/div/list_filter/div[2]/table/tbody/tr[2]/td/table/tbody/tr/td/table/tbody/tr[4]/td/table/tbody/tr/td[4]/input")
    #element1 = driver.find_element(By.CLASS_NAME, "filerTableInput.form-control")
    driver.implicitly_wait(5)
    element1.clear()
    element1.send_keys(eprid)
    time.sleep(2)
    element1 = driver.find_element(by=By.ID, value='test_filter_action_toolbar_run')
    element1.click()
    print("Run button clicked")
    time.sleep(40)
    cnt_of_rds = driver.find_element(By.XPATH, '/html/body/div[1]/div[1]/span/div/div[7]/div[2]/table/tbody/tr/td[2]/span[1]/div/span[2]/span[2]').text
    cnt_of_rds = ''.join(letter for letter in cnt_of_rds if letter.isalnum())
    print("Max Count of records: "+cnt_of_rds)
    outcome = {}
    outcome['Eprid'] =eprid
    outcome['TicketCount'] =cnt_of_rds
    output.append(outcome)
    print(output)
df = pd.DataFrame(output)
df.to_csv(r'C:\Users\KO40134144\PythonNew\servicenowscripts\incidentdump.csv',index=False)
driver.quit()
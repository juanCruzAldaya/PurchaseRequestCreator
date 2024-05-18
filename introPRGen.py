import os
import time
import glob
import pandas as pd
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.edge.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.edge.options import Options as EdgeOptions
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from selenium.common.exceptions import NoSuchElementException
import shutil
import openpyxl

def element_exists(driver, xpath):
    try:
        driver.find_element(By.XPATH, xpath)
        return True
    except NoSuchElementException:
        return False


def get_latest_file(folder_path: str) -> str:

    list_of_files = glob.glob(os.path.join(folder_path, '*'))
    latest_file = max(list_of_files, key=os.path.getctime)
    return latest_file







option = EdgeOptions()
option.add_argument("start-maximized")
driver = webdriver.Edge(options = option)

timeOut = 35

driver.get("https://ts.accenture.com/sites/hsalts/Documentation/Forms/AllItems.aspx")

os.system('cls')

driver.execute_script("window.scrollTo(0,document.body.scrollHeight)")
driver.execute_script("window.scrollTo(5,document.body.scrollHeight)")

WebDriverWait(driver, timeOut).until(EC.title_contains('Hispanic'))
#/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[2]/div[20]/div/div/div[2]/div[2]/div/div[1]/div[1]/span/span[2]/button




csvExists = element_exists(driver, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[2]/div[21]/div/div/div[2]/div[2]/div/div[1]/div[1]/span/span[1]/button')


print(csvExists)
if csvExists:
        userField= driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[2]/div[21]/div/div/div[2]/div[2]/div/div[1]/div[1]/span/span[1]/button')
        fieldName = userField.text

        if fieldName == 'PRToDo.csv': #IF EXISTS THE SCRIPT DELETE IT
                print(fieldName)
                try: 
                        WebDriverWait(driver, timeOut).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[2]/div[21]/div/div/div[2]/div[2]')))
                        itemCheck = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[2]/div[21]/div/div/div[2]/div[2]')
                        itemCheck.click()
                except TimeoutException:
                        print ("Snow failed, run script again...")
                time.sleep(2)
                ActionChains(driver).send_keys(Keys.DELETE).perform()

                try: 
                        WebDriverWait(driver, timeOut).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div/div/div/div[2]/div[2]/div/div[2]/div[2]/div/span[1]/button')))
                        confirmDelete = driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div/div[2]/div[2]/div/div[2]/div[2]/div/span[1]/button')
                        confirmDelete.click()
                        time.sleep(4)
                except TimeoutException:
                        print ("Snow failed, run script again...")


original_window = driver.current_window_handle
driver.switch_to.new_window("window")
newWindow = driver.current_window_handle
driver.switch_to.window(original_window)
driver.close()
driver.switch_to.window(newWindow)

driver.get('https://make.powerautomate.com/environments/Default-e0793d39-0939-496d-b129-198edd916feb/flows/shared/3fd38501-5ec6-438c-a436-fc160acd65c2/details')
time.sleep(10)
try: 
        WebDriverWait(driver, timeOut).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div[1]/div/div/div[2]/div[2]/div[2]/div/section/main/div[1]/div/div/div/div/div/div[1]/div[5]/button')))
        runFlow = driver.find_element(By.XPATH, '/html/body/div[1]/div/div[1]/div/div/div[2]/div[2]/div[2]/div/section/main/div[1]/div/div/div/div/div/div[1]/div[5]/button')
        runFlow.click()
except TimeoutException:
        print ("Snow failed, run script again...")

try: 
        WebDriverWait(driver, timeOut).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div/div/div/div/div[2]/div[2]/div/div[4]/div/div[2]/button[1]')))
        runFlow = driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div/div/div[2]/div[2]/div/div[4]/div/div[2]/button[1]')
        runFlow.click()
except TimeoutException:
        print ("Snow failed, run script again...")
WebDriverWait(driver, timeOut).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div[1]/div/div/div/div[2]/div[2]/div/div[4]/div/div/button')))

time.sleep(8)

original_window = driver.current_window_handle
driver.switch_to.new_window("window")
newWindow = driver.current_window_handle
driver.switch_to.window(original_window)
driver.close()
driver.switch_to.window(newWindow)


driver.get("https://ts.accenture.com/sites/hsalts/Documentation/Forms/AllItems.aspx")

os.system('cls')

driver.execute_script("window.scrollTo(5,document.body.scrollHeight)")
driver.execute_script("window.scrollTo(5,document.body.scrollHeight)")

try:

        WebDriverWait(driver, timeOut).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[2]/div[20]/div/div/div[2]/div[2]/div/div[1]/div[1]/span/span[2]/button')))
        userField= driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[2]/div[20]/div/div/div[2]/div[2]/div/div[1]/div[1]/span/span[2]/button')
        fieldName = userField.text

        if fieldName == 'PRToDo.csv': #IF EXISTS THE SCRIPT DOWNLOAD IT
                try: 
                        WebDriverWait(driver, timeOut).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[2]/div[20]/div/div/div[1]/div')))
                        itemCheck = driver.find_element(By.XPATH, '/html/body/div[1]/div/div/div[2]/div[1]/div[2]/div[2]/div/div[2]/div[2]/div[2]/div[2]/div[1]/div/div/div/div/div/div[3]/div/div/div/div/div[2]/div/div/div/div/div[2]/div[20]/div/div/div[1]/div')
                
                        ActionChains(driver).context_click(itemCheck).perform()

                        try:
                                
                                WebDriverWait(driver, timeOut).until(EC.presence_of_element_located((By.XPATH, '/html/body/div[3]/div/div/div/div/div/div/ul/li[12]/button')))
                                downloadItem = driver.find_element(By.XPATH, '/html/body/div[3]/div/div/div/div/div/div/ul/li[12]/button')
                                downloadItem.click()
                        except TimeoutException:
                                print ("Snow failed, run script again...")
                except TimeoutException:
                        print ("Snow failed, run script again...")
                time.sleep(1)

except TimeoutException:
        print ("Snow failed, run script again...")

time.sleep(4)

source_file = get_latest_file("C:\\Users\\juan.cruz.aldaya\\Downloads")
destination_folder = os.path.dirname(os.path.abspath(__file__))
destination_file = os.path.join(destination_folder, "PRToDo.csv")

shutil.copy(source_file, destination_file)

csv_file_path = destination_file
df = pd.read_csv(csv_file_path)

# Create an Excel writer object
excel_file_path = os.path.join(destination_folder, "PRSinHacerFS.xlsx")
book = openpyxl.load_workbook(excel_file_path)


# selected_rows = df.iloc[]

writer = pd.ExcelWriter(excel_file_path, engine="openpyxl", mode="a", if_sheet_exists="overlay")

# Write the selected rows to the Excel file starting from cell B2 (2nd row, 2nd column)
df.to_excel(writer, sheet_name="PRSinHacerFS", index=False,header=False, startrow=1)

# Save the Excel file
writer.close()

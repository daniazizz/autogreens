import gspread
from google.oauth2.service_account import Credentials
from time import sleep
import random
from selenium import webdriver
from selenium.webdriver.edge.service import Service as EdgeService
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from openpyxl import Workbook
import datetime


# Load credentials from JSON file
creds = Credentials.from_service_account_file('autogreens-key.json')

# Define the scope
scope = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive.file',
    'https://www.googleapis.com/auth/drive'
]
creds = creds.with_scopes(scope)

# Authorize the client
client = gspread.authorize(creds)

# Open the Google Sheet
sheet = client.open('AUTOGREENS').sheet1

# Create (Add Data)
def create_row(data):
    sheet.append_row(data)

# Read (Retrieve Data)
def read_data():
    return sheet.get_all_records()

# Update (Modify Data)
def update_cell(row, col, new_value):
    sheet.update_cell(row, col, new_value)

# Delete (Remove Data)
def delete_row(row):
    sheet.delete_rows(row)
    

MC_PRICE_COL = 4
GREENYARD_PRICE_COL = 5
PRIJS_VERSHIL_COL = 6
VP_COL = 7
MARGE_COL = 8
LAST_UPDATE_COL = 9

import json

# Load the JSON file
with open('config.json', 'r') as file:
    config = json.load(file)

# Access the parameters
BROWSER_DRIVER_PATH = config.get('browser_driver_path')
GY_USERNAME = config.get('gy_username')
GY_PASSWORD = config.get('gy_password')
MC_USERNAME = config.get('mc_username')
MC_SHOP_ID = config.get('mc_shop_id')
MC_PASSWORD = config.get('mc_password')




data = read_data()
print(data)
# update_cell(2, 2, '25')

def human_sleep(min_time=1, max_time=3):
    sleep_time = random.uniform(min_time, max_time)
    sleep(sleep_time)

def init_eos():
    # Path to the Edge WebDriver
    browser_driver_path = BROWSER_DRIVER_PATH  # Adjust the path to your Edge WebDriver

    # Initialize the Edge WebDriver
    service = EdgeService(executable_path=browser_driver_path)
    driver = webdriver.Edge(service=service)



    # Step 1: Log in to the website
    driver.get("https://eos.firstinfresh.be/login")
    human_sleep(2, 4)

    # Enter username
    username_input = driver.find_element(By.XPATH, "/html/body/div[2]/div/div/div/form/div[1]/div[1]/input")
    ActionChains(driver).move_to_element(username_input).click().perform()
    human_sleep(1, 2)
    username_input.send_keys(GY_USERNAME)  # Replace with your username
    human_sleep(1, 3)

    # Enter password
    password_input = driver.find_element(By.XPATH, "/html/body/div[2]/div/div/div/form/div[1]/div[2]/input")
    ActionChains(driver).move_to_element(password_input).click().perform()
    human_sleep(1, 2)
    password_input.send_keys(GY_PASSWORD)  # Replace with your password
    human_sleep(1, 3)

    # Submit the form
    login_button = driver.find_element(By.XPATH, "/html/body/div[2]/div/div/div/form/div[2]/input")
    ActionChains(driver).move_to_element(login_button).click().perform()
    human_sleep(3, 5)
    
    return driver

driver = init_eos()




i = 2
for e in data:
    print(e)
    # Step 2: Navigate to the desired page
    driver.get(f"https://eos.firstinfresh.be/shop/item/{e.get('GY-REF')}")
    human_sleep(2, 4)

    # Step 3: Scrape the required information
    data_element = driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div/div[2]/div/div[2]/div[6]/table/tbody/tr[2]/td[2]")
    scraped_data = data_element.text
    print(scraped_data)

    update_cell(i, GREENYARD_PRICE_COL, scraped_data)
    i+=1

driver.quit()   


def init_mc():
    # Path to the Edge WebDriver
    edge_driver_path = BROWSER_DRIVER_PATH  # Adjust the path to your Edge WebDriver

    # Initialize the Edge WebDriver
    service = EdgeService(executable_path=edge_driver_path)
    driver = webdriver.Edge(service=service)



    # Step 1: Log in to the website
    driver.get("https://mycadencier.carrefour.eu/client/#!/login")
    human_sleep(2, 4)

    # Enter username
    username_input = driver.find_element(By.XPATH, "/html/body/ui-view/login/div/div/div/form/div/div[1]/input")
    ActionChains(driver).move_to_element(username_input).click().perform()
    human_sleep(1, 2)
    username_input.send_keys(MC_USERNAME)  # Replace with your username
    human_sleep(1, 3)
    
    # Enter store id
    store_id_input = driver.find_element(By.XPATH, "/html/body/ui-view/login/div/div/div/form/div/div[2]/input")
    ActionChains(driver).move_to_element(store_id_input).click().perform()
    human_sleep(1, 2)
    store_id_input.send_keys(MC_SHOP_ID)  # Replace with your username
    human_sleep(1, 3)

    # Enter password
    password_input = driver.find_element(By.XPATH, "/html/body/ui-view/login/div/div/div/form/div/div[3]/input")
    ActionChains(driver).move_to_element(password_input).click().perform()
    human_sleep(1, 2)
    password_input.send_keys(MC_PASSWORD)  # Replace with your password
    human_sleep(1, 3)

    # Submit the form
    login_button = driver.find_element(By.XPATH, "/html/body/ui-view/login/div/div/div/form/div/button")
    ActionChains(driver).move_to_element(login_button).click().perform()
    human_sleep(3, 5)
    
    # click category
    login_button = driver.find_element(By.XPATH, "/html/body/ui-view/app/div/home/div/div[1]/div/div[2]/div/i")
    ActionChains(driver).move_to_element(login_button).click().perform()
    human_sleep(3, 5)
    
    
    
    
    
    return driver

def format_euro(number):
    # Format the number with a comma as the decimal separator and two decimal places
    formatted_number = number.replace('.', ',')
    
    # Append the euro symbol if it's not already there
    if not formatted_number.endswith(" €"):
        formatted_number += " €"
    
    return formatted_number



driver = init_mc()



i = 2
for e in data:
    print(e)
    # Enter password
    search_input = driver.find_element(By.XPATH, "/html/body/ui-view/app/div/order/div/div[2]/div[2]/div[1]/div[2]/div[1]/div[2]/input")
    
    ActionChains(driver).move_to_element(search_input).click().perform()
    human_sleep(1, 2)
    search_input.clear()
    human_sleep(1, 2)
    search_input.send_keys(e.get('MC-REF')) 
    human_sleep(1, 3)
    
    vp_data_element = driver.find_element(By.XPATH, "/html/body/ui-view/app/div/order/div/div[2]/div[2]/div[3]/div[2]/table/tbody/tr[1]/td[3]/span")
    vp_scraped_data = vp_data_element.text
    update_cell(i, VP_COL, format_euro(vp_scraped_data))
    
    # click prODUCT
    product_button = driver.find_element(By.XPATH, "/html/body/ui-view/app/div/order/div/div[2]/div[2]/div[3]/div[1]/table/tbody/tr[1]/td")
    ActionChains(driver).move_to_element(product_button).click().perform()
    
    
    
    human_sleep(2, 4)

    # Step 3: Scrape the required information
    data_element = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div[1]/div[1]/div/ul[1]/li[1]/span[1]/strong")
    scraped_data = data_element.text
    print(scraped_data)

    update_cell(i, MC_PRICE_COL, format_euro(scraped_data))
    
    ct = datetime.datetime.now()
    update_cell(i, LAST_UPDATE_COL, str(ct))
    
    # click close
    close_button = driver.find_element(By.XPATH, "/html/body/div[1]/div/div/div/div[1]/div[1]/div/p/i")
    ActionChains(driver).move_to_element(close_button).click().perform()
    
    i+=1
    

driver.quit()




sheet.sort((PRIJS_VERSHIL_COL,'des'))


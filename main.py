import time
import os
import configparser
import glob
import requests
from zipfile import ZipFile
from io import BytesIO

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import pandas as pd

config = configparser.ConfigParser()
config.read('config.ini')

def login(email, password, driver):
    driver.get("https://dash.readme.com/login")
    
    try:
        # Wait for the email input field to be visible
        WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, "email")))

        # Input email and password
        driver.find_element(By.NAME, "email").send_keys(email)
        driver.find_element(By.NAME, "password").send_keys(password)

        # Wait for the login button to become clickable
        login_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(@class, 'Button_primary') and contains(., 'Log In')]"))
        )

        login_button.click()
        time.sleep(20)  # Wait for login process to complete
        
    except Exception as e:
        print("Exception occurred during login:", e)

    return driver

def download_chromedriver():
    print("Downloading chromedriver...")
    url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
    response = requests.get(url)
    version_number = response.text.strip()
    chromedriver_url = f"https://chromedriver.storage.googleapis.com/{version_number}/chromedriver_win32.zip"
    response = requests.get(chromedriver_url)
    with ZipFile(BytesIO(response.content)) as zip_file:
        zip_file.extractall()

def download_csv(driver):
    time.sleep(20)
    driver.get("https://dash.readme.com/project/api-beta-deepgram/v1.0/metrics/api-calls")
    try:
        # Click on the "Export CSV" button
        export_csv_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.XPATH, "//button[contains(., 'Export CSV')]"))
        )
        print("Export CSV button found:", export_csv_button.get_attribute('outerHTML'))
        export_csv_button.click()
        print("Export CSV button clicked")
        # Wait for the file to download
        time.sleep(10)  # Adjust the time according to your download speed

    except TimeoutException:
        print("Timed out waiting for Export CSV button to be clickable.")

    finally:
        driver.quit()

def add_to_master_csv(new_csv_file):
    # Check if the master.csv file exists
    if not os.path.exists('master.xlsx'):
        # If it doesn't exist, create a new master.csv with the content of the new CSV file
        df = pd.read_csv(new_csv_file)
        # Write the DataFrame to a new master.csv file with sheet name as 'Sheet1'
        df.to_excel('master.xlsx', index=False, sheet_name='Sheet1')
    else:
      print("Appending new sheet to existing Excel file...")
      # If the master.xlsx already exists, open it and append a new sheet with the content of the new CSV file
      with pd.ExcelWriter('master.xlsx', mode='a', engine='openpyxl') as writer:
          df_new = pd.read_csv(new_csv_file)  # Read the new CSV file into a DataFrame
          sheet_name = f'Data_{pd.to_datetime(df_new["time"][0]).date()}'
          # Write the DataFrame to the master.xlsx file with a new sheet name based on the date
          df_new.to_excel(writer, index=False, sheet_name=sheet_name)

def main():
    email = os.environ.get('EMAIL')
    password = os.environ.get('PASSWORD')

    options = Options()
    options.add_argument("--headless")
    if not os.path.exists('chromedriver'):
        download_chromedriver()
    service = Service('chromedriver')
    service.start()
    driver = webdriver.Remote(service.service_url, options=options)
    
    driver = login(email, password, driver)
    download_csv(driver)

    # Find the latest downloaded CSV file
    directory = os.getcwd()
    list_of_files = glob.glob(os.path.join(directory, '*.csv'))
    if not list_of_files:
        print(f"No CSV files found in {directory}")
        return

    # Get the latest downloaded CSV file
    latest_file = max(list_of_files, key=os.path.getmtime)
    print(f"Latest file: {latest_file}")

    # Add the contents of the latest downloaded CSV file to master.csv
    add_to_master_csv(latest_file)

    driver.quit()

if __name__ == "__main__":
    main()

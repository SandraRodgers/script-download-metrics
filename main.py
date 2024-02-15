import time
import os
import platform
import stat
import glob
import configparser
import requests
import zipfile
from zipfile import ZipFile

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

import requests
import os
import zipfile

def download_chromedriver(chrome_version):

  # Get latest stable version info from API
  api_url = "https://www.googleapis.com/storage/v1/b/chromium-browser-snapshots/o?prefix=Chrome-for-Testing%2Fstable%2F"
  response = requests.get(api_url)
  latest_info = response.json()["items"][0]
  
  # Extract version and platform 
  latest_version = latest_info["name"].split("/")[-1]
  platform = "linux64" # Set platform
  
  # Construct download URL
  download_url = f"https://storage.googleapis.com/chromium-browser-snapshots/Chrome-for-Testing/stable/{latest_version}/chromedriver_{platform}.zip"

  # Download and extract ChromeDriver
  response = requests.get(download_url)
  with open("chromedriver.zip", "wb") as f:
    f.write(response.content)
  
  with zipfile.ZipFile("chromedriver.zip", "r") as zip_ref:
    zip_ref.extractall()  

  # Set permissions
  chromedriver_path = "./chromedriver"
  os.chmod(chromedriver_path, 0o755)

  print("ChromeDriver downloaded successfully!")
  return chromedriver_path


# def download_chromedriver(chrome_version):
#     # Determine the operating system
#     system = platform.system().lower()
    
#     # Determine the chromedriver URL based on the operating system
#     if system == "windows":
#         chromedriver_url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
#         response = requests.get(chromedriver_url)
#         version_number = response.text.strip()
#         print("windows", version_number)
#         chromedriver_url = f"https://chromedriver.storage.googleapis.com/{version_number}/chromedriver_win32.zip"
#     elif system == "darwin":
#         chromedriver_url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
#         response = requests.get(chromedriver_url)
#         version_number = response.text.strip()
#         print("darwin", version_number)
#         chromedriver_url = f"https://chromedriver.storage.googleapis.com/{version_number}/chromedriver_mac64.zip"
#     else:
#         # Linux
#         # chromedriver_url = "https://chromedriver.storage.googleapis.com/LATEST_RELEASE"
#         chromedriver_url = "https://storage.googleapis.com/chrome-for-testing-public/121.0.6167.85/linux64/chrome-linux64.zip"
#         response = requests.get(chromedriver_url)
#         version_number = response.text.strip()
#         chromedriver_url = f"https://chromedriver.storage.googleapis.com/{version_number}/chromedriver_linux64.zip"

#     # Download chromedriver zip file
#     print("Downloading chromedriver...")
#     response = requests.get(chromedriver_url)
#     with open("chromedriver.zip", "wb") as f:
#         f.write(response.content)
#     print(f"Wrote to {os.path.abspath('chromedriver.zip')}")
#     print("Chromedriver download complete.")
#     time.sleep(10)
#     # Verify the file path where the script expects to find the "chromedriver.zip" file
#     chromedriver_path = os.path.abspath("chromedriver.zip")
#     print(f"Expected file path: {chromedriver_path}")  # Print the expected file path

#     # Check if the file exists at the specified path
#     if not os.path.exists(chromedriver_path):
#         print(f"File not found at: {chromedriver_path}")
#         return None
    
#     # Extract chromedriver
#     extracted_dir = os.path.dirname(chromedriver_path)
#     with zipfile.ZipFile(chromedriver_path, "r") as zip_ref:
#         zip_ref.extractall(extracted_dir)  

#     # Get path to extracted chromedriver
#     chromedriver_exe = os.path.join(extracted_dir, "chromedriver")

#     # Set permissions
#     os.chmod(chromedriver_exe, 0o755) 

#     # Return extracted exe path
#     return chromedriver_exe



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
    chrome_version = os.environ.get('CHROME_VERSION')

    options = Options()
    options.add_argument("--headless")

    chromedriver_path = download_chromedriver(chrome_version)
    
    if not os.path.exists(chromedriver_path):
        print("chromedriver not found!")
        return
    
    service = Service(chromedriver_path)
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

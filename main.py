import os
import glob
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv

load_dotenv()

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
        time.sleep(10)  # Wait for login process to complete
        
    except Exception as e:
        print("Exception occurred during login:", e)

    return driver

def download_csv(driver):
  print("Downloading CSV file...")
    # Navigate to the webpage containing the button you want to click
  driver.get("https://dash.readme.com/project/api-beta-deepgram/v1.0/metrics/api-calls")
  time.sleep(10)
  # Find and click the button that triggers the file download
  download_button = driver.find_element(By.XPATH, "//button[contains(., 'Export CSV')]")

  download_button.click()
  time.sleep(10)

def add_to_master_csv(new_csv_file):
  # Check if the master.csv file exists
  if not os.path.exists(f"{os.getenv('DIRECTORY')}master.xlsx"):
      # If it doesn't exist, create a new master.xlsx with the content of the new CSV file
      new_data = pd.read_csv(new_csv_file)
      # Write the DataFrame to a new master.csv file with sheet name as 'Sheet1'
      new_data.to_excel(f"{os.getenv('DIRECTORY')}master.xlsx", index=False, sheet_name='Sheet1')
  else:
    print("Appending new sheet to existing Excel file...")
    # If the master.xlsx already exists, open it and append a new sheet with the content of the new CSV file
    with pd.ExcelWriter(f"{os.getenv('DIRECTORY')}master.xlsx", mode='a', engine='openpyxl') as writer:
        data = pd.read_csv(new_csv_file)  # Read the new CSV file into a DataFrame
        sheet_name = f'Data_{pd.to_datetime(data["time"][0]).date()}'
        # Write the DataFrame to the master.xlsx file with a new sheet name based on the date
        data.to_excel(writer, index=False, sheet_name=sheet_name)

def main():
  # Set up Chrome options to automatically download files to the default downloads folder
  chrome_options = webdriver.ChromeOptions()
  chrome_options.add_experimental_option("prefs", {
    "download.default_directory": f"{os.getenv('DIRECTORY')}files",
    "download.prompt_for_download": False,
    "download.directory_upgrade": True
  })

  email = os.getenv('EMAIL')
  password = os.getenv('PASSWORD')

  # Initialize the Chrome webdriver with the options
  driver = webdriver.Chrome(options=chrome_options)
  driver = login(email, password, driver)
  
  # Download the CSV file
  download_csv(driver)
  print("CSV file downloaded")
  
  # Find the latest downloaded CSV file
  directory = f"{os.getenv('DIRECTORY')}files/"
  list_of_files = glob.glob(os.path.join(directory, '*.csv'))
  print(list_of_files)
  if not list_of_files:
      print(f"No CSV files found in {directory}")
      return

  # Get the latest downloaded CSV file
  latest_file = max(list_of_files, key=os.path.getmtime)
  print(f"Latest file: {latest_file}")

  # Add the contents of the latest downloaded CSV file to master.csv
  add_to_master_csv(latest_file)

  # Wait for 10 seconds before quitting the webdriver
  time.sleep(10)
  driver.quit()

if __name__ == "__main__":
      main()


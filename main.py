import os
import time
import glob
import json
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account

# --- 1. CLOUD CONFIGURATION ---
PORTAL_URL = "https://motoindia.nuralsales.app/login" # Ensure this is your actual URL
USERNAME = os.getenv("PORTAL_USER")
PASSWORD = os.getenv("PORTAL_PASS")
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID")
JSON_KEY_DATA = os.getenv("GCP_JSON_KEY")

DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

# --- 2. DATE CALCULATION ---
today = datetime.now()
date_str = today.strftime("%d%m%Y") 

# --- 3. BROWSER SETUP ---
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

prefs = {"download.default_directory": DOWNLOAD_DIR, "download.prompt_for_download": False}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
actions = ActionChains(driver)

try:
    print(f"Logging into portal as {USERNAME}...")
    driver.get(PORTAL_URL)
    time.sleep(5)

    # Login
    driver.find_element(By.XPATH, "//input[@placeholder='Enter Username']").send_keys(USERNAME)
    driver.find_element(By.XPATH, "//input[@placeholder='Enter Password']").send_keys(PASSWORD)
    login_btn = driver.find_element(By.XPATH, "//div[text()='LOGIN']")
    driver.execute_script("arguments[0].click();", login_btn)
    time.sleep(10) 

    # Navigate and Export
    print("Navigating to Reports...")
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//img[@alt='reports']"))
    time.sleep(5)
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//a[@href='/isp-sales-report']"))
    time.sleep(10) 

    print(f"Setting dates for {date_str}...")
    calendars = driver.find_elements(By.XPATH, "//*[local-name()='svg' and @data-testid='CalendarTodayIcon']")
    for cal in calendars:
        actions.move_to_element(cal).click().send_keys(date_str).send_keys(Keys.ENTER).perform()
        time.sleep(1)
    
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'SEARCH')]"))
    time.sleep(20) 
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'EXPORT EXCEL')]"))
    time.sleep(20) 

    # --- 4. GOOGLE DRIVE UPLOAD (Fixed for Quota Issues) ---
    print("Uploading to Google Drive...")
    list_of_files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx"))
    if not list_of_files:
        raise Exception("Download failed: No Excel file found.")
    latest_file = max(list_of_files, key=os.path.getctime)

    creds_info = json.loads(JSON_KEY_DATA)
    creds = service_account.Credentials.from_service_account_info(creds_info)
    service = build('drive', 'v3', credentials=creds)

    DRIVE_FILE_NAME = "Daily_ISP_Sales_Report.xlsx"
    
    # Search specifically in your folder
    query = f"name='{DRIVE_FILE_NAME}' and '{DRIVE_FOLDER_ID}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id)", supportsAllDrives=True).execute()
    files = results.get('files', [])

    media = MediaFileUpload(latest_file, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    if files:
        # Update existing file
        file_id = files[0]['id']
        service.files().update(fileId=file_id, media_body=media, supportsAllDrives=True).execute()
        print(f"Success: Updated {DRIVE_FILE_NAME}")
    else:
        # Create new file inside your folder to use YOUR quota
        file_metadata = {'name': DRIVE_FILE_NAME, 'parents': [DRIVE_FOLDER_ID]}
        service.files().create(body=file_metadata, media_body=media, fields='id', supportsAllDrives=True).execute()
        print(f"Success: Created {DRIVE_FILE_NAME}")

except Exception as e:
    print(f"ERROR: {e}")
finally:
    driver.quit()

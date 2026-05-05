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

# --- 1. CONFIGURATION ---
PORTAL_URL = "https://motoindia.nuralsales.app/login" 
USERNAME = os.getenv("PORTAL_USER")
PASSWORD = os.getenv("PORTAL_PASS")
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID")
JSON_KEY_DATA = os.getenv("GCP_JSON_KEY")

DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

# --- 2. BROWSER SETUP ---
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_experimental_option("prefs", {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True
})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
actions = ActionChains(driver)

try:
    print(f"Logging in as {USERNAME}...")
    driver.get(PORTAL_URL)
    time.sleep(10) # Increased initial wait

    driver.find_element(By.XPATH, "//input[@placeholder='Enter Username']").send_keys(USERNAME)
    driver.find_element(By.XPATH, "//input[@placeholder='Enter Password']").send_keys(PASSWORD)
    login_btn = driver.find_element(By.XPATH, "//div[text()='LOGIN']")
    driver.execute_script("arguments[0].click();", login_btn)
    time.sleep(15) 

    print("Navigating to Reports...")
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//img[@alt='reports']"))
    time.sleep(5)
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//a[@href='/isp-sales-report']"))
    time.sleep(15) 

    date_str = datetime.now().strftime("%d%m%Y")
    print(f"Setting dates for {date_str}...")
    calendars = driver.find_elements(By.XPATH, "//*[local-name()='svg' and @data-testid='CalendarTodayIcon']")
    for cal in calendars:
        actions.move_to_element(cal).click().send_keys(date_str).send_keys(Keys.ENTER).perform()
        time.sleep(2)
    
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'SEARCH')]"))
    time.sleep(25) 
    
    print("Triggering Export...")
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'EXPORT EXCEL')]"))
    
    # --- 3. THE HEAVY WAIT LOGIC (Fixes 'Empty Sequence' Error) ---
    print("Waiting for file to appear in cloud storage (Max 5 mins)...")
    wait_time = 0
    max_wait = 300 # 5 Minutes
    latest_file = None

    while wait_time < max_wait:
        files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx"))
        # Ensure no '.crdownload' temp files are present
        temp_files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload"))
        
        if files and not temp_files:
            latest_file = max(files, key=os.path.getctime)
            if os.path.getsize(latest_file) > 0:
                print(f"Success: Found {os.path.basename(latest_file)} ({os.path.getsize(latest_file)} bytes)")
                break
        
        time.sleep(10)
        wait_time += 10
        if wait_time % 60 == 0:
            print(f"Still waiting for portal to generate file... ({wait_time}s)")

    if not latest_file:
        raise Exception("PORTAL ERROR: The portal never started the download.")

    # --- 4. UPLOAD ---
    print("Connecting to Google Drive...")
    creds = service_account.Credentials.from_service_account_info(json.loads(JSON_KEY_DATA))
    service = build('drive', 'v3', credentials=creds)

    DRIVE_FILE_NAME = "Daily_ISP_Sales_Report.xlsx"
    
    # Clean up old file to prevent quota check errors
    query = f"name='{DRIVE_FILE_NAME}' and '{DRIVE_FOLDER_ID}' in parents"
    existing = service.files().list(q=query).execute().get('files', [])
    for f in existing:
        service.files().delete(fileId=f['id']).execute()
        print("Removed old file version.")

    media = MediaFileUpload(latest_file, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)
    file_metadata = {'name': DRIVE_FILE_NAME, 'parents': [DRIVE_FOLDER_ID]}
    
    service.files().create(body=file_metadata, media_body=media, fields='id', supportsAllDrives=True).execute()
    print(f"REPORT UPLOADED SUCCESSFULLY TO GOOGLE DRIVE")

except Exception as e:
    print(f"CRITICAL ERROR: {e}")
finally:
    driver.quit()

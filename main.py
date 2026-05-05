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
PORTAL_URL = "https://motoindia.nuralsales.app/login" # Ensure this is updated
USERNAME = os.getenv("PORTAL_USER")
PASSWORD = os.getenv("PORTAL_PASS")
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID")
JSON_KEY_DATA = os.getenv("GCP_JSON_KEY")

DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

# --- 2. BROWSER SETUP (Optimized for Downloads) ---
chrome_options = Options()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "directory_upgrade": True,
    "safebrowsing.enabled": True
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
actions = ActionChains(driver)

try:
    print(f"Logging in as {USERNAME}...")
    driver.get(PORTAL_URL)
    time.sleep(5)

    driver.find_element(By.XPATH, "//input[@placeholder='Enter Username']").send_keys(USERNAME)
    driver.find_element(By.XPATH, "//input[@placeholder='Enter Password']").send_keys(PASSWORD)
    login_btn = driver.find_element(By.XPATH, "//div[text()='LOGIN']")
    driver.execute_script("arguments[0].click();", login_btn)
    time.sleep(10) 

    print("Navigating to Reports...")
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//img[@alt='reports']"))
    time.sleep(5)
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//a[@href='/isp-sales-report']"))
    time.sleep(10) 

    date_str = datetime.now().strftime("%d%m%Y")
    print(f"Selecting dates for {date_str}...")
    calendars = driver.find_elements(By.XPATH, "//*[local-name()='svg' and @data-testid='CalendarTodayIcon']")
    for cal in calendars:
        actions.move_to_element(cal).click().send_keys(date_str).send_keys(Keys.ENTER).perform()
        time.sleep(1)
    
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'SEARCH')]"))
    time.sleep(25) # Wait for portal to process large data search
    
    print("Exporting Excel (Large File Mode)...")
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'EXPORT EXCEL')]"))
    
    # --- DYNAMIC WAIT FOR LARGE DOWNLOADS ---
    print("Waiting for download to finish (Max 3 minutes)...")
    timeout = 180 # 3 minutes for 12MB+ files
    seconds = 0
    latest_file = None
    
    while seconds < timeout:
        # Check if there are any temporary '.crdownload' files still active
        temp_files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload"))
        files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx"))
        
        if files and not temp_files:
            latest_file = max(files, key=os.path.getctime)
            size = os.path.getsize(latest_file)
            if size > 0:
                print(f"Verified: {os.path.basename(latest_file)} is ready ({round(size/1024/1024, 2)} MB)")
                break
        
        time.sleep(5)
        seconds += 5
        if seconds % 30 == 0:
            print(f"...still waiting ({seconds}s elapsed)...")

    if not latest_file or os.path.getsize(latest_file) == 0:
        raise Exception("Download timed out or file is empty.")

    # --- 3. GOOGLE DRIVE UPLOAD (Resumable for Large Files) ---
    print("Uploading to Google Drive...")
    creds_info = json.loads(JSON_KEY_DATA)
    creds = service_account.Credentials.from_service_account_info(creds_info)
    service = build('drive', 'v3', credentials=creds)

    DRIVE_FILE_NAME = "Daily_ISP_Sales_Report.xlsx"
    query = f"name='{DRIVE_FILE_NAME}' and '{DRIVE_FOLDER_ID}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id)").execute()
    files = results.get('files', [])

    # Using resumable=True is mandatory for 10MB+ files to prevent timeouts
    media = MediaFileUpload(latest_file, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)

    if files:
        file_id = files[0]['id']
        service.files().update(fileId=file_id, media_body=media).execute()
        print(f"Success: Updated {DRIVE_FILE_NAME} in your Drive.")
    else:
        file_metadata = {'name': DRIVE_FILE_NAME, 'parents': [DRIVE_FOLDER_ID]}
        service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print(f"Success: Created {DRIVE_FILE_NAME} in your Drive.")

except Exception as e:
    print(f"ERROR: {e}")
finally:
    driver.quit()

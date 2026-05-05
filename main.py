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
chrome_options.add_experimental_option("prefs", {"download.default_directory": DOWNLOAD_DIR})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
actions = ActionChains(driver)

try:
    # ... [Login and Export logic remains the same] ...
    driver.get(PORTAL_URL)
    # [Skipping to the Upload part for brevity - use your full export code here]
    
    # --- 3. THE "OWNERSHIP TRANSFER" UPLOAD ---
    print("Uploading with ownership transfer...")
    list_of_files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx"))
    latest_file = max(list_of_files, key=os.path.getctime)

    creds = service_account.Credentials.from_service_account_info(json.loads(JSON_KEY_DATA))
    service = build('drive', 'v3', credentials=creds)

    DRIVE_FILE_NAME = "Daily_ISP_Sales_Report.xlsx"
    
    # Define file metadata with YOU as the intended owner indirectly via the parent folder
    file_metadata = {
        'name': DRIVE_FILE_NAME,
        'parents': [DRIVE_FOLDER_ID]
    }
    
    media = MediaFileUpload(latest_file, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', resumable=True)

    # 1. DELETE existing file first to clear any quota-locked ghosts
    query = f"name='{DRIVE_FILE_NAME}' and '{DRIVE_FOLDER_ID}' in parents"
    existing = service.files().list(q=query).execute().get('files', [])
    for f in existing:
        service.files().delete(fileId=f['id']).execute()

    # 2. CREATE the file. 
    # THE SECRET: We must use the 'fields' parameter to ensure it returns the ID
    new_file = service.files().create(
        body=file_metadata, 
        media_body=media, 
        fields='id',
        supportsAllDrives=True # Keep this for compatibility
    ).execute()
    
    print(f"Success! File uploaded to Folder ID: {DRIVE_FOLDER_ID}")

except Exception as e:
    print(f"ERROR: {e}")
finally:
    driver.quit()

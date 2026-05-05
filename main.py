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
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials

# --- 1. CONFIGURATION ---
PORTAL_URL = "https://motoindia.nuralsales.app/login"  # Ensure this is correct
USERNAME = os.getenv("PORTAL_USER")
PASSWORD = os.getenv("PORTAL_PASS")
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID")
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
    "download.prompt_for_download": False
})

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
actions = ActionChains(driver)

try:
    print(f"Logging into portal as {USERNAME}...")
    driver.get(PORTAL_URL)
    time.sleep(10)

    driver.find_element(By.XPATH, "//input[@placeholder='Enter Username']").send_keys(USERNAME)
    driver.find_element(By.XPATH, "//input[@placeholder='Enter Password']").send_keys(PASSWORD)
    login_btn = driver.find_element(By.XPATH, "//div[text()='LOGIN']")
    driver.execute_script("arguments[0].click();", login_btn)
    time.sleep(15)

    print("Navigating to ISP Sales Reports...")
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//img[@alt='reports']"))
    time.sleep(5)
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//a[@href='/isp-sales-report']"))
    time.sleep(15)

    date_str = datetime.now().strftime("%d%m%Y")
    print(f"Selecting date: {date_str}...")
    calendars = driver.find_elements(By.XPATH, "//*[local-name()='svg' and @data-testid='CalendarTodayIcon']")
    for cal in calendars:
        actions.move_to_element(cal).click().send_keys(date_str).send_keys(Keys.ENTER).perform()
        time.sleep(2)

    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'SEARCH')]"))
    time.sleep(25)

    print("Exporting Excel file...")
    driver.execute_script("arguments[0].click();", driver.find_element(By.XPATH, "//button[contains(., 'EXPORT EXCEL')]"))

    # Dynamic wait for large files (12MB+)
    print("Waiting for download to finalize...")
    timeout = 300  # 5 minutes
    elapsed = 0
    latest_file = None
    while elapsed < timeout:
        files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx"))
        temp = glob.glob(os.path.join(DOWNLOAD_DIR, "*.crdownload"))
        if files and not temp:
            latest_file = max(files, key=os.path.getctime)
            if os.path.getsize(latest_file) > 0:
                print(f"Verified: {os.path.basename(latest_file)} is ready.")
                break
        time.sleep(10)
        elapsed += 10

    # --- 2a. DOWNLOAD-TIMEOUT GUARD ---
    if latest_file is None:
        raise RuntimeError(
            f"Download timed out after {timeout}s. "
            f"Files in directory: {os.listdir(DOWNLOAD_DIR)}"
        )

    # --- 3. UPLOAD AS ASHISH (NO QUOTA LIMITS) ---
    print("Authenticating with Personal OAuth Token...")
    creds = Credentials(
        None,
        refresh_token=os.getenv("GOOGLE_REFRESH_TOKEN"),
        token_uri="https://oauth2.googleapis.com/token",
        client_id=os.getenv("GOOGLE_CLIENT_ID"),
        client_secret=os.getenv("GOOGLE_CLIENT_SECRET")
    )
    creds.refresh(Request())
    service = build('drive', 'v3', credentials=creds)

    DRIVE_FILE_NAME = "Daily_ISP_Sales_Report.xlsx"

    # Overwrite existing .xlsx logic
    query = f"name='{DRIVE_FILE_NAME}' and '{DRIVE_FOLDER_ID}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id)").execute()
    existing_files = results.get('files', [])

    media = MediaFileUpload(
        latest_file,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=True
    )

    if existing_files:
        file_id = existing_files[0]['id']
        service.files().update(fileId=file_id, media_body=media).execute()
        print(f"SUCCESS: Updated {DRIVE_FILE_NAME} in your Drive.")
    else:
        file_metadata = {'name': DRIVE_FILE_NAME, 'parents': [DRIVE_FOLDER_ID]}
        service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print(f"SUCCESS: Created {DRIVE_FILE_NAME} in your Drive.")

    # --- 4. ALSO MAINTAIN A GOOGLE SHEET FOR THE DASHBOARD ---
    # The dashboard reads this Sheet via its published-to-web CSV URL.
    print("Maintaining Google Sheet for dashboard...")
    SHEET_NAME = "Daily_ISP_Sales_Report_Live"

    sheet_query = (
        f"name='{SHEET_NAME}' "
        f"and '{DRIVE_FOLDER_ID}' in parents "
        f"and mimeType='application/vnd.google-apps.spreadsheet' "
        f"and trashed=false"
    )
    sheet_results = service.files().list(q=sheet_query, fields="files(id)").execute()
    sheet_existing = sheet_results.get('files', [])

    sheet_media = MediaFileUpload(
        latest_file,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        resumable=True
    )

    if sheet_existing:
        sheet_id = sheet_existing[0]['id']
        service.files().update(fileId=sheet_id, media_body=sheet_media).execute()
        print(f"SUCCESS: Updated dashboard Sheet ({sheet_id}).")

        # Self-check: confirm conversion still works (catches Drive API quirks)
        verify = service.files().get(fileId=sheet_id, fields="mimeType").execute()
        if verify['mimeType'] != 'application/vnd.google-apps.spreadsheet':
            raise RuntimeError(
                f"Sheet conversion broke - mimeType is now {verify['mimeType']}. "
                f"Drive update no longer auto-converts xlsx to Sheets. "
                f"Switch to Option B (Sheets API). Sheet ID: {sheet_id}"
            )
    else:
        sheet_metadata = {
            'name': SHEET_NAME,
            'parents': [DRIVE_FOLDER_ID],
            'mimeType': 'application/vnd.google-apps.spreadsheet'
        }
        new_sheet = service.files().create(
            body=sheet_metadata,
            media_body=sheet_media,
            fields='id'
        ).execute()
        new_id = new_sheet['id']
        print(f"SUCCESS: Created dashboard Sheet ({new_id}).")
        print("=" * 70)
        print("ACTION REQUIRED (one-time setup):")
        print(f"1. Open: https://docs.google.com/spreadsheets/d/{new_id}")
        print("2. File > Share > Publish to web")
        print("3. Choose 'Entire document', format 'Comma-separated values (.csv)'")
        print("4. Click Publish, copy the CSV URL, share it with Claude.")
        print("=" * 70)

except Exception as e:
    print(f"CRITICAL ERROR: {e}")
finally:
    driver.quit()

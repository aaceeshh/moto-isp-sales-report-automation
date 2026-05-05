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

# --- 1. CLOUD CONFIGURATION (using GitHub Secrets) ---
# These variables will pull from the 'Secrets' you set up in Step 2
PORTAL_URL = "https://motoindia.nuralsales.app/login" # Update this to your actual URL
USERNAME = os.getenv("PORTAL_USER")
PASSWORD = os.getenv("PORTAL_PASS")
DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID")
JSON_KEY_DATA = os.getenv("GCP_JSON_KEY")

DOWNLOAD_DIR = os.path.join(os.getcwd(), "downloads")
if not os.path.exists(DOWNLOAD_DIR):
    os.makedirs(DOWNLOAD_DIR)

# --- 2. DATE CALCULATION ---
today = datetime.now()
start_date = today.strftime("%d%m%Y") 
end_date = today.strftime("%d%m%Y")   

# --- 3. HEADLESS BROWSER SETUP ---
chrome_options = Options()
chrome_options.add_argument("--headless") # Required for Cloud running
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.add_argument("--window-size=1920,1080")

prefs = {
    "download.default_directory": DOWNLOAD_DIR,
    "download.prompt_for_download": False,
    "directory_upgrade": True
}
chrome_options.add_experimental_option("prefs", prefs)

driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
actions = ActionChains(driver)

try:
    print(f"Starting Portal Login for {USERNAME}...")
    driver.get(PORTAL_URL)
    time.sleep(5)

    # Login Phase
    driver.find_element(By.XPATH, "//input[@placeholder='Enter Username']").send_keys(USERNAME)
    driver.find_element(By.XPATH, "//input[@placeholder='Enter Password']").send_keys(PASSWORD)
    login_btn = driver.find_element(By.XPATH, "//div[text()='LOGIN']")
    driver.execute_script("arguments[0].click();", login_btn)
    time.sleep(10) 

    # Navigation
    print("Navigating to Reports...")
    reports_icon = driver.find_element(By.XPATH, "//img[@alt='reports']")
    driver.execute_script("arguments[0].click();", reports_icon)
    time.sleep(5)
    
    report_link = driver.find_element(By.XPATH, "//a[@href='/isp-sales-report']")
    driver.execute_script("arguments[0].click();", report_link)
    time.sleep(10) 

    # Date Selection
    print(f"Selecting Dates: {start_date}")
    calendar_icons = driver.find_elements(By.XPATH, "//*[local-name()='svg' and @data-testid='CalendarTodayIcon']")
    
    for i, date_val in enumerate([start_date, end_date]):
        actions.move_to_element(calendar_icons[i]).click().perform()
        time.sleep(1)
        actions.send_keys(date_val).send_keys(Keys.ENTER).perform()
    
    # Search and Export
    print("Clicking Search & Export...")
    search_btn = driver.find_element(By.XPATH, "//button[contains(., 'SEARCH')]")
    driver.execute_script("arguments[0].click();", search_btn)
    time.sleep(25) 

    export_btn = driver.find_element(By.XPATH, "//button[contains(., 'EXPORT EXCEL')]")
    driver.execute_script("arguments[0].click();", export_btn)
    time.sleep(15)
    print("Download triggered successfully.")

    # --- 4. GOOGLE DRIVE UPLOAD ---
    print("Connecting to Google Drive...")
    
    # Find the file
    list_of_files = glob.glob(os.path.join(DOWNLOAD_DIR, "*.xlsx"))
    if not list_of_files:
        raise Exception("No Excel file found in download directory!")
    latest_file = max(list_of_files, key=os.path.getctime)

    # Authenticate using the Secret JSON string
    info = json.loads(JSON_KEY_DATA)
    creds = service_account.Credentials.from_service_account_info(info)
    service = build('drive', 'v3', credentials=creds)

    # Replace logic
    DRIVE_FILE_NAME = "Daily_ISP_Sales_Report.xlsx"
    query = f"name='{DRIVE_FILE_NAME}' and '{DRIVE_FOLDER_ID}' in parents and trashed=false"
    results = service.files().list(q=query, fields="files(id)").execute()
    files = results.get('files', [])

    media = MediaFileUpload(latest_file, mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    if files:
        file_id = files[0]['id']
        service.files().update(fileId=file_id, media_body=media).execute()
        print(f"Done! Updated existing file: {DRIVE_FILE_NAME}")
    else:
        file_metadata = {'name': DRIVE_FILE_NAME, 'parents': [DRIVE_FOLDER_ID]}
        service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print(f"Done! Created new file: {DRIVE_FILE_NAME}")

except Exception as e:
    print(f"ERROR: {e}")
finally:
    driver.quit()

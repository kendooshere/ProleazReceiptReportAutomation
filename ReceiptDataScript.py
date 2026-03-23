import os
import time
import shutil
import pandas as pd 
import gspread
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.library import ChromeDriverManager
from selenium.webdriver.common.keys import Keys
from dotenv import load_dotenv
from datetime import date
from google.oauth2.service_account import Credentials

load_dotenv()
CREDENTIALS = os.getenv("CREDENTIALS")
BUFFER_SHEET = os.getenv("BUFFER_SHEET")
USERNAME = os.getenv("USER_NAME")
PASSWORD = os.getenv("PASS_WORD")
DOWNLOAD_PATH = os.getenv("DOWNLOAD_PATH")
PORTAL_URL = os.getenv("PORTAL_URL")

def clear_receipt_folder(path):
    if not os.path.exists(path):
        os.makedirs(path)
        return
    
    for file in os.listdir(path):
        file_path = os.path.join(path, file)
        if os.path.isfile(file_path):
            os.remove(file_path)

def wait_for_download(download_path, timeout=180):
    time_out = time.time() + timeout
    while time.time() < time_out:
        files = os.listdir(download_path)

        if any(f.endswith(".crdownload") for f in files):
            time.sleep(1)
            continue

        excel_files = [f for f in files if f.endswith(".xlsx")]

        if len(excel_files) == 1:
            return os.path.join(download_path, excel_files[0])
            
    raise TimeoutError("Unable to download file due to timeout!")


def buffer_sheet_upload(df):
    # --- UPDATED SECTION FOR FORMULA COMPATIBILITY ---
    
    # 1. Ensure 'Amount' (Column K) is numeric so SUMIFS can add it
    if 'Amount' in df.columns:
        # Removes commas/spaces and converts to float
        df['Amount'] = pd.to_numeric(df['Amount'].astype(str).str.replace(r'[^\d.]', '', regex=True), errors='coerce')

    # 2. Standardize Dates (Column A / Transaction Date) to ISO format
    # This is the "Universal" format that UK/Indian sheets recognize as a Date Value
    date_cols = ['Created Date', 'Transaction Date', 'Chq Date', 'PDC Matured On']
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
            # Using YYYY-MM-DD ensures no Month-Name (Sep/Oct) errors
            df[col] = df[col].dt.strftime('%Y-%m-%d')

    # Cleanup for JSON compatibility
    df = df.fillna("").replace([float("inf"), float("-inf")], "")
    # --------------------------------------------------
            
    scopes = ['https://www.googleapis.com/auth/spreadsheets']
    creds = Credentials.from_service_account_file(CREDENTIALS, scopes=scopes)
    client = gspread.authorize(creds)
    
    sh = client.open_by_key(BUFFER_SHEET)
    
    try:
        sheet = sh.get_worksheet(1)
    except Exception as e:
        print("Error: Could not find the second worksheet. Ensure the spreadsheet has at least two tabs.")
        raise e

    sheet.clear()
    
    # Prepare data for upload
    data = [df.columns.values.tolist()] + df.values.tolist()
    
    # value_input_option='USER_ENTERED' removes the ' prefix and triggers date recognition
    sheet.update(data, value_input_option='USER_ENTERED')
    
    print("Data uploaded to buffer sheet (Subsheet 2) successfully!")
    
    
def download_report():
    clear_receipt_folder(DOWNLOAD_PATH)

    chrome_options = webdriver.ChromeOptions()
    prefs  = {
        "download.default_directory" : DOWNLOAD_PATH,
        "download.prompt_for_download" : False,
        "directory_upgrade" : True,
        "safebrowsing.enabled": True,
        "safebrowsing.disable_download_protection": True
    }

    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.add_argument("--safebrowsing-disable-download-protection")

    driver = webdriver.Chrome(service = Service(ChromeDriverManager().install()), options = chrome_options)
    driver.command_executor._commands["send_command"] = ("POST", "/session/$sessionId/chromium/send_command")
    driver.execute("send_command", {
            "cmd": "Page.setDownloadBehavior",
            "params": {"behavior": "allow", "downloadPath": DOWNLOAD_PATH}
        })
        
    driver.get(PORTAL_URL)
    today = date.today()

    fToday = today.strftime("%d-%b-%Y")

    try:
        driver.find_element(By.ID,"txtusr").send_keys(USERNAME)
        driver.find_element(By.ID, "txtpswd").send_keys(PASSWORD)
        driver.find_element(By.ID, "btnlogin").click()

        print("Browser launched and login is done in the portal")
        time.sleep(3)
        driver.get(f"{PORTAL_URL}/LeaseManagement/Report/ReceiptReport")
        start_date = "01-Jan-2020" 
        from_date = driver.find_element(By.ID, "FromDate")
        driver.execute_script("""arguments[0].value = arguments[1];
                              arguments[0].dispatchEvent(new Event ('change'));
                              arguments[0].dispatchEvent(new Event ('blur'));""", from_date, start_date)
        print("From Date is now entered!")
        to_date = driver.find_element(By.ID, "ToDate")
        driver.execute_script("""arguments[0].value = arguments[1];
                              arguments[0].dispatchEvent(new Event ('change'));
                              arguments[0].dispatchEvent(new Event ('blur'));""", to_date, fToday)
        print("To Date is now entered!")
        time.sleep(3)
        
        driver.find_element(By.ID, "btnExport").click()
        print("Export button Clicked. Waiting for file download.....")
        receipt_path = wait_for_download(DOWNLOAD_PATH, 300)
        print("File downloaded at :", receipt_path)
        driver.get(f"{PORTAL_URL}/Login/Logout")
        df = pd.read_excel(receipt_path, skiprows=1)
        print("Data read successful, proceeding to upload....")
        
        buffer_sheet_upload(df)

    except Exception as e:
        print("An error occured:", e)
        raise

    finally:
        driver.quit()
        print("Browser Closed!")

if __name__ == "__main__":
    download_report()

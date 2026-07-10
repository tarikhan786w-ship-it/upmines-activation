# -*- coding: utf-8 -*-
"""
UPMINES OSTP-ISTP CHECKER - API VERSION
"""
import os
import shutil
import tempfile
import time
import threading
import uuid
from datetime import datetime
from fastapi import FastAPI, File, UploadFile, BackgroundTasks, HTTPException
from fastapi.responses import FileResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import easyocr
import requests
import json
import hashlib
import socket
import uuid as uuid_lib

# ============================================
# 🔑 GITHUB CONFIGURATION
# ============================================

GITHUB_TOKEN = "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"
GITHUB_CREDENTIALS_URL = "https://raw.githubusercontent.com/tarikhan786w-ship-it/upmines-activation/refs/heads/main/credentials.json"
GITHUB_ACTIVATION_URL = "https://raw.githubusercontent.com/tarikhan786w-ship-it/upmines-activation/refs/heads/main/activated_users.json"

# ============================================
# ACTIVATION SYSTEM
# ============================================

class ActivationManager:
    def __init__(self):
        self.github_url = GITHUB_ACTIVATION_URL
        self.github_token = GITHUB_TOKEN
        self.is_activated = False
        self.user_uuid = None
        self.expiry_date = None
    
    def get_uuid(self):
        try:
            mac = ':'.join([uuid_lib.UUID(int=uuid_lib.getnode()).hex[-12:][e:e+2] for e in range(0, 12, 2)])
            hostname = socket.gethostname()
            unique = hashlib.md5(f"{hostname}_{mac}".encode()).hexdigest()[:8].upper()
            self.user_uuid = f"{unique[:4]}-{unique[4:]}"
            return self.user_uuid
        except:
            return None
    
    def check_github(self):
        try:
            uuid = self.get_uuid()
            if not uuid:
                return False
            
            headers = {
                'Authorization': f'token {self.github_token}',
                'Accept': 'application/vnd.github.v3.raw'
            }
            response = requests.get(self.github_url, headers=headers, timeout=10)
            
            if response.status_code == 200:
                data = response.json()
                users = data.get('users', {})
                
                if uuid not in users:
                    self.is_activated = False
                    return False
                
                user_data = users[uuid]
                expiry = user_data.get('expiry_date')
                
                if expiry:
                    exp_dt = datetime.strptime(expiry, "%Y-%m-%d")
                    if datetime.now() > exp_dt:
                        self.is_activated = False
                        return False
                
                self.is_activated = True
                self.expiry_date = expiry
                return True
            
            return False
        except:
            return False

# ============================================
# FASTAPI APP
# ============================================

app = FastAPI(title="UPMINES OSTP-ISTP CHECKER API", 
              description="Upload Excel file, get ISTP validation results")

# CORS (Allow all origins)
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ============================================
# GLOBALS
# ============================================

USERNAME = ""
PASSWORD = ""
BASE_URL = "http://upmines.upsdc.gov.in/TRANSPORTER/PrintTRANSPORTERFormVehicleCheckValidOrNot.aspx?eId="
INPUT_FILE = "LOADOSTP&ISTP.xlsx"
OUTPUT_FILE = "final_output_fixed.xlsx"

# ============================================
# FETCH CREDENTIALS
# ============================================

def fetch_credentials():
    global USERNAME, PASSWORD
    try:
        headers = {
            'Authorization': f'token {GITHUB_TOKEN}',
            'Accept': 'application/vnd.github.v3.raw'
        }
        response = requests.get(GITHUB_CREDENTIALS_URL, headers=headers, timeout=10)
        
        if response.status_code == 200:
            data = response.json()
            USERNAME = data.get('username', '').strip()
            PASSWORD = data.get('password', '').strip()
            return USERNAME and PASSWORD
        return False
    except:
        return False

# ============================================
# SELENIUM HELPER FUNCTIONS
# ============================================

def get_driver():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-automation", "enable-logging"])
    chrome_options.add_experimental_option('useAutomationExtension', False)
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_argument("--silent")
    
    # ChromeDriver path (use absolute path)
    chromedriver_path = r"C:\Users\Tarik\AppData\Local\Programs\Python\Python314\Scripts\chromedriver.exe"
    service = Service(chromedriver_path)
    driver = webdriver.Chrome(service=service, options=chrome_options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver

def wait_for_element(driver, by, value, timeout=15):
    try:
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
    except TimeoutException:
        return None

def handle_alert(driver):
    try:
        alert = driver.switch_to.alert
        alert_text = alert.text
        alert.accept()
        time.sleep(1)
        return alert_text
    except:
        return None

def login(driver):
    driver.get("https://upmines.upsdc.gov.in/DefaultLicense.aspx")
    time.sleep(3)
    handle_alert(driver)
    
    username_field = wait_for_element(driver, By.ID, "ContentPlaceHolder1_txtAadharNumber", 20)
    if username_field:
        username_field.send_keys(USERNAME)
    
    password_field = wait_for_element(driver, By.ID, "ContentPlaceHolder1_txtPassword", 10)
    if password_field:
        password_field.send_keys(PASSWORD)
    
    # OCR for captcha
    reader = easyocr.Reader(['en'])
    captcha_img = wait_for_element(driver, By.ID, "Captcha", 15)
    if captcha_img:
        captcha_text = "".join(reader.readtext(captcha_img.screenshot_as_png, detail=0)).replace(" ", "").upper()
        captcha_input = wait_for_element(driver, By.ID, "ContentPlaceHolder1_txtCaptcha", 10)
        if captcha_input:
            captcha_input.send_keys(captcha_text)
            
            login_btn = wait_for_element(driver, By.ID, "ContentPlaceHolder1_btn_captcha", 10)
            if login_btn:
                driver.execute_script("arguments[0].click();", login_btn)
                time.sleep(5)
                handle_alert(driver)

# ============================================
# PROCESS SINGLE ROW
# ============================================

def process_single_row(driver, row_data, index):
    try:
        handle_alert(driver)
        
        ostp_value = str(row_data['OSTP']).split('.')[0]
        istp_value = str(row_data['ISTP']).split('.')[0]
        
        time.sleep(2)
        
        try:
            dropdown = Select(wait_for_element(driver, By.ID, "ContentPlaceHolder1_ddl_LicenseeID", 10))
            if dropdown:
                dropdown.select_by_index(1)
        except:
            pass
        
        try:
            radio_btn = wait_for_element(driver, By.ID, "ContentPlaceHolder1_RbtWise_1", 5)
            if radio_btn:
                driver.execute_script("arguments[0].click();", radio_btn)
        except:
            pass
        
        time.sleep(1)
        
        ostp_box = wait_for_element(driver, By.ID, "ContentPlaceHolder1_txtOSTP", 10)
        if ostp_box:
            driver.execute_script(f"arguments[0].value = '{ostp_value}';", ostp_box)
            driver.execute_script("arguments[0].dispatchEvent(new Event('change', {bubbles: true}));", ostp_box)
        
        full_link = BASE_URL + istp_value
        istp_box = wait_for_element(driver, By.ID, "ContentPlaceHolder1_txtISTP", 10)
        if istp_box:
            driver.execute_script(f"arguments[0].value = '{full_link}';", istp_box)
            driver.execute_script("arguments[0].dispatchEvent(new Event('input', {bubbles: true}));", istp_box)
            driver.execute_script("arguments[0].dispatchEvent(new Event('change', {bubbles: true}));", istp_box)
        
        time.sleep(1)
        
        proceed_btn = wait_for_element(driver, By.ID, "ContentPlaceHolder1_btnProceed1", 10)
        if proceed_btn:
            driver.execute_script("arguments[0].click();", proceed_btn)
            time.sleep(3)
        
        # Get result
        alert_text = handle_alert(driver)
        if alert_text:
            return f"Alert: {alert_text}"
        
        try:
            error_lbl = WebDriverWait(driver, 5).until(
                EC.visibility_of_element_located((By.ID, "ContentPlaceHolder1_ErrorLbl"))
            )
            result_text = error_lbl.text.strip()
            if result_text:
                return result_text
        except:
            pass
        
        try:
            result_text = driver.execute_script(
                "return document.getElementById('ContentPlaceHolder1_ErrorLbl')?.innerText || '';"
            ).strip()
            if result_text:
                return result_text
        except:
            pass
        
        return "No message found"
        
    except Exception as e:
        return f"Error: {str(e)[:100]}"

# ============================================
# API ENDPOINTS
# ============================================

@app.get("/")
def read_root():
    return {
        "service": "UPMINES OSTP-ISTP CHECKER API",
        "status": "active",
        "endpoints": {
            "/upload": "POST - Upload Excel file for processing",
            "/status/{task_id}": "GET - Check processing status",
            "/download/{task_id}": "GET - Download processed file",
            "/health": "GET - Health check"
        }
    }

@app.get("/health")
def health_check():
    activation = ActivationManager()
    activated = activation.check_github()
    return {
        "status": "healthy",
        "activated": activated,
        "uuid": activation.user_uuid,
        "expiry": activation.expiry_date
    }

@app.post("/upload")
async def upload_file(file: UploadFile = File(...), background_tasks: BackgroundTasks = BackgroundTasks()):
    # Check activation
    activation = ActivationManager()
    if not activation.check_github():
        raise HTTPException(
            status_code=403, 
            detail=f"License not activated. UUID: {activation.user_uuid}"
        )
    
    # Check credentials
    if not fetch_credentials():
        raise HTTPException(status_code=500, detail="Failed to fetch credentials from GitHub")
    
    # Validate file
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(status_code=400, detail="Only Excel files (.xlsx, .xls) are supported")
    
    # Generate task ID
    task_id = str(uuid.uuid4())[:8]
    task_dir = os.path.join("tasks", task_id)
    os.makedirs(task_dir, exist_ok=True)
    
    # Save uploaded file
    file_path = os.path.join(task_dir, "input.xlsx")
    with open(file_path, "wb") as f:
        shutil.copyfileobj(file.file, f)
    
    # Process in background
    background_tasks.add_task(process_file, task_id, task_dir, file_path)
    
    return JSONResponse(content={
        "task_id": task_id,
        "status": "processing",
        "message": "File uploaded successfully. Processing in background.",
        "status_url": f"/status/{task_id}",
        "download_url": f"/download/{task_id}"
    })

@app.get("/status/{task_id}")
def get_status(task_id: str):
    task_dir = os.path.join("tasks", task_id)
    status_file = os.path.join(task_dir, "status.json")
    
    if not os.path.exists(task_dir):
        raise HTTPException(status_code=404, detail="Task not found")
    
    if os.path.exists(status_file):
        with open(status_file, "r") as f:
            return json.load(f)
    
    return {"task_id": task_id, "status": "processing"}

@app.get("/download/{task_id}")
def download_file(task_id: str):
    task_dir = os.path.join("tasks", task_id)
    output_file = os.path.join(task_dir, "output.xlsx")
    
    if not os.path.exists(output_file):
        raise HTTPException(status_code=404, detail="Output file not found or still processing")
    
    return FileResponse(
        path=output_file,
        filename=f"upmines_result_{task_id}.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ============================================
# BACKGROUND PROCESSING
# ============================================

def process_file(task_id, task_dir, file_path):
    status_file = os.path.join(task_dir, "status.json")
    output_file = os.path.join(task_dir, "output.xlsx")
    
    try:
        # Update status
        with open(status_file, "w") as f:
            json.dump({"task_id": task_id, "status": "processing", "progress": 0}, f)
        
        # Read Excel
        df = pd.read_excel(file_path, dtype=str)
        if "Form_Result" not in df.columns:
            df["Form_Result"] = ""
        
        total_rows = len(df)
        
        # Initialize driver
        driver = get_driver()
        login(driver)
        
        driver.get("https://upmines.upsdc.gov.in/licensee/Apply_for_addQuantity_ByeMM11.aspx")
        time.sleep(5)
        handle_alert(driver)
        
        # Process each row
        for index, row in df.iterrows():
            result = process_single_row(driver, row, index)
            df.at[index, "Form_Result"] = result
            df.to_excel(output_file, index=False)
            
            # Update progress
            progress = int((index + 1) / total_rows * 100)
            with open(status_file, "w") as f:
                json.dump({
                    "task_id": task_id, 
                    "status": "processing", 
                    "progress": progress,
                    "current_row": index + 1,
                    "total_rows": total_rows
                }, f)
        
        # Cleanup driver
        try:
            driver.quit()
        except:
            pass
        
        # Final status
        with open(status_file, "w") as f:
            json.dump({
                "task_id": task_id, 
                "status": "completed", 
                "progress": 100,
                "total_rows": total_rows,
                "output_file": "output.xlsx",
                "download_url": f"/download/{task_id}"
            }, f)
        
    except Exception as e:
        error_msg = str(e)
        with open(status_file, "w") as f:
            json.dump({
                "task_id": task_id, 
                "status": "failed", 
                "error": error_msg[:200]
            }, f)

# ============================================
# STARTUP
# ============================================

if __name__ == "__main__":
    import uvicorn
    # Create tasks directory
    os.makedirs("tasks", exist_ok=True)
    
    # Check activation on startup
    activation = ActivationManager()
    print(f"🔐 Activation Status: {activation.check_github()}")
    print(f"📱 UUID: {activation.user_uuid}")
    
    # Start server
    uvicorn.run(app, host="0.0.0.0", port=8000)

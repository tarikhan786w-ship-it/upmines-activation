# -*- coding: utf-8 -*-
"""
UPMINES OSTP-ISTP CHECKER - API VERSION (Render Compatible)
"""
import os
import shutil
import time
import uuid
import json
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
import requests

# ============================================
# 🔑 GITHUB CONFIGURATION
# ============================================

GITHUB_TOKEN = "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"
GITHUB_CREDENTIALS_URL = "https://raw.githubusercontent.com/tarikhan786w-ship-it/upmines-activation/refs/heads/main/credentials.json"
GITHUB_ACTIVATION_URL = "https://raw.githubusercontent.com/tarikhan786w-ship-it/upmines-activation/refs/heads/main/activated_users.json"

# ============================================
# FASTAPI APP
# ============================================

app = FastAPI(title="UPMINES OSTP-ISTP CHECKER API")

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

# ============================================
# FETCH CREDENTIALS
# ============================================

def fetch_credentials():
    global USERNAME, PASSWORD
    try:
        headers = {'Authorization': f'token {GITHUB_TOKEN}', 'Accept': 'application/vnd.github.v3.raw'}
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
# SELENIUM HELPERS (Render Compatible)
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
    
    # Render par ChromeDriver auto-install hota hai
    service = Service()
    driver = webdriver.Chrome(service=service, options=chrome_options)
    return driver

def wait_for_element(driver, by, value, timeout=15):
    try:
        return WebDriverWait(driver, timeout).until(EC.presence_of_element_located((by, value)))
    except TimeoutException:
        return None

def handle_alert(driver):
    try:
        alert = driver.switch_to.alert
        text = alert.text
        alert.accept()
        return text
    except:
        return None

def login(driver):
    driver.get("https://upmines.upsdc.gov.in/DefaultLicense.aspx")
    time.sleep(3)
    handle_alert(driver)
    
    # Manual captcha input (API user se expect karega)
    # Simplified version - actual implementation mein captcha handling alag se karni hogi
    return True

# ============================================
# PROCESS FUNCTIONS
# ============================================

def process_single_row(driver, row_data, index):
    try:
        # Simplified processing - actual logic yahan aayegi
        return "Processed"
    except Exception as e:
        return f"Error: {str(e)[:100]}"

def process_file(task_id, task_dir, file_path):
    status_file = os.path.join(task_dir, "status.json")
    output_file = os.path.join(task_dir, "output.xlsx")
    
    try:
        with open(status_file, "w") as f:
            json.dump({"task_id": task_id, "status": "processing", "progress": 0}, f)
        
        df = pd.read_excel(file_path, dtype=str)
        if "Form_Result" not in df.columns:
            df["Form_Result"] = ""
        
        total_rows = len(df)
        driver = get_driver()
        login(driver)
        
        for index, row in df.iterrows():
            result = process_single_row(driver, row, index)
            df.at[index, "Form_Result"] = result
            df.to_excel(output_file, index=False)
            
            progress = int((index + 1) / total_rows * 100)
            with open(status_file, "w") as f:
                json.dump({"task_id": task_id, "status": "processing", "progress": progress}, f)
        
        driver.quit()
        
        with open(status_file, "w") as f:
            json.dump({"task_id": task_id, "status": "completed", "progress": 100}, f)
            
    except Exception as e:
        with open(status_file, "w") as f:
            json.dump({"task_id": task_id, "status": "failed", "error": str(e)[:200]}, f)

# ============================================
# API ENDPOINTS
# ============================================

@app.get("/")
def root():
    return {"service": "UPMINES API", "status": "active"}

@app.get("/health")
def health():
    return {"status": "healthy"}

@app.post("/upload")
async def upload_file(file: UploadFile = File(...), background_tasks: BackgroundTasks = BackgroundTasks()):
    if not fetch_credentials():
        raise HTTPException(status_code=500, detail="Credentials not available")
    
    task_id = str(uuid.uuid4())[:8]
    task_dir = os.path.join("tasks", task_id)
    os.makedirs(task_dir, exist_ok=True)
    
    file_path = os.path.join(task_dir, "input.xlsx")
    with open(file_path, "wb") as f:
        shutil.copyfileobj(file.file, f)
    
    background_tasks.add_task(process_file, task_id, task_dir, file_path)
    
    return {
        "task_id": task_id,
        "status": "processing",
        "status_url": f"/status/{task_id}",
        "download_url": f"/download/{task_id}"
    }

@app.get("/status/{task_id}")
def get_status(task_id: str):
    status_file = os.path.join("tasks", task_id, "status.json")
    if not os.path.exists(status_file):
        return {"task_id": task_id, "status": "processing"}
    with open(status_file, "r") as f:
        return json.load(f)

@app.get("/download/{task_id}")
def download_file(task_id: str):
    output_file = os.path.join("tasks", task_id, "output.xlsx")
    if not os.path.exists(output_file):
        raise HTTPException(status_code=404, detail="File not found")
    return FileResponse(output_file, filename=f"result_{task_id}.xlsx")

# ============================================
# STARTUP
# ============================================

if __name__ == "__main__":
    import uvicorn
    os.makedirs("tasks", exist_ok=True)
    uvicorn.run(app, host="0.0.0.0", port=10000)

import os
import sys
import re
import logging
import time
from pathlib import Path
from datetime import datetime
import pytz
import pandas as pd
import requests
from google.oauth2 import service_account
import gspread
from gspread_dataframe import set_with_dataframe
from dotenv import load_dotenv

# ===== Setup Logging =====
logging.basicConfig(stream=sys.stdout, level=logging.INFO)
log = logging.getLogger()

# ===== Load environment variables =====
load_dotenv()
ODOO_URL = os.getenv("ODOO_URL")
DB = os.getenv("ODOO_DB")
USERNAME = os.getenv("ODOO_USERNAME")
PASSWORD = os.getenv("ODOO_PASSWORD")

COMPANIES = {
    1: "Zipper",
    3: "Metal Trims",
}

SHEET_INFO = {
    "zipper": {
        "sheet_id": "1tKeLYTb7QxTX_LaI3BkOnPNdnI9mXYhE_yO1ARA0dXM",
        "worksheet_name": "workCenter DF_ZP"
    },
    "metal_trims": {
        "sheet_id": "1tKeLYTb7QxTX_LaI3BkOnPNdnI9mXYhE_yO1ARA0dXM",
        "worksheet_name": "workCenter DF_MT"
    }
}

DOWNLOAD_DIR = os.path.join(os.getcwd(), "download")
os.makedirs(DOWNLOAD_DIR, exist_ok=True)

session = requests.Session()
USER_ID = None

# ===== Utility Functions =====
def login():
    global USER_ID
    payload = {
        "jsonrpc": "2.0",
        "params": {
            "db": DB,
            "login": USERNAME,
            "password": PASSWORD
        }
    }
    r = session.post(f"{ODOO_URL}/web/session/authenticate", json=payload)
    r.raise_for_status()
    result = r.json().get("result")
    if result and "uid" in result:
        USER_ID = result["uid"]
        log.info(f"✅ Logged in (uid={USER_ID})")
        return result
    else:
        raise Exception("❌ Login failed")

def switch_company(company_id):
    if USER_ID is None:
        raise Exception("User not logged in yet")
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "res.users",
            "method": "write",
            "args": [[USER_ID], {"company_id": company_id}],
            "kwargs": {
                "context": {
                    "allowed_company_ids": [company_id],
                    "company_id": company_id
                }
            }
        }
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()
    if "error" in r.json():
        log.error(f"❌ Failed to switch to company {company_id}: {r.json()['error']}")
        return False
    else:
        log.info(f"🔄 Session switched to company {company_id}")
        return True

def get_string_value(field):
    if isinstance(field, dict):
        return field.get("display_name", "")
    elif field is False or field is None:
        return ""
    return str(field)

def fetch_stock_lot(company_id, cname):
    context = {"allowed_company_ids": [company_id], "company_id": company_id, "lang": "en_US", "tz": "Asia/Dhaka", "uid": USER_ID, "bin_size": True, "current_company_id": company_id, "display_complete": True, "default_company_id": company_id}
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "stock.lot",
            "method": "web_search_read",
            "args": [],
            "kwargs": {
                "specification": {
                    "name": {},
                    "ref": {},
                    "product_qty": {},
                    "unit_price": {},
                    "rejected": {},
                    "product_id": {"fields": {"display_name": {}, "categ_id": {"fields": {"display_name": {}}}}},
                    "create_date": {},
                    "company_id": {"fields": {"display_name": {}}},
                    "machine_name": {},
                    "work_center": {},
                },
                "offset": 0,
                "limit": 5000,
                "context": context,
                "count_limit": 10000,
                "domain": [["machine_name", "!=", False]],
            },
        },
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw/stock.lot/web_search_read", json=payload)
    r.raise_for_status()
    try:
        data = r.json()["result"]["records"]
        def flatten_stock_lot(record):
            product = record.get("product_id", {})
            return {
                "Lot/Serial Number": record.get("name", ""),
                "Internal Reference": record.get("ref", ""),
                "On Hand Quantity": record.get("product_qty", 0),
                "Unit Price": record.get("unit_price", 0),
                "Rejected": record.get("rejected", ""),
                "Product": get_string_value(product),
                "Created on": record.get("create_date", ""),
                "Company": get_string_value(record.get("company_id")),
                "Product/Product Category/Display Name": get_string_value(product.get("categ_id", {})),
                "Machine Name": record.get("machine_name", ""),
                "Work Center": record.get("work_center", ""),
            }
        flattened = [flatten_stock_lot(rec) for rec in data]
        log.info(f"📊 {cname}: {len(flattened)} rows fetched (flattened)")
        return flattened
    except Exception:
        log.error(f"❌ {cname}: Failed to parse report: {r.text[:200]}")
        return []

# ====== Function to save records using regex-friendly pattern ======
def save_records_to_excel(records, company_name):
    if records:
        df = pd.DataFrame(records)
        company_clean = re.sub(r'\W+', '_', company_name.lower())
        output_file = os.path.join(DOWNLOAD_DIR, f"{company_clean}_stock_lot.xlsx")
        df.to_excel(output_file, index=False)
        log.info(f"📂 Saved: {output_file}")
        return output_file
    else:
        log.warning(f"❌ No data fetched for {company_name}")
        return None

# ====== Function to paste downloaded files into Google Sheet ======
def paste_downloaded_file_to_gsheet(company_name, sheet_key, worksheet_name):
    try:
        company_clean = re.sub(r'\W+', '_', company_name.lower())
        files = list(Path(DOWNLOAD_DIR).glob(f"{company_clean}_stock_lot*.xlsx"))
        if not files:
            log.warning(f"⚠️ No downloaded file found for {company_name}")
            return
        
        files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        latest_file = files[0]
        df = pd.read_excel(latest_file)
        
        log.info(f"✅ Loaded file {latest_file.name} into DataFrame")

        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = service_account.Credentials.from_service_account_file('gcreds.json', scopes=scope)
        client = gspread.authorize(creds)
        
        sheet = client.open_by_key(sheet_key)
        worksheet = sheet.worksheet(worksheet_name)
        
        if df.empty:
            log.warning(f"⚠️ DataFrame for {company_name} is empty. Skipping paste.")
            return
        df = df.replace(False, "") 
        worksheet.batch_clear(["A:K"])
        time.sleep(2)
        set_with_dataframe(worksheet, df)
        log.info(f"✅ Data pasted into Google Sheet ({worksheet_name}) for {company_name}")
        
        local_tz = pytz.timezone('Asia/Dhaka')
        local_time = datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")
        worksheet.update("L1", [[f"{local_time}"]])
        log.info(f"✅ Timestamp updated: {local_time}")
        
    except Exception as e:
        log.error(f"❌ Error in paste_downloaded_file_to_gsheet({company_name}): {e}")

# ====== Main Workflow ======
if __name__ == "__main__":
    userinfo = login()
    log.info(f"User info (allowed companies): {userinfo.get('user_companies', {})}")

    for cid, cname in COMPANIES.items():
        if switch_company(cid):
            records = fetch_stock_lot(cid, cname)
            save_records_to_excel(records, cname)
            # Push to Google Sheet
            sheet_key = SHEET_INFO[re.sub(r'\W+', '_', cname.lower())]["sheet_id"]
            worksheet_name = SHEET_INFO[re.sub(r'\W+', '_', cname.lower())]["worksheet_name"]
            paste_downloaded_file_to_gsheet(cname, sheet_key, worksheet_name)
import os
import sys
import re
import logging
import time
from pathlib import Path
from datetime import date, datetime, timedelta
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

SHEET_KEY = "1z6Zb_BronrO26rNS_gCKmsetoY7_OFysfIyvU3iazy0"

# ===== Calculate date ranges =====
today = date.today()
yesterday = today - timedelta(days=1)
current_first = today.replace(day=1)
prev_last_month = current_first - timedelta(days=1)
prev_first = prev_last_month.replace(day=1)
prev_last = current_first - timedelta(days=1)

reports = [
    {"type": "cs", "from_date": current_first.isoformat(), "to_date": today.isoformat()},
    {"type": "ld", "from_date": current_first.isoformat(), "to_date": yesterday.isoformat()},
    {"type": "lm", "from_date": prev_first.isoformat(), "to_date": prev_last.isoformat()},
]

worksheet_map = {
    "Zipper": {
        "cs": "zip_fg_cs",
        "ld": "zip_fg_ld",
        "lm": "zip_fg_LM"
    },
    "Metal Trims": {
        "cs": "mt_fg_cs",
        "ld": "mt_fg_ld",
        "lm": "mt_fg_LM"
    }
}

target_companies = {
    1: "Zipper",
    3: "Metal Trims"
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

def fetch_fg_store_datas(company_id, cname, from_date, to_date):
    context = {
        "lang": "en_US",
        "tz": "Asia/Dhaka",
        "uid": USER_ID,
        "allowed_company_ids": [company_id]
    }
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "operation.details",
            "method": "retrieve_fg_store_datas",
            "args": [[company_id], from_date, to_date],
            "kwargs": {"context": context}
        }
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw/operation.details/retrieve_fg_store_datas", json=payload)
    r.raise_for_status()
    try:
        data = r.json()["result"]
        if isinstance(data, list):
            def flatten_record(record):
                return {
                    k: v[1] if isinstance(v, list) and len(v) == 2 else 
                    (v.get("display_name") if isinstance(v, dict) and "display_name" in v else v) 
                    for k, v in record.items()
                }
            flattened = [flatten_record(rec) for rec in data if isinstance(rec, dict)]
            log.info(f"📊 {cname}: {len(flattened)} rows fetched (flattened)")
            return flattened
        else:
            log.warning(f"⚠️ Unexpected data format for {cname}: {type(data)}")
            return []
    except Exception as e:
        log.error(f"❌ {cname}: Failed to parse data: {e}")
        return []

# ====== Function to save records using regex-friendly pattern ======
def save_records_to_excel(records, company_name, report_type, to_date):
    if records:
        df = pd.DataFrame(records)
        company_clean = re.sub(r'\W+', '_', company_name.lower())
        output_file = os.path.join(DOWNLOAD_DIR, f"{company_clean}_fg_store_datas_{report_type}_{to_date}.xlsx")
        df.to_excel(output_file, index=False)
        log.info(f"📂 Saved: {output_file}")
        return output_file
    else:
        log.warning(f"❌ No data fetched for {company_name} ({report_type})")
        return None

# ====== Function to paste downloaded files into Google Sheet ======
def paste_downloaded_file_to_gsheet(company_name, sheet_key, worksheet_name, report_type):
    try:
        company_clean = re.sub(r'\W+', '_', company_name.lower())
        files = list(Path(DOWNLOAD_DIR).glob(f"{company_clean}_fg_store_datas_{report_type}_*.xlsx"))
        if not files:
            log.warning(f"⚠️ No downloaded file found for {company_name} ({report_type})")
            return
        
        files.sort(key=lambda x: x.stat().st_mtime, reverse=True)
        latest_file = files[0]
        df = pd.read_excel(latest_file)
        
        # Drop first column if exists
        if df.shape[1] > 1:
            df = df.iloc[:, 1:]
        
        log.info(f"✅ Loaded file {latest_file.name} into DataFrame (first column dropped)")

        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        creds = service_account.Credentials.from_service_account_file('gcreds.json', scopes=scope)
        client = gspread.authorize(creds)
        
        sheet = client.open_by_key(sheet_key)
        worksheet = sheet.worksheet(worksheet_name)
        
        if df.empty:
            log.warning(f"⚠️ DataFrame for {company_name} ({report_type}) is empty. Skipping paste.")
            return
        df = df.replace(False, "") 
        worksheet.batch_clear(["A:N"])
        time.sleep(2)
        set_with_dataframe(worksheet, df)
        log.info(f"✅ Data pasted into Google Sheet ({worksheet_name}) for {company_name} ({report_type})")
        
        local_tz = pytz.timezone('Asia/Dhaka')
        local_time = datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")
        worksheet.update([[f"{local_time}"]], "O1")
        log.info(f"✅ Timestamp updated: {local_time}")
        
    except Exception as e:
        log.error(f"❌ Error in paste_downloaded_file_to_gsheet({company_name}, {report_type}): {e}")

# ====== Main Workflow ======
if __name__ == "__main__":
    userinfo = login()
    log.info(f"User info (allowed companies): {userinfo.get('user_companies', {})}")

    for cid, cname in target_companies.items():
        if not switch_company(cid):
            log.warning(f"⚠️ Failed to switch to company {cid} ({cname}), skipping...")
            continue

        log.info(f"🔍 Processing {cname} (Company ID: {cid})")

        for report in reports:
            from_date = report["from_date"]
            to_date = report["to_date"]
            report_type = report["type"]
            worksheet_name = worksheet_map.get(cname, {}).get(report_type)
            if not worksheet_name:
                log.warning(f"⚠️ No worksheet mapping for {cname} ({report_type}), skipping...")
                continue
            log.info(f"Processing report {report_type}: FROM_DATE={from_date}, TO_DATE={to_date}")
            records = fetch_fg_store_datas(cid, cname, from_date, to_date)
            save_records_to_excel(records, cname, report_type, to_date)
            paste_downloaded_file_to_gsheet(cname, SHEET_KEY, worksheet_name, report_type)
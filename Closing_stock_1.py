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

COMPANIES = {
    1: "Zipper",
    3: "Metal Trims",
}

SHEET_INFO = {
    "zipper": {
        "sheet_id": "1z6Zb_BronrO26rNS_gCKmsetoY7_OFysfIyvU3iazy0",
        "worksheet_name": "Closing_stock_1_zip"
    },
    "metal_trims": {
        "sheet_id": "1z6Zb_BronrO26rNS_gCKmsetoY7_OFysfIyvU3iazy0",
        "worksheet_name": "Closing_stock_1_mt"
    }
}

# ===== Default: previous month 1st to last day, ignoring env for TO_DATE =====
today = date.today()
from_date_env = os.getenv("FROM_DATE", "").strip()
to_date_env = os.getenv("TO_DATE", "").strip()

first_this_month = today.replace(day=1)
last_prev_month = first_this_month - timedelta(days=1)
first_prev_month = last_prev_month.replace(day=1)

FROM_DATE = from_date_env if from_date_env else first_prev_month.isoformat()
# Always set TO_DATE as last day of previous month, ignoring env if set
TO_DATE = last_prev_month.isoformat()

log.info(f"Using FROM_DATE={FROM_DATE}, TO_DATE={TO_DATE}")

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

def create_forecast_wizard(company_id, from_date, to_date):
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "stock.forecast.report",
            "method": "create",
            "args": [{"from_date": from_date, "to_date": to_date}],
            "kwargs": {"context": {"allowed_company_ids": [company_id], "company_id": company_id}}
        }
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()
    wizard_id = r.json()["result"]
    log.info(f"🪄 Created wizard {wizard_id} for company {company_id}")
    return wizard_id

def compute_forecast(company_id, wizard_id):
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "stock.forecast.report",
            "method": "print_date_wise_stock_register",
            "args": [[wizard_id]],
            "kwargs": {
                "context": {
                    "lang": "en_US",
                    "tz": "Asia/Dhaka",
                    "uid": USER_ID,
                    "allowed_company_ids": [company_id],
                    "company_id": company_id
                }
            }
        }
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_button", json=payload)
    r.raise_for_status()
    log.info(f"⚡ Forecast computed for wizard {wizard_id} (company {company_id})")
    return r.json()

def fetch_opening_closing(company_id, cname, wizard_id):
    context = {
        "allowed_company_ids": [company_id],
        "company_id": company_id,
        "active_model": "stock.forecast.report",
        "active_id": wizard_id,
        "active_ids": [wizard_id]
    }
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "stock.opening.closing",
            "method": "web_search_read",
            "args": [],
            "kwargs": {
                "specification": {
                    "product_category": {"fields": {"display_name": {}}},
                    "classification_id": {"fields": {"display_name": {}}},
                    "cloing_qty": {},
                    "cloing_value": {},
                    "lot_id": {"fields": {"display_name": {}}},
                    "issue_qty": {},
                    "issue_value": {},
                    "product_id": {"fields": {"display_name": {}}},
                    "pr_code": {},
                    "landed_cost": {},
                    "opening_qty": {},
                    "opening_value": {},
                    "po_type": {},
                    "lot_price": {},
                    "parent_category": {"fields": {"display_name": {}}},
                    "pur_price": {},
                    "receive_date": {},
                    "receive_qty": {},
                    "receive_value": {},
                    "rejected": {},
                    "shipment_mode": {},
                    "product_uom": {"fields": {"display_name": {}}},
                    "partner_id": {"fields": {"display_name": {}}},
                    "po_number": {},
                    "product_type": {"fields": {"display_name": {}}},
                    "item_category": {"fields": {"display_name": {}}},
                },
                "offset": 0,
                "limit": 5000,
                "context": context,
                "count_limit": 10000,
                "domain": [["product_id.categ_id.complete_name", "ilike", "All / RM"]],
            },
        },
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()
    try:
        data = r.json()["result"]["records"]
        def flatten_record(record):
            return {k: v.get("display_name") if isinstance(v, dict) and "display_name" in v else v for k, v in record.items()}
        flattened = [flatten_record(rec) for rec in data]
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
        output_file = os.path.join(DOWNLOAD_DIR, f"{company_clean}_opening_closing_{TO_DATE}.xlsx")
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
        files = list(Path(DOWNLOAD_DIR).glob(f"{company_clean}_opening_closing_*.xlsx"))
        if not files:
            log.warning(f"⚠️ No downloaded file found for {company_name}")
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
            log.warning(f"⚠️ DataFrame for {company_name} is empty. Skipping paste.")
            return
        df = df.replace(False, "") 
        worksheet.batch_clear(["A:AA"])
        time.sleep(2)
        set_with_dataframe(worksheet, df)
        log.info(f"✅ Data pasted into Google Sheet ({worksheet_name}) for {company_name}")
        
        local_tz = pytz.timezone('Asia/Dhaka')
        local_time = datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")
        worksheet.update([[f"{local_time}"]], "AA2")
        log.info(f"✅ Timestamp updated: {local_time}")
        
    except Exception as e:
        log.error(f"❌ Error in paste_downloaded_file_to_gsheet({company_name}): {e}")

# ====== Main Workflow ======
if __name__ == "__main__":
    userinfo = login()
    log.info(f"User info (allowed companies): {userinfo.get('user_companies', {})}")

    for cid, cname in COMPANIES.items():
        log.info(f"\n🚀 Processing company: {cname} (ID={cid})")
        success = False

        for attempt in range(1, 2):  # Retry up to 30 times for this company
            try:
                if switch_company(cid):
                    wiz_id = create_forecast_wizard(cid, FROM_DATE, TO_DATE)
                    compute_forecast(cid, wiz_id)
                    records = fetch_opening_closing(cid, cname,wiz_id)
                    save_records_to_excel(records, cname)

                    # Push to Google Sheet
                    sheet_key = SHEET_INFO[re.sub(r'\W+', '_', cname.lower())]["sheet_id"]
                    worksheet_name = SHEET_INFO[re.sub(r'\W+', '_', cname.lower())]["worksheet_name"]
                    paste_downloaded_file_to_gsheet(cname, sheet_key, worksheet_name)

                    log.info(f"✅ Completed successfully for {cname} (Attempt {attempt})")
                    success = True
                    break

            except Exception as e:
                log.warning(f"⚠️ Attempt {attempt}/2 failed for {cname}: {e}")
                if attempt < 2:
                    wait_time = min(60, 5 * attempt)  # backoff delay, max 60s
                    log.info(f"🔁 Retrying {cname} in {wait_time}s...")
                    time.sleep(wait_time)
                else:
                    log.error(f"❌ Max retries reached for {cname}. Moving to next company.")

        if not success:
            log.error(f"🚫 Skipping {cname} after 2 failed attempts.\n")

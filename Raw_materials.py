import os
import sys
import time
import logging
import pytz
import re
from pathlib import Path
from datetime import date, datetime
import requests
import pandas as pd
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

# ===== Companies & Google Sheet info =====
COMPANIES = {
    1: "Zipper",
    3: "Metal Trims",
}

SHEET_INFO = {
    "zipper": {
        "sheet_id": "1z6Zb_BronrO26rNS_gCKmsetoY7_OFysfIyvU3iazy0",
        "worksheet_name": "STD_ITEM_STOCK"
    },
    "metal_trims": {
        "sheet_id": "1kD4iCUqEAQsE_CLuv3dFSFNSjD2Hj2dTrE40deGZaK0",
        "worksheet_name": "odoo_data"
    }
}

# ===== Default date =====
today = date.today()
FROM_DATE = os.getenv("FROM_DATE", today.replace(day=1).isoformat())
TO_DATE = os.getenv("TO_DATE", today.isoformat())

log.info(f"Using FROM_DATE={FROM_DATE}, TO_DATE={TO_DATE}")

session = requests.Session()
USER_ID = None

# ===== Login =====
def login():
    global USER_ID
    payload = {"jsonrpc": "2.0", "params": {"db": DB, "login": USERNAME, "password": PASSWORD}}
    r = session.post(f"{ODOO_URL}/web/session/authenticate", json=payload)
    r.raise_for_status()
    result = r.json().get("result")
    if result and "uid" in result:
        USER_ID = result["uid"]
        log.info(f"✅ Logged in (uid={USER_ID})")
        return result
    else:
        raise Exception("❌ Login failed")

# ===== Switch company =====
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
            "kwargs": {"context": {"allowed_company_ids": [company_id], "company_id": company_id}}
        }
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()
    if "error" in r.json():
        log.error(f"❌ Failed to switch to company {company_id}: {r.json()['error']}")
        return False
    log.info(f"🔄 Session switched to company {company_id}")
    return True

# ===== Fetch Raw Material Products =====
def fetch_raw_materials(company_id, cname):
    context = {"lang": "en_US","tz": "Asia/Dhaka","uid": USER_ID,"allowed_company_ids": [company_id],"current_company_id": company_id}
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "product.template",
            "method": "web_search_read",
            "args": [],
            "kwargs": {
                "specification": {
                    "categ_type": {"fields": {"display_name": {}}},
                    "default_code": {},
                    "name": {},
                    "categ_id": {"fields": {"display_name": {}}},
                    "qty_available": {},
                    "generic_name": {}
                },
                "offset": 0,
                "limit": 10000,
                "context": context,
                "count_limit": 100000,
                "domain": [
                    "&",
                    ["categ_id", "ilike", "ALL / RM /"],
                    ["default_code", "ilike", "R_"]
                ],
            },
        },
    }

    r = session.post(f"{ODOO_URL}/web/dataset/call_kw/product.template/web_search_read", json=payload)
    r.raise_for_status()
    try:
        records = r.json()["result"]["records"]
        def flatten(record):
            return {k: (v.get("display_name") if isinstance(v, dict) and "display_name" in v else v) for k,v in record.items()}
        flattened = [flatten(rec) for rec in records]
        log.info(f"📦 {cname}: {len(flattened)} raw material product rows fetched")
        return flattened
    except Exception:
        log.error(f"❌ {cname}: Failed to parse product data: {r.text[:200]}")
        return []

# ===== Save to Excel & Paste to Google Sheet =====
def save_and_paste_to_sheet(records, cname):
    if not records:
        log.warning(f"❌ No data for {cname}")
        return
    df = pd.DataFrame(records)
    file_name = f"{cname.lower().replace(' ','_')}_raw_materials_{today.isoformat()}.xlsx"
    df.to_excel(file_name, index=False)
    log.info(f"📂 Saved: {file_name}")

    # Google Sheet
    scope = ["https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"]
    creds = service_account.Credentials.from_service_account_file('gcreds.json', scopes=scope)
    client = gspread.authorize(creds)

    sheet_key = SHEET_INFO[re.sub(r'\W+', '_', cname.lower())]["sheet_id"]
    worksheet_name = SHEET_INFO[re.sub(r'\W+', '_', cname.lower())]["worksheet_name"]
    worksheet = client.open_by_key(sheet_key).worksheet(worksheet_name)

    if df.empty:
        log.warning("Skip: DataFrame empty, not pasting.")
        return

    worksheet.clear()
    time.sleep(4)
    set_with_dataframe(worksheet, df)
    worksheet.update('G1', [['Date']])
    local_tz = pytz.timezone('Asia/Dhaka')
    local_time = datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")
    worksheet.update('G2', [[f"{local_time}"]])
    log.info(f"✅ Data pasted to {worksheet_name} with timestamp {local_time}")

# ===== Main =====
if __name__ == "__main__":
    userinfo = login()
    log.info(f"User info: {userinfo.get('user_companies',{})}")
    for cid, cname in COMPANIES.items():
        if switch_company(cid):
            records = fetch_raw_materials(cid, cname)
            save_and_paste_to_sheet(records, cname)

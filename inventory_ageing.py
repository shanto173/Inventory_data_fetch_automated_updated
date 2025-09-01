import requests
import json
import re
import logging
import sys
import os
from datetime import date, datetime, timedelta
import calendar
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2 import service_account
import pandas as pd
import pytz
from dotenv import load_dotenv
load_dotenv()
logging.basicConfig(stream=sys.stdout, level=logging.INFO)
log = logging.getLogger()

# ========= CONFIG ==========
ODOO_URL = os.getenv("ODOO_URL")
DB = os.getenv("ODOO_DB")
USERNAME = os.getenv("ODOO_USERNAME")
PASSWORD = os.getenv("ODOO_PASSWORD")

COMPANIES = {
    1: "Zipper",
    3: "Metal Trims",
}

# ---------------------- DATE HANDLING ----------------------
TO_DATE = os.getenv("TO_DATE")
today = date.today()
if not TO_DATE:
    TO_DATE = today.isoformat()
FROM_DATE = False  # First of month, only needed by wizard

print(f"📅 Report date: {TO_DATE}")

# ---------------------- ODOO SESSION ----------------------
session = requests.Session()
USER_ID = None

def login():
    global USER_ID
    payload = {"jsonrpc": "2.0","params":{"db": DB,"login": USERNAME,"password": PASSWORD}}
    r = session.post(f"{ODOO_URL}/web/session/authenticate", json=payload)
    r.raise_for_status()
    result = r.json().get("result")
    if result and "uid" in result:
        USER_ID = result["uid"]
        print(f"✅ Logged in (uid={USER_ID})")
        return result
    else:
        raise Exception("❌ Login failed")

def switch_company(company_id):
    payload = {"jsonrpc":"2.0","method":"call","params":{"model":"res.users",
        "method":"write","args":[[USER_ID],{"company_id":company_id}],
        "kwargs":{"context":{"allowed_company_ids":[company_id],"company_id":company_id}}}}
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()
    if "error" in r.json():
        print(f"❌ Failed to switch to company {company_id}")
        return False
    else:
        print(f"🔄 Session switched to company {company_id}")
        return True

def create_forecast_wizard(company_id, from_date, to_date):
    payload = {"jsonrpc":"2.0","method":"call","params":{
        "model":"stock.forecast.report",
        "method":"create",
        "args":[{"from_date":from_date,"to_date":to_date}],
        "kwargs":{"context":{"allowed_company_ids":[company_id],"company_id":company_id}}
    }}
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()
    wizard_id = r.json()["result"]
    print(f"🪄 Created wizard {wizard_id} for company {company_id}")
    return wizard_id

def compute_forecast(company_id, wizard_id):
    payload = {"jsonrpc":"2.0","method":"call","params":{
        "model":"stock.forecast.report",
        "method":"print_date_wise_stock_register",
        "args":[[wizard_id]],
        "kwargs":{"context":{"lang":"en_US","tz":"Asia/Dhaka","uid":USER_ID,"allowed_company_ids":[company_id],"company_id":company_id}}
    }}
    r = session.post(f"{ODOO_URL}/web/dataset/call_button", json=payload)
    r.raise_for_status()
    print(f"⚡ Forecast computed for wizard {wizard_id} (company {company_id})")
    return r.json()

def fetch_opening_closing(company_id, cname):
    payload = {"jsonrpc":"2.0","method":"call","params":{
        "model":"stock.opening.closing",
        "method":"web_search_read",
        "args":[],
        "kwargs":{"specification":{},"offset":0,"limit":5000,"context":{"allowed_company_ids":[company_id],"company_id":company_id},"count_limit":10000}
    }}
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()
    try:
        data = r.json()["result"]["records"]
        flattened = []
        for rec in data:
            flat = {k:(v.get("display_name") if isinstance(v, dict) and "display_name" in v else v) for k,v in rec.items()}
            flattened.append(flat)
        print(f"📊 {cname}: {len(flattened)} rows fetched")
        return flattened
    except Exception:
        print(f"❌ {cname}: Failed to parse report")
        return []

# ---------------------- GOOGLE SHEETS ----------------------
SHEETS_CONFIG = {
    1: {"key":"1z6Zb_BronrO26rNS_gCKmsetoY7_OFysfIyvU3iazy0","sheet":"age_ZIP"},
    3: {"key":"1z6Zb_BronrO26rNS_gCKmsetoY7_OFysfIyvU3iazy0","sheet":"age_MT"}
}

def paste_to_sheet(df, company_id):
    if df.empty:
        print("⚠️ DataFrame empty, skipping paste.")
        return
    cfg = SHEETS_CONFIG[company_id]
    creds = service_account.Credentials.from_service_account_file('gcreds.json', scopes=[
        "https://www.googleapis.com/auth/spreadsheets","https://www.googleapis.com/auth/drive"])
    client = gspread.authorize(creds)
    sheet = client.open_by_key(cfg["key"])
    worksheet = sheet.worksheet(cfg["sheet"])
    worksheet.clear()
    time.sleep(2)
    set_with_dataframe(worksheet, df)
    local_time = datetime.now(pytz.timezone("Asia/Dhaka")).strftime("%Y-%m-%d %H:%M:%S")
    worksheet.update("W2", [[local_time]])
    print(f"✅ Data pasted for company {company_id}, timestamp {local_time}")

# ---------------------- MAIN ----------------------
if __name__ == "__main__":
    userinfo = login()
    print("User info:", userinfo.get("user_companies", {}))

    for cid, cname in COMPANIES.items():
        if switch_company(cid):
            wiz_id = create_forecast_wizard(cid, FROM_DATE, TO_DATE)
            compute_forecast(cid, wiz_id)
            records = fetch_opening_closing(cid, cname)
            if records:
                df = pd.DataFrame(records)
                output_file = f"{cname.lower().replace(' ','_')}_opening_closing_{TO_DATE}.xlsx"
                df.to_excel(output_file, index=False)
                print(f"📂 Saved: {output_file}")
                paste_to_sheet(df, cid)
            else:
                print(f"❌ No data fetched for {cname}")

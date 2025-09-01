import requests
import json
import re
import logging
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

MODEL = "ppc.report"
REPORT_BUTTON_METHOD = "action_generate_xlsx_report"

# 🔹 Configurable
REPORT_TYPE = "pslc"   # e.g. "pw_ppc", "pw_summary", "pw_buyer"

# ---------------------- DATE HANDLING ----------------------
from_date = os.getenv("FROM_DATE")
to_date = os.getenv("TO_DATE")

if not from_date or not to_date:
    today = date.today()
    first_day = today.replace(day=1)
    last_day = today.replace(day=calendar.monthrange(today.year, today.month)[1])
    DATE_FROM = first_day.strftime("%Y-%m-%d")
    DATE_TO = last_day.strftime("%Y-%m-%d")
else:
    DATE_FROM = from_date
    DATE_TO = to_date

print(f"📅 Report period: {DATE_FROM} → {DATE_TO}")

# ---------------------- GOOGLE SHEETS ----------------------
SHEET_ID = "1acV7UrmC8ogC54byMrKRTaD9i1b1Cf9QZ-H1qHU5ZZc"
COMPANY_SHEETS = {
    1: {"sheet": "Zip_Pending_order", "clear_range": "A2:AD", "timestamp_cell": "C1"},
    3: {"sheet": "MT_Pending_order", "clear_range": "A2:AD", "timestamp_cell": "C1"},
}

scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = service_account.Credentials.from_service_account_file("gcreds.json", scopes=scope)
client = gspread.authorize(creds)
print("✅ Google Sheets authorized")

# ========= START SESSION ==========
session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"})

# ---------------------- LOGIN ----------------------
login_url = f"{ODOO_URL}/web/session/authenticate"
login_payload = {
    "jsonrpc": "2.0",
    "params": {"db": DB, "login": USERNAME, "password": PASSWORD}
}
resp = session.post(login_url, json=login_payload)
resp.raise_for_status()
uid = resp.json().get("result", {}).get("uid")
if not uid:
    raise Exception(f"❌ Login failed: {resp.text}")
print("✅ Logged in, UID =", uid)

# ---------------------- CSRF TOKEN ----------------------
resp = session.get(f"{ODOO_URL}/web")
match = re.search(r'var odoo = {\s*csrf_token: "([A-Za-z0-9]+)"', resp.text)
csrf_token = match.group(1) if match else None
if not csrf_token:
    raise Exception("❌ Failed to extract CSRF token")
print("✅ CSRF token =", csrf_token)

# ---------------------- GENERATE & DOWNLOAD ----------------------
COMPANIES = {
    1: "Zipper",
    3: "Metal_Trims"
}

def generate_and_download(company_id, company_name):
    print(f"\n🔹 Processing company: {company_name} (ID={company_id})")

    # Step 3: Onchange
    onchange_url = f"{ODOO_URL}/web/dataset/call_kw/{MODEL}/onchange"
    onchange_payload = {
        "id": 1, "jsonrpc": "2.0", "method": "call",
        "params": {
            "model": MODEL, "method": "onchange",
            "args": [[], {}, [], {
                "report_type": {}, "date_from": {}, "date_to": {},
                "all_buyer_list": {"fields": {"display_name": {}}},
                "all_Customer": {"fields": {"display_name": {}}}
            }],
            "kwargs": {"context": {"lang": "en_US","tz": "Asia/Dhaka","uid": uid,"allowed_company_ids":[company_id]}}
        }
    }
    resp = session.post(onchange_url, json=onchange_payload)
    resp.raise_for_status()
    print("✅ Onchange defaults received")

    # Step 4: Save wizard
    web_save_url = f"{ODOO_URL}/web/dataset/call_kw/{MODEL}/web_save"
    web_save_payload = {
        "id": 2, "jsonrpc": "2.0", "method": "call",
        "params": {
            "model": MODEL, "method": "web_save",
            "args": [[], {"report_type": REPORT_TYPE, "date_from": DATE_FROM, "date_to": DATE_TO, "all_buyer_list": [], "all_Customer": []}],
            "kwargs": {
                "context": {"lang": "en_US","tz":"Asia/Dhaka","uid": uid,"allowed_company_ids":[company_id]},
                "specification": {"report_type": {}, "date_from": {}, "date_to": {},
                                  "all_buyer_list":{"fields":{"display_name":{}}},
                                  "all_Customer":{"fields":{"display_name":{}}}}
            }
        }
    }
    resp = session.post(web_save_url, json=web_save_payload)
    resp.raise_for_status()
    wizard_id = resp.json().get("result", [{}])[0].get("id")
    if not wizard_id:
        raise Exception(f"❌ Wizard creation failed: {resp.text}")
    print("✅ Wizard saved, ID =", wizard_id)

    # Step 5: Trigger report generation
    call_button_url = f"{ODOO_URL}/web/dataset/call_button"
    call_button_payload = {
        "id": 3, "jsonrpc": "2.0", "method": "call",
        "params": {"model": MODEL, "method": REPORT_BUTTON_METHOD,
                   "args": [[wizard_id]],
                   "kwargs":{"context":{"lang":"en_US","tz":"Asia/Dhaka","uid":uid,"allowed_company_ids":[company_id]}}}
    }
    resp = session.post(call_button_url, json=call_button_payload)
    resp.raise_for_status()
    report_info = resp.json().get("result", {})
    report_name = report_info.get("report_name")
    if not report_name:
        raise Exception(f"❌ Failed to generate report: {resp.text}")
    print("✅ Report generated:", report_name)

    # Step 6: Download
    download_url = f"{ODOO_URL}/report/download"
    options = {"date_from": DATE_FROM, "date_to": DATE_TO, "company_id": company_id}
    context = {"lang": "en_US", "tz": "Asia/Dhaka","uid": uid,"allowed_company_ids":[company_id]}
    report_path = f"/report/xlsx/{report_name}/{wizard_id}?options={json.dumps(options)}&context={json.dumps(context)}"
    download_payload = {"data": json.dumps([report_path, "xlsx"]),"context": json.dumps(context),
                        "token":"dummy","csrf_token": csrf_token}
    headers = {"X-CSRF-Token": csrf_token,"Referer":f"{ODOO_URL}/web"}
    resp = session.post(download_url, data=download_payload, headers=headers, timeout=60)

    if resp.status_code == 200 and "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" in resp.headers.get("content-type",""):
        filename = f"{company_name}_{REPORT_TYPE}_{DATE_FROM}_to_{DATE_TO}.xlsx"
        with open(filename, "wb") as f: f.write(resp.content)
        print(f"✅ Report downloaded for {company_name}: {filename}")

        # ---------------------- PASTE TO GOOGLE SHEETS ----------------------
        sheet_cfg = COMPANY_SHEETS[company_id]
        worksheet = client.open_by_key(SHEET_ID).worksheet(sheet_cfg["sheet"])
        worksheet.batch_clear([sheet_cfg["clear_range"]])
        df = pd.read_excel(filename)
        if not df.empty:
            set_with_dataframe(worksheet, df, row=2, col=1)
            timestamp = datetime.now(pytz.timezone("Asia/Dhaka")).strftime("%Y-%m-%d %H:%M:%S")
            worksheet.update(sheet_cfg["timestamp_cell"], [[timestamp]])
            print(f"✅ {company_name} data pasted to sheet, timestamp: {timestamp}")
        else:
            print(f"⚠️ No data to paste for {company_name}")

    else:
        print(f"❌ Download failed for {company_name}", resp.status_code, resp.text[:500])


# ---------------------- RUN ----------------------
for cid, cname in COMPANIES.items():
    try:
        generate_and_download(cid, cname)
    except Exception as e:
        print(f"❌ Error for {cname}: {e}")

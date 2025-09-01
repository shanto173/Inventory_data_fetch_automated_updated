import requests, json, re, os, time
import pandas as pd
import pytz
from datetime import datetime, date, timedelta
import calendar
import gspread
from gspread_dataframe import set_with_dataframe
from google.oauth2 import service_account

# ========= CONFIG ==========
ODOO_URL = os.getenv("ODOO_URL")
USERNAME = os.getenv("ODOO_USERNAME")
PASSWORD = os.getenv("ODOO_PASSWORD")
DB = os.getenv("ODOO_DB")

MODEL = "ppc.report"
REPORT_BUTTON_METHOD = "action_generate_xlsx_report"
REPORT_TYPE = "pslc"

# Google Sheet config
SHEET_ID = "1acV7UrmC8ogC54byMrKRTaD9i1b1Cf9QZ-H1qHU5ZZc"
COMPANIES = {
    1: {"name": "Zipper", "sheet": "Zip_Pending_order", "clear_range": "A2:AD", "timestamp_cell": "C1"},
    3: {"name": "Metal_Trims", "sheet": "MT_Pending_order", "clear_range": "A2:AD", "timestamp_cell": "C1"},
}

# ========= DATE HANDLING ==========
from_date = os.getenv("FROM_DATE")
to_date = os.getenv("TO_DATE")

if not from_date or not to_date:
    today = date.today()
    first_day = today.replace(day=1)
    last_day = today.replace(day=calendar.monthrange(today.year, today.month)[1])
    FROM_DATE = first_day.strftime("%Y-%m-%d")
    TO_DATE = last_day.strftime("%Y-%m-%d")
else:
    FROM_DATE, TO_DATE = from_date, to_date

print(f"📅 Report period: {FROM_DATE} → {TO_DATE}")

# ========= GOOGLE AUTH ==========
scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
creds = service_account.Credentials.from_service_account_file("gcreds.json", scopes=scope)
client = gspread.authorize(creds)
print("✅ Google Sheets authorized")

# ========= ODOO SESSION ==========
session = requests.Session()
session.headers.update({"User-Agent": "Mozilla/5.0"})

# Step 1: Login
resp = session.post(f"{ODOO_URL}/web/session/authenticate", json={
    "jsonrpc": "2.0",
    "params": {"db": DB, "login": USERNAME, "password": PASSWORD}
})
resp.raise_for_status()
uid = resp.json().get("result", {}).get("uid")
if not uid:
    raise Exception("❌ Login failed")
print("✅ Logged in as UID", uid)

# Step 2: CSRF token
resp = session.get(f"{ODOO_URL}/web")
csrf_token = re.search(r'var odoo = {\s*csrf_token: "([A-Za-z0-9]+)"', resp.text).group(1)
print("✅ CSRF token =", csrf_token)

# ========= MAIN FUNCTION ==========
def generate_and_upload(company_id, company_cfg):
    name = company_cfg["name"]
    sheet_name = company_cfg["sheet"]
    clear_range = company_cfg["clear_range"]
    ts_cell = company_cfg["timestamp_cell"]

    print(f"\n🔹 Processing {name} (ID={company_id})")

    # Step 3: Onchange
    session.post(f"{ODOO_URL}/web/dataset/call_kw/{MODEL}/onchange", json={
        "id": 1,
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": MODEL,
            "method": "onchange",
            "args": [[], {}, [], {}],
            "kwargs": {"context": {"lang": "en_US", "tz": "Asia/Dhaka", "uid": uid, "allowed_company_ids": [company_id]}}
        }
    })

    # Step 4: Save wizard
    resp = session.post(f"{ODOO_URL}/web/dataset/call_kw/{MODEL}/web_save", json={
        "id": 2,
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": MODEL,
            "method": "web_save",
            "args": [[], {
                "report_type": REPORT_TYPE,
                "date_from": FROM_DATE,
                "date_to": TO_DATE,
                "all_buyer_list": [],
                "all_Customer": []
            }],
            "kwargs": {
                "context": {"lang": "en_US", "tz": "Asia/Dhaka", "uid": uid, "allowed_company_ids": [company_id]},
                "specification": {
                    "report_type": {},
                    "date_from": {},
                    "date_to": {},
                    "all_buyer_list": {"fields": {"display_name": {}}},
                    "all_Customer": {"fields": {"display_name": {}}}
                }
            }
        }
    })

    if "result" not in resp.json():
        raise Exception(f"❌ Wizard creation failed: {resp.text}")

    wizard_id = resp.json()["result"][0]["id"]

    # Step 5: Trigger report
    resp = session.post(f"{ODOO_URL}/web/dataset/call_button", json={
        "id": 3,
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": MODEL,
            "method": REPORT_BUTTON_METHOD,
            "args": [[wizard_id]],
            "kwargs": {"context": {"lang": "en_US", "tz": "Asia/Dhaka", "uid": uid, "allowed_company_ids": [company_id]}}
        }
    })

    report_name = resp.json().get("result", {}).get("report_name")
    if not report_name:
        raise Exception(f"❌ Failed to generate report: {resp.text}")

    # Step 6: Download XLSX
    report_path = f"/report/xlsx/{report_name}/{wizard_id}?options={json.dumps({'date_from': FROM_DATE, 'date_to': TO_DATE, 'company_id': company_id})}&context={json.dumps({'lang':'en_US','tz':'Asia/Dhaka','uid':uid,'allowed_company_ids':[company_id]})}"
    resp = session.post(f"{ODOO_URL}/report/download",
                        data={"data": json.dumps([report_path, "xlsx"]), "context": json.dumps({}), "token": "dummy", "csrf_token": csrf_token},
                        headers={"X-CSRF-Token": csrf_token, "Referer": f"{ODOO_URL}/web"}, timeout=60)
    filename = f"{name}_{REPORT_TYPE}_{FROM_DATE}_to_{TO_DATE}.xlsx"
    with open(filename, "wb") as f:
        f.write(resp.content)
    print(f"✅ XLSX saved: {filename}")

    # Step 7: Load → DataFrame → Paste to Sheet
    df = pd.read_excel(filename)
    if df.empty:
        print("⚠️ No data, skipping sheet update")
        return

    worksheet = client.open_by_key(SHEET_ID).worksheet(sheet_name)
    worksheet.batch_clear([clear_range])
    time.sleep(3)
    set_with_dataframe(worksheet, df, row=2, col=1)

    local_time = datetime.now(pytz.timezone("Asia/Dhaka")).strftime("%Y-%m-%d %H:%M:%S")
    worksheet.update(ts_cell, [[local_time]])
    print(f"✅ {name} data pasted to {sheet_name}, timestamp: {local_time}")


# Run all companies
for cid, cfg in COMPANIES.items():
    try:
        generate_and_upload(cid, cfg)
    except Exception as e:
        print(f"❌ Error for {cfg['name']}: {e}")

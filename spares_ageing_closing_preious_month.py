import requests
import logging
import sys
import os
from datetime import date, datetime, timedelta
import gspread
from gspread_dataframe import set_with_dataframe
import pandas as pd
import pytz
from dotenv import load_dotenv
import time

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

today = date.today()

# TO_DATE = last day of previous month
first_day_this_month = today.replace(day=1)
last_day_prev_month = first_day_this_month - timedelta(days=1)
TO_DATE = last_day_prev_month.isoformat()
FROM_DATE = False

print("To date (last day of previous month):", TO_DATE)

session = requests.Session()
USER_ID = None

# ========= LABEL MAPPING ==========
LABELS = {
    "parent_category": "Product",
    "product_category": "Category",
    "product_id": "Item",
    "lot_id": "Invoice",
    "receive_date": "Receive Date",
    "shipment_mode": "Shipmnet Mode",
    "slot_1": "0-30",
    "slot_2": "31-60",
    "slot_3": "61-90",
    "slot_4": "91-180",
    "slot_5": "181-365",
    "slot_6": "365+",
    "duration": "Duration",
    "cloing_qty": "Quantity",
    "cloing_value": "Value",
    "landed_cost": "Landed Cost",
    "lot_price": "Price",
    "pur_price": "Pur Price",
    "rejected": "Rejected",
    "company_id": "Company",
    "lot_id/work_center": "Invoice/Work Center",
    "product_id/work_center": "Item/Work Center",
}

# ========= LOGIN ==========
def login():
    global USER_ID
    payload = {"jsonrpc": "2.0", "params": {"db": DB, "login": USERNAME, "password": PASSWORD}}
    r = session.post(f"{ODOO_URL}/web/session/authenticate", json=payload)
    r.raise_for_status()
    result = r.json().get("result")
    if result and "uid" in result:
        USER_ID = result["uid"]
        print(f"✅ Logged in (uid={USER_ID})")
        return result
    else:
        raise Exception("❌ Login failed")

# ========= SWITCH COMPANY ==========
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
            "kwargs": {"context": {"allowed_company_ids": [company_id], "company_id": company_id}},
        },
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()
    if "error" in r.json():
        print(f"❌ Failed to switch to company {company_id}: {r.json()['error']}")
        return False
    else:
        print(f"🔄 Session switched to company {company_id}")
        return True

# ========= CREATE AGEING WIZARD ==========
def create_ageing_wizard(company_id, from_date, to_date):
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "stock.forecast.report",
            "method": "web_save",
            "args": [[], {
                "report_type": "ageing",
                "report_for": "spare",
                "all_iteam_list": [],
                "from_date": from_date,
                "to_date": to_date
            }],
            "kwargs": {
                "context": {"lang": "en_US", "tz": "Asia/Dhaka", "uid": USER_ID,
                            "allowed_company_ids": [company_id], "company_id": company_id},
                "specification": {
                    "report_type": {},
                    "report_for": {},
                    "all_iteam_list": {"fields": {"display_name": {}}},
                    "from_date": {},
                    "to_date": {},
                },
            },
        },
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw/stock.forecast.report/web_save", json=payload)
    r.raise_for_status()
    result = r.json().get("result", [])
    if isinstance(result, list) and result:
        wiz_id = result[0]["id"]
        print(f"🪄 Ageing wizard {wiz_id} created for company {company_id}")
        return wiz_id
    else:
        raise Exception(f"❌ Failed to create ageing wizard: {r.text}")

# ========= COMPUTE AGEING ==========
def compute_ageing(company_id, wizard_id):
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "stock.forecast.report",
            "method": "print_date_wise_stock_register",
            "args": [[wizard_id]],
            "kwargs": {"context": {"lang": "en_US", "tz": "Asia/Dhaka",
                                   "uid": USER_ID,
                                   "allowed_company_ids": [company_id],
                                   "company_id": company_id}},
        },
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_button", json=payload)
    r.raise_for_status()
    result = r.json()
    if "error" in result:
        print(f"❌ Error computing ageing for {company_id}: {result['error']}")
    else:
        print(f"⚡ Ageing computed for wizard {wizard_id} (company {company_id})")
    return result

# ========= FETCH AGEING REPORT ==========
def fetch_ageing(company_id, cname, wizard_id):
    context = {"allowed_company_ids": [company_id], "company_id": company_id,
               "active_model": "stock.forecast.report", "active_id": wizard_id, "active_ids": [wizard_id]}
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "stock.ageing",
            "method": "web_search_read",
            "args": [],
            "kwargs": {
                "specification": {
                    "parent_category": {"fields": {"display_name": {}}},
                    "product_category": {"fields": {"display_name": {}}},
                    "product_id": {"fields": {"display_name": {}, "work_center": {}}},
                    "lot_id": {"fields": {"display_name": {}, "work_center": {}}},
                    "receive_date": {},
                    "shipment_mode": {},
                    "slot_1": {},
                    "slot_2": {},
                    "slot_3": {},
                    "slot_4": {},
                    "slot_5": {},
                    "slot_6": {},
                    "duration": {},
                    "cloing_qty": {},
                    "cloing_value": {},
                    "landed_cost": {},
                    "lot_price": {},
                    "pur_price": {},
                    "rejected": {},
                    "company_id": {"fields": {"display_name": {}}},
                },
                "offset": 0,
                "limit": 5000,
                "context": context,
                "count_limit": 10000,
                "domain": [["product_id.categ_id.complete_name", "ilike", "All / Spare Parts"]],
            },
        },
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()
    try:
        data = r.json()["result"]["records"]
        def flatten(record):
            flat = {}
            for k, v in record.items():
                if isinstance(v, dict) and "display_name" in v:
                    flat[LABELS.get(k, k)] = v["display_name"]
                    if k == "lot_id" and "work_center" in v:
                        flat["Invoice/Work Center"] = v["work_center"]
                    elif k == "product_id" and "work_center" in v:
                        flat["Item/Work Center"] = v["work_center"]
                else:
                    flat[LABELS.get(k, k)] = v
            return flat
        flattened = [flatten(rec) for rec in data]
        print(f"📊 {cname}: {len(flattened)} ageing rows fetched")
        return flattened
    except Exception:
        print(f"❌ {cname}: Failed to parse ageing report:", r.text[:200])
        return []

# ========= GOOGLE SHEETS CONFIG ==========
SHEET_KEY = "1P8vMDw-rFZtzOV162wlWeXVnHbx_bwnRDMJ-0cSaFjU"
SHEET_NAMES = {
    1: "spares_ageing closing_Zip",   # Zipper
    3: "spares_ageing closing_MT",    # Metal Trims
}

# ========= MAIN ==========
if __name__ == "__main__":
    userinfo = login()
    print("User info (allowed companies):", userinfo.get("user_companies", {}))

    for cid, cname in COMPANIES.items():
        print(f"\n🚀 Processing company: {cname} (ID={cid})")
        success = False

        for attempt in range(1, 3):
            try:
                if switch_company(cid):
                    wiz_id = create_ageing_wizard(cid, FROM_DATE, TO_DATE)
                    compute_ageing(cid, wiz_id)
                    records = fetch_ageing(cid, cname, wiz_id)

                    if records:
                        df = pd.DataFrame(records)
                        # Drop first column (parent_category grouping)
                        df = df.iloc[:, 1:]
                        output_file = f"{cname.lower().replace(' ', '_')}_spares_ageing_closing_{TO_DATE}.xlsx"
                        df.to_excel(output_file, index=False)
                        print(f"📂 Saved: {output_file}")

                        # ========= GOOGLE SHEETS ==========
                        try:
                            client = gspread.service_account(filename="gcreds.json")
                            sheet = client.open_by_key(SHEET_KEY)
                            worksheet = sheet.worksheet(SHEET_NAMES[cid])

                            worksheet.clear()
                            set_with_dataframe(worksheet, df)
                            local_tz = pytz.timezone("Asia/Dhaka")
                            local_time = datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")
                            worksheet.update([[f"{local_time}"]], "W2")
                            print(f"✅ Data pasted & timestamp updated: {local_time}")

                        except Exception as e:
                            raise Exception(f"Google Sheets paste failed: {e}")

                    else:
                        raise Exception(f"No ageing data fetched for {cname}")

                    success = True
                    print(f"✅ Completed successfully for {cname} (Attempt {attempt})")
                    break

            except Exception as e:
                print(f"⚠️ Attempt {attempt}/30 failed for {cname}: {e}")
                if attempt < 2:
                    wait_time = 5
                    print(f"🔁 Retrying {cname} in {wait_time}s...")
                    time.sleep(wait_time)
                else:
                    print(f"❌ Max retries reached for {cname}. Skipping to next company.")

        if not success:
            print(f"🚫 Skipping {cname} after 2 failed attempts.\n")

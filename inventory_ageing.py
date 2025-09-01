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

log = logging.getLogger()
log.setLevel(logging.INFO)

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

FROM_DATE = False  # no start date needed
TO_DATE = os.getenv("TO_DATE")  # fetch from GitHub Actions input

if not TO_DATE:
    TO_DATE = today.isoformat()  # fallback to today

session = requests.Session()
USER_ID = None

# ========= LOGIN ==========
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
        print(f"❌ Failed to switch to company {company_id}: {r.json()['error']}")
        return False
    else:
        print(f"🔄 Session switched to company {company_id}")
        return True


# ========= CREATE FORECAST WIZARD ==========
def create_forecast_wizard(company_id, from_date, to_date):
    payload = {
        "jsonrpc": "2.0",
        "method": "call",
        "params": {
            "model": "stock.forecast.report",
            "method": "create",
            "args": [{
                "from_date": from_date,
                "to_date": to_date,
            }],
            "kwargs": {
                "context": {
                    "allowed_company_ids": [company_id],
                    "company_id": company_id,
                }
            }
        }
    }
    r = session.post(f"{ODOO_URL}/web/dataset/call_kw", json=payload)
    r.raise_for_status()
    wizard_id = r.json()["result"]
    print(f"🪄 Created wizard {wizard_id} for company {company_id}")
    return wizard_id


# ========= COMPUTE FORECAST ==========
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
    result = r.json()
    if "error" in result:
        print(f"❌ Error computing forecast for {company_id}: {result['error']}")
    else:
        print(f"⚡ Forecast computed for wizard {wizard_id} (company {company_id})")
    return result


# ========= FETCH REPORT ==========
def fetch_opening_closing(company_id, cname):
    context = {
        "allowed_company_ids": [company_id],
        "company_id": company_id,
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
                "context": {
                    **context,
                    "active_model": "stock.forecast.report",
                    "active_id": 0,
                    "active_ids": [0],
                },
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
            flat = {}
            for k, v in record.items():
                if isinstance(v, dict) and "display_name" in v:
                    flat[k] = v["display_name"]
                else:
                    flat[k] = v
            return flat

        flattened = [flatten_record(rec) for rec in data]
        print(f"📊 {cname}: {len(flattened)} rows fetched (flattened)")
        return flattened
    except Exception:
        print(f"❌ {cname}: Failed to parse report:", r.text[:200])
        return []


# ========= MAIN ==========
if __name__ == "__main__":
    userinfo = login()
    print("User info (allowed companies):", userinfo.get("user_companies", {}))

    for cid, cname in COMPANIES.items():
        if switch_company(cid):
            wiz_id = create_forecast_wizard(cid, FROM_DATE, TO_DATE)
            compute_forecast(cid, wiz_id)
            records = fetch_opening_closing(cid, cname)

            if records:
                df = pd.DataFrame(records)
                output_file = f"{cname.lower().replace(' ', '_')}_opening_closing_{today.isoformat()}.xlsx"
                df.to_excel(output_file, index=False)
                print(f"📂 Saved: {output_file}")

                # ---------- GOOGLE SHEET PASTING ----------
                try:
                    if cname == "Zipper":
                        creds = service_account.Credentials.from_service_account_file('gcreds.json', scopes=[
                            "https://www.googleapis.com/auth/spreadsheets",
                            "https://www.googleapis.com/auth/drive",
                        ])
                        client = gspread.authorize(creds)
                        sheet = client.open_by_key("1z6Zb_BronrO26rNS_gCKmsetoY7_OFysfIyvU3iazy0")
                        worksheet = sheet.worksheet("age_ZIP")

                    elif cname == "Metal Trims":
                        creds = service_account.Credentials.from_service_account_file('gcreds.json', scopes=[
                            "https://www.googleapis.com/auth/spreadsheets",
                            "https://www.googleapis.com/auth/drive",
                        ])
                        client = gspread.authorize(creds)
                        sheet = client.open_by_key("1z6Zb_BronrO26rNS_gCKmsetoY7_OFysfIyvU3iazy0")
                        worksheet = sheet.worksheet("age_MT")

                    if not df.empty:
                        worksheet.clear()
                        time.sleep(2)
                        set_with_dataframe(worksheet, df)
                        local_tz = pytz.timezone('Asia/Dhaka')
                        local_time = datetime.now(local_tz).strftime("%Y-%m-%d %H:%M:%S")
                        worksheet.update("W2", [[f"{local_time}"]])
                        log.info(f"✅ Data pasted & timestamp updated: {local_time}")
                    else:
                        print("Skip: DataFrame is empty, not pasting to sheet.")
                except Exception as e:
                    log.error(f"❌ Error while pasting to Google Sheets: {e}")
            else:
                print(f"❌ No data fetched for {cname}")

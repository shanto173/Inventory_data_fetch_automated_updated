# Inventory Data Fetch Automated

An automated Python-based system that fetches inventory and stock data from an **Odoo ERP** instance and uploads the reports to **Google Sheets**. The entire pipeline is orchestrated via **GitHub Actions**, allowing scheduled or manual execution without any local setup.

---

## 📋 Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Repository Structure](#repository-structure)
- [Prerequisites](#prerequisites)
- [Setup & Configuration](#setup--configuration)
  - [GitHub Secrets](#github-secrets)
  - [Google Service Account](#google-service-account)
- [Usage](#usage)
  - [Running via GitHub Actions](#running-via-github-actions)
  - [Running Locally](#running-locally)
- [Scripts Reference](#scripts-reference)
- [Output](#output)
- [License](#license)

---

## Overview

This project connects to an Odoo ERP backend using its JSON-RPC API, pulls inventory-related data for multiple companies (Zipper and Metal Trims), processes the data with `pandas`, and then writes the results into designated Google Sheets worksheets. It supports date-range filtering so you can pull data for any period.

---

## Features

- 🔄 **Multi-company support** – fetches data for both *Zipper* (Company ID 1) and *Metal Trims* (Company ID 3)
- 📅 **Date-range filtering** – accepts optional `FROM_DATE` and `TO_DATE` inputs; defaults to the current month
- 📊 **Google Sheets integration** – automatically writes data to pre-configured worksheets and stamps a timestamp
- 💾 **Local Excel backup** – saves downloaded data as `.xlsx` files inside a `download/` directory
- ⚡ **GitHub Actions automation** – supports `workflow_dispatch` with a script selector, from-date and to-date inputs
- 🔁 **Retry logic** – scripts retry on transient failures before giving up

---

## Repository Structure

```
├── .github/
│   └── workflows/
│       └── main.yml                        # GitHub Actions workflow
├── download/                               # Auto-generated Excel backups
├── Closing_stock.py                        # Current-month opening/closing stock (RM category)
├── Closing_stock_1.py                      # Alternate closing stock variant
├── Closing_stock_last_day.py               # Closing stock for last day of month
├── Consumption_stock_Apr24_till.py         # Consumption stock from April 2024 onwards
├── Consumption_stock_mar24_till.py         # Consumption stock from March 2024 onwards
├── Fg_stock.py                             # Finished goods stock
├── MT_spares.py                            # Metal Trims spares stock
├── Raw_materials.py                        # Raw materials product list with available qty
├── Relese_inovice_summary.py               # Released invoice summary
├── Sep_inovice_summary.py                  # September invoice summary
├── Spares_stock.py                         # Spares stock report
├── inovice_summary.py                      # Invoice summary report
├── inventory_ageing.py                     # Inventory ageing report (current)
├── inventory_ageing_1.py                   # Inventory ageing variant
├── inventory_ageing_last_day.py            # Inventory ageing as of last day
├── pending_invoice_last_month.py           # Pending invoices from last month
├── pending_slider.py                       # Pending slider/delivery report
├── spares_ageing.py                        # Spares ageing report
├── spares_ageing_closing_preious_month.py  # Spares ageing closing (previous month)
├── spares_workcenter_df.py                 # Spares by work center
├── unuseable_stock.py                      # Unusable/dead stock report
├── .gitignore
└── LICENSE
```

---

## Prerequisites

| Requirement | Version |
|---|---|
| Python | 3.11+ |
| pip packages | `requests`, `pandas`, `gspread`, `gspread-dataframe`, `google-auth`, `google-auth-oauthlib`, `google-auth-httplib2`, `openpyxl`, `pytz`, `python-dotenv` |
| Odoo ERP | Accessible instance with JSON-RPC enabled |
| Google Cloud | Service Account with Sheets + Drive API access |

---

## Setup & Configuration

### GitHub Secrets

Add the following secrets to your repository (**Settings → Secrets and variables → Actions**):

| Secret Name | Description |
|---|---|
| `ODOO_URL` | Base URL of your Odoo instance (e.g. `https://your-odoo.com`) |
| `ODOO_DB` | Odoo database name |
| `ODOO_USERNAME` | Odoo login username / email |
| `ODOO_PASSWORD` | Odoo login password |
| `GOOGLE_CREDENTIALS_BASE64` | Base64-encoded Google service account JSON key |

To encode your Google credentials:

```bash
base64 -w 0 your-service-account.json
```

Copy the output and paste it as the `GOOGLE_CREDENTIALS_BASE64` secret.

### Google Service Account

1. Create a project in [Google Cloud Console](https://console.cloud.google.com/).
2. Enable the **Google Sheets API** and **Google Drive API**.
3. Create a **Service Account** and download the JSON key.
4. Share each target Google Sheet with the service account email (editor access).

### Local `.env` File (for local runs)

Create a `.env` file in the project root:

```env
ODOO_URL=https://your-odoo-instance.com
ODOO_DB=your_database
ODOO_USERNAME=your@email.com
ODOO_PASSWORD=your_password
FROM_DATE=2024-01-01   # optional
TO_DATE=2024-01-31     # optional
```

Also place your `gcreds.json` (Google service account key) in the project root.

---

## Usage

### Running via GitHub Actions

1. Go to **Actions** → **Odoo Stock Report Automation**.
2. Click **Run workflow**.
3. Fill in the inputs:
   - **Script(s) to run**: choose a specific script or `All` to run every script.
   - **Start date** (`FROM_DATE`): optional, format `YYYY-MM-DD`.
   - **End date** (`TO_DATE`): optional, defaults to today.
4. Click **Run workflow**.

### Running Locally

Install dependencies:

```bash
pip install requests pandas gspread gspread-dataframe google-auth \
            google-auth-oauthlib google-auth-httplib2 openpyxl pytz python-dotenv
```

Run a specific script:

```bash
python Closing_stock.py
```

Or run all scripts sequentially:

```bash
for script in Closing_stock.py Raw_materials.py inventory_ageing.py; do
    python "$script"
done
```

---

## Scripts Reference

| Script | Description |
|---|---|
| `Closing_stock.py` | Fetches opening/closing stock for RM category for the configured date range |
| `Closing_stock_1.py` | Alternate closing stock report variant |
| `Closing_stock_last_day.py` | Closing stock snapshot for the last day of the month |
| `Raw_materials.py` | Fetches raw material products with current available quantity |
| `inventory_ageing.py` | Generates inventory ageing analysis for the current period |
| `inventory_ageing_1.py` | Inventory ageing variant report |
| `inventory_ageing_last_day.py` | Inventory ageing as of the last day |
| `Fg_stock.py` | Finished goods stock report |
| `Spares_stock.py` | Spares inventory stock report |
| `MT_spares.py` | Metal Trims spares-specific report |
| `spares_ageing.py` | Ageing analysis for spares inventory |
| `spares_ageing_closing_preious_month.py` | Spares ageing closing for the previous month |
| `spares_workcenter_df.py` | Spares breakdown by work center |
| `inovice_summary.py` | Invoice summary report |
| `Relese_inovice_summary.py` | Released invoice summary |
| `Sep_inovice_summary.py` | September invoice summary |
| `pending_invoice_last_month.py` | Pending invoices from last month |
| `pending_slider.py` | Pending slider/delivery status report |
| `Consumption_stock_mar24_till.py` | Consumption stock from March 2024 to date |
| `Consumption_stock_Apr24_till.py` | Consumption stock from April 2024 to date |
| `unuseable_stock.py` | Identifies and reports unusable/dead stock items |

---

## Output

Each script produces two outputs:

1. **Excel file** – saved locally in the `download/` directory, named with the company and date (e.g. `zipper_opening_closing_2025-04-01.xlsx`).
2. **Google Sheet** – data is written to a pre-configured worksheet. A timestamp (Asia/Dhaka timezone) is written to track the last update time.

---

## License

This project is licensed under the terms of the [LICENSE](LICENSE) file included in this repository.

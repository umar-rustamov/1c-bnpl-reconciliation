# -*- coding: utf-8 -*-
"""Live Google-Sheet report for automate_order_ekvayring.py's batch runs.

Replaces the old end-of-run Excel file (report_writer.py): instead of one
attachment sent when the whole batch finishes, this appends ONE ROW to a
shared Google Sheet immediately after each order/application is processed,
so the sheet fills in live as the run goes and you can watch it update.

The target sheet's URL is in credentials\\.env as REPORT_SPREADSHEET_URL
(shared with the same service account as the other sheets -
see sheets_client.py). Reuses that module's credentials/scopes and its
429-and-connection retry wrapper.

One row per order, columns below. Prices are per-item: "1C price before" /
"after" are the main item's Tsena before and after this run set it; Item
2-5 price are the additional items' set prices (for multi-item orders).
"""
import datetime
import re

import gspread
from google.oauth2.service_account import Credentials

from env_config import load_env
from sheets_client import DEFAULT_CREDENTIALS_PATH, SCOPES, _with_retry

_env = load_env()
REPORT_SPREADSHEET_URL = _env.get("REPORT_SPREADSHEET_URL", "")


def _spreadsheet_id(url_or_id):
    """Accept either a full Google Sheets URL or a bare spreadsheet id."""
    m = re.search(r"/spreadsheets/d/([A-Za-z0-9_-]+)", url_or_id or "")
    return m.group(1) if m else (url_or_id or "").strip()


REPORT_SPREADSHEET_ID = _spreadsheet_id(REPORT_SPREADSHEET_URL)

# (Column header written to row 1, dict key read from each row).
COLUMNS = [
    ("Report time", "logged_at"),
    ("ID", "id"),
    ("Contract (1C №)", "contract"),
    ("Order date", "order_date"),
    ("1C name", "name_1c"),
    ("Source name", "name_source"),
    ("Source price", "price_source"),
    ("1C price before", "price_before"),
    ("1C price after (item 1)", "price_after"),
    ("Item 2 price", "item2_price"),
    ("Item 3 price", "item3_price"),
    ("Item 4 price", "item4_price"),
    ("Item 5 price", "item5_price"),
    ("Bonus item", "bonus"),
    ("Quantity", "quantity"),
    ("Status", "status"),
    ("Terminal", "terminal"),
    ("Соглашение", "soglashenie"),
    ("Category", "category"),
    ("Commission", "commission"),
    ("Comment", "comment"),
]
_HEADERS = [label for label, _key in COLUMNS]


def is_configured():
    return bool(REPORT_SPREADSHEET_ID)


def _client():
    creds = Credentials.from_service_account_file(str(DEFAULT_CREDENTIALS_PATH), scopes=SCOPES)
    return gspread.authorize(creds)


def get_report_worksheet():
    """Open the report spreadsheet's first worksheet and make sure its
    header row exists (writes _HEADERS into row 1 if the sheet is empty).
    Returns the worksheet, or raises if REPORT_SPREADSHEET_URL isn't set."""
    if not is_configured():
        raise RuntimeError("REPORT_SPREADSHEET_URL is not set in credentials\\.env - can't write the report sheet.")
    ws = _client().open_by_key(REPORT_SPREADSHEET_ID).sheet1
    first_row = _with_retry(ws.row_values, 1)
    if [c.strip() for c in first_row][: len(_HEADERS)] != _HEADERS:
        # Empty (or stale) header row - (re)write the headers. Only ever
        # touches row 1; existing data rows below are left as-is.
        # Keyword args on purpose: gspread 6.x swapped update()'s positional
        # order (values, range_name), so naming them avoids the footgun.
        _with_retry(ws.update, range_name="A1", values=[_HEADERS])
    return ws


def append_row(ws, row):
    """Append one report row (a dict keyed by COLUMNS' second elements) to
    the sheet, in column order, stamping the current time into "Report
    time". Missing keys render blank."""
    row = dict(row)
    row.setdefault("logged_at", datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    values = []
    for _label, key in COLUMNS:
        v = row.get(key, "")
        values.append("" if v is None else v)
    _with_retry(ws.append_row, values, value_input_option="USER_ENTERED")

# -*- coding: utf-8 -*-
"""Google Sheets access for the Alif Shop reconciliation flow - a sibling
to sheets_client.py (IMAN Sverka) but for TWO different sheets:

  - the "Alif Nasiya" applications sheet (ALIF_SPREADSHEET_ID) - one row
    per physical item sold through Alif Shop's installment ("nasiya")
    product, cross-referenced against business.alif.uz (see alif_client.py)
    for each application's current status.
  - the commissions sheet (COMMISSIONS_SPREADSHEET_ID) - a 3-level product
    category hierarchy with the % Alif Shop takes as commission for each
    leaf category, subtracted from Alif's listed price before it goes into
    1C (see get_commission_for_category).

Both are shared with the SAME service account already used for IMAN Sverka
(<service-account>@<gcp-project>.iam.gserviceaccount.com, see
sheets_client.py's docstring for the one-time setup) - Google Sheets access
is per-sheet-share, not per-credentials-file, so no separate service
account is needed. The Alif Nasiya sheet needs Editor (this writes status/
price/colors back); the commissions sheet only needs Viewer.

Both spreadsheet IDs and both worksheet gids can be overridden via
credentials\\.env (ALIF_SPREADSHEET_ID, ALIF_WORKSHEET_GID,
COMMISSIONS_SPREADSHEET_ID, COMMISSIONS_WORKSHEET_GID - see env_config.py
and .env.example).

Column layout of the Alif Nasiya sheet (row 1 = header, data from row 2),
confirmed live:
    A  Дата (sale date)
    B  Application ID (business.alif.uz's own id for this application) -
       NOT unique per row: a multi-item application (2+ products bought in
       one Alif Nasiya application) spans one row per physical item, all
       sharing the same id here (confirmed live: application 1000002 has 2
       rows, one per Honor X7d color variant) - see
       find_all_rows_for_application.
    C  Status - a STRICT one-of-list dropdown (confirmed live via the
       sheet's own data validation rule): "waiting", "success", "approved",
       "Cancelled" - exact case, no leading space. Plenty of historical
       rows have variants with a leading space (" waiting", " success") or
       raw legacy API codes ("waiting_for_qabz", "approved_by_qabz") from
       before the dropdown was made strict - reading tolerates those
       (compares stripped/lowercased), but this module always WRITES one
       of the three clean validated values above.
    D  Item name - already pre-populated (matches both Alif's own "SKU"
       text, see alif_client.py, and 1C's Номенклатура) - read-only here,
       never written.
    E  Kolichestvo - already pre-populated; read-only here. Rows with
       kolichestvo > 1 are handled: each such row's units get claimed WHOLE
       by a single Alif Business item whose own kolichestvo covers them,
       and F/G below get the same per-unit price for kol=1 and kol>=2 rows
       alike (see _match_alifshop_items_to_sheet_rows and
       reconcile_alifshop_application in automate_order_ekvayring.py).
    F  Сумма 1С (post-commission price) - written once reconciled.
    G  same value as F - confirmed live, every existing "success" row has
       F == G exactly; this module always writes the same amount to both.
    H  Склад/filial - already pre-populated, read-only here.
    I, J  unused/hidden columns - not touched.
    K  1C order number - ALREADY pre-populated (unlike the IMAN flow,
       there's no need to search anything for this - it's given directly
       per row).

Row coloring: NOT conditional formatting (confirmed live - this sheet has
zero conditionalFormats rules) - existing "success"/"Cancelled" rows are
plain solid cell background colors (confirmed live by reading real rows:
pure green {"green": 1} for success, pure red {"red": 1} for Cancelled)
applied directly, not computed by a rule - so this module replicates that
by explicitly setting the row's background after writing its final status.

Commissions sheet columns: category_first, category_second, category_third
(the leaf category - matches Alif's scraped item category text exactly,
confirmed live) and Комиссия (commission as a fraction, e.g. 0.1 for 10% -
confirmed live via UNFORMATTED_VALUE). Leaf category names are usually
unique but NOT always (confirmed live: "Чистящие средства" appears under
two different parent categories with different commissions, 0.1 vs 0.15) -
looked up by leaf name alone (Alif's detail page doesn't expose the parent
categories), but raises rather than silently picking one if a leaf name
matches rows that disagree on the commission.
"""
import re
import time

import gspread
import requests
from google.oauth2.service_account import Credentials

from env_config import load_env
from sheets_client import DEFAULT_CREDENTIALS_PATH, SCOPES, parse_sheet_date

_env = load_env()
# ALIF_SPREADSHEET_ID / ALIF_WORKSHEET_GID / COMMISSIONS_SPREADSHEET_ID /
# COMMISSIONS_WORKSHEET_GID in credentials\.env set these (see
# env_config.py and .env.example).
ALIF_SPREADSHEET_ID = _env.get("ALIF_SPREADSHEET_ID", "")
ALIF_WORKSHEET_GID = int(_env.get("ALIF_WORKSHEET_GID", 0))
COMMISSIONS_SPREADSHEET_ID = _env.get("COMMISSIONS_SPREADSHEET_ID", "")
COMMISSIONS_WORKSHEET_GID = int(_env.get("COMMISSIONS_WORKSHEET_GID", 0))

COL_DATE = 1
COL_APPLICATION_ID = 2
COL_STATUS = 3
COL_ITEM_NAME = 4
COL_KOLICHESTVO = 5
COL_SUMMA_1 = 6
COL_SUMMA_2 = 7
COL_SKLAD = 8
COL_ORDER_NUMBER = 11  # K

FIRST_DATA_ROW = 2
ROW_COLOR_RANGE = "A{row}:M{row}"  # matches the sheet's real column count (header goes A..M, confirmed live)

STATUS_WAITING = "waiting"
STATUS_SUCCESS = "success"
STATUS_CANCELLED = "Cancelled"

ROW_COLOR_SUCCESS = {"red": 0.0, "green": 1.0, "blue": 0.0}
ROW_COLOR_CANCELLED = {"red": 1.0, "green": 0.0, "blue": 0.0}


def _with_retry(fn, *args, retries=4, backoff=20, **kwargs):
    """Same reasoning as sheets_client._with_retry (429 backoff, PLUS
    transient connection errors) - kept as its own small copy rather than
    importing that module's private helper, so this file stays
    self-contained (same pattern emit_transactions/ already uses for its
    own separate copy of lib1c.py)."""
    for attempt in range(1, retries + 1):
        try:
            return fn(*args, **kwargs)
        except gspread.exceptions.APIError as e:
            status = getattr(getattr(e, "response", None), "status_code", None)
            if status != 429 or attempt == retries:
                raise
            wait = backoff * attempt
            print(f"   (Google Sheets API rate limit hit, waiting {wait}s and retrying, attempt {attempt}/{retries})")
            time.sleep(wait)
        except requests.exceptions.RequestException as e:
            # A raw transport-level failure (confirmed live during a batch
            # run: 'RemoteDisconnected: Remote end closed connection
            # without response') - no HTTP status to check (unlike the 429
            # case above), but always transient by nature, so just retry
            # the same way. Previously uncaught here, this crashed the
            # WHOLE batch loop outright even though the 1C side of the
            # application it happened on had already finished successfully
            # moments earlier.
            if attempt == retries:
                raise
            wait = backoff * attempt
            print(f"   (Google Sheets connection hiccup ({e!r}), waiting {wait}s and retrying, "
                  f"attempt {attempt}/{retries})")
            time.sleep(wait)


def _client():
    creds = Credentials.from_service_account_file(str(DEFAULT_CREDENTIALS_PATH), scopes=SCOPES)
    return gspread.authorize(creds)


def get_alif_worksheet():
    return _client().open_by_key(ALIF_SPREADSHEET_ID).get_worksheet_by_id(ALIF_WORKSHEET_GID)


def get_commissions_worksheet():
    return _client().open_by_key(COMMISSIONS_SPREADSHEET_ID).get_worksheet_by_id(COMMISSIONS_WORKSHEET_GID)


def _row_dict(row_number, row_values):
    def cell(col_index):
        i = col_index - 1
        return row_values[i].strip() if i < len(row_values) else ""

    kolichestvo_raw = cell(COL_KOLICHESTVO)
    try:
        # Some historical rows have garbage/shifted data in this column
        # (confirmed live: sheet row 1116 has '31.12.1899' here, alongside
        # other clearly-corrupted cells on that same row) - treat as
        # unknown rather than crashing the whole scan; the actual reconcile
        # flow independently requires kolichestvo == 1 later anyway, so an
        # unparseable value here just naturally fails that check.
        kolichestvo = float(kolichestvo_raw) if kolichestvo_raw else None
    except ValueError:
        kolichestvo = None
    return {
        "row": row_number,
        "status": cell(COL_STATUS),
        "item_name": cell(COL_ITEM_NAME),
        "kolichestvo": kolichestvo,
        "order_number": cell(COL_ORDER_NUMBER),
    }


def find_all_rows_for_application(ws, application_id, start_row=FIRST_DATA_ROW):
    """Return every sheet row (not just the first) for this Alif Business
    application id (column B, exact match) - a multi-item application spans
    one row per physical item (confirmed live: application 1000002 has 2
    rows, one per Honor X7d color variant, sharing one 1C order
    ON-26/00101) - mirrors sheets_client.find_all_rows_for_order for the
    IMAN flow. Returns [{row, status, item_name, kolichestvo,
    order_number}, ...] in sheet order."""
    last_row = max(ws.row_count, start_row)
    rows = _with_retry(ws.get, f"A{start_row}:K{last_row}")
    results = []
    for offset, row_values in enumerate(rows):
        b = row_values[COL_APPLICATION_ID - 1].strip() if len(row_values) >= COL_APPLICATION_ID else ""
        if b == str(application_id):
            results.append(_row_dict(start_row + offset, row_values))
    return results


def find_application_row(ws, application_id):
    """Find the FIRST row for this Alif Business application id - returns
    {row, status, item_name, kolichestvo, order_number}, or None if not
    found. Used only for the initial "is this application already
    resolved" check (see reconcile_alifshop_application) - the actual
    reconcile logic uses find_all_rows_for_application, since a multi-item
    application needs every one of its rows, not just the first."""
    rows = find_all_rows_for_application(ws, application_id)
    return rows[0] if rows else None


def _is_resolved(status):
    """True if this row's status is already a final outcome (success or
    Cancelled) - tolerates the historical leading-space/legacy-code
    variants (see this module's docstring), same as
    reconcile_alifshop_application's own existing_status check."""
    s = (status or "").strip().lower()
    return s.startswith("success") or s.startswith("cancel")


def find_next_unresolved_application(ws, start_row=FIRST_DATA_ROW, since_date=None):
    """Return the first row (top to bottom) whose application id isn't
    already resolved (blank status = never checked, or 'waiting' = checked
    before but Alif hadn't finalized it yet - both need (re)checking) -
    analogous to sheets_client.find_next_unchecked_row for the IMAN flow.
    None if there's nothing to do.

    since_date (a date, or a 'DD.MM.YYYY' string): rows whose Дата (column
    A) is earlier than this are skipped rather than stopping the scan -
    same reasoning as the IMAN flow's --since.

    Returns {row, application_id, status, order_number}."""
    if isinstance(since_date, str):
        since_date = parse_sheet_date(since_date)

    last_row = max(ws.row_count, start_row)
    rows = _with_retry(ws.get, f"A{start_row}:K{last_row}")

    for offset, row_values in enumerate(rows):
        info = _row_dict(start_row + offset, row_values)
        application_id = row_values[COL_APPLICATION_ID - 1].strip() if len(row_values) >= COL_APPLICATION_ID else ""
        if not application_id or _is_resolved(info["status"]):
            continue
        if since_date is not None:
            date_text = row_values[COL_DATE - 1].strip() if len(row_values) >= COL_DATE else ""
            row_date = parse_sheet_date(date_text)
            if row_date is None or row_date < since_date:
                continue
        return {"row": info["row"], "application_id": application_id, "status": info["status"],
                "order_number": info["order_number"]}
    return None


def count_unresolved_applications(ws, start_row=FIRST_DATA_ROW, since_date=None):
    """How many rows currently qualify as "not resolved yet" (same rule as
    find_next_unresolved_application) - a cheap upfront count for the batch
    loop's progress display, counting ROWS (a multi-item application's rows
    all count individually - same simplification the IMAN batch report
    already makes, since a successful multi-item reconcile resolves all of
    an application's rows at once anyway)."""
    if isinstance(since_date, str):
        since_date = parse_sheet_date(since_date)

    last_row = max(ws.row_count, start_row)
    rows = _with_retry(ws.get, f"A{start_row}:K{last_row}")

    count = 0
    for row_values in rows:
        application_id = row_values[COL_APPLICATION_ID - 1].strip() if len(row_values) >= COL_APPLICATION_ID else ""
        status = row_values[COL_STATUS - 1].strip() if len(row_values) >= COL_STATUS else ""
        if not application_id or _is_resolved(status):
            continue
        if since_date is not None:
            date_text = row_values[COL_DATE - 1].strip() if len(row_values) >= COL_DATE else ""
            row_date = parse_sheet_date(date_text)
            if row_date is None or row_date < since_date:
                continue
        count += 1
    return count


def _normalize_category_text(text):
    """Squash a category name down to bare lowercase letters/digits (no
    spaces, hyphens, or other punctuation at all) for comparison - matches
    Alif's scraped category text against the commissions sheet's leaf
    category even when the two systems disagree on more than just casing.
    Confirmed live twice now: 'Смартфоны SAMSUNG' vs 'Смартфоны Samsung'
    (casing only - a plain .lower() would've caught that one), and
    'радар-детекторы' vs 'Радардетекторы' (Alif hyphenates the compound
    word, the sheet doesn't - a mere .lower() does NOT catch this, since
    the word boundary itself differs: 'радар-детекторы'.lower() is 'радар
    детекторы' after replacing the hyphen with a space, which still isn't
    'радардетекторы'; stripping the separator entirely, not just
    lowercasing, is what's needed)."""
    return re.sub(r"[^a-z0-9а-яё]", "", (text or "").lower())


def get_commission_for_category(commissions_ws, category):
    """Look up the commission fraction (e.g. 0.1 for 10%) for a leaf
    category name (matches Alif's scraped item category text against the
    commissions sheet's leaf category using _normalize_category_text -
    tolerates whitespace, casing, and hyphenation differences between the
    two systems, see that function's docstring for confirmed live
    examples). Raises if the category isn't found, or if it matches more
    than one row that DISAGREES on the commission (see this module's
    docstring - a real, if rare, case) - never silently guesses."""
    last_row = max(commissions_ws.row_count, FIRST_DATA_ROW)
    rows = _with_retry(commissions_ws.get, f"A{FIRST_DATA_ROW}:D{last_row}",
                        value_render_option="UNFORMATTED_VALUE")
    normalized_category = _normalize_category_text(category)
    matches = []
    for row_values in rows:
        if len(row_values) < 4:
            continue
        leaf = str(row_values[2])
        if _normalize_category_text(leaf) == normalized_category:
            matches.append(row_values[3])
    if not matches:
        raise ValueError(
            f"Category {category!r} not found in the commissions sheet - add it there before this can "
            "compute a price for it."
        )
    distinct = {float(m) for m in matches}
    if len(distinct) > 1:
        raise ValueError(
            f"Category {category!r} matches {len(matches)} rows in the commissions sheet with DIFFERENT "
            f"commissions ({sorted(distinct)}) - it exists under more than one parent category and Alif's "
            "own page doesn't say which one this item is under, so this can't pick the right one automatically. "
            "Resolve the ambiguity in the commissions sheet (or by hand) before reconciling this item."
        )
    return float(matches[0])


def mark_waiting(ws, row):
    """Application isn't resolved yet (Alif status = 'Yetkazib berishga') -
    just record that we checked, no color (unprocessed/pending rows aren't
    colored at all, confirmed live - only success/Cancelled are)."""
    _with_retry(ws.update_cell, row, COL_STATUS, STATUS_WAITING)


def mark_cancelled(ws, row):
    _with_retry(ws.update_cell, row, COL_STATUS, STATUS_CANCELLED)
    _with_retry(ws.format, ROW_COLOR_RANGE.format(row=row), {"backgroundColor": ROW_COLOR_CANCELLED})


def mark_success(ws, row, price):
    _with_retry(ws.update_cell, row, COL_SUMMA_1, price)
    _with_retry(ws.update_cell, row, COL_SUMMA_2, price)
    _with_retry(ws.update_cell, row, COL_STATUS, STATUS_SUCCESS)
    _with_retry(ws.format, ROW_COLOR_RANGE.format(row=row), {"backgroundColor": ROW_COLOR_SUCCESS})

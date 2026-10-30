# -*- coding: utf-8 -*-
"""Google Sheets access for the "IMAN Sverka" reconciliation workbook
(spreadsheet id from IMAN_SPREADSHEET_ID in credentials\\.env),
specifically the "IMANUM STORE" tab.

Column layout of that tab (row 1 = header, data from row 2), confirmed live:
    A  Дата Реализации          - IMAN's own record of the sale date
    B  Номер ИМАН               - IMAN's transaction id
    C  Статус ИМАН на момент отчета
    D  1С Номенклатура          - expected product name (compare against 1C)
    E  Бухгалтерия номенклатура
    F  Маркировка
    G  СуммаИМАН                - the amount IMAN says is correct (source of
                                   truth - this is what gets written into 1C)
    H  Номер Заказ 1            - the 1C "Заказ клиента" document number
    I  Сумма 1С                 - amount currently in 1C. BLANK exactly on
                                   rows nobody has reconciled yet - this is
                                   the "not checked" signal this module looks
                                   for. This script fills it in once done.
    J  (formula, =I-G)          - the diff, should recalculate itself off I
                                   once both are filled in - EXCEPT this
                                   formula was apparently never dragged down
                                   to newer rows (confirmed live: genuinely
                                   blank, not just displaying blank, for
                                   every row in the low-1970s+ range), so
                                   mark_sum_1c below also writes the =I-G
                                   formula into J itself if it's missing.
                                   Never touches J if it already has
                                   something (formula or otherwise).

--------------------------------------------------------------------- setup
This uses a Google service account (not OAuth) so it can run unattended.
One-time setup (only the sheet owner can do this - it's an access-control change):

  1. https://console.cloud.google.com/ -> select/create a project.
  2. APIs & Services -> Library -> enable "Google Sheets API".
  3. APIs & Services -> Credentials -> Create Credentials -> Service account.
     Give it any name (e.g. "sverka-bot"); no roles needed.
  4. Open the created service account -> Keys -> Add key -> Create new key ->
     JSON. This downloads a .json file.
  5. Save that file as credentials\\service_account.json in this folder
     (<project folder>\\credentials\\service_account.json) -
     create the "credentials" folder if it doesn't exist yet.
  6. Open the JSON file, copy the "client_email" value (looks like
     ...@...iam.gserviceaccount.com).
  7. Open the IMAN Sverka sheet -> Share -> paste that email -> give it
     "Editor" access (needed to write "Сумма 1С" back) -> Send/Share.

Once that's done, get_worksheet() below just works, no browser/login needed.
"""
import re
import time
from datetime import datetime
from pathlib import Path

import gspread
import requests
from google.oauth2.service_account import Credentials

from env_config import CREDENTIALS_DIR, load_env

_env = load_env()
# IMAN_SPREADSHEET_ID / IMAN_WORKSHEET_NAME in credentials\.env set these
# (see env_config.py and .env.example).
SPREADSHEET_ID = _env.get("IMAN_SPREADSHEET_ID", "")
DEFAULT_WORKSHEET = _env.get("IMAN_WORKSHEET_NAME", "IMANUM STORE")
DEFAULT_CREDENTIALS_PATH = CREDENTIALS_DIR / "service_account.json"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]

COL_DATA_REALIZATSII = 1
COL_NOMER_IMAN = 2
COL_STATUS_IMAN = 3
COL_1C_NOMENKLATURA = 4
COL_BUKH_NOMENKLATURA = 5
COL_MARKIROVKA = 6
COL_SUMMA_IMAN = 7
COL_NOMER_ZAKAZ = 8
COL_SUMMA_1C = 9
COL_DIFF = 10

HEADER_ROW = 1
FIRST_DATA_ROW = 2


def get_worksheet(worksheet_name=DEFAULT_WORKSHEET, credentials_path=DEFAULT_CREDENTIALS_PATH,
                   spreadsheet_id=SPREADSHEET_ID):
    credentials_path = Path(credentials_path)
    if not credentials_path.exists():
        raise FileNotFoundError(
            f"No service account key at {credentials_path} - see the setup steps in this "
            "file's module docstring (sheets_client.py) to create one."
        )
    creds = Credentials.from_service_account_file(str(credentials_path), scopes=SCOPES)
    client = gspread.authorize(creds)
    return client.open_by_key(spreadsheet_id).worksheet(worksheet_name)


def parse_sheet_amount(raw):
    """Parse the Google Sheet's display format for amounts, e.g. '2 705 080,00'
    or '2,705,080.00' or a plain number - handles both possible locale
    renderings gspread might hand back depending on the cell's number format,
    plus MALFORMED cells with several thousands separators mixed together
    (confirmed live: '2,549.000.00' and '5199.000.00' - inconsistent hand
    entry, both meaning 2 549 000 / 5 199 000). Returns None for a blank
    cell; raises ValueError only for genuinely unparseable text."""
    if raw is None:
        return None
    if isinstance(raw, (int, float)):
        return float(raw)
    s = str(raw).strip().replace("\xa0", "").replace(" ", "")
    if not s:
        return None

    seps = [c for c in s if c in ".,"]
    if len(seps) >= 2:
        # More than one separator -> all but the last are thousands
        # groupers (the sheet is just inconsistent about , vs . - e.g.
        # '2,549.000.00', '2,705,080.00', '2.705.080,00'). The LAST
        # separator is the decimal point ONLY if it's followed by 1-2 digits
        # (e.g. '.00'); otherwise it's a thousands separator too (e.g.
        # '5199.000' -> 5199000).
        last_sep = max(s.rfind(","), s.rfind("."))
        tail = s[last_sep + 1:]
        if len(tail) in (1, 2) and tail.isdigit():
            int_part = re.sub(r"[.,]", "", s[:last_sep])
            s = f"{int_part}.{tail}"
        else:
            s = re.sub(r"[.,]", "", s)
    elif "," in s:
        # Ambiguous: '2705080,00' (decimal) vs '2,705,080' (thousands). A
        # trailing ",dd" (exactly 2 digits) is treated as decimal - matches
        # every amount seen in this sheet so far.
        if re.search(r",\d{2}$", s):
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    return float(s)


def parse_sheet_date(raw):
    """Parse Дата Реализации's display format (DD.MM.YYYY, e.g. '15.07.2026').
    Returns None if blank or unparseable."""
    s = (raw or "").strip()
    if not s:
        return None
    try:
        return datetime.strptime(s, "%d.%m.%Y").date()
    except ValueError:
        return None


def _col_letter(col_index):
    letters = ""
    n = col_index
    while n:
        n, r = divmod(n - 1, 26)
        letters = chr(65 + r) + letters
    return letters


def _with_retry(fn, *args, retries=4, backoff=20, **kwargs):
    """Retry a gspread call with backoff if it hits the Sheets API's
    per-minute read/write quota (HTTP 429) - confirmed live this is a real
    risk for the batch loop: a synthetic no-delay test of ~10 orders hit
    '429 Quota exceeded for... Read requests per minute' well before it hit
    ~50 total API calls. The combined-range reads below cut each order's
    call count from ~5 down to ~1, but a long batch run (or this quota
    being shared with something else) can still add up - wait and retry a
    few times rather than letting one transient 429 kill the whole run.

    Also retries a raw transport-level failure (requests.exceptions.
    RequestException - e.g. 'RemoteDisconnected: Remote end closed
    connection without response') the same way - confirmed live during an
    Alif Shop batch run (see alif_sheets_client.py's copy of this function):
    uncaught, this crashed the WHOLE batch loop outright even though the 1C
    side of the application it happened on had already finished
    successfully moments earlier. No HTTP status to check here (unlike the
    429 case above), but always transient by nature, so just retry."""
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
            if attempt == retries:
                raise
            wait = backoff * attempt
            print(f"   (Google Sheets connection hiccup ({e!r}), waiting {wait}s and retrying, "
                  f"attempt {attempt}/{retries})")
            time.sleep(wait)


def _read_columns_a_to_i(ws, start_row=FIRST_DATA_ROW):
    """One combined range read (columns A:I) instead of separate
    col_values() calls per column - each col_values() call is its own
    Sheets API request, and doing 4-5 of them on every single loop
    iteration is exactly what triggered the 429 above. Returns rows
    (row-major, one list per sheet row from start_row onward) using the
    same FORMATTED display values col_values() returns, so
    parse_sheet_amount/parse_sheet_date don't need to change."""
    last_row = max(ws.row_count, start_row)
    return _with_retry(ws.get, f"A{start_row}:I{last_row}")


def _cell(row_values, col_index_1based):
    """0-based lookup into one row of _read_columns_a_to_i's result
    (col_index_1based is one of the module's COL_* constants, e.g.
    COL_NOMER_ZAKAZ) - safe against ragged rows (Sheets trims trailing
    empty cells from what get() returns)."""
    i = col_index_1based - 1
    return row_values[i].strip() if i < len(row_values) else ""


def find_text_numbers_in_column(ws, col_index, start_row=FIRST_DATA_ROW):
    """Some rows have this column (typically G, СуммаИМАН) stored as literal
    text (e.g. '2 705 080,00') instead of a real number - confirmed live,
    this breaks the J column's diff formula (=I-G) with #VALUE!, even though
    parse_sheet_amount reads the text version just fine. Returns
    [(row, raw_text, parsed_float), ...] for every such cell - use
    fix_text_numbers_in_column to actually convert them."""
    letter = _col_letter(col_index)
    last_row = max(ws.row_count, start_row)
    values = _with_retry(ws.get, f"{letter}{start_row}:{letter}{last_row}", value_render_option="UNFORMATTED_VALUE")
    found = []
    for i, row_vals in enumerate(values):
        if not row_vals:
            continue
        val = row_vals[0]
        if isinstance(val, str) and val.strip():
            parsed = parse_sheet_amount(val)
            if parsed is not None:
                found.append((start_row + i, val, parsed))
    return found


def fix_text_numbers_in_column(ws, col_index, start_row=FIRST_DATA_ROW):
    """Convert every text-stored number in this column to a real number, in
    place (one batched write). Returns the same list find_text_numbers_in_column
    would have, i.e. what was fixed."""
    to_fix = find_text_numbers_in_column(ws, col_index, start_row=start_row)
    if to_fix:
        cells = [gspread.cell.Cell(row, col_index, value) for row, _raw, value in to_fix]
        _with_retry(ws.update_cells, cells)
    return to_fix


def _ensure_numeric_cell(ws, row, col_index, expected_value):
    """Check a single cell's underlying (unformatted) value and convert it
    to a real number in place if it's currently stored as text. Returns
    True if a fix was made. Cheap (one small API call) - called on whatever
    row find_next_unchecked_row is about to hand back, so this class of
    problem self-heals as rows get processed, without needing a separate
    bulk-scan pass each time."""
    letter = _col_letter(col_index)
    raw = _with_retry(ws.get, f"{letter}{row}", value_render_option="UNFORMATTED_VALUE")
    current = raw[0][0] if raw and raw[0] else None
    if isinstance(current, str):
        _with_retry(ws.update_cell, row, col_index, expected_value)
        return True
    return False


def find_next_unchecked_row(ws, start_row=FIRST_DATA_ROW, since_date=None):
    """Return the first row (top to bottom) with a 1C order number (H) but no
    Сумма 1С (I) yet - i.e. not reconciled. None if there's nothing to do.

    since_date (a date, or a 'DD.MM.YYYY' string): if given, rows whose
    Дата Реализации (A) is earlier than this are skipped rather than
    stopping the scan - use this to skip past old backlog rows (e.g. ones
    that will never be found in 1C) instead of getting stuck re-attempting
    the same old row every run. A row with no parseable date is skipped too
    when since_date is set, since there's no way to tell if it's in range.

    Returns a dict: {row, order_number, product_name, expected_amount,
    g_cell_was_text} - the last one is True if СуммаИМАН for this row was
    just auto-converted from text to a real number (see
    _ensure_numeric_cell). Uses ONE combined range read (see
    _read_columns_a_to_i) - this used to be four separate col_values()
    calls, which is what hit the Sheets API's read quota during a batch run
    (confirmed live)."""
    if isinstance(since_date, str):
        since_date = parse_sheet_date(since_date)

    rows = _read_columns_a_to_i(ws, start_row)

    for offset, row_values in enumerate(rows):
        order_number = _cell(row_values, COL_NOMER_ZAKAZ)
        sum_1c = _cell(row_values, COL_SUMMA_1C)
        if not order_number or sum_1c:
            continue
        if since_date is not None:
            row_date = parse_sheet_date(_cell(row_values, COL_DATA_REALIZATSII))
            if row_date is None or row_date < since_date:
                continue
        row = start_row + offset
        raw_amount = _cell(row_values, COL_SUMMA_IMAN)
        amount_error = None
        amount_note = None
        try:
            expected_amount = parse_sheet_amount(raw_amount)
        except (ValueError, TypeError):
            # A single malformed amount cell (e.g. '5199.000.00' -
            # inconsistent thousand separators) must NOT abort the whole
            # batch scan. Pass it through as a skippable row (expected_amount
            # None + amount_error set) so reconcile_order logs it and the
            # batch moves on to the next contract, instead of raising up to
            # the batch loop's "Sheets outage -> stop everything" handler.
            expected_amount = None
            amount_error = str(raw_amount)
        else:
            # This sheet's amounts are normally "2 705 080,00" (one comma);
            # anything with 2+ separators was either copy-pasted with
            # grouping or malformed (e.g. '2,549.000.00'). It parses fine
            # now (digit-preserving - see parse_sheet_amount), but surface
            # the raw-vs-parsed in the report Comment so a wrong read is
            # easy to spot-check.
            if expected_amount is not None and sum(c in ".," for c in str(raw_amount)) >= 2:
                amount_note = f"source СуммаИМАН {str(raw_amount).strip()!r} read as {expected_amount:g}"
        g_cell_was_text = False
        if expected_amount is not None:
            g_cell_was_text = _ensure_numeric_cell(ws, row, COL_SUMMA_IMAN, expected_amount)
        return {
            "row": row,
            "order_number": order_number,
            "product_name": _cell(row_values, COL_1C_NOMENKLATURA),
            "expected_amount": expected_amount,
            "amount_error": amount_error,
            "amount_note": amount_note,
            "g_cell_was_text": g_cell_was_text,
        }
    return None


def count_unchecked_rows(ws, start_row=FIRST_DATA_ROW, since_date=None):
    """How many rows currently qualify as "not reconciled yet" (same rule as
    find_next_unchecked_row: order number present, Сумма 1С blank, and
    on/after since_date if given) - a cheap upfront count for the batch
    loop's progress display, not itself used to decide what to process
    (that's still find_next_unchecked_row, called fresh each iteration
    since rows get filled in as the loop goes). Shares the same combined
    range read as find_next_unchecked_row - see its docstring."""
    if isinstance(since_date, str):
        since_date = parse_sheet_date(since_date)

    rows = _read_columns_a_to_i(ws, start_row)

    count = 0
    for row_values in rows:
        order_number = _cell(row_values, COL_NOMER_ZAKAZ)
        sum_1c = _cell(row_values, COL_SUMMA_1C)
        if not order_number or sum_1c:
            continue
        if since_date is not None:
            row_date = parse_sheet_date(_cell(row_values, COL_DATA_REALIZATSII))
            if row_date is None or row_date < since_date:
                continue
        count += 1
    return count


def find_all_rows_for_order(ws, order_number, start_row=FIRST_DATA_ROW):
    """Return every sheet row (not just the first) whose Номер Заказ 1 (H)
    matches order_number, regardless of whether Сумма 1С is already filled
    in. For a multi-item order, each 1C Товары row corresponds to its own
    separate sheet row sharing the same order number - this is how
    automate_order_ekvayring's determine_item_price_pairs figures out which
    price belongs to which item. Returns [{row, product_name,
    expected_amount}, ...] in sheet order (not sorted/matched to anything -
    that's the caller's job). One combined range read (see
    _read_columns_a_to_i), same reasoning as find_next_unchecked_row."""
    rows = _read_columns_a_to_i(ws, start_row)

    results = []
    for offset, row_values in enumerate(rows):
        if _cell(row_values, COL_NOMER_ZAKAZ) != order_number:
            continue
        row = start_row + offset
        expected_amount = parse_sheet_amount(_cell(row_values, COL_SUMMA_IMAN))
        if expected_amount is not None:
            _ensure_numeric_cell(ws, row, COL_SUMMA_IMAN, expected_amount)
        results.append({
            "row": row,
            "product_name": _cell(row_values, COL_1C_NOMENKLATURA),
            "expected_amount": expected_amount,
        })
    return results


def mark_sum_1c(ws, row, amount):
    """Write the confirmed 1C amount back into Сумма 1С for this row, and
    make sure the diff column (J, =I-G) actually exists for this row - see
    the module docstring: confirmed live, that formula was apparently only
    ever entered by hand up to some older row and never dragged down to
    newer ones, so plenty of rows have a genuinely EMPTY J that writing to I
    alone does nothing for. Only writes J's formula if it's currently
    blank - never overwrites an existing formula or value there."""
    _with_retry(ws.update_cell, row, COL_SUMMA_1C, amount)

    letter_j = _col_letter(COL_DIFF)
    existing = _with_retry(ws.get, f"{letter_j}{row}", value_render_option="FORMULA")
    has_content = bool(existing and existing[0] and existing[0][0] not in (None, ""))
    if not has_content:
        letter_i, letter_g = _col_letter(COL_SUMMA_1C), _col_letter(COL_SUMMA_IMAN)
        _with_retry(ws.update_cell, row, COL_DIFF, f"={letter_i}{row}-{letter_g}{row}")

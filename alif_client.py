# -*- coding: utf-8 -*-
"""business.alif.uz ("Alif Business") web client - a sibling to the IMAN
Sverka flow (see automate_order_ekvayring.py, sheets_client.py) but the
"source of truth" for what's been sold is a web dashboard instead of a
desktop app, so this uses Playwright (browser automation) instead of
pywinauto/UIA.

Two SEPARATE tenants/accounts share this one module (same site, different
logins - confirmed live these are genuinely different business accounts,
not just different sections of one account):
  - "alifshop" - the Alif Shop marketplace reconciliation flow (application
    status/detail lookup by id, see get_application_status/
    get_application_detail below).
  - "alif_moliya" - the Alif Moliya flow, which has no per-order API/page
    to query at all - instead it works off a downloadable "Tovarlar" Excel
    report from /nasiya/reports (see get_report_row_last_generated/
    set_report_date_range/generate_and_download_report below).
Every function that logs in takes a `tenant` argument (default "alifshop"
for the functions that existed before alif_moliya was added, so existing
call sites don't need to change) - it selects both which credentials to
use and which session-state file to persist/reuse.

Credentials come from credentials\\.env (<tenant>_login=<phone digits>,
<tenant>_password=<password>, e.g. alifshop_login/alifshop_password and
alif_moliya_login/alif_moliya_password) - same "local secrets file, never
printed or committed" pattern as credentials\\service_account.json for
Google Sheets.

Login flow, confirmed live while building this:
  - The phone number field is a masked/formatted input (renders as
    "(99) 999-99-99") - filling it with .fill() (a direct DOM value set)
    silently does NOT work, the form behaves as if the field is empty/
    invalid. Real keystroke-by-keystroke typing (Locator.press_sequentially)
    is what actually works - confirmed live, .fill() failed every time,
    press_sequentially succeeded.
  - Login also failed at least once even with press_sequentially, with no
    visible error - retrying (same credentials, no code change) succeeded.
    Since this is a REAL business account (not a disposable test one),
    hammering the login endpoint on every run to debug this further wasn't
    worth the risk of tripping a fraud/rate-limit flag on the account - so
    this saves the authenticated session (cookies) to
    credentials\\alif_session.json after a successful login and reuses it
    on every subsequent call, only logging in again when there's no saved
    session or it's no longer valid (redirects to /auth/login). This also
    means normal usage only hits the login form rarely, not once per CLI
    invocation.
  - Headless Chromium's own analytics/fingerprinting showed up as
    "HeadlessChrome" in the page's own collected client hints - possibly
    related to the intermittent failure above (bot detection). Runs
    headED (headless=False) instead, same as this whole project already
    assumes a visible interactive Windows desktop (see lib1c.py/pywinauto).
"""
import datetime
import re
import time
from pathlib import Path

from openpyxl import load_workbook
from playwright.sync_api import sync_playwright

from env_config import CREDENTIALS_DIR, ENV_PATH, load_env
from lib1c import parse_ru_number

DEFAULT_TENANT = "alifshop"
# alifshop's session file predates the multi-tenant generalization and was
# already in use as "alif_session.json" (not "alifshop_session.json") -
# kept as an explicit override so existing saved sessions/logins aren't
# orphaned. Any other tenant gets the obvious "<tenant>_session.json".
_SESSION_FILENAMES = {
    "alifshop": "alif_session.json",
}

# ALIF_BUSINESS_BASE_URL in credentials\.env overrides this - see
# env_config.py's docstring on why there's still a hardcoded fallback.
BASE_URL = load_env().get("ALIF_BUSINESS_BASE_URL", "https://business.alif.uz")
LOGIN_URL = f"{BASE_URL}/auth/login"
APPLICATIONS_URL = f"{BASE_URL}/nasiya/applications"
REPORTS_URL = f"{BASE_URL}/nasiya/reports"


def _session_path_for(tenant):
    return CREDENTIALS_DIR / _SESSION_FILENAMES.get(tenant, f"{tenant}_session.json")


def _load_env(path=ENV_PATH):
    env = load_env(path)
    if not env:
        raise FileNotFoundError(
            f"No (or empty) {path} found - create it with <tenant>_login=<phone digits, no country code> "
            "and <tenant>_password=<password> for each tenant this is used with (see this module's "
            "docstring)."
        )
    return env


def _is_logged_in(page):
    """After navigating somewhere under /nasiya, we're logged in unless the
    app redirected us to /auth/login (confirmed live - that's exactly what
    an expired/missing session does)."""
    return "/auth/login" not in page.url


def _wait_until_navigated_away_from_login(page, timeout_s=20):
    """This SPA has continuous background network activity (analytics
    beacons, live polling - confirmed live), so wait_for_load_state
    ("networkidle") can hang well past the actual page being ready. Poll
    the URL directly instead - cheap and doesn't depend on network traffic
    ever going quiet."""
    end = time.time() + timeout_s
    while time.time() < end:
        if "/auth/login" not in page.url:
            return
        time.sleep(0.3)


def _perform_login(page, tenant):
    env = _load_env()
    login_key, password_key = f"{tenant}_login", f"{tenant}_password"
    for key in (login_key, password_key):
        if key not in env:
            raise ValueError(f"{ENV_PATH} is missing required key {key!r} for tenant {tenant!r}")
    page.goto(BASE_URL, wait_until="domcontentloaded", timeout=30000)
    page.wait_for_timeout(1200)  # let the SPA finish its initial client-side render
    phone_input = page.locator('input[type="text"]').first
    password_input = page.locator('input[type="password"]').first
    # .fill() does NOT work on this masked phone field - see module
    # docstring. press_sequentially (real keystrokes) does.
    phone_input.click()
    phone_input.press_sequentially(env[login_key], delay=30)
    password_input.click()
    password_input.press_sequentially(env[password_key], delay=30)
    page.click('button:has-text("Kirish")')
    _wait_until_navigated_away_from_login(page)
    page.wait_for_timeout(1000)
    if not _is_logged_in(page):
        raise RuntimeError(
            f"Alif Business login (tenant {tenant!r}) did not succeed (still on /auth/login after "
            "submitting) - this can happen intermittently even with correct credentials (confirmed live, "
            "for both tenants built so far). Check credentials\\.env is correct, and that nothing on the "
            "account (e.g. a security prompt) needs handling by hand, then try again."
        )


def get_authenticated_context(playwright, headless=False, tenant=DEFAULT_TENANT):
    """Returns a fresh Playwright BrowserContext already logged into the
    given Alif Business tenant (see this module's docstring for the
    tenants this is used with) - reuses that tenant's saved session
    (credentials\\<tenant>_session.json, see _session_path_for) if it's
    still valid, only doing a real login (see _perform_login) when there's
    no saved session or it's expired. Caller is responsible for closing the
    returned browser (context.browser)."""
    session_path = _session_path_for(tenant)
    browser = playwright.chromium.launch(headless=headless)

    if session_path.exists():
        context = browser.new_context(storage_state=str(session_path))
        page = context.new_page()
        page.goto(APPLICATIONS_URL, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(1200)
        if _is_logged_in(page):
            page.close()
            return context
        print(f"   (saved Alif Business session for tenant {tenant!r} has expired - logging in again)")
        page.close()
        context.close()

    context = browser.new_context()
    page = context.new_page()
    _perform_login(page, tenant)
    CREDENTIALS_DIR.mkdir(exist_ok=True)
    context.storage_state(path=str(session_path))
    page.close()
    return context


# Statuses confirmed live on the search-results page's "HOLAT" column /
# detail page's status badge - Uzbek text, matched as substrings since the
# badge sometimes has surrounding whitespace/other markup.
STATUS_WAITING = "Yetkazib berishga"     # out for delivery - not resolved yet, keep re-checking later
STATUS_READY = "Rasmiylashtirildi"       # finalized - safe to read its items/price and reconcile in 1C
STATUS_CONFIRMED = "Tasdiqlandi"         # confirmed/approved - also safe to reconcile (see reconcile_alifshop_application)
STATUS_CANCELLED = "Bekor qilindi"       # cancelled - mark the sheet row Cancelled/red, nothing to reconcile

# The statuses on which it's safe to read the items and reconcile in 1C.
STATUSES_READY_TO_RECONCILE = (STATUS_READY, STATUS_CONFIRMED)


def get_application_status(context, application_id):
    """Look up an application's current status via the search-results page
    (https://business.alif.uz/nasiya/applications?q=<id>) - confirmed live,
    the HOLAT column there shows the same status text as the detail page's
    own badge, and it's one request instead of also loading the (heavier)
    detail page just to check status. Returns the raw status text (one of
    the STATUS_* constants above, or possibly something else not seen yet -
    callers should not assume it's always one of those three)."""
    page = context.new_page()
    try:
        page.goto(f"{APPLICATIONS_URL}?q={application_id}&page=1", wait_until="domcontentloaded", timeout=30000)

        # Wait for the search-results row for THIS application to actually
        # render, rather than a fixed short sleep - same render-race class
        # as get_application_detail (see its comment): the results table's
        # rows load via a separate request after the page shell, so a blind
        # sleep could read an empty table and wrongly report "not found in
        # search results". Poll until the matching row appears. Only after
        # the timeout do we treat "no row" as a genuine not-found.
        deadline = time.time() + 15
        matching = []
        while time.time() < deadline:
            rows = page.query_selector_all("table tbody tr")
            matching = [r for r in rows if application_id in r.inner_text()]
            if matching:
                break
            page.wait_for_timeout(500)
        if not matching:
            raise RuntimeError(f"Application {application_id!r} not found in the Alif Business search results.")

        row_text = matching[0].inner_text()
        for status in (STATUS_WAITING, STATUS_READY, STATUS_CONFIRMED, STATUS_CANCELLED):
            if status in row_text:
                return status
        raise RuntimeError(
            f"Application {application_id!r}'s status wasn't recognized (row text: {row_text!r}) - "
            "this is a status this script hasn't seen before, check it by hand."
        )
    finally:
        page.close()


def get_application_detail(context, application_id):
    """Read the application's detail page
    (https://business.alif.uz/nasiya/applications/<id>) - returns
    {"items": [{"name", "category", "price", "kolichestvo"}]}.

    Deliberately does NOT also scrape an overall application status from
    this page - confirmed live, there's no reliably-classed status badge
    element here (a class-based search for badge/status/tag/chip elements
    found none), and a plain text search for the known status strings
    matched an UNRELATED "Yetkazib berish..." section elsewhere on the page
    (a delivery-tracking sub-section, not the application's own status) -
    got a wrong answer for an application that was actually
    "Rasmiylashtirildi". get_application_status (the search-results page)
    is the reliable source for status - call that first, same order the
    real workflow checks things in (status, then items only once it's
    "Rasmiylashtirildi").

    The Tovarlar table's row structure (confirmed live via outerHTML dump):
    4 <td>s - NOMI (two stacked <p>s: display name, then the leaf category
    used for the commission lookup, see alif_sheets_client.py), SKU (the
    1C-matching product name - confirmed against the sheet's "1С
    Номенклатура" column), NARXI (price), SONI (quantity - initially missed
    this column entirely since it's the 4th cell, out of the visible area
    in a screenshot that only showed a horizontal scrollbar hinting at
    more)."""
    page = context.new_page()
    try:
        page.goto(f"{APPLICATIONS_URL}/{application_id}", wait_until="domcontentloaded", timeout=30000)
        if not _is_logged_in(page):
            raise RuntimeError("Session was logged out while reading the detail page - check by hand.")

        # Wait for the Tovarlar table's DATA row(s) to actually render, not
        # just its header - confirmed live (batch run, application
        # 1000001): the header row ("NOMI/SKU/NARXI/SONI") appears a moment
        # BEFORE the data row populates (it loads via a separate request),
        # so a fixed short sleep could read the right table with the header
        # present but ZERO data rows, spuriously reporting "0 items" and
        # skipping an application that genuinely has one. Poll until a
        # matching table has a >=4-cell body row (same "wait for the real
        # thing, not a fixed sleep" pattern as launch_1c's responsiveness
        # wait). Falls through after the timeout so a genuinely empty page
        # still returns [] (handled safely upstream) rather than hanging.
        deadline = time.time() + 15
        tovarlar = None
        while time.time() < deadline:
            for table in page.query_selector_all("table"):
                header_text = table.inner_text()
                if "NOMI" not in header_text or "NARXI" not in header_text:
                    continue
                if any(len(r.query_selector_all("td")) >= 4 for r in table.query_selector_all("tbody tr")):
                    tovarlar = table
                break  # first NOMI/NARXI table is the Tovarlar table, checked or not
            if tovarlar is not None:
                break
            page.wait_for_timeout(500)

        items = []
        if tovarlar is not None:
            for row in tovarlar.query_selector_all("tbody tr"):
                cells = row.query_selector_all("td")
                if len(cells) < 4:
                    continue
                nomi_cell, sku_cell, narxi_cell, soni_cell = cells[0], cells[1], cells[2], cells[3]
                name_lines = nomi_cell.query_selector_all("p")
                category = name_lines[1].inner_text().strip() if len(name_lines) > 1 else ""
                name = sku_cell.inner_text().strip()
                price = _parse_uzs_amount(narxi_cell.inner_text().strip())
                kolichestvo = _parse_soni(soni_cell.inner_text().strip())
                items.append({"name": name, "category": category, "price": price, "kolichestvo": kolichestvo})

        return {"items": items}
    finally:
        page.close()


def _parse_soni(text):
    """Parse the SONI (quantity) cell - plain integer text, confirmed live
    (e.g. '1')."""
    text = text.strip()
    return float(text) if text else None


def _parse_uzs_amount(text):
    """Parse Alif Business's amount display (e.g. '829 000 so'm') to a
    float."""
    cleaned = re.sub(r"[^\d,.]", "", text.replace("\xa0", "").replace(" ", ""))
    if not cleaned:
        return None
    if "," in cleaned and "." not in cleaned:
        cleaned = cleaned.replace(",", ".")
    return float(cleaned)


# --------------------------------------------------- Alif Moliya reports

REPORT_TOVARLAR = "Tovarlar"

# The reports page (/nasiya/reports) is a totally different UI shape from
# the applications pages above - a table of report TYPES (Tovarlar,
# Qaytarilgan tovarlar, Yetkazib berishlar, Solishtirma Dalolatnomasi),
# each with its own date-range picker, a "Yaratish" (generate) button, a
# "So'nggi generatsiya" (last generated) timestamp, and a "Yuklash"
# (download) button. Alif Moliya has no per-order API/page to query at all
# (unlike Alif Shop's application status/detail lookups above) - this is
# the ONLY way to get at what it sold, via the downloadable Tovarlar
# report (see automate_order_ekvayring.py's Alif Moliya flow).
#
# Confirmed live while building this:
#   - The page shows a loading skeleton for several seconds before the
#     real table renders (an empty/placeholder table briefly READS as 15+
#     blank rows, not the real ~4) - wait at least ~9s after navigating
#     before trusting anything read from this page.
#   - The date-range input (e.g. value "13-07-2026 – 18-07-2026") is
#     read-only - typing into it (even via real keystrokes) does NOTHING,
#     confirmed live. The only way to change it is through its calendar
#     popup, clicking actual day cells.
#   - The calendar's prev/next month-navigation buttons do NOT respond to
#     a JS-dispatched .click() (e.g. via page.evaluate calling
#     element.click()) - confirmed live, this closed the popup instead of
#     navigating (React's synthetic event system apparently requires a
#     real/trusted pointer event, not an untrusted DOM click). A real
#     Playwright mouse click (page.mouse.click at the button's own
#     getBoundingClientRect() center) DOES work - confirmed live, this is
#     what _navigate_calendar_to/_click_calendar_day both use throughout.
#   - Day cells for the CURRENTLY-displayed month vs. adjacent-month
#     padding days (e.g. the tail end of June shown before July 1st) can
#     show the SAME day NUMBER at once (e.g. two "29"s visible - June's and
#     July's) - the out-of-month one is styled with an "opacity-50" class
#     on its wrapping button, confirmed live via outerHTML - _click_
#     calendar_day filters those out rather than risking clicking the
#     wrong one.
#   - "Yaratish" (generate) sometimes shows its usual success feedback but
#     the "So'nggi generatsiya" timestamp does NOT actually update -
#     confirmed by the user manually reproducing this live - a second
#     click is needed. generate_report retries rather than trusting one
#     click ever worked.

_MONTH_NAMES_RU = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль",
                   "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
_MONTH_NUM = {name: i + 1 for i, name in enumerate(_MONTH_NAMES_RU)}

_MONTHS_JS_ARRAY = "[" + ",".join(f'"{m}"' for m in _MONTH_NAMES_RU) + "]"

_CALENDAR_HEADER_JS = f"""
() => {{
    const monthNames = {_MONTHS_JS_ARRAY};
    let header = null;
    document.querySelectorAll('button').forEach(el => {{
        const t = (el.textContent||'').trim();
        if (t.length < 20 && monthNames.some(m => t.includes(m))) {{
            if (!header || header.textContent.trim().length > t.length) header = el;
        }}
    }});
    return header ? header.textContent.trim() : null;
}}
"""

_CALENDAR_NAV_RECTS_JS = f"""
() => {{
    const monthNames = {_MONTHS_JS_ARRAY};
    let header = null;
    document.querySelectorAll('button').forEach(el => {{
        const t = (el.textContent||'').trim();
        if (t.length < 20 && monthNames.some(m => t.includes(m))) {{
            if (!header || header.textContent.trim().length > t.length) header = el;
        }}
    }});
    if (!header) return null;
    const level2 = header.parentElement.parentElement;
    const buttons = Array.from(level2.querySelectorAll('button'));
    const idx = buttons.indexOf(header);
    const rectOf = (el) => {{ const r = el.getBoundingClientRect(); return {{x: r.x + r.width / 2, y: r.y + r.height / 2}}; }};
    return {{prev: buttons[idx - 1] ? rectOf(buttons[idx - 1]) : null, next: buttons[idx + 1] ? rectOf(buttons[idx + 1]) : null}};
}}
"""

_CALENDAR_DAY_RECT_JS = f"""
(day) => {{
    const monthNames = {_MONTHS_JS_ARRAY};
    let header = null;
    document.querySelectorAll('button').forEach(el => {{
        const t = (el.textContent||'').trim();
        if (t.length < 20 && monthNames.some(m => t.includes(m))) {{
            if (!header || header.textContent.trim().length > t.length) header = el;
        }}
    }});
    if (!header) return null;
    let container = header;
    for (let i = 0; i < 4 && container.parentElement; i++) container = container.parentElement;
    const buttons = Array.from(container.querySelectorAll('button'));
    const target = buttons.find(b => {{
        const span = b.querySelector('span');
        const t = span ? span.textContent.trim() : '';
        return t === String(day) && !b.className.includes('opacity-50');
    }});
    if (!target) return null;
    const r = target.getBoundingClientRect();
    return {{x: r.x + r.width / 2, y: r.y + r.height / 2}};
}}
"""


def _report_row(page, report_name):
    rows = page.query_selector_all("table tr")
    for row in rows[1:]:
        cells = row.query_selector_all("td")
        if cells and cells[0].inner_text().strip() == report_name:
            return row
    raise RuntimeError(f"Report row {report_name!r} not found on {REPORTS_URL} - check the name/page by hand.")


def _parse_calendar_header(text):
    if not text:
        raise RuntimeError("Could not find the calendar's month/year header - check the page layout by hand.")
    name, _, year = text.rpartition(" ")
    month = _MONTH_NUM.get(name.strip())
    if month is None or not year.strip().isdigit():
        raise RuntimeError(f"Unrecognized calendar header text: {text!r} - check the page layout by hand.")
    return month, int(year)


def _navigate_calendar_to(page, target_month, target_year, max_steps=36):
    """Click the calendar's prev/next month arrows (real mouse clicks, see
    this section's module-level notes) until its header shows
    target_month/target_year. max_steps is a generous 3-year bound - this
    should only ever need a handful of clicks in practice (report ranges
    here span at most a few days)."""
    for _ in range(max_steps):
        month, year = _parse_calendar_header(page.evaluate(_CALENDAR_HEADER_JS))
        if (month, year) == (target_month, target_year):
            return
        rects = page.evaluate(_CALENDAR_NAV_RECTS_JS)
        direction = "next" if (target_year, target_month) > (year, month) else "prev"
        rect = rects.get(direction) if rects else None
        if not rect:
            raise RuntimeError(f"Calendar {direction!r} button not found while navigating to {target_month}/{target_year}.")
        page.mouse.click(rect["x"], rect["y"])
        page.wait_for_timeout(300)
    raise RuntimeError(f"Could not navigate the calendar to {target_month}/{target_year} after {max_steps} steps.")


def _click_calendar_day(page, day):
    rect = page.evaluate(_CALENDAR_DAY_RECT_JS, day)
    if not rect:
        raise RuntimeError(
            f"Day {day} not found (or only found as a grayed-out adjacent-month day) in the currently "
            "visible calendar month."
        )
    page.mouse.click(rect["x"], rect["y"])
    page.wait_for_timeout(300)


def _parse_date_range_value(text):
    """Parse the date-range input's displayed value (e.g. '13-07-2026 –
    18-07-2026') into (start_date, end_date) - tolerant of the exact dash
    character used as separator (confirmed live it's an en-dash, but
    comparing by re-parsing rather than exact string equality is more
    robust than assuming that never changes)."""
    # The string itself uses plain '-' WITHIN each date (DD-MM-YYYY) and a
    # dash-like separator BETWEEN the two dates - match each date as a
    # whole DD-MM-YYYY group rather than splitting on every '-', which
    # would also split the dates themselves apart.
    match = re.match(r"^(\d{2}-\d{2}-\d{4})\s*[^\d\s]\s*(\d{2}-\d{2}-\d{4})$", text.strip())
    if not match:
        raise RuntimeError(f"Could not parse date-range input value {text!r}.")
    return (datetime.datetime.strptime(match.group(1), "%d-%m-%Y").date(),
            datetime.datetime.strptime(match.group(2), "%d-%m-%Y").date())


def set_report_date_range(page, report_name, start_date, end_date):
    """Set a /nasiya/reports row's date-range picker to [start_date,
    end_date] (both datetime.date) by clicking through its calendar popup -
    see this section's module-level notes for why (read-only input, needs
    real mouse clicks not JS ones)."""
    row = _report_row(page, report_name)
    date_input = row.query_selector('input[type="text"]')
    if date_input is None:
        raise RuntimeError(f"Report row {report_name!r} has no date-range input - check the page layout by hand.")
    date_input.click()
    page.wait_for_timeout(500)

    _navigate_calendar_to(page, start_date.month, start_date.year)
    _click_calendar_day(page, start_date.day)
    _navigate_calendar_to(page, end_date.month, end_date.year)
    _click_calendar_day(page, end_date.day)

    actual_start, actual_end = _parse_date_range_value(date_input.get_attribute("value") or "")
    if (actual_start, actual_end) != (start_date, end_date):
        raise RuntimeError(
            f"Date range reads {actual_start}..{actual_end} after setting it, expected "
            f"{start_date}..{end_date} - check by hand."
        )


def generate_report(page, report_name, max_attempts=3, wait_after_click_s=6.0):
    """Click 'Yaratish' for this report row and confirm the 'So'nggi
    generatsiya' timestamp actually changed - confirmed live (reproduced by
    hand) that a single click sometimes leaves the timestamp exactly as it
    was, needing a second click. Retries rather than trusting the first
    click worked."""
    row = _report_row(page, report_name)
    before = row.query_selector_all("td")[2].inner_text()

    for attempt in range(1, max_attempts + 1):
        row = _report_row(page, report_name)
        yaratish_btn = row.query_selector('button:has-text("Yaratish")')
        if yaratish_btn is None:
            raise RuntimeError(f"'Yaratish' button not found for report {report_name!r}.")
        yaratish_btn.click()
        page.wait_for_timeout(int(wait_after_click_s * 1000))
        row = _report_row(page, report_name)
        after = row.query_selector_all("td")[2].inner_text()
        if after != before:
            return after
        print(f"   ('So'nggi generatsiya' for {report_name!r} did not change after attempt {attempt}/"
              f"{max_attempts} - retrying, confirmed live this sometimes needs a second click)")

    raise RuntimeError(
        f"'So'nggi generatsiya' for {report_name!r} never changed after {max_attempts} Yaratish clicks - "
        "check the report page by hand."
    )


def download_report(page, report_name, download_dir):
    """Click 'Yuklash' for this report row and save the downloaded file,
    returning its path."""
    row = _report_row(page, report_name)
    yuklash_btn = row.query_selector('button:has-text("Yuklash")')
    if yuklash_btn is None:
        raise RuntimeError(f"'Yuklash' button not found for report {report_name!r}.")
    with page.expect_download(timeout=30000) as download_info:
        yuklash_btn.click()
    download = download_info.value
    download_dir = Path(download_dir)
    download_dir.mkdir(parents=True, exist_ok=True)
    dest = download_dir / download.suggested_filename
    download.save_as(str(dest))
    return dest


def get_tovarlar_report(playwright, start_date, end_date, tenant="alif_moliya", download_dir=None,
                         headless=False):
    """End-to-end: log into `tenant`, navigate to /nasiya/reports, set the
    Tovarlar report's date range to [start_date, end_date], generate it
    (retrying until the timestamp actually updates), download it, and
    return the saved file's path. Caller does NOT need to manage the
    browser/context - this owns and closes its own."""
    download_dir = Path(download_dir) if download_dir else (CREDENTIALS_DIR.parent / "downloads")
    context = get_authenticated_context(playwright, headless=headless, tenant=tenant)
    try:
        page = context.new_page()
        page.goto(REPORTS_URL, wait_until="domcontentloaded", timeout=30000)
        # Confirmed live: this page shows a loading skeleton (looks like
        # 15+ blank rows) for several seconds before the real ~4-row table
        # renders - trusting anything read before that finishes reads
        # garbage.
        page.wait_for_timeout(9000)
        set_report_date_range(page, REPORT_TOVARLAR, start_date, end_date)
        generate_report(page, REPORT_TOVARLAR)
        path = download_report(page, REPORT_TOVARLAR, download_dir)
        page.close()
        return path
    finally:
        context.browser.close()


# Header text -> internal key. Fields we actually consume downstream; other
# columns in the file (Магазин, Регион, ИНН, IMEI, Маркировка, etc.) are
# ignored. Keyed by header name because Alif has already re-ordered these
# once - confirmed live: a "Цена (за вычетом партнерской комиссии)" column
# was inserted between "Цена" and "SKU" (August 2026), which shifted SKU
# and Количество one column to the right and crashed the old positional
# parser (row[7] was suddenly product text, not a float).
_TOVARLAR_COLUMN_MAP = {
    "Дата": "_date_text",
    "Id заявки": "application_id",
    "Клиент": "client",
    "Название товара": "product_name",
    "Цена": "_price_text",
    "SKU": "sku",
    "Количество": "_kolichestvo_raw",
}


def parse_tovarlar_report(path):
    """Parse a downloaded Tovarlar report (.xlsx, see get_tovarlar_report)
    into a list of {date, application_id, client, product_name, price,
    sku, kolichestvo} dicts, one per data row (oldest first, matching the
    file's own order).

    Fields are read by HEADER NAME (see _TOVARLAR_COLUMN_MAP) - Alif has
    reordered and inserted columns before, and will again. "Дата" is
    DD.MM.YYYY TEXT (confirmed live, NOT a real Excel date). "Цена" is
    comma-decimal text (e.g. '3799000,00'), same format lib1c.parse_ru_number
    already handles. "SKU" is the SPECIFIC product name that matches 1C's
    Номенклатура (unlike "Название товара", which is a broad/generic
    display name - the same generic-vs-specific split as Alif Shop's
    application detail page, see get_application_detail) - use SKU, not
    "Название товара", for matching against 1C. Columns beyond these
    (Магазин/Регион/Район/ИНН/Период/Дата первого платежа/Дата одобрения/
    IMEI/Маркировка/... and the newly-inserted "Цена (за вычетом партнерской
    комиссии)") aren't needed for reconciliation and aren't returned."""
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.active
    rows_iter = ws.iter_rows(values_only=True)
    header = next(rows_iter, None) or ()
    header_texts = [str(c).strip() if c is not None else "" for c in header]
    col_by_key = {}
    missing = []
    for header_text, key in _TOVARLAR_COLUMN_MAP.items():
        try:
            col_by_key[key] = header_texts.index(header_text)
        except ValueError:
            missing.append(header_text)
    if missing:
        wb.close()
        raise RuntimeError(
            f"Tovarlar report {path!r} is missing required column(s): {missing}. Actual header row: "
            f"{header_texts}. Alif may have renamed these columns - update _TOVARLAR_COLUMN_MAP.")

    def cell(row, key):
        idx = col_by_key[key]
        return row[idx] if idx < len(row) else None

    rows = []
    for row in rows_iter:
        if row is None or cell(row, "_date_text") is None:
            continue
        try:
            date = datetime.datetime.strptime(str(cell(row, "_date_text")).strip(), "%d.%m.%Y").date()
        except ValueError:
            continue
        price_raw = cell(row, "_price_text")
        kol_raw = cell(row, "_kolichestvo_raw")
        rows.append({
            "date": date,
            "application_id": str(cell(row, "application_id")).strip() if cell(row, "application_id") is not None else "",
            "client": str(cell(row, "client")).strip() if cell(row, "client") is not None else "",
            "product_name": str(cell(row, "product_name")).strip() if cell(row, "product_name") is not None else "",
            "price": parse_ru_number(str(price_raw)) if price_raw is not None else None,
            "sku": str(cell(row, "sku")).strip() if cell(row, "sku") is not None else "",
            "kolichestvo": float(kol_raw) if kol_raw is not None else None,
        })
    wb.close()
    return rows

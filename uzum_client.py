# -*- coding: utf-8 -*-
"""merchant.uzumnasiya.uz ("Uzum Nasiya Business") web client - a sibling
to alif_client.py's Alif Moliya support, but for a completely different
site/vendor. Like Alif Moliya, there's no per-order API/page to query at
all - this works off a downloadable CSV report ("Отчет по филиалам с
датами отмененных") covering whatever date range is requested (see
get_cancel_dates_report/parse_report), used by automate_order_ekvayring.py's
Uzum Nasiya flow the same way Alif Moliya's Tovarlar report is.

Credentials come from credentials\\.env (uzum_nasiya_login,
uzum_nasiya_password) - same pattern as every other secret in this
project. Base URL is also configurable there (UZUM_NASIYA_BASE_URL),
falling back to the real one if not set (see env_config.py's docstring).

Login flow, confirmed live while building this: a plain (unmasked)
ID/password form - "ID компании" (company id) then "Пароль" (password),
submitted via a "Продолжить" button. Unlike Alif Business's masked phone
field, .fill() would likely work here too, but press_sequentially (real
keystrokes) is used for consistency/already-confirmed reliability with the
rest of this project's web automation.

The reports page (/merchant/reports) uses Naive UI (Vue) components:
  - "Период" (period) and "Отчеты" (report type) are custom dropdowns
    (`.n-base-selection-label` - NOT native <select>s) - open by clicking
    that label element, then click the option text (a plain page.click
    ("text=...") works, confirmed live).
  - The date range picker (opened by clicking the "Дата начала" input) is
    a Naive UI daterange panel showing TWO month calendars side by side,
    classed ".n-date-panel-calendar--start" (left) and "...--end" (right).
    Confirmed live this is a bit surprising: which VISUAL panel you click
    in does NOT determine start vs end - the component tracks that by
    CLICK ORDER (1st click = start, 2nd = end), regardless of which panel
    column it lands in. So both target days can be (and here, always are)
    clicked in the SAME (start) panel, one after the other, avoiding any
    need to figure out which physical panel to use for which date.
  - Each panel has its own independent month navigation
    (.n-date-panel-month__prev/__next, plus __fast-prev/__fast-next for
    whole years) - confirmed live the panel ALWAYS opens showing [today's
    month, today's month + 1] in its two columns, so month navigation here
    is computed as an integer offset from datetime.date.today() rather
    than by parsing the panel's own abbreviated Russian month/year header
    text (e.g. "июл. 2026") into a month number, which would need a
    hand-maintained abbreviation table for no real benefit.
  - Day cells are `.n-date-panel-date` elements; out-of-month padding days
    get an extra `n-date-panel-date--excluded` class (same
    "exclude greyed-out adjacent-month days" idea as Alif Moliya's
    calendar, different class name).
  - Clicking "Подтвердить" (Confirm) commits the range; "Скачать отчеты"
    (Download reports) then triggers the actual CSV download.
"""
import csv
import datetime
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

from env_config import CREDENTIALS_DIR, load_env

_env = load_env()
BASE_URL = _env.get("UZUM_NASIYA_BASE_URL", "https://merchant.uzumnasiya.uz")
REPORTS_URL = f"{BASE_URL}/merchant/reports"

SESSION_STATE_PATH = CREDENTIALS_DIR / "uzum_nasiya_session.json"

REPORT_NAME = "Отчет по филиалам с датами отмененных"


def _is_logged_in(page):
    return "/auth/login" not in page.url


def _perform_login(page):
    login = _env.get("uzum_nasiya_login")
    password = _env.get("uzum_nasiya_password")
    if not login or not password:
        raise ValueError(
            "uzum_nasiya_login / uzum_nasiya_password not found in credentials\\.env - can't log in."
        )
    page.goto(BASE_URL, wait_until="networkidle", timeout=30000)
    page.wait_for_timeout(1500)
    id_input = page.locator('input[type="text"]').first
    password_input = page.locator('input[type="password"]').first
    id_input.click()
    id_input.press_sequentially(login, delay=30)
    password_input.click()
    password_input.press_sequentially(password, delay=30)
    page.click('button:has-text("Продолжить")')

    end = time.time() + 20
    while time.time() < end and "/auth/login" in page.url:
        time.sleep(0.3)
    if not _is_logged_in(page):
        raise RuntimeError(
            "Uzum Nasiya login did not succeed (still on /auth/login after submitting) - check "
            "credentials\\.env is correct, then try again."
        )


def get_authenticated_context(playwright, headless=False):
    """Returns a fresh Playwright BrowserContext already logged into Uzum
    Nasiya Business - reuses the saved session (credentials\\
    uzum_nasiya_session.json) if it's still valid, only doing a real login
    when there's no saved session or it's expired. Caller is responsible
    for closing the returned browser (context.browser)."""
    browser = playwright.chromium.launch(headless=headless)

    if SESSION_STATE_PATH.exists():
        context = browser.new_context(storage_state=str(SESSION_STATE_PATH))
        page = context.new_page()
        page.goto(REPORTS_URL, wait_until="networkidle", timeout=30000)
        page.wait_for_timeout(1000)
        if _is_logged_in(page):
            page.close()
            return context
        print("   (saved Uzum Nasiya session has expired - logging in again)")
        page.close()
        context.close()

    context = browser.new_context()
    page = context.new_page()
    _perform_login(page)
    CREDENTIALS_DIR.mkdir(exist_ok=True)
    context.storage_state(path=str(SESSION_STATE_PATH))
    page.close()
    return context


def _click_day_in_panel(page, panel_selector, day):
    handle = page.evaluate_handle(
        """([panelSelector, day]) => {
            const panel = document.querySelector(panelSelector);
            if (!panel) return null;
            const cells = Array.from(panel.querySelectorAll('.n-date-panel-date'));
            return cells.find(c => c.textContent.trim() === String(day) && !c.className.includes('excluded'));
        }""",
        [panel_selector, day],
    ).as_element()
    if handle is None:
        raise RuntimeError(f"Day {day} not found (or only as an excluded adjacent-month day) in "
                            f"{panel_selector} - check the page by hand.")
    handle.click()


def _navigate_panel(page, panel_selector, months_delta):
    if months_delta == 0:
        return
    direction_cls = "next" if months_delta > 0 else "prev"
    for _ in range(abs(months_delta)):
        page.click(f"{panel_selector} .n-date-panel-month__{direction_cls}")
        page.wait_for_timeout(200)


def set_report_date_range(page, start_date, end_date):
    """Set the reports page to Период=Произвольный период, Отчеты=
    REPORT_NAME, and the date range to [start_date, end_date] (both
    datetime.date) - see this module's docstring for the Naive UI
    component quirks this navigates around."""
    labels = page.query_selector_all(".n-base-selection-label")
    if len(labels) < 2:
        raise RuntimeError("Could not find the Период/Отчеты dropdown triggers - check the page by hand.")
    labels[0].click()
    page.wait_for_timeout(400)
    page.click("text=Произвольный период")
    page.wait_for_timeout(400)

    labels = page.query_selector_all(".n-base-selection-label")
    labels[1].click()
    page.wait_for_timeout(400)
    page.click(f"text={REPORT_NAME}")
    page.wait_for_timeout(400)

    date_start_input = page.locator('input[placeholder="Дата начала"]')
    date_start_input.click()
    page.wait_for_timeout(400)

    today = datetime.date.today()
    baseline = today.month + today.year * 12
    start_delta = (start_date.month + start_date.year * 12) - baseline
    _navigate_panel(page, ".n-date-panel-calendar--start", start_delta)
    _click_day_in_panel(page, ".n-date-panel-calendar--start", start_date.day)
    page.wait_for_timeout(300)

    end_delta = (end_date.month + end_date.year * 12) - (baseline + start_delta)
    _navigate_panel(page, ".n-date-panel-calendar--start", end_delta)
    _click_day_in_panel(page, ".n-date-panel-calendar--start", end_date.day)
    page.wait_for_timeout(300)

    page.click("button:has-text('Подтвердить')")
    page.wait_for_timeout(500)

    actual_start = date_start_input.input_value()
    actual_end = page.locator('input[placeholder="Дата окончания"]').input_value()
    if actual_start != start_date.isoformat() or actual_end != end_date.isoformat():
        raise RuntimeError(
            f"Date range reads {actual_start}..{actual_end} after setting it, expected "
            f"{start_date.isoformat()}..{end_date.isoformat()} - check by hand."
        )


def download_report(page, download_dir):
    download_dir = Path(download_dir)
    download_dir.mkdir(parents=True, exist_ok=True)
    with page.expect_download(timeout=30000) as download_info:
        page.click("button:has-text('Скачать отчеты')")
    download = download_info.value
    dest = download_dir / download.suggested_filename
    download.save_as(str(dest))
    return dest


def get_cancel_dates_report(playwright, start_date, end_date, download_dir=None, headless=False):
    """End-to-end: log in, navigate to /merchant/reports, set the period/
    report-type/date-range, download the CSV, and return its path. Caller
    does NOT need to manage the browser/context - this owns and closes its
    own."""
    download_dir = Path(download_dir) if download_dir else (CREDENTIALS_DIR.parent / "downloads")
    context = get_authenticated_context(playwright, headless=headless)
    try:
        page = context.new_page()
        page.goto(REPORTS_URL, wait_until="networkidle", timeout=30000)
        page.wait_for_timeout(1500)
        set_report_date_range(page, start_date, end_date)
        path = download_report(page, download_dir)
        page.close()
        return path
    finally:
        context.browser.close()


def parse_report(path):
    """Parse the downloaded CSV (semicolon-delimited, UTF-8 with a BOM -
    confirmed live) into [{"date", "product_name", "kolichestvo",
    "price"}, ...]. Column layout confirmed live from a real download:
        5   Создан - 'YYYY-MM-DD HH:MM:SS' text (only the date part is
            used - matching against 1C's own order date, same idea as
            Alif Moliya's report)
        10  Наименование Товара - matches 1C's Номенклатура (confirmed
            live against a real order)
        11  Количество
        21  Стоимость одного товара с НДС и комиссией - the price to use
            (already commission-inclusive, unlike Alif Shop/Moliya there's
            no separate commission subtraction step needed here)
    Amounts here are plain period-decimal text (e.g. '2349000.00') -
    NOT the comma-decimal format lib1c.parse_ru_number handles elsewhere in
    this project, so parsed with a plain float() instead."""
    rows = []
    with open(path, encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f, delimiter=";")
        for i, row in enumerate(reader):
            if i == 0 or len(row) < 22:
                continue
            created_text = row[5].strip()
            if not created_text:
                continue
            try:
                date = datetime.datetime.strptime(created_text.split(" ")[0], "%Y-%m-%d").date()
            except ValueError:
                continue
            kolichestvo_text = row[11].strip()
            price_text = row[21].strip()
            rows.append({
                "date": date,
                "product_name": row[10].strip(),
                "kolichestvo": float(kolichestvo_text) if kolichestvo_text else None,
                "price": float(price_text) if price_text else None,
            })
    return rows

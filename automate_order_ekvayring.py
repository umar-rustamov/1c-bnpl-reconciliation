# -*- coding: utf-8 -*-
"""Automates the IMAN Sverka reconciliation workflow in the 1C:Enterprise
desktop client ("<Company> / Управление торговлей"), driven by the
"IMANUM STORE" tab of the IMAN Sverka Google Sheet (see sheets_client.py).

Prerequisite: none, as far as 1C itself goes - main() calls
launch_1c.ensure_1c_open() before anything else, which launches the 1C
thin client and logs in automatically (credentials\\.env's 1c_login/
1c_password) if it isn't already open, or does nothing if it is (see
launch_1c.py). Either way, this script then navigates from wherever the
app currently is via Продажи -> Заказы клиентов. A Google service account
key must exist at credentials\\service_account.json - see
sheets_client.py's module docstring for the one-time setup steps.

Two modes: process just the next unreconciled row (default, interactive
confirm() prompts before every post/write), or `--since DD.MM.YYYY --yes`
to loop over and reconcile EVERY unreconciled row from that date forward in
one run, unattended, printing a progress line per order and a final report
(see process_all_unchecked_orders). Either way, each order goes through the
same flow below.

Flow, per sheet row whose "Номер Заказ 1" is filled in but "Сумма
1С" is still blank (= not reconciled yet):
  1. Find that order in 1C's Заказы клиентов list. It can be filed under
     either the "IMANUM marketplace" or "IMAN PAY" Соглашение - both are
     tried in turn (see AGREEMENTS). In the rare case it's under neither,
     this is treated as an exceptional case: skip it, leave the sheet
     untouched, and let a human look at it.
  2. Ctrl+Shift+S -> "Связанные документы". This MUST show exactly two rows
     (the order itself, and one "Реализация товаров и услуг" under it), and
     BOTH must show 1C's posted/green-checkmark icon - if either condition
     fails (missing row, an unexpected extra row, or a document that isn't
     actually posted yet), stop on this order without changing anything.
  3. Open the order (Заказ клиента), Товары tab. Which price applies to
     which row (see determine_item_price_pairs) depends on the shape found:
       - Exactly one item row with Количество = 1: matched directly against
         the sheet row's "1С Номенклатура"/"СуммаИМАН".
       - More than one item row (or a single row with Количество > 1), and
         the sheet has exactly one row per physical unit sharing the same
         "Номер Заказ 1": every such row is looked up (see
         sheets_client.find_all_rows_for_order) and each 1C item claims as
         many sheet rows as its own Количество calls for, by fuzzy product
         name (see _match_items_to_sheet_rows_by_quantity) - covers both the
         plain "N items, N sheet rows, quantity 1 each" case and 1C
         consolidating several identically-priced physical units of the
         SAME product into one row with Количество > 1 (confirmed live:
         MA-26/00101 - 2 Samsung TVs at the same price shown as one row
         with Количество=2, matched against 2 separate sheet rows). All
         sheet rows one item claims must agree on СуммаИМАН - 1C only has
         one Цена (a per-unit price) for the whole line.
       - Exactly two item rows (each Количество = 1) but only ONE sheet row
         for the order (a "bonus item" bundle - e.g. a phone with free/near-
         token-priced earbuds bundled in): see resolve_bonus_item_prices.
         The smaller-priced item is treated as the bonus and its price is
         left exactly as-is; the other item's price is recomputed as the
         sheet's total minus the bonus item's price, so the two 1C rows
         still sum to what the sheet says. If they already sum to the
         sheet's total, neither is touched.
     Product names are allowed to be phrased differently between the two
     systems (e.g. "Galaxy A17 — 8/256/Gray" vs "Galaxy A17 8/256GB Gray
     (Серый)"), so matching uses fuzzy token-coverage and only asks a human
     when that's ambiguous, never on a clean match or a clear mismatch
     either way (see names_match_or_confirm) - except the bonus item in a
     bundle, whose name is never checked against the sheet (the sheet only
     ever names the main product). Once confirmed, each item's Вид цены is
     set to "<произвольная>" and Цена to its resolved price - unless it's
     already exactly that, in which case that item is left untouched - then
     Ctrl+Enter to post & close.
  4. Same price fix(es) on the Реализация document.
  5. Back on the order, "Создать на основании" -> "Эквайринговая операция",
     set its date to the Реализация's date (found in step 2) and Терминал to
     "Iman Pay" (EKVAYRING_TERMINAL - a required field, distinct from the
     order's own Соглашение; confirmed live posting fails without it), then
     Ctrl+Enter.
  6. Write the confirmed amount back into the sheet's "Сумма 1С" column, so
     the next run naturally picks the next unreconciled row.

IMPORTANT - read before running unattended:
  - The green-checkmark "posted" detection (lib1c.row_has_green_checkmark)
    is a pixel-color heuristic, not a real API - the Related Documents tree
    exposes no UIA state for this at all. It was verified against one real
    posted pair; if it ever looks wrong, capture and inspect the grid image
    the way scratch diagnostics did while building this (ImageGrab on the
    grid's rectangle()) before trusting it further.
  - Multi-item orders are handled (see determine_item_price_pairs) but only
    for the shapes seen so far: total Количество across all item rows
    equal to the sheet's row count for that order (each item claiming as
    many sheet rows as its own Количество - see
    _match_items_to_sheet_rows_by_quantity, this also covers the plain
    "every item quantity 1, one sheet row per item" case), or exactly two
    quantity-1 items against one shared "bonus item bundle" sheet row (see
    resolve_bonus_item_prices). Any other count mismatch, or sheet rows
    that disagree on price for the same item, stops the script rather than
    guessing. Any dialog this hasn't seen before also still stops the
    script - don't add handling for new cases speculatively; ask first.
  - This clicks real buttons in the real 1C client with synthetic OS input.
    Every post/save/write step pauses for a y/n confirm() unless run with
    --yes - do not remove those without adding equivalent safety.
  - --manual restores the original single-order, manually-typed-in-values
    flow (order number / соглашение / price / terminal all typed in, one
    price applied to every item row) - for edge cases the sheet-driven flow
    above still refuses to handle (e.g. variable quantities, or a sheet/1C
    item-count mismatch).
"""

import argparse
import datetime
import re
import sys
import time

import alif_client
import alif_sheets_client
import launch_1c
import report_sheet
import telegram_logger
import uzum_client
from playwright.sync_api import sync_playwright
from lib1c import (
    capture_element_image,
    ensure_checkbox_checked,
    extract_date,
    focus,
    force_foreground,
    get_stable_dialog,
    get_window,
    nearest_control_right_of_checkbox,
    nearest_control_right_of_text,
    paste_text,
    pick_suggestion,
    product_name_coverage,
    read_grid_via_clipboard,
    read_item_rows,
    row_has_green_checkmark,
    send_hotkey,
    set_clipboard_text,
    strip_ui_markup,
    type_digits,
)
from pywinauto import keyboard, mouse
from sheets_client import (
    DEFAULT_WORKSHEET,
    FIRST_DATA_ROW,
    count_unchecked_rows,
    find_all_rows_for_order,
    find_next_unchecked_row,
    get_worksheet,
    mark_sum_1c,
    parse_sheet_date,
)

# Both of these are specific to the "IMANUM STORE" sheet this script
# reconciles (see sheets_client.DEFAULT_WORKSHEET) - a different sheet (e.g.
# a future "IMANUM SAMSUNG" flow) may need different values for both and
# should NOT just inherit these as generic defaults.
AGREEMENTS = ("IMANUM marketplace", "IMAN PAY")

# Терминал on the Эквайринговая операция form - a separate field from the
# order's own Соглашение (which can be either "IMANUM marketplace" or "IMAN
# PAY", see AGREEMENTS above). Confirmed live: for this reconciliation flow
# it's always "Iman Pay" regardless of which Соглашение the order matched,
# and it's a required field - posting without it fails with 'Поле
# "Эквайринговый терминал" не заполнено'.
EKVAYRING_TERMINAL = "Iman Pay"

# Named bundles of (worksheet, agreements, terminal), one per sheet/workflow
# this script can reconcile - --profile is required (see main()) rather than
# defaulting to one of these, specifically so adding more here later (e.g. a
# future "IMANUM SAMSUNG" flow with its own agreements/terminal) can never
# result in a run silently using the wrong sheet's values. Add a new entry
# here when connecting a different sheet instead of overriding the
# individual --worksheet/--agreements/--terminal flags by hand each run.

# Соглашение/Терминал for the Alif Shop flow (see reconcile_alifshop_application)
# - confirmed live the same way AGREEMENTS/EKVAYRING_TERMINAL above were:
# searched the Соглашение combo's own suggestions for "alif" and confirmed
# a real known order (ON-26/00102) is actually found under it; opened a
# throwaway Эквайринговая операция form and read the Терминал combo's
# suggestions for "alif" the same way.
ALIFSHOP_AGREEMENTS = ("Alifshop Marketplace",)
ALIFSHOP_TERMINAL = "Alif Shop"

# Соглашение/Терминал for the Alif Moliya flow (see reconcile_alifmoliya_order)
# - Соглашение confirmed live the same way as the others above (found in
# the Соглашение combo's own "alif" suggestions, confirmed a real known
# order - ON-26/00102 - was findable under it... actually confirmed under
# a DIFFERENT known order for THIS profile, see reconcile_alifmoliya_order's
# docstring). Терминал "Alif Moliya" is what the user specified directly;
# unlike ALIFSHOP_TERMINAL this hasn't independently been confirmed against
# the Терминал combo's own suggestions yet - create_ekvayring_operation
# already re-reads and verifies whatever gets set there regardless, so a
# wrong value here fails loudly (not silently) the first time this runs.
ALIFMOLIYA_AGREEMENTS = ("Alif Moliya",)
ALIFMOLIYA_TERMINAL = "Alif Moliya"

# Соглашение/Терминал for the Uzum Nasiya flow (see
# reconcile_uzumnasiya_order) - confirmed live the Соглашение combo shows
# "Uzum Nasiya" exactly (seen directly in the Заказы клиентов filter with
# real open orders under it). Терминал "Uzum Nasiya" is what the user
# specified directly, same "create_ekvayring_operation re-verifies it
# regardless, so a wrong value fails loudly not silently" reasoning as
# ALIFMOLIYA_TERMINAL above.
UZUM_AGREEMENTS = ("Uzum Nasiya",)
UZUM_TERMINAL = "Uzum Nasiya"

PROFILES = {
    "imanum": {
        "worksheet": DEFAULT_WORKSHEET,
        "agreements": AGREEMENTS,
        "terminal": EKVAYRING_TERMINAL,
    },
    # No "worksheet" key - this flow reads/writes the Alif Nasiya and
    # commissions Google Sheets via alif_sheets_client.py instead of
    # sheets_client.py's IMAN-specific get_worksheet, so main()'s
    # --application handling never touches profile["worksheet"].
    "alif-shop": {
        "agreements": ALIFSHOP_AGREEMENTS,
        "terminal": ALIFSHOP_TERMINAL,
    },
    # No "worksheet" key either - there's no Google Sheet at all for this
    # profile (see reconcile_alifmoliya_order's docstring) - the source of
    # truth is 1C's own "Все открытые" filter plus a downloadable Excel
    # report from business.alif.uz/nasiya/reports (see alif_client.py).
    "alif-moliya": {
        "agreements": ALIFMOLIYA_AGREEMENTS,
        "terminal": ALIFMOLIYA_TERMINAL,
    },
    # No "worksheet" key either - same reasoning as alif-moliya (see its
    # comment above): the source of truth is 1C's own "Все открытые"
    # filter plus a downloadable CSV report from
    # merchant.uzumnasiya.uz/merchant/reports (see uzum_client.py).
    "uzum-nasiya": {
        "agreements": UZUM_AGREEMENTS,
        "terminal": UZUM_TERMINAL,
    },
}


def confirm(prompt):
    # Block for a decision that can be answered from EITHER the terminal OR
    # the Telegram group (a /continue, /skip or /quit slash command) -
    # whichever comes first (see telegram_logger.prompt_and_wait). This is
    # what lets an unattended batch (which still hits confirm() for e.g. a
    # below-threshold product-name match, see names_match_or_confirm) be
    # steered from a phone instead of silently hanging on the terminal.
    cmd = telegram_logger.prompt_and_wait(
        f"⏸ Paused - needs a decision:\n{prompt}\n\nReply /continue, /skip or /quit here (or answer in the terminal).",
        f"{prompt} [Enter=continue, s=skip, q=quit] ",
    )
    if cmd == "quit":
        print("Aborted by user.")
        sys.exit(0)
    return cmd != "skip"


# Dialogs we know how to answer automatically. Matched two ways (either is
# enough - message text capture is sometimes flaky depending on timing, so
# button-set shape is a reliable fallback):
#   - a substring of the dialog's message text
#   - the exact set of button labels on it
# -> the button text to click.
KNOWN_DIALOGS = [
    # NOTE: 1C's "OK" button is literally Latin "OK" (U+004F U+004B), not
    # Cyrillic "ОК" - confirmed by reading the actual button text via UIA.
    # Только "Отмена"/"Закрыть" are real Cyrillic. Do not "fix" this typo.
    {"message_contains": "скидки", "click": "OK"},  # "Автоматические скидки (наценки) не рассчитаны. Рассчитать...?"
    {"buttons_are": {"OK", "Отмена"}, "click": "OK"},
    {"buttons_are": {"OK", "Отмена", "Закрыть"}, "click": "OK"},
]


def guard_no_surprise_dialog(win, poll_seconds=4.0):
    """Call after any click that might have popped an unexpected dialog
    (e.g. the app's own 'Завершить работу с программой?' quit prompt, or the
    'Рассчитать автоматические скидки?' prompt on posting). Polls for a few
    seconds - some dialogs (posting validation) take a moment to appear over
    the server round-trip. Known dialogs (see KNOWN_DIALOGS) are answered
    automatically; anything else is shown and handed to the human instead of
    silently clicking through."""
    end = time.time() + poll_seconds
    while time.time() < end:
        hwnd, dlg = get_stable_dialog(win)
        if dlg is not None:
            try:
                texts = [t.window_text().strip() for t in dlg.descendants(control_type="Text")]
                texts = [t for t in texts if t]
                buttons = [b for b in dlg.descendants(control_type="Button") if b.window_text().strip()]
            except Exception:
                texts, buttons = [], []

            message = " ".join(texts).lower()
            button_labels = {b.window_text().strip() for b in buttons}

            for rule in KNOWN_DIALOGS:
                is_match = (
                    ("message_contains" in rule and rule["message_contains"] in message)
                    or ("buttons_are" in rule and button_labels == rule["buttons_are"])
                )
                if not is_match:
                    continue
                target = next((b for b in buttons if b.window_text().strip() == rule["click"]), None)
                if target:
                    label = texts[0] if texts else f"buttons={sorted(button_labels)}"
                    print(f"-> known dialog ({label!r}) - clicking {rule['click']!r}")
                    target.click_input()
                    time.sleep(0.6)
                    focus(win)
                    return

            print("!! A dialog appeared that this script did not expect:")
            for t in texts:
                print(f"   {t}")
            for b in buttons:
                print(f"   [button] {b.window_text().strip()}")
            print("!! Go look at the 1C window NOW and handle it by hand.")
            detail = texts[0] if texts else f"buttons={sorted(button_labels)}"
            cmd = telegram_logger.prompt_and_wait(
                f"⚠ An unexpected 1C dialog appeared and needs handling by hand:\n{detail}\n\n"
                "Deal with it in 1C, then reply /continue here (or /quit to stop the run).",
                "Press Enter once you've dealt with it (or the window is back to normal)... ",
            )
            if cmd == "quit":
                print("Aborted by user.")
                sys.exit(0)
            focus(win)
            return
        time.sleep(0.3)


def click_button(win, text, exact=True):
    for b in win.descendants(control_type="Button"):
        t = b.window_text().strip()
        if (t == text) if exact else (text in t):
            b.click_input()
            time.sleep(0.4)
            return True
    return False


def click_tabitem(win, text, exact=False):
    for t in win.descendants(control_type="TabItem"):
        tt = t.window_text().strip()
        if (tt == text) if exact else tt.startswith(text):
            t.click_input()
            time.sleep(0.5)
            return True
    return False


def click_menuitem(win, text):
    for m in win.descendants(control_type="MenuItem"):
        if m.window_text().strip() == text:
            m.click_input()
            time.sleep(0.5)
            return True
    return False


# ---------------------------------------------------------------- navigation

def open_zakazy_klientov(win):
    focus(win)
    print("-> Продажи")
    click_tabitem(win, "Продажи")
    time.sleep(0.5)
    focus(win)
    print("-> Заказы клиентов")
    if not click_menuitem(win, "Заказы клиентов"):
        raise RuntimeError('"Заказы клиентов" menu item not found - is Продажи section open?')
    time.sleep(0.8)
    focus(win)


def set_soglashenie_filter(win, soglashenie):
    """Just the Соглашение filter, on its own - used for every agreement
    retry after the first (see find_order_under_agreements). Split out from
    set_filters so retrying with a different Соглашение doesn't also
    re-touch Проведен/Текущее состояние every time, which don't change
    between agreement attempts within the same order's search."""
    focus(win)
    print(f"-> Соглашение = {soglashenie}")
    combo, cb = nearest_control_right_of_checkbox(win, "Соглашение:", "ComboBox")
    if not combo:
        raise RuntimeError('Could not locate "Соглашение" filter control')
    ensure_checkbox_checked(cb)
    paste_text(combo, soglashenie)
    if not pick_suggestion(win, soglashenie):
        print(f'   (no exact suggestion match for "{soglashenie}" - leaving typed text, check it manually)')
    keyboard.send_keys("{TAB}")
    time.sleep(0.8)
    focus(win)


def clear_tekushee_sostoyanie(win):
    """Empty the Текущее состояние filter, so the list shows orders in EVERY
    state instead of only open ones.

    This is what searching for a specific order wants. "Все открытые" hides
    an order the moment it is fully reconciled, so a later run looking that
    order up finds nothing and reports not_found_in_1c - which reads as "this
    order doesn't exist in 1C" when the truth is the exact opposite: it is
    finished. Clearing the filter is what lets the "already has an
    Эквайринговая операция" check (see open_and_read_related_documents)
    actually reach those orders and skip them for the right reason.

    "Все открытые" is kept ONLY where that filter is itself the source of
    truth for which orders to process at all - the alif-moliya / uzum-nasiya
    batch listings, which pass it explicitly.

    An empty field shows 1C's own gray "Все" placeholder. get_value() reads
    '' for it, which is how this verifies the clear landed - window_text() on
    these filter Edits returns their LABEL, not their value, so it can't be
    used for the check.

    Clearing has to go through the little ✕ 1C draws at the right-hand end of
    the field. Emptying the TEXT does not work on this field at all -
    measured live, all three of ctrl+A/Delete/Tab, End/shift+Home/Delete/Enter
    and a run of Backspaces left it still reading "Все открытые": it's a
    choice field, and 1C just restores the previous value when the text it's
    left with isn't one. The ✕ is drawn by 1C inside the Edit rather than
    exposed as its own element (the field reports no children at all), so it
    can only be hit positionally - hence the offsets, each re-checked, and
    each kept well clear of the ▾ dropdown button that sits about 28px in
    (opening that instead would leave a popup hanging over the form)."""
    edits = [e for e in win.descendants(control_type="Edit")
             if e.window_text().strip() == "Текущее состояние:"]
    if not edits:
        print("   (Текущее состояние filter not found - leaving it alone)")
        return
    field = edits[0]
    if not field.get_value().strip():
        print("-> Текущее состояние already empty (Все) - leaving it")
        return
    print("-> clearing Текущее состояние (so orders in ANY state can be found)")
    for offset in (14, 11, 17):
        r = field.rectangle()
        mouse.click(button="left", coords=(r.right - offset, (r.top + r.bottom) // 2))
        time.sleep(0.9)
        focus(win)
        if not field.get_value().strip():
            return
    left = field.get_value().strip()
    print(f"   (Текущее состояние still reads {left!r} - clear it by hand; an order that's "
          "already fully reconciled won't be findable while it's set)")


def set_filters(win, soglashenie, tekushee_sostoyanie=None):
    """Full filter setup: Соглашение + Проведен + Текущее состояние. Only
    needed once per order search (see find_order_under_agreements) - use
    set_soglashenie_filter alone for subsequent agreement retries.

    tekushee_sostoyanie=None (the default) CLEARS the Текущее состояние
    filter rather than setting one - see clear_tekushee_sostoyanie for why
    that, and not "Все открытые", is the right default for looking an order
    up. Pass a value explicitly to filter the list down (the alif-moliya /
    uzum-nasiya batch listings pass "Все открытые", which is their whole
    definition of what to process)."""
    set_soglashenie_filter(win, soglashenie)

    print("-> Проведен = Да")
    combo2, cb2 = nearest_control_right_of_checkbox(win, "Проведен:", "ComboBox")
    if combo2:
        ensure_checkbox_checked(cb2)
        current = combo2.window_text().strip()
        if current != "Да":
            paste_text(combo2, "Да")
            pick_suggestion(win, "Да")
            keyboard.send_keys("{TAB}")
            time.sleep(0.6)
            focus(win)

    if not tekushee_sostoyanie:
        clear_tekushee_sostoyanie(win)
        return

    print(f"-> Текущее состояние = {tekushee_sostoyanie}")
    edits = [e for e in win.descendants(control_type="Edit") if e.window_text().strip() == "Текущее состояние:"]
    if edits:
        paste_text(edits[0], tekushee_sostoyanie)
        pick_suggestion(win, tekushee_sostoyanie)
        keyboard.send_keys("{TAB}")
        time.sleep(0.8)
        focus(win)


# The magnifier button sitting immediately right of the quick-search box in
# a list form's command bar. Used purely as a positional anchor for that box
# (see get_quick_search_box) - the box itself has no accessible name.
QUICK_SEARCH_BUTTON = "Управление поиском"


def get_quick_search_box(win):
    """Return the list form's quick-search field - the "Поиск (Ctrl+F)" box in
    the Заказы клиентов command bar - or None if that form isn't in front.

    This all used to go through the Ctrl+F HOTKEY instead. Confirmed live that
    the hotkey simply does not reach this 1C build's list form via synthesized
    input: after clicking a grid cell and sending ^f, the box was still empty
    and unfocused, so the pasted order number went into the GRID rather than
    the search, and the {ENTER} meant to APPLY the search instead OPENED THE
    SELECTED ROW as a document tab. That swapped the list form for a document
    form - which is why the grid then vanished from the UIA tree on the very
    next look ("Order grid (Table) not found"), on every order rather than
    intermittently, and why crashed runs kept leaving stray "Заказ клиента
    ..." tabs behind in 1C. Clicking the box directly involves no hotkey at
    all and is deterministic.

    The box has no accessible name of its own ("Поиск (Ctrl+F)" is only
    placeholder text), so it is anchored to the one neighbour that does have a
    stable one: the "Управление поиском" magnifier button immediately to its
    right, on the same row. Rightmost qualifying Edit wins - the filter fields
    further left live on their own rows and can't match the same-row test."""
    buttons = [b for b in win.descendants(control_type="Button")
               if b.window_text().strip() == QUICK_SEARCH_BUTTON]
    if not buttons:
        return None
    br = buttons[0].rectangle()
    mid = (br.top + br.bottom) // 2
    best = None
    for e in win.descendants(control_type="Edit"):
        r = e.rectangle()
        if r.right <= br.left + 5 and r.top <= mid <= r.bottom and sane_rect(r):
            if best is None or r.right > best.rectangle().right:
                best = e
    return best


def search_order(win, order_number):
    """Filter the Заказы клиентов list down to one order using the command
    bar's quick-search box.

    The {ENTER} that applies the search is only ever sent while that box
    genuinely holds keyboard focus. That guard is the whole point: an Enter
    landing on the GRID instead opens whichever row is selected as a document,
    which is exactly how this used to derail (see get_quick_search_box).
    has_keyboard_focus() is reliable for this - confirmed live: False before
    the click, True after. The box's own text can NOT be used for the check:
    1C reports it as '' through UIA no matter what has been typed into it."""
    focus(win)
    print(f"-> quick-search: {order_number}")
    box = get_quick_search_box(win)
    if box is None:
        raise RuntimeError('Quick-search box not found - the "Заказы клиентов" list is not the form '
                           "in front (a document tab left open over it does exactly this)")
    paste_text(box, order_number)
    if not box.has_keyboard_focus():
        raise RuntimeError("quick-search box lost focus before the search could be applied - not "
                           "pressing Enter, since it would open a document instead of searching")
    keyboard.send_keys("{ENTER}")
    time.sleep(1.0)
    focus(win)

    # Hand keyboard focus back to the GRID now that the search has been
    # applied. Everything downstream assumes the list - not the search box -
    # is what holds focus: open_and_read_related_documents' Ctrl+Shift+S in
    # particular does nothing at all when the search box still has it. The
    # old Ctrl+F flow got this for free (it had to click into the grid before
    # it could send the hotkey); doing the search through the box directly
    # does not, so click a real cell explicitly. After a search that matched,
    # the first row IS the order being looked for. A degenerate-rect child is
    # never clicked (see sane_rect) - that lands on screen (0,0), i.e. some
    # unrelated control, rather than on the grid.
    tables = win.descendants(control_type="Table")
    if tables:
        cell = next((k for k in tables[0].children() if sane_rect(k.rectangle())), None)
        if cell is not None:
            cell.click_input()
            time.sleep(0.3)
            focus(win)


def clear_search(win):
    """Clear any leftover text in the quick-search box, so the grid shows ALL
    rows matching the current filters again. Confirmed live this is needed for
    batch mode (read_open_orders_with_dates): a search term left in that box
    from a previous manual search or a previous single-order run (e.g. an
    unrelated 'ON-26/00103') filters the grid down to zero rows, which the
    batch then wrongly reads as "no open orders to process". Same box and the
    same focus guard as search_order - Ctrl+A + Delete instead of pasting,
    then Enter to re-apply the now-empty search."""
    focus(win)
    box = get_quick_search_box(win)
    if box is None:
        return   # not a list form in front - nothing to clear
    print("-> clearing quick-search box (in case a stale search is filtering the grid)")
    box.click_input()
    time.sleep(0.2)
    if not box.has_keyboard_focus():
        print("   (couldn't focus the quick-search box - leaving it as it is)")
        return
    send_hotkey("ctrl", "a")
    time.sleep(0.1)
    keyboard.send_keys("{DELETE}")
    time.sleep(0.1)
    keyboard.send_keys("{ENTER}")
    time.sleep(1.0)
    focus(win)


def order_found_in_grid(win, order_number, timeout=2.5, poll=0.25):
    """After search_order(), check whether the order actually got selected.
    Polls for a bit rather than checking once - confirmed live that right
    after switching Соглашение (which requeries the list from the server)
    the grid can still be mid-reload when the search fires, so a same-instant
    check can read stale/empty content even though the order genuinely is
    there. Ctrl+F's search also silently finds nothing (no error) if the
    order truly isn't under the currently-filtered Соглашение, which is the
    real "not found" case this is ultimately trying to distinguish."""
    end = time.time() + timeout
    while time.time() < end:
        focus(win)
        if any(
            strip_ui_markup(c.window_text()).strip().startswith(order_number)
            for c in win.descendants(control_type="Custom")
        ):
            return True
        time.sleep(poll)
    return False


def find_order_under_agreements(win, order_number, agreements=AGREEMENTS, attempts_per_agreement=2):
    """Try each Соглашение in turn until the order shows up in the (now
    re-filtered) Заказы клиентов grid. Returns the matching Соглашение name,
    or None if it wasn't found under any of them - which per the reconcile
    workflow is the rare exceptional case, not a bug to retry harder on.
    Re-runs search_order a couple of times per agreement first though (see
    order_found_in_grid's docstring on why the first attempt can race).

    Only the FIRST agreement attempt does the full set_filters (Проведен +
    Текущее состояние + Соглашение) - retrying with the next agreement
    after a miss only touches Соглашение, since the other two don't change
    within one order's search and re-touching them every retry was pure
    waste (per-field UI interaction, not free)."""
    for i, soglashenie in enumerate(agreements):
        if i == 0:
            set_filters(win, soglashenie)
        else:
            set_soglashenie_filter(win, soglashenie)
        found = False
        for attempt in range(1, attempts_per_agreement + 1):
            search_order(win, order_number)
            if order_found_in_grid(win, order_number):
                found = True
                break
            print(f"   ({order_number} not visible yet under {soglashenie!r}, retry {attempt}/{attempts_per_agreement})")
            # Put every row back before searching again - retrying against a
            # list the previous miss already filtered down to zero rows would
            # find nothing no matter what.
            clear_search(win)
            time.sleep(0.8)
        if found:
            print(f"   {order_number} found under Соглашение = {soglashenie!r}")
            return soglashenie
        print(f"   {order_number} not found under Соглашение = {soglashenie!r}")
    return None


def find_order_smart(win, order_number, agreements, session_state):
    """Like find_order_under_agreements, but avoids redundant filter resets
    ACROSS an entire batch run's worth of orders (see
    process_all_unchecked_orders), not just within one order's own agreement
    retries: Проведен + Текущее состояние are only set once for the WHOLE
    run (the very first order) - every order after that just searches under
    whatever Соглашение happens to be left set from the previous order's
    search, only switching agreements (Соглашение only, never touching the
    other two) if it's not found there, and only falling back to a full
    set_filters reset (in case Проведен/Текущее состояние got changed by
    something unexpected) if it's STILL not found under any agreement even
    after trying all of them without a reset.

    session_state is a plain dict the CALLER creates once and passes the
    SAME instance into every call this run - {"initialized": bool} - this
    function mutates it, that's the whole point (a fresh dict each call
    would defeat the purpose and always take the "first order" path)."""
    if not session_state.get("initialized"):
        agreement = find_order_under_agreements(win, order_number, agreements)
        session_state["initialized"] = True
        return agreement

    search_order(win, order_number)
    if order_found_in_grid(win, order_number):
        return "current"  # caller only checks truthiness, doesn't need to know which one

    print(f"   {order_number} not found under the current Соглашение - trying each agreement in turn "
          "(Проведен/Текущее состояние left as they are)")
    for soglashenie in agreements:
        set_soglashenie_filter(win, soglashenie)
        found = False
        for attempt in range(1, 3):
            search_order(win, order_number)
            if order_found_in_grid(win, order_number):
                found = True
                break
            time.sleep(0.8)
        if found:
            print(f"   {order_number} found under Соглашение = {soglashenie!r}")
            return soglashenie
        print(f"   {order_number} not found under Соглашение = {soglashenie!r}")

    print(f"   {order_number} not found under any Соглашение even without resetting Проведен/Текущее "
          "состояние - re-checking those too in case something changed them, before giving up on this order")
    return find_order_under_agreements(win, order_number, agreements)


def sane_rect(r, min_w=50, min_h=15):
    """Guard against acting on a degenerate (0,0,0,0) or otherwise bogus
    rectangle - this is exactly what caused a stray double-click to land on
    the window's system-menu icon instead of a grid row."""
    return (r.right - r.left) >= min_w and (r.bottom - r.top) >= min_h


def _read_posted_marks(win, grid, realizatsiya_index, attempts=3, pause=0.5):
    """(order_posted, realizatsiya_posted), read off the Связанные документы
    icons and re-read before a negative is believed.

    This read is a SCREENSHOT of that grid (row_has_green_checkmark counts
    green pixels), so it reports "not posted" for any reason those pixels
    aren't there to be read - the panel still painting, or another window
    sitting over 1C at that instant, since ImageGrab captures whatever is on
    screen at those coordinates rather than the window's own buffer.

    Confirmed live on QQ-26/00101: both rows came back False mid-batch, and
    both read True on the very next look at the same unchanged order - whose
    Заказ клиента is posted BY DEFINITION anyway, the list it was found in
    being filtered Проведен=Да. So re-focus and look again before reporting a
    negative, and see open_and_read_related_documents for why a negative that
    survives all the attempts still doesn't stop the order."""
    order_posted = realizatsiya_posted = False
    for attempt in range(1, attempts + 1):
        focus(win)
        time.sleep(pause)
        image = capture_element_image(grid)
        order_posted = row_has_green_checkmark(image, row_index=0)
        realizatsiya_posted = row_has_green_checkmark(image, row_index=realizatsiya_index)
        if order_posted and realizatsiya_posted:
            return True, True
        if attempt < attempts:
            print(f"   (posted-checkmark read came back order={order_posted}, "
                  f"Реализация={realizatsiya_posted} - re-reading [{attempt}/{attempts}])")
    return order_posted, realizatsiya_posted


def open_and_read_related_documents(win, order_number, attempts=3):
    """Select the order row, Ctrl+Shift+S, and read the resulting
    'Связанные документы' grid via clipboard. Requires the order to have
    exactly one Реализация under it, plus - optionally - the Эквайринговая
    операция a previous run already created. A missing Реализация, or
    anything else unexpected under there, is a stop condition for this order
    (per the reconcile workflow), not a retryable error, so this returns None
    rather than raising. Only genuine technical failures (the grid never
    loading) raise.

    The posted (green) checkmark is REPORTED but no longer stops anything -
    see _read_posted_marks for why that read can't be trusted to refuse an
    order on.

    Returns {"date": <Реализация date>, "order_date": <the order's own date>,
    "grid": <DataGrid element>, "ekvayring": <text of the existing
    Эквайринговая операция row, or "">}, or None if this order should be
    skipped.

    "ekvayring" is what keeps an already-reconciled order from being done a
    second time. It used to be impossible to even see: the check demanded
    EXACTLY two rows, so the third row every successful run leaves behind
    made each later run fail that order with "Expected exactly 2 related
    documents" - reported as a check failure when the truth was the exact
    opposite, the order was already finished. Callers skip on it (and the
    Alif Shop flow finishes writing its sheet row - see
    reconcile_alifshop_application) instead of fixing prices again and
    posting a second Эквайринговая операция."""
    for attempt in range(1, attempts + 1):
        focus(win)
        print(f"-> selecting row, Ctrl+Shift+S (Связанные документы) [attempt {attempt}]")
        cells = [c for c in win.descendants() if strip_ui_markup(c.window_text()).strip().startswith(order_number)]
        if cells:
            cells[0].click_input()
            time.sleep(0.3)
        send_hotkey("ctrl", "shift", "s")
        time.sleep(1.4)
        focus(win)

        grids = win.descendants(control_type="DataGrid")
        grids = [g for g in grids if sane_rect(g.rectangle())]
        if not grids:
            print("   (no valid Связанные документы grid yet, retrying)")
            time.sleep(0.6)
            continue

        grid = grids[0]
        rows = read_grid_via_clipboard(win, grid)
        if not any(order_number in text for _, text in rows):
            print(f"   (grid content didn't look right yet: {rows!r}, retrying)")
            time.sleep(0.6)
            continue

        print("   Related documents rows read:")
        for indent, text in rows:
            print(f"     {'  ' * indent}{text}")

        # Classify the rows by what each one IS, rather than by how many
        # there are: a finished order legitimately carries a third row (its
        # own Эквайринговая операция), and a count-based check reads that
        # perfectly normal state as a failure.
        order_row, child_rows = rows[0], rows[1:]
        realizatsiya_rows = [r for r in child_rows if "Реализац" in r[1]]
        ekvayring_rows = [r for r in child_rows if "Эквайринг" in r[1]]
        unexpected = [r for r in child_rows
                      if r not in realizatsiya_rows and r not in ekvayring_rows]
        if unexpected:
            print(f"!! Unexpected related document(s) under the order: {[t for _, t in unexpected]} - "
                  "stopping on this order.")
            return None
        if len(realizatsiya_rows) != 1:
            print(f"!! Expected exactly one 'Реализация товаров и услуг' under the order, found "
                  f"{len(realizatsiya_rows)} - stopping on this order.")
            return None
        if len(ekvayring_rows) > 1:
            print(f"!! Found {len(ekvayring_rows)} Эквайринговые операции under the order - stopping "
                  "on this order, a human needs to look at that.")
            return None

        realizatsiya_row = realizatsiya_rows[0]
        date = extract_date(realizatsiya_row[1])
        # The report wants the ORDER's own date (the "Заказ клиента ... от
        # DD.MM.YYYY ..." line, row 0), which is distinct from the
        # Реализация date used for the Эквайринговая операция.
        order_date = extract_date(order_row[1])

        if ekvayring_rows:
            # Nothing left to do in 1C - say so and let the caller decide.
            # Deliberately no green-checkmark check here: that gate exists to
            # protect documents about to be CHANGED, and nothing is going to
            # be changed on this order.
            print(f"   already reconciled - this order already has: {ekvayring_rows[0][1]}")
            return {"date": date, "order_date": order_date, "grid": grid,
                    "ekvayring": ekvayring_rows[0][1]}

        order_posted, realizatsiya_posted = _read_posted_marks(win, grid, rows.index(realizatsiya_row))
        print(f"   posted (green checkmark) - order: {order_posted}, Реализация: {realizatsiya_posted}")
        if not (order_posted and realizatsiya_posted):
            # Not a stop condition. This mark is only a visual read of an
            # icon and it produces false negatives (see _read_posted_marks),
            # so refusing an order on it meant refusing to create the
            # Эквайринговая операция for a perfectly good one. What actually
            # guards this order is elsewhere and is not pixel-based: the list
            # it was found in is filtered Проведен=Да, the Реализация exists
            # and is checked above, every price fix verifies the value it
            # wrote, and each post_and_close verifies the document closed.
            print("   (couldn't confirm the green checkmark on both documents - continuing anyway, see "
                  "_read_posted_marks; the price and posting steps below verify themselves)")

        print(f"   Реализация date extracted: {date}; order date: {order_date}")
        return {"date": date, "order_date": order_date, "grid": grid, "ekvayring": ""}

    raise RuntimeError(
        f'Could not get a valid "Связанные документы" view for {order_number} after {attempts} attempts. '
        "Open it by hand and re-run from here if needed."
    )


def _open_document_title(win):
    """Text of the currently active document tab, or '' if none is open -
    both Заказ клиента и Реализация forms have a highlighted 'Провести и
    закрыть' button (the plain 'Связанные документы' list view never does).
    NOTE: the top document-tab-strip items (this document, 'Связанные
    документы', 'Заказы клиентов', ...) are exposed via UIA as plain 'Pane'
    controls, NOT 'TabItem' - 'TabItem' is only the *sub*-tabs inside a
    document (Основное/Товары/Дополнительно). Mixing these up silently finds
    nothing."""
    if not any(b.window_text().strip() == "Провести и закрыть" for b in win.descendants(control_type="Button")):
        return ""
    for pane in win.descendants(control_type="Pane"):
        t = pane.window_text().strip()
        if t.startswith(("Заказ клиента", "Реализация")):
            return t
    return "?"


def open_related_row(win, grid, row_index, expect_title_startswith, attempts=4):
    """row_index: 0 = order (bold/top) line, 1 = realizatsiya line.

    The double click's target coordinates are computed from the grid's live
    rectangle plus a measured row-height offset - this is inherently a guess
    (the related-documents tree exposes no per-row UIA elements to click by
    reference, see lib1c's module docstring). It has also intermittently
    missed entirely when the 1C window didn't *actually* have OS foreground
    at the moment of the click (a background/terminal-launched script can
    lose that race even after SetForegroundWindow reports success), or
    landed on the wrong row. So: verify the *expected* document form
    actually opened afterward, and retry the click if not, instead of
    silently continuing to edit whatever happens to be open."""
    for attempt in range(1, attempts + 1):
        focus(win)
        r = grid.rectangle()
        if not sane_rect(r):
            raise RuntimeError(f"Related-documents grid rect looks bogus ({r}) - refusing to click blindly")
        y = r.top + 19 + row_index * 31
        x = r.left + 250
        if not force_foreground(win):
            raise RuntimeError("could not foreground 1C window")
        mouse.move(coords=(x, y))
        time.sleep(0.1)
        mouse.double_click(button="left", coords=(x, y))
        time.sleep(1.3)
        focus(win)

        title = _open_document_title(win)
        if title.startswith(expect_title_startswith):
            return
        print(f"   (expected a {expect_title_startswith!r} document, got {title!r} - retrying) [attempt {attempt}]")
        time.sleep(0.5)

    raise RuntimeError(
        f"Double-clicking related-documents row {row_index} did not open a "
        f"{expect_title_startswith!r} document after {attempts} attempts. "
        "Open it by hand and re-run from here if needed."
    )


# ------------------------------------------------------------- price editing

VID_TSENY_ARBITRARY = "<произвольная>"


def set_price_type_arbitrary(win, vidtseny_cell, target_text=VID_TSENY_ARBITRARY, max_steps=30, open_attempts=6):
    """Open the 'Вид цены' picklist for this cell and select the actual
    "<произвольная>" item, verifying the real popup contents after every
    single scroll step rather than trusting any fixed count or any one
    interaction to have worked.

    An earlier version blindly pressed Down 30 times assuming that always
    reached the list's last item, then Enter. Confirmed live this is WRONG
    and dangerous: the real list has dozens of entries (per-channel price
    schedules like "36 месяцев", "Huawei brand shop", "Uzum market", ...), so
    Down-30 landed on an arbitrary unrelated one ("36 месяцев") instead -
    which is a computed/locked price schedule, and the code then silently
    failed to overwrite Цена with it, posting a real order with a materially
    wrong price.

    Three more things confirmed live while fixing that:
      - The Down arrow key does NOTHING to this popup at all (tested
        directly: the visible item set is 100% unchanged after repeated
        presses) - PageDown is what actually scrolls it, one row at a time.
      - Getting the popup to even open is itself flaky: identical
        double-clicks at the identical coordinates sometimes open it and
        sometimes silently do nothing, with no single click pattern (single
        vs double, left-of-center vs near the dropdown arrow) found to be
        100% reliable in this environment. This retries opening it several
        times and raises clearly if it never opens, rather than proceeding
        against a popup that isn't actually there.
      - "<произвольная>" itself never appears as that literal text among the
        popup's ListItem elements - screenshotted live, it's visibly the
        last entry in the list, but win.descendants() reports its text as
        an empty string. Almost certainly the same markup-stripping this
        app applies to search-highlighted cell text elsewhere (see
        lib1c.strip_ui_markup) is kicking in here too: "<произвольная>" on
        its own looks exactly like one opening pseudo-XML tag with no
        closing tag, so whatever strips "<b><colorstyle -46>...</></>" for
        highlighting is likely consuming this whole string the same way.
        The item still renders as plain readable "<произвольная>" text once
        actually committed into the grid cell (confirmed live) - it's only
        the popup's own accessible name for the item that's affected. So:
        match on "the last real (sane-rect'd, not a virtualization
        placeholder) ListItem with empty text" instead of matching literal
        target_text.
    """
    r = vidtseny_cell.rectangle()
    x, y = r.left + 65, (r.top + r.bottom) // 2

    def find_target():
        candidates = [
            li for li in win.descendants(control_type="ListItem")
            if li.window_text().strip() == "" and sane_rect(li.rectangle())
        ]
        return candidates[-1] if candidates else None

    opened = False
    for attempt in range(1, open_attempts + 1):
        force_foreground(win)
        mouse.move(coords=(x, y))
        time.sleep(0.15)
        mouse.double_click(button="left", coords=(x, y))
        time.sleep(0.8)
        if win.descendants(control_type="ListItem"):
            opened = True
            break
        print(f"   (Вид цены dropdown didn't open, retry {attempt}/{open_attempts})")
        time.sleep(0.4)

    if not opened:
        raise RuntimeError("Could not open the Вид цены dropdown after several attempts - check by hand.")

    target = find_target()
    steps = 0
    while target is None and steps < max_steps:
        keyboard.send_keys("{PGDN}")
        time.sleep(0.2)
        target = find_target()
        steps += 1

    if target is None:
        # Escape here closes the WHOLE document tab in this app (not just the
        # dropdown) when nothing turned out to be open - confirmed live. Only
        # send it if we know we actually have something open to back out of.
        if win.descendants(control_type="ListItem"):
            keyboard.send_keys("{ESC}")
            time.sleep(0.3)
        raise RuntimeError(
            f'"{target_text}" not found in the Вид цены list after {max_steps} PageDown presses - '
            "aborting without changing it, check by hand."
        )

    # PageDown only scrolls the popup's VIEW - confirmed live, it does not
    # move whatever item is actually keyboard-selected, so pressing Enter at
    # this point commits whatever was selected before scrolling even
    # started (the ORIGINAL price type), not the now-visible target. Click
    # the target element itself instead - same principle as everywhere else
    # in this app that a genuine click is what actually commits a choice.
    # The click alone highlights it as the pending choice but doesn't always
    # fully close the popup/commit it into the cell - an explicit Enter
    # right after is what actually confirms and closes it.
    target.click_input()
    time.sleep(0.3)
    keyboard.send_keys("{ENTER}")
    time.sleep(0.5)


def set_grid_cell_value(win, cell, text):
    """A single click on a grid cell only *selects* it - it does not reliably
    enter edit mode (confirmed the hard way: rows left un-typed, or got
    garbled values from a half-committed previous state). A genuine OS-level
    double click does reliably enter edit mode, same as opening a row or a
    dropdown elsewhere in this grid."""
    r = cell.rectangle()
    x, y = (r.left + r.right) // 2, (r.top + r.bottom) // 2
    force_foreground(win)
    mouse.move(coords=(x, y))
    time.sleep(0.1)
    mouse.double_click(button="left", coords=(x, y))
    time.sleep(0.4)
    send_hotkey("ctrl", "a")
    time.sleep(0.15)
    keyboard.send_keys(text)
    time.sleep(0.2)
    keyboard.send_keys("{ENTER}")
    time.sleep(0.4)


def fix_realizatsiya_prices(win, price):
    """Original multi-row-capable price fixer, kept for --manual only: it
    applies the SAME price to every item row, which is wrong for a genuine
    multi-item order with different prices per item. The sheet-driven flow
    uses determine_item_price_pairs + apply_price_fix_for_item instead,
    which looks up each item's own correct price from the sheet."""
    focus(win)
    print("-> Товары tab")
    click_tabitem(win, "Товары")
    time.sleep(0.5)
    focus(win)

    tables = win.descendants(control_type="Table")
    if not tables:
        raise RuntimeError("Товары grid not found on Реализация document")
    table = tables[0]

    n_rows = len([k for k in table.children() if k.window_text().strip().endswith(" Цена")])
    print(f"   {n_rows} item row(s) found")

    for i in range(n_rows):
        kids = table.children()
        vidtseny_cells = [k for k in kids if "Вид цены" in k.window_text()]
        price_cells = [k for k in kids if k.window_text().strip().endswith(" Цена")]
        if i >= len(vidtseny_cells) or i >= len(price_cells):
            break

        current_vid = vidtseny_cells[i].window_text()
        print(f"   row {i + 1}: Вид цены currently {current_vid!r}")
        if "произвольн" not in current_vid.lower():
            set_price_type_arbitrary(win, vidtseny_cells[i])
            guard_no_surprise_dialog(win)
            # re-fetch after the dropdown interaction, layout may have shifted
            kids = table.children()
            price_cells = [k for k in kids if k.window_text().strip().endswith(" Цена")]

        print(f"   row {i + 1}: setting Цена = {price}")
        set_grid_cell_value(win, price_cells[i], str(price))
        focus(win)

        kids = table.children()
        price_cells_after = [k for k in kids if k.window_text().strip().endswith(" Цена")]
        new_value = price_cells_after[i].window_text().strip() if i < len(price_cells_after) else "?"
        print(f"   row {i + 1}: Цена now reads {new_value!r}")
        if str(price) not in new_value:
            print(f"   !! that does not contain {price!r} - this row may not have taken the new value, check it by hand")


# A run-scoped note bag for how a price got entered: when the primary method
# (typing) fails verification and the clipboard-paste fallback is used
# instead, apply_price_fix_for_item records it here so the reconcile
# functions can surface it in the report's Comment column. Reset at the
# start of each order (see _reset_price_notes) and drained into that order's
# comment (see _price_notes_suffix).
_price_fallback_notes = []


def _reset_price_notes():
    _price_fallback_notes.clear()


def _price_notes_suffix():
    return "; ".join(_price_fallback_notes)


def _combine_detail(*parts):
    """Join non-empty comment fragments with '; ', or None if all empty -
    for a _status(..., detail=...) that reads cleanly in the report."""
    return "; ".join(p for p in parts if p) or None


def apply_price_fix_for_item(win, item, price):
    """Set Вид цены to <произвольная> (if not already) and Цена to `price`
    for ONE already-read item row (see lib1c.read_item_rows) - used for
    both single- and multi-item documents alike (a single-item document is
    just the n=1 case).

    A line whose Цена ALREADY equals `price` is returned untouched, price
    type included - nothing is clicked, typed or re-verified on it. Every
    profile's order and Реализация handling goes through this one function,
    so that is the single place "don't rewrite a price that is already
    right" needed to live. Raises if the price doesn't verify afterward,
    rather than silently continuing - this is what would have posted a
    real order with a materially wrong price, live, before this check
    existed (see set_price_type_arbitrary's docstring for the full story on
    why the price-type picklist needed fixing first).

    Цена is the column immediately right of Вид цены, and Right-arrow from
    a selected cell reliably scrolls the next one into view and selects it
    (confirmed live - NOT the same as being off-screen, which is what made
    an earlier direct double-click-by-rectangle attempt on Цена fail: its
    rectangle can extend past the currently visible/scrolled window area).
    But being *selected* is not the same as being in EDIT mode - confirmed
    live, typing digits directly onto a merely-selected cell does NOTHING
    AT ALL, not even partial text. F2 is what actually enters edit mode
    here (this grid's equivalent of a double-click), after which Ctrl+A +
    typing + Enter works exactly as expected.

    UIA cell references can go stale after an edit (the row's own column
    layout can shift, e.g. Вид цены changing whether "% авт."/"Сумма авт."
    columns are shown - confirmed live), so this re-fetches ALL rows and
    relocates the same one by its stable "index" (position in the grid)
    after every interaction, rather than reusing cell references across
    edits."""
    def refetch(index):
        rows = read_item_rows(win)
        match = next((r for r in rows if r["index"] == index), None)
        if match is None:
            raise RuntimeError(f"Item row at index {index} disappeared after an edit - check by hand.")
        return match

    index = item["index"]
    print(f"   [item {index}] Номенклатура={item['nomenklatura']!r} Количество={item['kolichestvo']} "
          f"Вид цены={item['vid_tseny']!r} Цена={item['tsena']}")

    if item["tsena"] is not None and abs(item["tsena"] - float(price)) < 0.5:
        # The price is ALREADY what this was going to set it to, so there is
        # nothing to write - whatever Вид цены currently says. This used to
        # additionally require Вид цены to already read <произвольная>, which
        # meant a line that was already correct still got put through the
        # price-type picklist and re-typed with the identical number. That is
        # the flakiest, most dangerous interaction in this whole script (see
        # set_price_type_arbitrary: the popup opens only intermittently, and
        # an early version of it landed on a locked price schedule and posted
        # a real order with a materially wrong price), so running it for no
        # gain was pure risk.
        #
        # The trade-off, deliberately accepted: a line left on a price-list
        # Вид цены stays bound to that list rather than being pinned to
        # <произвольная>. Its value is correct now and posting doesn't change
        # it; only editing that price list later would, and that is a far
        # smaller exposure than re-driving the picklist on every already-good
        # line.
        if "произвольн" in item["vid_tseny"].lower():
            print(f"   [item {index}] already <произвольная> at Цена={price} - nothing to change")
        else:
            print(f"   [item {index}] Цена is already {price} (Вид цены={item['vid_tseny']!r}) - "
                  "nothing to change, leaving the price type alone too")
        return item

    if "произвольн" not in item["vid_tseny"].lower():
        set_price_type_arbitrary(win, item["vid_tseny_cell"])
        guard_no_surprise_dialog(win)
        time.sleep(0.5)  # let 1C fully settle after the price-type change before reading/clicking again
        item = refetch(index)
        if "произвольн" not in item["vid_tseny"].lower():
            raise RuntimeError(
                f"Item {index}'s Вид цены still reads {item['vid_tseny']!r} after trying to set it to "
                f"{VID_TSENY_ARBITRARY!r} - refusing to touch Цена on the wrong price type."
            )
        print(f"   [item {index}] Вид цены confirmed as {item['vid_tseny']!r} (Цена currently {item['tsena']!r})")

    print(f"   [item {index}] setting Цена = {price}")
    # Confirmed live (batch mode against a real order): this click+Right+F2+
    # type+Enter sequence can silently do nothing on the first try even
    # though it works reliably when a human is pacing the run with confirm()
    # prompts in between - fully unattended, there's less incidental settle
    # time between switching Вид цены and clicking back into the row. Retry
    # a few times, re-fetching fresh cell references each time (not just
    # re-reading the value), before giving up.
    price_text = _format_price_for_1c(price)
    updated = None
    used_paste = False
    for attempt in range(1, 4):
        # PRIMARY method: type the value (keyboard.send_keys). FALLBACK
        # method (attempts 2-3): clipboard paste - confirmed live that
        # typing a long number can silently drop digits (ON-26/00104: Цена
        # 3300429 registered as 33.0), whereas paste lands the whole value
        # atomically (lib1c's own docstring notes send_keys "does not
        # reliably type"). Which method got the value in is recorded so the
        # report Comment can flag when the paste fallback was needed. The
        # clipboard is set BEFORE entering edit mode - 1C can transiently
        # lock the clipboard once a grid cell is in edit mode.
        use_clipboard = attempt >= 2
        if use_clipboard:
            try:
                set_clipboard_text(price_text)
                time.sleep(0.15)
            except Exception as e:  # noqa: BLE001 - clipboard locked/unavailable: type instead this attempt
                print(f"   [item {index}] couldn't set clipboard ({e!r}) - typing instead this attempt")
                use_clipboard = False
        vt_rect = item["vid_tseny_cell"].rectangle()
        force_foreground(win)
        mouse.click(button="left", coords=((vt_rect.left + vt_rect.right) // 2, (vt_rect.top + vt_rect.bottom) // 2))
        time.sleep(0.3)
        keyboard.send_keys("{RIGHT}")
        time.sleep(0.3)
        keyboard.send_keys("{F2}")
        time.sleep(0.3)
        send_hotkey("ctrl", "a")
        time.sleep(0.15)
        if use_clipboard:
            send_hotkey("ctrl", "v")
        else:
            keyboard.send_keys(price_text)
        time.sleep(0.2)
        keyboard.send_keys("{ENTER}")
        time.sleep(0.5)
        focus(win)

        updated = refetch(index)
        if updated["tsena"] is not None and abs(updated["tsena"] - float(price)) < 0.5:
            used_paste = use_clipboard
            break
        method = "clipboard-paste" if use_clipboard else "typing"
        print(f"   [item {index}] Цена still reads {updated['tsena']!r} after attempt {attempt}/3 "
              f"({method}) - retrying" + (" with clipboard-paste" if attempt == 1 else ""))
        item = updated
        time.sleep(0.5)

    if used_paste:
        note = (f"item {index} ({item['nomenklatura']}) price set via clipboard-paste fallback "
                "(typing dropped digits)")
        print(f"   [item {index}] -> {note}")
        _price_fallback_notes.append(note)

    print(f"   [item {index}] Цена now reads {updated['tsena']!r}")
    if updated["tsena"] is None or abs(updated["tsena"] - float(price)) > 0.5:
        raise RuntimeError(
            f"Item {index}'s Цена reads {updated['tsena']!r} after setting it to {price!r} - it did not take "
            "after 3 attempts. Refusing to post a document with an unverified price."
        )
    return updated


class MultiItemMatchError(RuntimeError):
    """Raised by determine_item_price_pairs when a multi-item document's
    rows can't be confidently paired with sheet rows - callers should treat
    this as a stop condition for the whole order, same as
    MultipleItemRowsError used to be treated before multi-item support
    existed."""


def resolve_bonus_item_prices(items, sheet_row):
    """For a 2-item order where the sheet only has ONE row for it (a
    combined total, not itemized) - confirmed live pattern: one item is the
    real product, the other a "bonus" bundled in (e.g. free/near-free
    wireless earbuds with a phone), often priced at a token amount like 1
    so'm. The sheet only ever describes the main product by name.

    Identifies the bonus item as whichever of the two has the SMALLER
    current price. If the two items' current prices already sum to the
    sheet's total, both are left exactly as-is (target_price = current
    price, which apply_price_fix_for_item then simply skips - confirmed
    live case: 9,398,000 + 1,000 = 9,399,000, already correct). Otherwise
    the bonus item's price is kept exactly as-is and the main item's price
    is recomputed as total - bonus_price, so the 1C total still adds up to
    what the sheet says (confirmed live case: main 355,000 + bonus 1 !=
    sheet's 355,000, so main gets corrected to 354,999).

    Returns [(main_item, info), (bonus_item, info)] in the same (item,
    sheet_row) shape determine_item_price_pairs uses elsewhere - both
    share the SAME underlying sheet row (there's only one), marked with
    skip_name_check=True for the bonus item (the sheet's product_name
    describes the main item, not this one - a fuzzy match against it would
    be nonsense) - reconcile_order dedupes by row number before writing
    Сумма 1С back so the shared total isn't written twice."""
    a, b = items
    total = sheet_row["expected_amount"]
    a_price, b_price = a["tsena"] or 0, b["tsena"] or 0
    bonus, main = (a, b) if a_price <= b_price else (b, a)
    bonus_price, main_price = (a_price, b_price) if a_price <= b_price else (b_price, a_price)
    current_total = a_price + b_price

    if abs(current_total - total) < 0.5:
        print(f"   bonus-item bundle: main={main['nomenklatura']!r} ({main_price}), "
              f"bonus={bonus['nomenklatura']!r} ({bonus_price}) - already sum to sheet total {total}, "
              "leaving both as-is")
        main_target, bonus_target = main_price, bonus_price
    else:
        main_target = total - bonus_price
        bonus_target = bonus_price
        print(f"   bonus-item bundle: main={main['nomenklatura']!r} ({main_price}), "
              f"bonus={bonus['nomenklatura']!r} ({bonus_price}) - sum {current_total} != sheet total {total}, "
              f"recomputing main price as {total} - {bonus_price} = {main_target}")

    return [
        (main, {"row": sheet_row["row"], "rows": [sheet_row["row"]], "product_name": sheet_row["product_name"],
                "expected_amount": total, "target_price": main_target, "skip_name_check": False}),
        (bonus, {"row": sheet_row["row"], "rows": [sheet_row["row"]], "product_name": sheet_row["product_name"],
                 "expected_amount": total, "target_price": bonus_target, "skip_name_check": True}),
    ]


def _match_items_to_sheet_rows_by_quantity(items, sheet_rows):
    """Pair each 1C item row against however many sheet rows its own
    Количество calls for. Covers both the plain "N items, N sheet rows,
    quantity 1 each" case AND a newer one confirmed live: 1C sometimes
    consolidates several physical units of the SAME product into one
    Товары row with Количество > 1 (e.g. MA-26/00101: 2 identically-priced
    Samsung TVs sold together show up in 1C as a single row with
    Количество=2), while IMAN's sheet always lists one row per physical
    unit sold (so that order has 2 sheet rows for the TV, both
    СуммаИМАН=6,950,000, plus 1 more for the second product on the order).

    Greedy, same reasoning as elsewhere in this module: for each item (in
    grid order), claim whichever `kolichestvo` still-unclaimed sheet rows
    score highest (fuzzy, see product_name_coverage) against its own name -
    cheap, and for the handful of rows per order actually seen, equivalent
    to an optimal assignment.

    1C only has a single Цена for the whole line (a per-unit price,
    multiplied by Количество for the line total) - so every sheet row one
    item claims must agree on СуммаИМАН, or there's no single per-unit price
    that would be correct. Raises MultiItemMatchError rather than guessing
    if they disagree."""
    remaining = list(sheet_rows)
    pairs = []
    for item in items:
        remaining.sort(key=lambda r: product_name_coverage(r["product_name"], item["nomenklatura"]), reverse=True)
        # item["kolichestvo"] comes from lib1c.parse_ru_number, which always
        # returns a float (e.g. 2.0, not 2) - confirmed live, slicing with
        # that raises "slice indices must be integers..." Количество is
        # always a whole count in practice, so round-then-int is safe.
        n = int(round(item["kolichestvo"]))
        claimed, remaining = remaining[:n], remaining[n:]
        amounts = {sr["expected_amount"] for sr in claimed}
        if len(amounts) != 1:
            raise MultiItemMatchError(
                f"item {item['index']} ({item['nomenklatura']!r}, Количество={item['kolichestvo']}) matched sheet "
                f"row(s) {[sr['row'] for sr in claimed]} but they don't all have the same СуммаИМАН "
                f"({[sr['expected_amount'] for sr in claimed]}) - 1C only has one Цена for the whole line, can't "
                "set a single per-unit price from mismatched amounts. Stopping."
            )
        per_unit = amounts.pop()
        rows = [sr["row"] for sr in claimed]
        pairs.append((item, {"row": rows[0], "rows": rows, "product_name": claimed[0]["product_name"],
                              "expected_amount": per_unit, "target_price": per_unit, "skip_name_check": False}))
    return pairs


def determine_item_price_pairs(win, ws, order_number, expected_product, expected_amount, sheet_row_number=None):
    """Figure out which price applies to which 1C Товары row on the
    currently-open document. Returns a list of (item, sheet_row) pairs, one
    per 1C item row, where sheet_row is {row, rows, product_name,
    expected_amount, target_price, skip_name_check} - rows is the list of
    sheet row number(s) this item's confirmed price gets written back to
    (usually one, except the quantity>1 case below); row is just rows[0]
    for convenience/logging. expected_amount is what the sheet says (used
    for the write-back); target_price is what actually gets set on THIS
    item - usually the same as expected_amount, except in the bonus-item
    case below, where it's this item's own adjusted share of a total shared
    with another item.

    Single item, quantity 1 (the simplest, most common case): matched
    directly against expected_product/expected_amount, no extra sheet
    lookup needed.

    Everything else (more than one item row, or a single item row with
    quantity > 1) needs its own per-item sheet lookup - re-queries the
    sheet for every row sharing this order number, then:
      - if it's exactly 2 items (each quantity 1) against exactly 1 sheet
        row: the "bonus item bundle" pattern - see resolve_bonus_item_prices.
      - otherwise, if total Количество across all items equals the sheet's
        row count for this order: pair by quantity, see
        _match_items_to_sheet_rows_by_quantity (this also covers the
        original "N items, N sheet rows, quantity 1 each" case - that's just
        the special case where every item's quantity is 1).
      - anything else - stop rather than guess.

    Requires ws for anything beyond the single-item-quantity-1 case (can't
    look anything up in --order mode, which has no sheet)."""
    items = read_item_rows(win)

    if len(items) == 1 and items[0]["kolichestvo"] == 1:
        item = items[0]
        return [(item, {"row": sheet_row_number, "rows": [sheet_row_number] if sheet_row_number is not None else [],
                         "product_name": expected_product, "expected_amount": expected_amount,
                         "target_price": expected_amount, "skip_name_check": False})]

    if ws is None:
        raise MultiItemMatchError(
            f"{order_number} has {len(items)} item row(s) and/or an item quantity != 1, but no Google Sheet is "
            "available to look up each item's own price (running in --order mode) - can't proceed."
        )

    sheet_rows = find_all_rows_for_order(ws, order_number)
    total_kolichestvo = sum(item["kolichestvo"] for item in items)

    if len(sheet_rows) == 1 and len(items) == 2 and all(item["kolichestvo"] == 1 for item in items):
        return resolve_bonus_item_prices(items, sheet_rows[0])

    if total_kolichestvo == len(sheet_rows):
        return _match_items_to_sheet_rows_by_quantity(items, sheet_rows)

    raise MultiItemMatchError(
        f"1C shows {len(items)} item row(s) (total Количество {total_kolichestvo}) but the sheet has "
        f"{len(sheet_rows)} row(s) for {order_number} - don't know how to pair these, stopping."
    )


def _format_price_for_1c(price):
    price = float(price)
    if abs(price - round(price)) < 1e-6:
        return str(int(round(price)))
    return f"{price:.2f}".replace(".", ",")


def names_match_or_confirm(sheet_name, c1_name, auto_threshold=0.85):
    """Fuzzy-compare the sheet's expected product name against 1C's actual
    Номенклатура. Auto-proceeds on a clean match, auto-refuses on a clear
    mismatch is NOT done here - anything below the confident-match threshold
    is shown to a human rather than guessed, since naming conventions vary
    wildly by product category in this data (see product_name_coverage's
    docstring) and a false-positive auto-match would apply the wrong price
    to the wrong item."""
    coverage = product_name_coverage(sheet_name, c1_name)
    print(f"   comparing product names (token coverage={coverage:.2f}):")
    print(f"     Google Sheet : {sheet_name}")
    print(f"     1C           : {c1_name}")
    if coverage >= auto_threshold:
        print("   -> looks like the same product, proceeding automatically")
        return True
    return confirm("   Not a clean match - are these actually the same product?")


def post_and_close(win, verify_closed_title_contains=None, retries=3):
    """Ctrl+Enter = press the highlighted default button (Провести и
    закрыть). If a known dialog (e.g. the discount/markup recalculation
    prompt) intercepts it, one Ctrl+Enter may not be enough to actually post
    - answering that dialog can leave the document open needing a second
    Ctrl+Enter. Retry a couple of times and verify the document's tab is
    actually gone before declaring success."""
    for attempt in range(1, retries + 1):
        focus(win)
        print(f"-> Ctrl+Enter (Провести и закрыть) [attempt {attempt}]")
        send_hotkey("ctrl", "enter")
        time.sleep(1.0)
        guard_no_surprise_dialog(win)

        if not verify_closed_title_contains:
            return
        focus(win)
        # document tab titles are 'Pane' controls, not 'TabItem' - see
        # _open_document_title's note.
        still_open = any(
            verify_closed_title_contains in p.window_text()
            for p in win.descendants(control_type="Pane")
        )
        if not still_open:
            return
        print("   document tab is still open after Ctrl+Enter, trying again")
        time.sleep(0.5)

    raise RuntimeError(
        f"Document containing {verify_closed_title_contains!r} did not close after {retries} Ctrl+Enter attempts - "
        "check it by hand."
    )


# --------------------------------------------------------------- ekvayring

def create_ekvayring_operation(win, date_str, terminal=None):
    focus(win)
    print('-> "Создать на основании" -> "Эквайринговая операция"')
    btns = [b for b in win.descendants(control_type="Button") if b.window_text().strip() == "Создать на основании"]
    icon_variant = [b for b in btns if b.rectangle().left < 900]
    target = icon_variant[0] if icon_variant else (btns[0] if btns else None)
    if not target:
        raise RuntimeError('"Создать на основании" button not found')
    target.click_input()
    time.sleep(0.6)

    if not click_menuitem(win, "Эквайринговая операция"):
        raise RuntimeError('"Эквайринговая операция" not in the "Создать на основании" menu')
    time.sleep(1.2)
    focus(win)
    guard_no_surprise_dialog(win)

    print(f"-> setting date to {date_str}")
    date_edit = nearest_control_right_of_text(win, "от:", "Edit")
    if not date_edit:
        raise RuntimeError("Date field not found on Эквайринговая операция form")
    ddmmyyyy = date_str.replace(".", "")
    type_digits(date_edit, ddmmyyyy)

    if terminal:
        print(f"-> setting Терминал = {terminal}")
        # Anchor on the "Терминал:" label rather than just grabbing
        # win.descendants(control_type="ComboBox")[0] - the form also has
        # Операция and Контрагент combos, and tree order isn't guaranteed to
        # put Терминал first. Same pattern as the date field just above.
        terminal_combo = nearest_control_right_of_text(win, "Терминал:", "ComboBox")
        if not terminal_combo:
            raise RuntimeError("Терминал combo not found")
        paste_text(terminal_combo, terminal)
        # This suggestion list shows entries like "Iman Pay (0)" - a suffix,
        # not an exact match for the plain "Iman Pay" text typed in, so
        # pick_suggestion's exact-text match never found it. Enter simply
        # accepts whatever's highlighted (the sole/best match), same as
        # confirmed working by hand.
        time.sleep(0.3)
        keyboard.send_keys("{ENTER}")
        time.sleep(0.5)
        focus(win)
        # Проведение validates this field is filled (confirmed live: posting
        # without it fails with 'Поле "Эквайринговый терминал" не заполнено')
        # - re-read it back so a bad/no suggestion match is caught here
        # rather than surfacing as a confusing post-time validation error.
        confirmed = terminal_combo.legacy_properties().get("Value", "").strip()
        if terminal.lower() not in confirmed.lower():
            raise RuntimeError(
                f"Терминал reads {confirmed!r} after trying to set it to {terminal!r} - "
                "it did not take, check by hand."
            )
    else:
        print("-> leaving Терминал as 1C auto-filled it (not part of this workflow)")

    _fill_statya_dds_if_needed(win)


STATYA_DDS_VALUE = "Поступление оплаты от клиента"


def _fill_statya_dds_if_needed(win, value=STATYA_DDS_VALUE):
    """The "Расшифровка платежа" sub-tab has a required "Статья ДДС" (cash-
    flow article) field that 1C leaves blank sometimes but not always -
    confirmed live, posting only fails ('Поле "Статья ДДС" не заполнено')
    on some runs, not every one, with no discernible pattern. Check and fill
    it proactively before ever attempting to post, rather than reacting to
    that validation failure after the fact. Same value as Операция
    ("Поступление оплаты от клиента") - the two are naturally the same
    classification for this kind of document.

    Despite looking like a combo/reference picker (text + "..." + icon,
    same as Терминал), this field's actual UIA control type is "Edit", not
    "ComboBox" - confirmed live, searching for a ComboBox here always came
    back empty, which is why an earlier version of this silently did
    nothing at all (no print, no error - just never found the field). Its
    value must also be read with .get_value(), not window_text() (which
    just echoes the field's label for Edit controls - same as the date
    field elsewhere in this app). It also sits further right of its label
    than nearest_control_right_of_text's default max_gap=40 allows (~89px
    here) - confirmed live - so that needs widening too.
    """
    if not click_tabitem(win, "Расшифровка платежа"):
        print("   (Расшифровка платежа tab not found - Статья ДДС not checked, may need setting by hand)")
        return
    time.sleep(0.4)
    focus(win)

    field = nearest_control_right_of_text(win, "Статья ДДС:", "Edit", max_gap=120)
    if not field:
        print("   (Статья ДДС field not found on Расшифровка платежа tab - may need setting by hand)")
        return

    current = field.get_value().strip()
    if current:
        print(f"   Статья ДДС already set to {current!r}, leaving as-is")
        return

    print(f"-> setting Статья ДДС = {value}")
    paste_text(field, value)
    time.sleep(0.3)
    keyboard.send_keys("{ENTER}")
    time.sleep(0.5)
    focus(win)

    confirmed = field.get_value().strip()
    if not confirmed:
        raise RuntimeError(
            f"Статья ДДС is still blank after trying to set it to {value!r} - it did not take, check by hand."
        )
    print(f"   Статья ДДС confirmed as {confirmed!r}")


def _reopen_related_documents(win, grid, order_number):
    """After posting a document, try clicking back to the still-open
    'Связанные документы' tab (cheap, and preserves the same grid element);
    fall back to a full re-search+reopen only if that tab is gone."""
    focus(win)
    related = [p for p in win.descendants(control_type="Pane") if "Связанные" in p.window_text()]
    if related:
        related[0].click_input()
        time.sleep(0.6)
        focus(win)
        return grid
    print("   Связанные документы tab isn't there anymore - reopening it via search")
    result = open_and_read_related_documents(win, order_number)
    return result["grid"] if result else None


def _has_open_document(win):
    return any(b.window_text().strip() == "Провести и закрыть" for b in win.descendants(control_type="Button"))


def _has_related_documents_tab(win):
    return any("Связанные" in p.window_text() for p in win.descendants(control_type="Pane"))


def close_back_to_zakazy_klientov(win, max_presses=6):
    """After finishing (or giving up on) one order, close back down to the
    Заказы клиентов list before starting the next one - confirmed live,
    Escape closes whatever document or Связанные документы tab is
    currently on top (this is the exact same behavior flagged as a hazard
    in set_price_type_arbitrary's docstring - deliberately used here as a
    feature instead). How many tabs need closing varies by exactly where
    reconcile_order stopped (0 to ~3), so this just presses Escape and
    re-checks rather than assuming a fixed count, stopping as soon as
    neither an open document nor a Связанные документы tab is detected any
    more (i.e. we're back on a plain list view)."""
    for _ in range(max_presses):
        focus(win)
        if not _has_open_document(win) and not _has_related_documents_tab(win):
            return
        keyboard.send_keys("{ESC}")
        time.sleep(0.5)
    focus(win)
    if _has_open_document(win) or _has_related_documents_tab(win):
        print("   (tabs may not have fully closed after several Escape presses - check by hand)")


# ------------------------------------------------------- sheet-driven flow

def _status(kind, reason, **extra):
    """Build the structured result reconcile_order returns. kind is one of:
      "done"    - fully completed (1C posted, sheet updated if ws given).
      "skipped" - stopped before changing anything meaningful in 1C (order
                  not found, related-docs check failed, names didn't match,
                  multi-item shape not handled, or a price fix's own
                  verification failed before any post happened) - safe to
                  just move on to the next order.
      "partial" - SOME 1C document was already posted with a new price
                  before this stopped - needs a human to finish or check by
                  hand, NOT safe to silently skip past in a batch run.
    reason is a short machine-readable code (used to group counts in the
    batch report) - the human-readable detail was already printed by the
    caller."""
    return {"status": kind, "reason": reason, **extra}


# Machine status -> human label for the Excel report's Status column.
_REPORT_STATUS_LABEL = {
    "done": "success",
    "skipped": "skipped",
    "partial": "partial (posted - needs a human)",
    "error": "error",
}


def _report_row(rep, contract, terminal, soglashenie, result):
    """Assemble one report row (see report_sheet.COLUMNS) from the per-order
    `rep` dict a reconcile_* function filled in as it ran, plus the batch
    loop's own knowledge (contract number, terminal, соглашение) and the
    final result status/reason. Any field the reconcile function didn't get
    to fill (e.g. it was skipped before reading the item) just stays blank -
    which is correct, that data genuinely wasn't reached."""
    rep = rep or {}
    reason = result.get("reason", "?")
    comment = reason
    if result.get("detail"):
        comment = f"{comment}: {result['detail']}"
    # Status label: a Cancelled application and a still-"waiting" one both
    # come back with kind "done"/"skipped", so map their specific reason to
    # a truthful label instead of a generic "success"/"skipped" (an
    # order/application that was cancelled must NOT read as "success").
    if reason.startswith("cancel"):
        status_label = "Cancelled"
    elif reason == "waiting_not_ready":
        status_label = "waiting"
    else:
        status_label = _REPORT_STATUS_LABEL.get(result.get("status"), result.get("status") or "?")
    return {
        "id": rep.get("id", ""),
        "contract": contract,
        "order_date": rep.get("order_date", ""),
        "name_1c": rep.get("name_1c", ""),
        "name_source": rep.get("name_source", ""),
        "price_source": rep.get("price_source", ""),
        "price_before": rep.get("price_before", ""),
        "price_after": rep.get("price_after", ""),
        "item2_price": rep.get("item2_price", ""),
        "item3_price": rep.get("item3_price", ""),
        "item4_price": rep.get("item4_price", ""),
        "item5_price": rep.get("item5_price", ""),
        "bonus": rep.get("bonus", ""),
        "quantity": rep.get("quantity", ""),
        "status": status_label,
        "terminal": terminal,
        "soglashenie": soglashenie,
        "category": rep.get("category", ""),
        "commission": rep.get("commission", ""),
        "comment": comment,
    }


def _open_report_ws():
    """Open the live report Google Sheet once at the start of a batch run
    (see report_sheet.py). Returns the worksheet, or None if it isn't
    configured or can't be opened - reporting is best-effort and must never
    stop the actual run, so a failure here just disables per-order logging
    for this run with a warning."""
    if not report_sheet.is_configured():
        print("   (REPORT_SPREADSHEET_URL not set - per-order report logging disabled for this run)")
        return None
    try:
        return report_sheet.get_report_worksheet()
    except Exception as e:  # noqa: BLE001 - never let report setup break the run
        print(f"   (couldn't open the report sheet: {e!r} - per-order report logging disabled for this run)")
        return None


def _log_report_row(report_ws, rep, contract, terminal, soglashenie, result):
    """Append one order's outcome to the live report sheet right after it's
    processed (see report_sheet.append_row). Best-effort - a reporting
    failure is caught and logged, never raised, so it can't mask or stop
    the actual reconciliation."""
    if report_ws is None:
        return
    try:
        report_sheet.append_row(report_ws, _report_row(rep, contract, terminal, soglashenie, result))
    except Exception as e:  # noqa: BLE001 - reporting is a side-channel
        print(f"   (couldn't append {contract} to the report sheet: {e!r} - continuing)")


def reconcile_order(row, agreements=AGREEMENTS, auto_confirm=False, ws=None, terminal=EKVAYRING_TERMINAL,
                     session_state=None, report=None):
    """Core reconcile logic for one row dict {order_number, expected_amount,
    product_name, row (sheet row number, optional)}. Used both by the
    sheet-driven process_next_unchecked_order() below and directly from the
    CLI (--order/--amount/--product) for testing this flow against a real
    order before Google Sheets credentials even exist - see main()'s
    --order handling. If ws is None, the final "write back to the sheet"
    step is skipped and the amount to enter by hand is printed instead.

    agreements and terminal both default to the IMANUM STORE values (see
    their module-level constants) but are parameters, not hardcoded, so a
    future different sheet/workflow can pass its own without touching this
    function.

    session_state, if given, is passed straight to find_order_smart to
    avoid redundant Проведен/Текущее состояние resets across MULTIPLE calls
    to reconcile_order sharing the same instance (see
    process_all_unchecked_orders) - leave it None for a one-off single-order
    run (--order/manual), which always does the full filter reset, same as
    before.

    Returns a dict via _status() above - callers (especially the batch
    loop) use "status" to decide whether it's safe to move on to the next
    order without further attention ("done"/"skipped") or not ("partial")."""
    def ok(prompt):
        return True if auto_confirm else confirm(prompt)

    rep = report if report is not None else {}  # filled as data is discovered (see report_sheet.py)
    _reset_price_notes()

    order_number = row["order_number"]
    expected_amount = row["expected_amount"]
    expected_product = row["product_name"]
    rep["name_source"] = expected_product
    rep["price_source"] = row.get("amount_error") or expected_amount
    if row.get("amount_error"):
        # find_next_unchecked_row couldn't parse this row's СуммаИМАН (see
        # its amount_error) - skip cleanly (logged in the report) instead of
        # letting the bad value blow up mid-reconcile.
        print(f"!! {order_number}: СуммаИМАН value {row['amount_error']!r} isn't a valid number - "
              "skipping this order (fix the cell and re-run).")
        return _status("skipped", "unparseable_amount", detail=row["amount_error"])
    print(f"=== order {order_number}, expected amount={expected_amount}, "
          f"expected product={expected_product!r} ===")

    win = get_window()
    focus(win)
    open_zakazy_klientov(win)

    agreement = find_order_smart(win, order_number, agreements, session_state if session_state is not None else {})
    if not agreement:
        print(f"!! {order_number} not found in 1C under any of {agreements} - "
              "this is the rare exceptional case, skipping (needs manual handling).")
        return _status("skipped", "not_found_in_1c")

    result = open_and_read_related_documents(win, order_number)
    if result is None:
        print(f"!! Related-documents check failed for {order_number} (see messages above) - "
              "stopping, nothing was changed.")
        return _status("skipped", "related_docs_check_failed")
    if result.get("ekvayring"):
        print(f"=== {order_number} already has an Эквайринговая операция "
              f"({result['ekvayring']}) - reconciled by an earlier run, nothing to do. ===")
        return _status("skipped", "already_has_ekvayring", detail=result["ekvayring"])
    realizatsiya_date, grid = result["date"], result["grid"]
    rep["order_date"] = result.get("order_date") or ""

    # --- Заказ клиента ---
    open_related_row(win, grid, row_index=0, expect_title_startswith="Заказ клиента")
    guard_no_surprise_dialog(win)
    click_tabitem(win, "Товары")
    time.sleep(0.5)
    focus(win)

    order_items = read_item_rows(win)
    kinds = [_classify_order_item(it["nomenklatura"]) for it in order_items]
    delivery = [it for it, k in zip(order_items, kinds) if k == "delivery"]
    split_note = ""
    if delivery:
        # Delivery split (same rule as the other profiles): the sheet's
        # СуммаИМАН is the FULL financed total, but 1C splits it into a
        # product line + a Доставка line, so the product line = sheet total
        # − the delivery line's current price (product + delivery == sheet
        # total). The note goes in the report's Comment column, not a tag.
        # Only a single product line + delivery is handled here.
        products = [it for it, k in zip(order_items, kinds) if k == "product"]
        if len(products) != 1:
            print(f"!! {order_number} has a Доставка line plus {len(products)} product line(s) - the "
                  "delivery-split rule only handles a single product line, stopping.")
            return _status("skipped", "delivery_multi_product_unhandled")
        product = products[0]
        if product["kolichestvo"] != 1:
            print(f"!! {order_number}'s product line quantity is {product['kolichestvo']}, not 1 - stopping.")
            return _status("skipped", "quantity_not_1")
        delivery_total = sum((d["tsena"] or 0) for d in delivery)
        product_target = round((expected_amount or 0) - delivery_total, 2)
        split_note = _split_delivery_note(order_number, expected_amount, delivery)
        if split_note:
            print(f"   {split_note}")
        row_num = row.get("row")
        pairs = [(product, {"row": row_num, "rows": [row_num] if row_num is not None else [],
                            "product_name": expected_product, "expected_amount": expected_amount,
                            "target_price": product_target, "skip_name_check": False})]
        pairs += [(d, {"row": None, "rows": [], "product_name": d["nomenklatura"], "expected_amount": 0,
                       "target_price": d["tsena"], "skip_name_check": True, "skip_price_fix": True,
                       "is_delivery": True}) for d in delivery]
    else:
        try:
            pairs = determine_item_price_pairs(win, ws, order_number, expected_product, expected_amount,
                                                sheet_row_number=row.get("row"))
        except MultiItemMatchError as e:
            print(f"!! {order_number}: {e} - stopping without changes.")
            return _status("skipped", "multi_item_unhandled", detail=str(e))

    # Fill the report from the matched pairs. The main (non-bonus) items go
    # in the primary/Item-2..5 price columns; a bonus item (skip_name_check
    # - the near-free bundled extra, see resolve_bonus_item_prices) goes in
    # the Bonus column. "price before" is the main item's original Tsena;
    # "price after" is what this run sets it to.
    main_pairs = [(it, sr) for it, sr in pairs if not sr.get("skip_name_check")]
    # A delivery line is skip_name_check too, but it isn't a "bonus" - keep
    # it out of the Bonus column (it's noted in the Comment instead).
    bonus_pairs = [(it, sr) for it, sr in pairs if sr.get("skip_name_check") and not sr.get("is_delivery")]
    if main_pairs:
        rep["name_1c"] = main_pairs[0][0]["nomenklatura"]
        rep["price_before"] = main_pairs[0][0]["tsena"]
        rep["price_after"] = main_pairs[0][1]["target_price"]
        rep["quantity"] = main_pairs[0][0]["kolichestvo"]
        for n, (_, sr) in enumerate(main_pairs[1:5], start=2):  # extra items -> Item 2..5 price columns
            rep[f"item{n}_price"] = sr["target_price"]
    if bonus_pairs:
        bi, bsr = bonus_pairs[0]
        rep["bonus"] = f"{bi['nomenklatura']} = {bsr['target_price']}"

    if len(pairs) > 1:
        print(f"   {order_number} has {len(pairs)} items - matched each to its own sheet row:")
        for item, sheet_row in pairs:
            print(f"     [{item['index']}] {item['nomenklatura']!r} (Количество={item['kolichestvo']}) <-> "
                  f"sheet row(s) {sheet_row['rows']} {sheet_row['product_name']!r} = {sheet_row['expected_amount']}")

    for item, sheet_row in pairs:
        if sheet_row.get("skip_name_check"):
            continue  # bonus item - the sheet's product_name describes the main item, not this one
        if not names_match_or_confirm(sheet_row["product_name"], item["nomenklatura"]):
            print(f"!! Product name mismatch on {order_number} - stopping without changes.")
            return _status("skipped", "name_mismatch")

    price_summary = ", ".join(str(sr["target_price"]) for _, sr in pairs)
    if not ok(f"Set {order_number}'s order price(s) to [{price_summary}] and post it?"):
        print("Left the order open without changing its price - stopping.")
        return _status("skipped", "user_declined_order_price")
    try:
        for item, sheet_row in pairs:
            if sheet_row.get("skip_price_fix"):
                continue  # delivery line - leave its existing price untouched
            apply_price_fix_for_item(win, item, sheet_row["target_price"])
    except RuntimeError as e:
        print(f"!! {order_number}'s order price fix failed verification: {e} - stopping without posting.")
        return _status("skipped", "order_price_fix_failed", detail=str(e))
    try:
        post_and_close(win, verify_closed_title_contains="Заказ клиента")
    except RuntimeError as e:
        print(f"!! {order_number}'s order did not post: {e} - check by hand.")
        return _status("partial", "order_post_failed", detail=str(e))

    # --- Реализация ---
    print("-> back to Связанные документы, opening Реализация (row 1)")
    grid = _reopen_related_documents(win, grid, order_number)
    if grid is None:
        print(f"!! Could not reopen Связанные документы for {order_number} after fixing the order price - "
              "stopping, check by hand (order price WAS already changed).")
        return _status("partial", "reopen_failed_after_order_post")

    open_related_row(win, grid, row_index=1, expect_title_startswith="Реализация")
    guard_no_surprise_dialog(win)
    click_tabitem(win, "Товары")
    time.sleep(0.5)
    focus(win)
    if not ok("Set the Реализация's price(s) the same way, and post it?"):
        print("Left the Реализация untouched (order price WAS already changed) - finish this by hand.")
        return _status("partial", "user_declined_realizatsiya_price")
    try:
        realiz_items = read_item_rows(win)
        if len(realiz_items) != len(pairs):
            raise RuntimeError(
                f"Реализация has {len(realiz_items)} item row(s), expected {len(pairs)} to match the order"
            )
        # Реализация's own Товары grid is read fresh here (its rows aren't
        # guaranteed to be in the same order as the order's), so re-pair by
        # comparing each item's name against the ORDER's own item names
        # (both are 1C text for the same underlying product - matches
        # near-exactly), NOT against sheet_row["product_name"] - for the
        # bonus-item bundle case that's identical across both pairs (the
        # sheet only ever names the main product), which would match the
        # wrong item's target_price to the wrong row.
        remaining_pairs = list(pairs)
        realiz_pairs = []
        for ritem in realiz_items:
            remaining_pairs.sort(
                key=lambda p: product_name_coverage(p[0]["nomenklatura"], ritem["nomenklatura"]), reverse=True)
            _, sheet_row = remaining_pairs.pop(0)
            realiz_pairs.append((ritem, sheet_row))
        for ritem, sheet_row in realiz_pairs:
            if sheet_row.get("skip_price_fix"):
                continue  # delivery line - leave its existing price untouched
            apply_price_fix_for_item(win, ritem, sheet_row["target_price"])
    except RuntimeError as e:
        print(f"!! {order_number}'s Реализация price fix failed: {e} - "
              "stopping (order price was already fixed, check by hand).")
        return _status("partial", "realizatsiya_price_fix_failed", detail=str(e))
    try:
        post_and_close(win, verify_closed_title_contains="Реализация")
    except RuntimeError as e:
        print(f"!! {order_number}'s Реализация did not post: {e} - order price already fixed, check by hand.")
        return _status("partial", "realizatsiya_post_failed", detail=str(e))

    # --- Эквайринговая операция ---
    print("-> back to Связанные документы, opening the order (row 0) for Эквайринговая операция")
    grid = _reopen_related_documents(win, grid, order_number)
    if grid is None:
        print(f"!! Could not reopen Связанные документы for {order_number} to create the Эквайринговая операция - "
              "both prices ARE already fixed, finish this one by hand.")
        return _status("partial", "reopen_failed_before_ekvayring")
    open_related_row(win, grid, row_index=0, expect_title_startswith="Заказ клиента")
    guard_no_surprise_dialog(win)

    try:
        create_ekvayring_operation(win, realizatsiya_date or "", terminal=terminal)
    except RuntimeError as e:
        print(f"!! {order_number}'s Эквайринговая операция setup failed: {e} - "
              "both prices are already fixed and posted; finish this by hand.")
        return _status("partial", "ekvayring_setup_failed", detail=str(e))

    if not ok("Эквайринговая операция form filled in correctly above (date and Терминал set)? Post & close it now?"):
        print("Left the Эквайринговая операция open without posting - "
              "both prices are already fixed and posted; finish this by hand.")
        return _status("partial", "user_declined_ekvayring")
    try:
        post_and_close(win, verify_closed_title_contains="Эквайринговая")
    except RuntimeError as e:
        print(f"!! {order_number}'s Эквайринговая операция did not post: {e} - "
              "both prices are already fixed and posted; finish this by hand.")
        return _status("partial", "ekvayring_post_failed", detail=str(e))

    # Dedupe by row: the bonus-item bundle case has two pairs sharing the
    # SAME sheet row (there's only one row for the pair of them) - without
    # this, the same row's Сумма 1С would get written twice (harmlessly
    # identical writes, but confusing in the confirm prompt/log either way).
    # The quantity>1 case goes the other way - ONE pair can carry SEVERAL
    # sheet rows (one 1C line covering multiple duplicate sheet rows, see
    # _match_items_to_sheet_rows_by_quantity) - every one of them gets the
    # same per-unit amount written back, since that's genuinely what each
    # of those physical units cost.
    seen_rows = set()
    write_targets = []
    for _, sheet_row in pairs:
        for r in sheet_row["rows"]:
            if r is None or r in seen_rows:
                continue
            seen_rows.add(r)
            write_targets.append((r, sheet_row["expected_amount"]))
    if ws is None or not write_targets:
        by_hand = ", ".join(f"row {r}={amt}" for r, amt in write_targets) or f"{expected_amount} (no row number)"
        print(f"=== Done: {order_number} reconciled in 1C. No sheet configured - "
              f"update Сумма 1С by hand: {by_hand}. ===")
        return _status("done", "no_sheet_writeback",
                       detail=_combine_detail(row.get("amount_note"), split_note, _price_notes_suffix()))
    targets_desc = ", ".join(f"row {r}={amt}" for r, amt in write_targets)
    if not ok(f"Write Сумма 1С for [{targets_desc}] now?"):
        print("Skipped writing back to the sheet - 1C side is fully done, update the sheet by hand.")
        return _status("partial", "user_declined_writeback")
    for r, amt in write_targets:
        mark_sum_1c(ws, r, amt)
    print(f"=== Done: {order_number} reconciled, sheet row(s) {[r for r, _ in write_targets]} updated. ===")
    return _status("done", "reconciled", rows=[r for r, _ in write_targets],
                   detail=_combine_detail(row.get("amount_note"), split_note, _price_notes_suffix()))


# ----------------------------------------------------------- Alif Shop flow

# A KIT is one Alif Business product that 1C and the sheet both stock as
# several separate nomenclatures - a split air conditioner is a single Alif
# item but an indoor block plus an outdoor block everywhere else. The
# companion lines each carry this nominal price and the line matching the
# Alif item carries all the rest, so the order still totals exactly the
# post-commission price. Confirmed against every kit already reconciled by
# hand, in the sheet AND in 1C itself - e.g. ON-26/00105 reads 3 394 029 on
# the Внутренний блок and 1 on the Наружный, totalling the 3 394 030 that
# Alif's 3 499 000 comes to after its 3% commission.
KIT_COMPANION_PRICE = 1.0


def _alifshop_write_targets(priced_pairs, companion_rows):
    """[(sheet row number, price to write into Сумма 1С), ...] for one
    application - the single place that decides what lands on which row, so
    the normal path and the already-reconciled-in-1C path can't drift apart.

    Without a kit this is just each alif_item's per-unit post-commission
    price on every row it covers. With one, the companion rows take
    KIT_COMPANION_PRICE each and the matched row takes the remainder, which
    keeps the row total equal to the post-commission price to the kopek."""
    companion_total = round(len(companion_rows) * KIT_COMPANION_PRICE, 2)
    targets = []
    for _alif_item, srs, price_1c in priced_pairs:
        price = round(price_1c - companion_total, 2) if companion_rows else price_1c
        targets.extend((sr["row"], price) for sr in srs)
    targets.extend((cr["row"], KIT_COMPANION_PRICE) for cr in companion_rows)
    return targets


def _match_alifshop_items_to_sheet_rows(alif_items, sheet_rows):
    """Greedy fuzzy-match Alif Business's item list for one application
    against the Alif Nasiya sheet's rows for that same application. Each
    alif_item covers one OR MORE sheet rows whose Kolichestvo values sum to
    exactly that alif_item's own Kolichestvo - confirmed live:
      - 1000002 (2 items kol=1 each, 2 sheet rows kol=1 each): 1:1 pairing.
      - 1000003 (1 item kol=4, 3 sheet rows kol=2+1+1=4):
        the single alif_item covers all 3 sheet rows.

    Returns ([(alif_item, [sheet_rows...]), ...], [companion_rows...]) - the
    pairs in alif_items order. A sheet row is claimed whole - never split
    across two alif_items - because F/G on that row (which we write later) is
    a single per-unit price, so all of a row's units must come from one
    alif_item.

    A KIT is the one case where the sheet legitimately carries MORE units
    than Alif Business does: a single Alif item that 1C and the sheet both
    stock as several separate nomenclatures. Confirmed live and repeatedly -
    a split air conditioner is one Alif product (1000004: "Внутренний блок
    ... Inv", 3 499 000) but two sheet rows, Внутренний блок + Наружный блок.
    Those leftover rows come back as companion_rows and are priced at a
    nominal KIT_COMPANION_PRICE each (see _alifshop_write_targets), NOT
    treated as a pairing failure - which is what used to happen, stopping the
    application with "shows 1.0 total unit(s) but the sheet's rows sum to
    2.0".

    Kits are only accepted in the shape every reconciled one in the sheet
    actually has: exactly one Alif item, Количество 1, covering exactly one
    sheet row. Anything wider (two Alif items AND leftover rows - which of
    them owns the leftover?) still raises rather than guessing at how to
    spread money across rows.

    Raises MultiItemMatchError if the totals don't line up in any other way,
    or the row kolichestvos can't sum exactly to each alif_item's."""
    total_alif = sum(item["kolichestvo"] for item in alif_items)
    total_sheet = sum((r["kolichestvo"] or 1) for r in sheet_rows)
    if total_sheet + 0.5 < total_alif:
        raise MultiItemMatchError(
            f"Alif Business shows {total_alif} total unit(s) but the sheet's rows sum to only "
            f"{total_sheet} for this application - don't know how to pair these, stopping.")
    if total_sheet > total_alif + 0.5 and len(alif_items) != 1:
        raise MultiItemMatchError(
            f"Alif Business shows {total_alif} total unit(s) but the sheet's rows sum to {total_sheet} "
            f"for this application, across {len(alif_items)} Alif items - can't tell which item the "
            "extra row(s) belong to, stopping.")

    remaining = list(sheet_rows)
    pairs = []
    for item in alif_items:
        # Prefer rows whose name fuzzy-matches this alif_item, and among
        # those prefer the LARGEST kolichestvo first (so a kol=2 row gets
        # claimed by a kol>=2 alif_item, not split into two 1-unit rows).
        need = item["kolichestvo"]
        remaining.sort(
            key=lambda r: (product_name_coverage(r["item_name"], item["name"]), (r["kolichestvo"] or 1)),
            reverse=True)
        covered = []
        remaining_need = need
        used_indices = []
        for i, r in enumerate(remaining):
            rn = r["kolichestvo"] or 1
            if rn > remaining_need + 0.5:
                continue  # too big for what's left
            covered.append(r)
            used_indices.append(i)
            remaining_need -= rn
            if remaining_need <= 0.5:
                break
        if remaining_need > 0.5:
            raise MultiItemMatchError(
                f"couldn't cover {item['name']!r} (Количество={item['kolichestvo']}) exactly with the "
                "available sheet rows' kolichestvos - stopping.")
        for i in sorted(used_indices, reverse=True):
            del remaining[i]
        pairs.append((item, covered))
    if remaining:
        # Leftovers are only ever the companion halves of a kit - see this
        # function's docstring - and only in the exact shape that's been
        # reconciled before. The name-coverage sort above is what puts the
        # Alif item's own match in `covered` and leaves the other halves here.
        if len(alif_items) != 1 or (alif_items[0]["kolichestvo"] or 1) != 1 or len(pairs[0][1]) != 1:
            raise MultiItemMatchError(
                f"after covering all Alif Business items, {len(remaining)} sheet row(s) still "
                "unassigned - stopping.")
        return pairs, remaining
    return pairs, []


def reconcile_alifshop_application(application_id, auto_confirm=False, agreements=ALIFSHOP_AGREEMENTS,
                                    terminal=ALIFSHOP_TERMINAL, session_state=None, report=None):
    """Core reconcile logic for one Alif Shop application (see --application
    in main()) - a sibling to reconcile_order (IMAN Sverka) above, sharing
    all the same low-level 1C helpers (find_order_under_agreements,
    open_and_read_related_documents, apply_price_fix_for_item,
    create_ekvayring_operation, post_and_close, ...), but everything about
    WHERE the expected product/price comes from is different:

      - IMAN's sheet already has "Сумма 1С" as ground truth per row; Alif
        Shop's sheet doesn't - the price has to be computed from what
        business.alif.uz says the item sold for, minus Alif's own
        commission % for that item's category (see alif_client.py,
        alif_sheets_client.py's module docstrings for how each of those is
        scraped/looked up).
      - The 1C order number is already given directly in the sheet (column
        K) - unlike IMAN, there's no need to search the sheet for it.
      - An application can be in a genuinely unresolved state (Alif's own
        "Yetkazib berishga" - out for delivery) that isn't a stop condition
        so much as "nothing to do yet, check again later" - recorded in
        the sheet as such (see alif_sheets_client.mark_waiting) rather than
        reported as skipped/needing attention.
      - Alif Business can also report an application Cancelled
        ("Bekor qilindi") - nothing to reconcile in 1C at all in that case,
        just record it (mark_cancelled, which also colors the row red).
      - A multi-item application spans several sheet rows sharing the same
        application id (one 1C order, several item rows) - see
        _match_alifshop_items_to_sheet_rows - and a successful reconcile
        writes the outcome back to ALL of them, same idea as the IMAN
        flow's multi-item write-back.

    Quantity > 1 is supported on all three sides (Alif Business items,
    sheet rows, 1C product lines) as long as unit totals line up:
    confirmed live (application 1000003: 1 Alif item kol=4, 3 sheet rows
    kol=2+1+1, 1C 3 rows kol=1+2+1 - all 4 units). Sheet write-back writes
    the per-unit post-commission price to F/G on EVERY covered sheet row,
    regardless of that row's kolichestvo (per instruction). Anything the
    unit-totals don't cover cleanly stops rather than guessing, same
    philosophy as reconcile_order throughout."""
    def ok(prompt):
        return True if auto_confirm else confirm(prompt)

    rep = report if report is not None else {}  # filled as data is discovered (see report_sheet.py)
    _reset_price_notes()
    rep["id"] = application_id

    ws = alif_sheets_client.get_alif_worksheet()
    rows = alif_sheets_client.find_all_rows_for_application(ws, application_id)
    if not rows:
        print(f"!! Application {application_id!r} not found in the Alif Nasiya sheet - check the id.")
        return _status("skipped", "application_not_found_in_sheet")

    order_number = rows[0]["order_number"]
    rep["contract"] = order_number

    existing_status = rows[0]["status"].strip().lower()
    if existing_status.startswith("cancel"):
        # Already Cancelled in the sheet - re-affirm it (re-write the
        # 'Cancelled' status AND re-paint the row(s) red) rather than doing
        # nothing. Confirmed live: a row can end up with the 'Cancelled'
        # text but WITHOUT the red fill - a prior run's paint step didn't
        # land, or a human set the dropdown by hand - and the old "nothing
        # to do" short-circuit left it unpainted (and unlogged).
        # mark_cancelled is idempotent, so re-applying is safe.
        for r in rows:
            alif_sheets_client.mark_cancelled(ws, r["row"])
        print(f"=== Application {application_id} (sheet row(s) {[r['row'] for r in rows]}) already "
              "Cancelled - re-affirmed the status and red highlight. ===")
        return _status("done", "cancelled", detail="was already Cancelled in the sheet")
    if existing_status.startswith("success"):
        print(f"=== Application {application_id} (sheet row {rows[0]['row']}) is already "
              f"{rows[0]['status']!r} - nothing to do. ===")
        return _status("done", "already_resolved")

    print(f"=== Alif Shop application {application_id}, sheet row(s) {[r['row'] for r in rows]}, "
          f"1C order {order_number!r} ===")

    with sync_playwright() as p:
        context = alif_client.get_authenticated_context(p)
        try:
            status = alif_client.get_application_status(context, application_id)
            print(f"-> Alif Business status: {status!r}")

            if status == alif_client.STATUS_CANCELLED:
                for r in rows:
                    alif_sheets_client.mark_cancelled(ws, r["row"])
                print(f"=== Done: application {application_id} is cancelled - sheet row(s) "
                      f"{[r['row'] for r in rows]} marked Cancelled (red). ===")
                return _status("done", "cancelled")

            if status == alif_client.STATUS_WAITING:
                for r in rows:
                    alif_sheets_client.mark_waiting(ws, r["row"])
                print(f"=== Application {application_id} is still {status!r} - sheet row(s) "
                      f"{[r['row'] for r in rows]} marked waiting, check again later. ===")
                return _status("skipped", "waiting_not_ready")

            if status not in alif_client.STATUSES_READY_TO_RECONCILE:
                print(f"!! Application {application_id}'s status {status!r} isn't one this script knows "
                      "how to act on yet - stopping.")
                return _status("skipped", "unrecognized_status", detail=status)

            # Both "Rasmiylashtirildi" and "Tasdiqlandi" are safe to
            # reconcile; when it's the less-usual Tasdiqlandi, note that in
            # the report Comment so it's visible after the fact.
            status_note = f"reconciled on status {status!r}" if status == alif_client.STATUS_CONFIRMED else ""

            detail = alif_client.get_application_detail(context, application_id)
        finally:
            context.browser.close()

    alif_items = detail["items"]
    try:
        item_pairs, companion_rows = _match_alifshop_items_to_sheet_rows(alif_items, rows)
    except MultiItemMatchError as e:
        print(f"!! Application {application_id}: {e}")
        return _status("skipped", "multi_item_unhandled", detail=str(e))
    if companion_rows:
        print(f"   application {application_id} is a KIT: one Alif item across "
              f"{1 + len(companion_rows)} nomenclatures - sheet row(s) "
              f"{[cr['row'] for cr in companion_rows]} are the companion half/halves and get "
              f"{KIT_COMPANION_PRICE} each:")
        for cr in companion_rows:
            print(f"     row {cr['row']}: {cr['item_name']!r}")

    if len(item_pairs) > 1 or any(len(srs) > 1 for _, srs in item_pairs):
        print(f"   application {application_id} has {len(item_pairs)} Alif item(s) - matched to sheet rows:")
        for alif_item, srs in item_pairs:
            row_desc = ", ".join(f"row {sr['row']} (kol={sr['kolichestvo']})" for sr in srs)
            print(f"     {alif_item['name']!r} (kol={alif_item['kolichestvo']}) <-> {row_desc}")

    order_numbers = {r["order_number"] for r in rows}
    if len(order_numbers) > 1:
        print(f"!! Application {application_id}'s sheet rows disagree on the 1C order number ({order_numbers}) - "
              "stopping.")
        return _status("skipped", "inconsistent_order_number")
    if not order_number:
        print(f"!! Sheet row(s) {[r['row'] for r in rows]} have no 1C order number (column K) - "
              "can't reconcile in 1C, stopping.")
        return _status("skipped", "no_1c_order_number")

    commissions_ws = alif_sheets_client.get_commissions_worksheet()
    priced_pairs = []  # [(alif_item, [sheet_rows], price_1c), ...] - price_1c is PER UNIT (post-commission)
    for alif_item, srs in item_pairs:
        try:
            commission = alif_sheets_client.get_commission_for_category(commissions_ws, alif_item["category"])
        except ValueError as e:
            print(f"!! {e}")
            return _status("skipped", "commission_lookup_failed", detail=str(e))
        price_1c = round(alif_item["price"] * (1 - commission), 2)
        print(f"   category {alif_item['category']!r} - Alif price {alif_item['price']} (per unit) - "
              f"{commission * 100:g}% commission = {price_1c} per unit "
              f"(alif_item kol={alif_item['kolichestvo']}, sheet rows {[sr['row'] for sr in srs]})")
        priced_pairs.append((alif_item, srs, price_1c))

    # Fill the report with the (possibly multi-) item data. The first
    # alif_item's post-commission per-unit price is "price after"; each
    # further alif_item's per-unit price goes in Item 2..5.
    first_item, _first_srs, first_price_1c = priced_pairs[0]
    rep["name_source"] = first_item["name"]
    rep["price_source"] = first_item["price"]
    rep["category"] = first_item["category"]
    # Commission column shows both the rate and the money it took off this
    # item (Alif's price minus what goes into 1C), e.g. "2.9% (62 321)".
    rep["commission"] = f"{commission * 100:g}% ({round(first_item['price'] - first_price_1c, 2)})"
    rep["price_after"] = first_price_1c
    rep["quantity"] = first_item.get("kolichestvo", "")
    for n, (_, _, p) in enumerate(priced_pairs[1:5], start=2):  # extra items -> Item 2..5 price columns
        rep[f"item{n}_price"] = p

    win = get_window()
    focus(win)
    open_zakazy_klientov(win)

    agreement = find_order_smart(win, order_number, agreements, session_state if session_state is not None else {})
    if not agreement:
        print(f"!! {order_number} not found in 1C under any of {agreements} - "
              "this is the rare exceptional case, skipping (needs manual handling).")
        return _status("skipped", "not_found_in_1c")

    result = open_and_read_related_documents(win, order_number)
    if result is None:
        print(f"!! Related-documents check failed for {order_number} (see messages above) - "
              "stopping, nothing was changed.")
        return _status("skipped", "related_docs_check_failed")
    if result.get("ekvayring"):
        # 1C is already finished for this application - an earlier run fixed
        # the prices and posted the Эквайринговая операция, and only the
        # sheet write-back never happened (interrupted run, declined prompt,
        # a Sheets API hiccup). So FINISH THAT rather than skipping: left
        # unresolved, the row gets picked up by every later --since run,
        # walks the whole Alif + 1C lookup again, and lands right back here,
        # forever. Same rows and the same per-unit post-commission prices the
        # normal path below would have written.
        write_targets = _alifshop_write_targets(priced_pairs, companion_rows)
        targets_desc = ", ".join(f"row {r}={amt}" for r, amt in write_targets)
        print(f"=== {order_number} already has an Эквайринговая операция ({result['ekvayring']}) - "
              "the 1C side was finished by an earlier run, only the sheet is behind. ===")
        if not write_targets:
            # Shouldn't happen (an application always has at least one item
            # and _match_alifshop_items_to_sheet_rows raises rather than
            # returning nothing), but writing zero rows and still reporting
            # "done" would be a lie in the batch report.
            print("!! No sheet rows to write for this application - skipping instead.")
            return _status("skipped", "already_has_ekvayring", detail=result["ekvayring"])
        if not ok(f"Write sheet row(s) back ({targets_desc}, status=success, rows colored green)?"):
            print("Skipped writing back to the sheet - 1C side is fully done, update the sheet by hand.")
            return _status("partial", "user_declined_writeback")
        for r, amt in write_targets:
            alif_sheets_client.mark_success(ws, r, amt)
        print(f"=== Done: application {application_id} / {order_number} was already reconciled in 1C, "
              f"sheet row(s) {[r for r, _ in write_targets]} updated. ===")
        return _status("done", "already_reconciled_in_1c", rows=[r for r, _ in write_targets],
                       detail=result["ekvayring"])
    realizatsiya_date, grid = result["date"], result["grid"]
    rep["order_date"] = result.get("order_date") or ""

    # --- Заказ клиента ---
    open_related_row(win, grid, row_index=0, expect_title_startswith="Заказ клиента")
    guard_no_surprise_dialog(win)
    click_tabitem(win, "Товары")
    time.sleep(0.5)
    focus(win)

    order_items = read_item_rows(win)
    if order_items:
        rep["name_1c"] = order_items[0]["nomenklatura"]
        rep["price_before"] = order_items[0]["tsena"]
    kinds = [_classify_order_item(oi["nomenklatura"]) for oi in order_items]
    delivery = [oi for oi, k in zip(order_items, kinds) if k == "delivery"]
    product_items = [oi for oi, k in zip(order_items, kinds) if k != "delivery"]

    # Total-units check (not row count): 1C can split the same product
    # across N rows with various Количество values (confirmed live: alif-
    # shop 1000003 has 1 Alif item kol=4 and 1C shows 3 rows with kol
    # 1+2+1=4). What must line up is the UNIT totals per side.
    alif_units = sum(ai["kolichestvo"] for ai, _, _ in priced_pairs)
    order_units = sum(oi["kolichestvo"] for oi in product_items)
    companion_lines = []
    if companion_rows:
        # KIT (see _match_alifshop_items_to_sheet_rows): 1C carries the same
        # extra nomenclatures the sheet does, so it legitimately shows MORE
        # units than Alif Business. The line whose name matches the Alif item
        # is the one that takes the money; the rest are companions at
        # KIT_COMPANION_PRICE. Checked strictly against the sheet's own
        # count - if the two disagree about how many pieces this kit has,
        # that's not something to spread money across on a guess.
        main_line = max(product_items,
                        key=lambda oi: product_name_coverage(priced_pairs[0][0]["name"],
                                                             oi["nomenklatura"]))
        companion_lines = [oi for oi in product_items if oi is not main_line]
        if len(companion_lines) != len(companion_rows):
            print(f"!! {order_number}: the sheet has {len(companion_rows)} companion row(s) for this kit "
                  f"but 1C has {len(companion_lines)} extra product line(s) - stopping.")
            return _status("skipped", "kit_shape_mismatch")
        if any((oi["kolichestvo"] or 1) != 1 for oi in product_items):
            print(f"!! {order_number}: a kit line has Количество > 1 - only 1-unit kit lines have been "
                  "reconciled before, stopping rather than guessing.")
            return _status("skipped", "kit_shape_mismatch")
        print(f"   kit in 1C: {main_line['nomenklatura']!r} takes the price, "
              f"{[oi['nomenklatura'] for oi in companion_lines]} get {KIT_COMPANION_PRICE} each")
    elif abs(alif_units - order_units) > 0.5:
        print(f"!! {order_number}: 1C shows {order_units} product unit(s) (across {len(product_items)} row(s), "
              f"plus {len(delivery)} delivery), but Alif Business says {alif_units} - stopping.")
        return _status("skipped", "multi_item_unhandled")

    # Delivery split (per instruction, same rule as the other profiles): the
    # product's post-commission per-unit price is the FULL financed per-unit
    # total, so the 1C product line = that minus the Доставка line's current
    # price (product + delivery == the total). Only a single product line +
    # delivery is handled. The note goes in the report's Comment column.
    delivery_total = sum((d["tsena"] or 0) for d in delivery)
    split_note = ""
    if delivery:
        if len(priced_pairs) != 1 or priced_pairs[0][0]["kolichestvo"] != 1 or len(product_items) != 1:
            print(f"!! {order_number} has a Доставка line but multiple product units - the delivery-split "
                  "rule only handles a single 1-unit product line, stopping.")
            return _status("skipped", "delivery_multi_product_unhandled")
        ai0, srs0, price0 = priced_pairs[0]
        priced_pairs[0] = (ai0, srs0, round((price0 or 0) - delivery_total, 2))
        rep["price_after"] = priced_pairs[0][2]
        split_note = _split_delivery_note(order_number, price0, delivery)
        if split_note:
            print(f"   {split_note}")

    # Match each 1C product line by fuzzy name against the alif_items in
    # priced_pairs. Each 1C line's Цена gets its matched alif_item's PER-UNIT
    # post-commission price (1C's Цена is per-unit, multiplied by Количество
    # for the line total). Multiple 1C lines can share the same alif_item -
    # that's the whole point of the multi-row shape.
    def find_alif_for(nomenklatura):
        return max(priced_pairs, key=lambda t: product_name_coverage(t[0]["name"], nomenklatura))

    companion_ids = {id(oi) for oi in companion_lines}
    companion_total = round(len(companion_lines) * KIT_COMPANION_PRICE, 2)
    order_triples = []  # [(order_item, alif_item, price_1c), ...]
    for order_item in product_items:
        if id(order_item) in companion_ids:
            # alif_item None marks a kit companion: it has no Alif item of
            # its own to be named after or priced from.
            order_triples.append((order_item, None, KIT_COMPANION_PRICE))
            continue
        alif_item, _srs, price_1c = find_alif_for(order_item["nomenklatura"])
        order_triples.append((order_item, alif_item, round(price_1c - companion_total, 2)))

    for order_item, alif_item, price_1c in order_triples:
        if alif_item is None:
            continue  # a kit companion is a DIFFERENT nomenclature by design - nothing to match it against
        if not names_match_or_confirm(alif_item["name"], order_item["nomenklatura"]):
            print(f"!! Product name mismatch on {order_number} - stopping without changes.")
            return _status("skipped", "name_mismatch")

    price_summary = ", ".join(str(price_1c) for _, _, price_1c in order_triples)
    if not ok(f"Set {order_number}'s order price(s) to [{price_summary}] (per unit) and post it?"):
        print("Left the order open without changing its price - stopping.")
        return _status("skipped", "user_declined_order_price")
    try:
        for order_item, _, price_1c in order_triples:
            apply_price_fix_for_item(win, order_item, price_1c)
    except RuntimeError as e:
        print(f"!! {order_number}'s order price fix failed verification: {e} - stopping without posting.")
        return _status("skipped", "order_price_fix_failed", detail=str(e))
    try:
        post_and_close(win, verify_closed_title_contains="Заказ клиента")
    except RuntimeError as e:
        print(f"!! {order_number}'s order did not post: {e} - check by hand.")
        return _status("partial", "order_post_failed", detail=str(e))

    # --- Реализация ---
    print("-> back to Связанные документы, opening Реализация (row 1)")
    grid = _reopen_related_documents(win, grid, order_number)
    if grid is None:
        print(f"!! Could not reopen Связанные документы for {order_number} after fixing the order price - "
              "stopping, check by hand (order price WAS already changed).")
        return _status("partial", "reopen_failed_after_order_post")

    open_related_row(win, grid, row_index=1, expect_title_startswith="Реализация")
    guard_no_surprise_dialog(win)
    click_tabitem(win, "Товары")
    time.sleep(0.5)
    focus(win)
    if not ok("Set the Реализация's price(s) the same way, and post it?"):
        print("Left the Реализация untouched (order price WAS already changed) - finish this by hand.")
        return _status("partial", "user_declined_realizatsiya_price")
    try:
        realiz_items = read_item_rows(win)
        # Delivery lines on the Реализация are left untouched - match only
        # the product rows to alif_items by fuzzy name (same idea as the
        # order side above). Verify unit totals line up.
        realiz_products = [ri for ri in realiz_items if _classify_order_item(ri["nomenklatura"]) != "delivery"]
        realiz_units = sum(ri["kolichestvo"] for ri in realiz_products)
        # A kit carries its companion nomenclatures here exactly as the order
        # does, so it legitimately shows more units than Alif Business - and
        # the companions must be priced the same way here too. Getting this
        # wrong is worse on this document than on the order: the order has
        # already been posted by now, so anything that fails here leaves the
        # application half-done.
        expected_units = alif_units + len(companion_lines)
        if abs(realiz_units - expected_units) > 0.5:
            raise RuntimeError(
                f"Реализация has {realiz_units} product unit(s) (across {len(realiz_products)} row(s)), "
                f"expected {expected_units} to match the order and Alif Business"
            )
        realiz_companion_ids = set()
        if companion_lines:
            realiz_main = max(realiz_products,
                              key=lambda ri: product_name_coverage(priced_pairs[0][0]["name"],
                                                                   ri["nomenklatura"]))
            realiz_companion_ids = {id(ri) for ri in realiz_products if ri is not realiz_main}
            if len(realiz_companion_ids) != len(companion_lines):
                raise RuntimeError(
                    f"Реализация has {len(realiz_companion_ids)} kit companion line(s) but the order had "
                    f"{len(companion_lines)}"
                )
            print(f"   kit on the Реализация: {realiz_main['nomenklatura']!r} takes the price, the other "
                  f"{len(realiz_companion_ids)} line(s) get {KIT_COMPANION_PRICE} each")
        for ritem in realiz_products:
            if id(ritem) in realiz_companion_ids:
                apply_price_fix_for_item(win, ritem, KIT_COMPANION_PRICE)
                continue
            alif_item, _srs, price_1c = find_alif_for(ritem["nomenklatura"])
            apply_price_fix_for_item(win, ritem, round(price_1c - companion_total, 2))
    except RuntimeError as e:
        print(f"!! {order_number}'s Реализация price fix failed: {e} - "
              "stopping (order price was already fixed, check by hand).")
        return _status("partial", "realizatsiya_price_fix_failed", detail=str(e))
    try:
        post_and_close(win, verify_closed_title_contains="Реализация")
    except RuntimeError as e:
        print(f"!! {order_number}'s Реализация did not post: {e} - order price already fixed, check by hand.")
        return _status("partial", "realizatsiya_post_failed", detail=str(e))

    # --- Эквайринговая операция ---
    print("-> back to Связанные документы, opening the order (row 0) for Эквайринговая операция")
    grid = _reopen_related_documents(win, grid, order_number)
    if grid is None:
        print(f"!! Could not reopen Связанные документы for {order_number} to create the Эквайринговая операция - "
              "both prices ARE already fixed, finish this one by hand.")
        return _status("partial", "reopen_failed_before_ekvayring")
    open_related_row(win, grid, row_index=0, expect_title_startswith="Заказ клиента")
    guard_no_surprise_dialog(win)

    try:
        create_ekvayring_operation(win, realizatsiya_date or "", terminal=terminal)
    except RuntimeError as e:
        print(f"!! {order_number}'s Эквайринговая операция setup failed: {e} - "
              "both prices are already fixed and posted; finish this by hand.")
        return _status("partial", "ekvayring_setup_failed", detail=str(e))

    if not ok("Эквайринговая операция form filled in correctly above (date and Терминал set)? Post & close it now?"):
        print("Left the Эквайринговая операция open without posting - "
              "both prices are already fixed and posted; finish this by hand.")
        return _status("partial", "user_declined_ekvayring")
    try:
        post_and_close(win, verify_closed_title_contains="Эквайринговая")
    except RuntimeError as e:
        print(f"!! {order_number}'s Эквайринговая операция did not post: {e} - "
              "both prices are already fixed and posted; finish this by hand.")
        return _status("partial", "ekvayring_post_failed", detail=str(e))

    # Each alif_item may span multiple sheet rows (its own kolichestvo shape) -
    # every one of those rows gets the SAME per-unit post-commission price
    # written to F/G (per instruction; F/G stays a per-unit column regardless
    # of the row's own Kolichestvo).
    write_targets = _alifshop_write_targets(priced_pairs, companion_rows)
    targets_desc = ", ".join(f"row {r}={amt}" for r, amt in write_targets)
    if not ok(f"Write sheet row(s) back ({targets_desc}, status=success, rows colored green)?"):
        print("Skipped writing back to the sheet - 1C side is fully done, update the sheet by hand.")
        return _status("partial", "user_declined_writeback")
    for r, amt in write_targets:
        alif_sheets_client.mark_success(ws, r, amt)
    print(f"=== Done: application {application_id} / {order_number} reconciled, sheet row(s) "
          f"{[r for r, _ in write_targets]} updated. ===")
    return _status("done", "reconciled", rows=[r for r, _ in write_targets],
                   detail=_combine_detail(status_note, split_note, _price_notes_suffix()))


def process_next_unchecked_order(ws, agreements=AGREEMENTS, auto_confirm=False, since_date=None,
                                  terminal=EKVAYRING_TERMINAL):
    row = find_next_unchecked_row(ws, since_date=since_date)
    if not row:
        since_note = f" on/after {since_date}" if since_date else ""
        print(f"Nothing to process - no unreconciled row{since_note}.")
        return
    if row.get("g_cell_was_text"):
        print(f"   (СуммаИМАН at sheet row {row['row']} was stored as text, not a number - converted it "
              "automatically so the diff formula in column J will compute correctly)")
    reconcile_order(row, agreements=agreements, auto_confirm=auto_confirm, ws=ws, terminal=terminal)


def process_all_unchecked_orders(ws, agreements=AGREEMENTS, since_date=None, terminal=EKVAYRING_TERMINAL):
    """Loop over EVERY unreconciled row from since_date forward (or the
    whole sheet if since_date is None), reconciling each in turn, and print
    a final summary report. Always runs with auto_confirm=True - there's no
    human to click through confirm() prompts for what could be dozens of
    orders (see main()'s --since handling, which requires --yes for this
    reason).

    Rows are never revisited: after each one (whatever the outcome - done,
    skipped, or partial), the search cursor advances past its sheet row
    number, so a row this run couldn't resolve doesn't get retried forever
    on every subsequent run - it's just reported and left for a human. Rows
    with no order number at all are already skipped for free by
    find_next_unchecked_row itself, same as always.

    A multi-item order can legitimately span several consecutive sheet rows
    sharing one order number (confirmed live) - a SUCCESSFUL multi-item
    reconcile writes Сумма 1С back to all of them at once, so the later
    ones naturally stop showing up as unreconciled and the loop skips past
    them for free. But if that order was instead skipped/partial/errored
    (no write-back happened), its other rows would otherwise resurface as
    if they were separate new orders, wastefully re-attempting the exact
    same 1C lookup and getting the exact same outcome - already-attempted
    order numbers are tracked this run and any further row for one is
    passed over without calling reconcile_order again.

    A genuinely unexpected exception from reconcile_order (anything it
    didn't already turn into a "partial" status itself) is caught here too,
    so one bad order can't take down the whole batch - it's counted as
    "error" and the loop moves on. Ctrl+C stops the loop early but still
    prints the report for everything done so far."""
    total = count_unchecked_rows(ws, since_date=since_date)
    since_note = f" on/after {since_date}" if since_date else ""
    print(f"=== Batch run: {total} unreconciled row(s){since_note} to process ===\n")

    win = get_window()
    done = []
    skipped = []
    partial = []
    errored = []
    report_ws = _open_report_ws()
    attempted_orders = {}  # order_number -> (reason, detail) from its first attempt this run
    # ONE shared instance passed to every reconcile_order call this run, so
    # find_order_smart only does the full Проведен/Текущее состояние reset
    # for the very first order - see its docstring.
    filter_session_state = {}
    cursor = FIRST_DATA_ROW
    i = 0

    try:
        while True:
            try:
                row = find_next_unchecked_row(ws, start_row=cursor, since_date=since_date)
            except Exception as e:  # noqa: BLE001 - a persistent Sheets API outage shouldn't crash the batch
                print(f"\n!! Could not fetch the next row from Google Sheets after retrying: {e!r} - "
                      "stopping the batch early, here's the report for what ran so far.\n")
                break
            if not row:
                break

            if row["order_number"] in attempted_orders:
                prior_status, prior_reason = attempted_orders[row["order_number"]]
                print(f"-- {row['order_number']} (sheet row {row['row']}) already attempted this run "
                      f"({prior_status}/{prior_reason}) - same order number, skipping without re-attempting 1C.\n")
                cursor = row["row"] + 1
                continue

            i += 1
            print(f"[{i}/{total}] --- {row['order_number']} (sheet row {row['row']}) ---")
            if row.get("g_cell_was_text"):
                print(f"   (СуммаИМАН at sheet row {row['row']} was stored as text, not a number - converted "
                      "it automatically so the diff formula in column J will compute correctly)")

            rep = {}
            try:
                result = reconcile_order(row, agreements=agreements, auto_confirm=True, ws=ws, terminal=terminal,
                                          session_state=filter_session_state, report=rep)
            except Exception as e:  # noqa: BLE001 - deliberately broad, see docstring
                print(f"!! Unexpected error on {row['order_number']}: {e!r} - moving on to the next order.")
                result = _status("error", "unexpected_exception", detail=repr(e))

            try:
                close_back_to_zakazy_klientov(win)
            except Exception as e:  # noqa: BLE001 - tab cleanup failing shouldn't kill the batch either
                print(f"   (tab cleanup after {row['order_number']} hit an error: {e!r} - continuing anyway)")

            _log_report_row(report_ws, rep, row["order_number"], terminal, "/".join(agreements), result)
            attempted_orders[row["order_number"]] = (result.get("status"), result.get("reason", "?"))

            entry = (row["order_number"], row["row"], result.get("reason", "?"), result.get("detail", ""))
            status = result.get("status")
            if status == "done":
                done.append(entry)
            elif status == "partial":
                partial.append(entry)
            elif status == "error":
                errored.append(entry)
            else:
                skipped.append(entry)

            cursor = row["row"] + 1
            print()
            telegram_logger.flush()
    except KeyboardInterrupt:
        print("\n!! Interrupted - stopping the batch early, here's the report for what ran so far.\n")

    with telegram_logger.batch():
        print("=== Batch report ===")
        print(f"Rows considered: {i} of {total} found at start (the sheet may have changed underneath if this "
              f"took a while)")
        print(f"Done: {len(done)}")
        print(f"Skipped (safe to ignore, nothing changed): {len(skipped)}")
        print(f"Partial (SOMETHING WAS POSTED - needs a human to finish/check!): {len(partial)}")
        print(f"Errored unexpectedly: {len(errored)}")

        def _print_group(title, entries):
            if not entries:
                return
            print(f"\n{title}:")
            for order_number, row_num, reason, detail in entries:
                suffix = f" ({detail})" if detail else ""
                print(f"  {order_number} (row {row_num}): {reason}{suffix}")

        _print_group("Skipped", skipped)
        _print_group("!! PARTIAL - needs attention", partial)
        _print_group("!! ERRORED - needs attention", errored)

    return {"done": done, "skipped": skipped, "partial": partial, "errored": errored}


def process_all_alifshop_applications(agreements=ALIFSHOP_AGREEMENTS, since_date=None, terminal=ALIFSHOP_TERMINAL):
    """Alif Shop's sibling to process_all_unchecked_orders above - loop over
    EVERY unresolved application (status blank = never checked, or
    'waiting' = checked before but not finalized yet - both get
    (re)checked) from since_date forward, reconciling each in turn, and
    print the same shape of final summary report. Always auto_confirm=True,
    same reasoning as the IMAN batch loop (no human to click through
    confirm() prompts for dozens of applications).

    A multi-item application spans several consecutive sheet rows sharing
    one application id (confirmed live: application 1000002) - already-
    attempted application ids are tracked this run so a later row for the
    same id is skipped rather than wastefully re-attempting the exact same
    Alif Business + 1C lookup (same reasoning attempted_orders has in the
    IMAN loop).

    Only ONE Соглашение exists for Alif Shop (see ALIFSHOP_AGREEMENTS) -
    unlike IMAN's two, so there's no "wrong agreement" to retry past, but
    find_order_under_agreements would still redo the FULL Проведен/Текущее
    состояние reset on every single application otherwise (i==0 always,
    with only one agreement) - reconcile_alifshop_application's own
    session_state parameter (routed through find_order_smart, see its
    docstring) avoids that across this whole run, same as the IMAN loop's
    filter_session_state."""
    ws = alif_sheets_client.get_alif_worksheet()
    total = alif_sheets_client.count_unresolved_applications(ws, since_date=since_date)
    since_note = f" on/after {since_date}" if since_date else ""
    print(f"=== Batch run: {total} unresolved Alif Shop row(s){since_note} to process ===\n")

    win = get_window()
    done = []
    skipped = []
    partial = []
    errored = []
    report_ws = _open_report_ws()
    attempted_applications = {}  # application_id -> (status, reason) from its first attempt this run
    filter_session_state = {}
    cursor = alif_sheets_client.FIRST_DATA_ROW
    i = 0

    try:
        while True:
            try:
                row = alif_sheets_client.find_next_unresolved_application(ws, start_row=cursor, since_date=since_date)
            except Exception as e:  # noqa: BLE001 - a persistent Sheets API outage shouldn't crash the batch
                print(f"\n!! Could not fetch the next row from Google Sheets after retrying: {e!r} - "
                      "stopping the batch early, here's the report for what ran so far.\n")
                break
            if not row:
                break

            application_id = row["application_id"]
            if application_id in attempted_applications:
                prior_status, prior_reason = attempted_applications[application_id]
                print(f"-- application {application_id} (sheet row {row['row']}) already attempted this run "
                      f"({prior_status}/{prior_reason}) - same application, skipping without re-attempting.\n")
                cursor = row["row"] + 1
                continue

            i += 1
            print(f"[{i}/{total}] --- application {application_id} (sheet row {row['row']}) ---")

            rep = {}
            try:
                result = reconcile_alifshop_application(application_id, auto_confirm=True, agreements=agreements,
                                                          terminal=terminal, session_state=filter_session_state,
                                                          report=rep)
            except Exception as e:  # noqa: BLE001 - deliberately broad, see process_all_unchecked_orders's docstring
                print(f"!! Unexpected error on application {application_id}: {e!r} - moving on to the next one.")
                result = _status("error", "unexpected_exception", detail=repr(e))

            try:
                close_back_to_zakazy_klientov(win)
            except Exception as e:  # noqa: BLE001 - tab cleanup failing shouldn't kill the batch either
                print(f"   (tab cleanup after application {application_id} hit an error: {e!r} - continuing anyway)")

            _log_report_row(report_ws, rep, rep.get("contract", ""), terminal, agreements[0], result)
            attempted_applications[application_id] = (result.get("status"), result.get("reason", "?"))

            entry = (application_id, row["row"], result.get("reason", "?"), result.get("detail", ""))
            status = result.get("status")
            if status == "done":
                done.append(entry)
            elif status == "partial":
                partial.append(entry)
            elif status == "error":
                errored.append(entry)
            else:
                skipped.append(entry)

            cursor = row["row"] + 1
            print()
            telegram_logger.flush()
    except KeyboardInterrupt:
        print("\n!! Interrupted - stopping the batch early, here's the report for what ran so far.\n")

    with telegram_logger.batch():
        print("=== Batch report ===")
        print(f"Rows considered: {i} of {total} found at start (the sheet may have changed underneath if this "
              f"took a while)")
        print(f"Done: {len(done)}")
        print(f"Skipped (safe to ignore, nothing changed): {len(skipped)}")
        print(f"Partial (SOMETHING WAS POSTED - needs a human to finish/check!): {len(partial)}")
        print(f"Errored unexpectedly: {len(errored)}")

        def _print_group(title, entries):
            if not entries:
                return
            print(f"\n{title}:")
            for application_id, row_num, reason, detail in entries:
                suffix = f" ({detail})" if detail else ""
                print(f"  application {application_id} (row {row_num}): {reason}{suffix}")

        _print_group("Skipped", skipped)
        _print_group("!! PARTIAL - needs attention", partial)
        _print_group("!! ERRORED - needs attention", errored)

    return {"done": done, "skipped": skipped, "partial": partial, "errored": errored}


# ---------------------------------------------------------- Alif Moliya flow

def _parse_grid_date_cell(text):
    """Parse one Заказы клиентов grid row's Дата cell - confirmed live, 1C
    shows a full 'DD.MM.YYYY' for most rows but only a bare 'H:MM' time (no
    date at all) for very recent ones. Per explicit instruction, bare-time
    rows are NOT yet confirmed and must NOT be treated as "today" - they're
    excluded from processing entirely (see read_open_orders_with_dates)
    rather than guessed at. Returns None for a bare-time cell."""
    text = text.strip()
    if re.fullmatch(r"\d{2}\.\d{2}\.\d{4}", text):
        return datetime.datetime.strptime(text, "%d.%m.%Y").date()
    if re.fullmatch(r"\d{1,2}:\d{2}", text):
        return None
    raise RuntimeError(f"Unrecognized grid Дата cell text: {text!r} - check the grid by hand.")


def read_open_orders_with_dates(win):
    """Read every order currently visible in the Заказы клиентов grid
    (after Соглашение/Проведен/Текущее состояние filters are already set) -
    returns (confirmed, unconfirmed):
      - confirmed: [(order_number, date), ...] for rows with a real
        DD.MM.YYYY date, in grid order.
      - unconfirmed: [order_number, ...] for rows showing only a bare time
        (no date) - per explicit instruction, these orders aren't
        confirmed yet and must be excluded from processing/report-range
        calculations entirely, not guessed to be "today".
    Confirmed live: the grid's virtualized 'Custom' cells for Номер and
    Дата come back in matching row order (same convention already relied
    on throughout this file, e.g. read_item_rows pairing price/
    nomenklatura cells by position) - this is small enough (confirmed
    live: Alif Moliya's whole 'Все открытые' list fits on one screen, no
    scrolling needed) that reading only the rendered rows is sufficient,
    unlike a full historical scan."""
    cells = [strip_ui_markup(c.window_text()).strip() for c in win.descendants(control_type="Custom")]
    order_texts = [c[: -len(" Номер")] for c in cells if c.endswith(" Номер")]
    date_texts = [c[: -len(" Дата")] for c in cells if c.endswith(" Дата")]
    if len(order_texts) != len(date_texts):
        raise RuntimeError(
            f"Grid has {len(order_texts)} Номер cell(s) but {len(date_texts)} Дата cell(s) - can't pair "
            "them reliably, check by hand."
        )
    confirmed, unconfirmed = [], []
    for order_number, date_text in zip(order_texts, date_texts):
        date = _parse_grid_date_cell(date_text)
        if date is None:
            unconfirmed.append(order_number)
        else:
            confirmed.append((order_number, date))
    return confirmed, unconfirmed


def _match_report_row_for_order(report_rows, order_date, item_nomenklatura, name_field="sku"):
    """Find the downloaded report row for a single-item order (Alif Moliya's
    Tovarlar report or Uzum Nasiya's CSV - see alif_client.
    parse_tovarlar_report / uzum_client.parse_report): filter to rows
    matching order_date exactly, then fuzzy-match by product name against
    the 1C item's own Номенклатура, same product_name_coverage approach
    used everywhere else in this file. name_field is which key in each row
    dict holds the 1C-matching product name - "sku" for Alif Moliya
    (distinct from its more generic "product_name" column, see that
    module's docstring), "product_name" for Uzum Nasiya (which has no such
    generic/specific split - confirmed live its one product name column
    already matches 1C directly). Returns (best_row, coverage) or
    (None, 0.0) if no row exists for that date at all."""
    same_day = [r for r in report_rows if r["date"] == order_date]
    if not same_day:
        return None, 0.0
    same_day.sort(key=lambda r: product_name_coverage(r[name_field], item_nomenklatura), reverse=True)
    best = same_day[0]
    return best, product_name_coverage(best[name_field], item_nomenklatura)


def _classify_order_item(nomenklatura):
    """A 1C Товары line is a delivery charge ("Доставка ..."), a free gift
    ("... (Подарок)", priced ~1 сум - the source never finances it), or a
    real product. Only real products are matched to (and priced from) the
    source; delivery and gift lines keep their existing 1C price (the
    delivery line's price is instead SUBTRACTED from the product's source
    total - see _split_delivery_note / build_alifmoliya_pairs)."""
    n = (nomenklatura or "").lower()
    if "доставка" in n:
        return "delivery"
    if "подарок" in n:
        return "gift"
    return "product"


def _split_delivery_note(order_number, source_total, delivery_items, gift_items=()):
    """Build the human note that goes in the report sheet's Comment column
    (NOT a Telegram tag) when an order had its delivery split out and/or a
    gift left as-is. Returns "" if there was neither."""
    bits = []
    if delivery_items:
        d_total = sum((d["tsena"] or 0) for d in delivery_items)
        names = ", ".join(d["nomenklatura"] for d in delivery_items)
        bits.append(f"delivery [{names}] = {d_total} split out of source total {source_total} "
                    f"(product set to {round((source_total or 0) - d_total, 2)})")
    for g in gift_items:
        bits.append(f"gift left as-is: {g['nomenklatura']} = {g['tsena']}")
    return "; ".join(bits)


def build_alifmoliya_pairs(order_items, report_rows, order_date):
    """Match a (possibly multi-item) Alif Moliya order to its Tovarlar report
    rows and work out each 1C item's target price. Returns a list of
    (order_item, plan) pairs in order_items order, where plan is a dict:
        target_price    - price to set on this line
        source_name     - the report SKU it matched (products) / its own name
        report_price    - the matched report row's raw price (products only)
        skip_name_check - True for gift/delivery (no report row to compare)
        skip_price_fix  - True for gift/delivery (leave their price alone)
        kind            - 'product' | 'gift' | 'delivery'
    plus the chosen report application_id (for logging), the delivery total
    and the gift total that were folded out of the first product (0 if none).

    Multi-item orders show up in the report as several rows sharing ONE
    application_id on the order's date (confirmed live), so this picks the
    report application whose line count equals the order's product-line
    count and whose SKUs best match, then assigns each product line to a
    distinct report row greedily by name coverage.

    Delivery split (per instruction): the report prices a delivery order's
    product at the FULL financed total, but 1C splits it into a product line
    + a Доставка line, so the product line must be (report total − delivery
    line's current price) to make product + delivery == the report total.
    Only handled for a single product line + delivery; a multi-product order
    that also has delivery is ambiguous and raises for manual handling.

    Raises MultiItemMatchError for any shape it can't handle (caller treats
    that as a clean skip, nothing changed)."""
    kinds = [_classify_order_item(it["nomenklatura"]) for it in order_items]
    products = [it for it, k in zip(order_items, kinds) if k == "product"]
    delivery = [it for it, k in zip(order_items, kinds) if k == "delivery"]

    if not products:
        raise MultiItemMatchError("order has no real product line (only delivery/gift) - nothing to price")

    same_day = [r for r in report_rows if r["date"] == order_date]
    if not same_day:
        raise MultiItemMatchError(f"no Tovarlar report row for {order_date}")

    # Match by TOTAL UNITS not row count: 1C can split the same product
    # across several rows (each with its own Serial number and its own
    # Количество), and the report can too - as long as the unit totals
    # agree, the pairing works. Confirmed live (SA-26/00101, alif moliya app
    # 1000005): 1C has 3 rows of Samsung A07 with Количество 1+2+1=4, the
    # report has correspondingly 4 units for the same SKU on that date.
    groups = {}
    for r in same_day:
        groups.setdefault(r["application_id"], []).append(r)
    product_units = sum(p["kolichestvo"] for p in products)

    def total_units(rs):
        return sum((r["kolichestvo"] or 0) for r in rs)

    candidates = [gid for gid, rs in groups.items() if abs(total_units(rs) - product_units) < 0.5]
    if not candidates:
        raise MultiItemMatchError(
            f"no report application on {order_date} has {product_units} unit(s) to match the order's "
            f"{product_units} product unit(s) - handle by hand")

    def assign(app_rows):
        """For each 1C product line (which may have Количество>1), draw
        enough units from the best-fuzzy-matching report row(s) to cover
        its own Количество. All units drawn for one 1C line must agree on
        per-unit price (they should, being units of the same SKU) - if not,
        this candidate application is rejected. Returns (item_id->matched
        report row, min coverage across the assignment) or (None, 0.0) if
        the candidate can't be satisfied."""
        pool = [{"r": r, "left": r["kolichestvo"] or 0} for r in app_rows]
        assigned = {}
        for it in products:
            need = it["kolichestvo"]
            pool.sort(key=lambda x: (product_name_coverage(x["r"]["sku"], it["nomenklatura"]), x["left"]),
                      reverse=True)
            collected = []
            remaining_need = need
            for x in pool:
                if x["left"] <= 0:
                    continue
                take = min(x["left"], remaining_need)
                collected.append((x, take))
                remaining_need -= take
                if remaining_need <= 0.5:
                    break
            if remaining_need > 0.5:
                return None, 0.0
            prices = {x["r"]["price"] for x, _ in collected}
            if len(prices) != 1:
                return None, 0.0
            for x, take in collected:
                x["left"] -= take
            assigned[id(it)] = collected[0][0]["r"]
        min_cov = min(product_name_coverage(assigned[id(it)]["sku"], it["nomenklatura"]) for it in products)
        return assigned, min_cov

    scored = []
    for gid in candidates:
        result, cov = assign(groups[gid])
        if result is not None:
            scored.append((gid, result, cov))
    if not scored:
        raise MultiItemMatchError(
            f"found {len(candidates)} report application(s) with the right unit total on {order_date}, but "
            "none could be assigned cleanly (mismatched per-unit prices, or product name coverage too low) "
            "- handle by hand")
    best_app, item_to_row, _ = max(scored, key=lambda x: x[2])

    gift = [it for it, k in zip(order_items, kinds) if k == "gift"]
    delivery_total = sum(d["tsena"] or 0 for d in delivery)
    gift_total = sum(g["tsena"] or 0 for g in gift)
    if delivery and len(products) != 1:
        raise MultiItemMatchError(
            f"order has a Доставка line plus {len(products)} product lines - the delivery-split rule only "
            "handles exactly one product line, handle this one by hand")

    # Fold the delivery line's price AND any gift line's price OUT of the
    # FIRST product line, so all the lines together still sum to the
    # products' report total while the gift/delivery lines keep their own
    # prices. Confirmed live (QQ-26/00102: 2 products + 1-сум gift): report
    # totals 500000 + 400000, so 1C should read 499999 + 400000 + 1 =
    # 900000 - same idea as the single-item bonus bundle in
    # resolve_bonus_item_prices, generalized to multi-item here.
    first_product_id = id(products[0])
    deduction = delivery_total + gift_total

    pairs = []
    for it, k in zip(order_items, kinds):
        if k == "product":
            r = item_to_row[id(it)]
            target = r["price"]
            if id(it) == first_product_id and deduction:
                target = round((r["price"] or 0) - deduction, 2)
            pairs.append((it, {"target_price": target, "source_name": r["sku"], "report_price": r["price"],
                               "skip_name_check": False, "skip_price_fix": False, "kind": "product"}))
        else:
            # gift / delivery: leave the line's existing price untouched (a
            # no-op re-apply) - their value was already folded out of the
            # first product above - but keep it in the pair list so the item
            # counts still line up with the Реализация's own grid.
            pairs.append((it, {"target_price": it["tsena"], "source_name": it["nomenklatura"],
                               "report_price": None, "skip_name_check": True, "skip_price_fix": True, "kind": k}))
    return pairs, best_app, delivery_total, gift_total


def reconcile_alifmoliya_order(order_number, report_rows=None, auto_confirm=False, agreements=ALIFMOLIYA_AGREEMENTS,
                                terminal=ALIFMOLIYA_TERMINAL, session_state=None, report=None):
    """Core reconcile logic for one Alif Moliya order - a sibling to
    reconcile_order (IMAN Sverka) and reconcile_alifshop_application above,
    sharing the same low-level 1C helpers, but there's no Google Sheet
    driving this AT ALL (unlike either of those):
      - the set of orders to reconcile is simply whatever's currently open
        (Проведен=Да, Текущее состояние=Все открытые) under Соглашение=
        Alif Moliya in 1C itself (see read_open_orders_with_dates) - once
        an order's Реализация posts, 1C naturally drops it out of "Все
        открытые" on its own, so nothing needs to be written back anywhere
        to mark it done.
      - there's no per-order price lookup either (Alif Shop has one via
        application id; IMAN's sheet has "Сумма ИМАН" directly) - instead,
        Alif Moliya publishes a downloadable "Tovarlar" Excel report (see
        alif_client.get_tovarlar_report/parse_tovarlar_report) covering
        whatever date range you ask for, and the order's OWN date (read
        off its Заказы клиентов grid row) is what's used to find its
        matching report row, then fuzzy-matched by product name the same
        way everything else in this file is.

    report_rows, if given, is a PRE-FETCHED Tovarlar report (see
    process_all_alifmoliya_orders, which fetches ONE report covering every
    open order's date range up front rather than one per order) - if None
    (the standalone --order case, see main()), this fetches its own
    narrow report covering just this order's own date ± 1 day.

    Only single-item, quantity-1 orders are handled so far (same starting
    point every other flow in this file began at) - anything else stops
    rather than guessing.

    session_state, if given, is passed straight to find_order_smart, same
    reasoning as reconcile_order/reconcile_alifshop_application - leave it
    None for a one-off single-order run."""
    def ok(prompt):
        return True if auto_confirm else confirm(prompt)

    rep = report if report is not None else {}  # filled as data is discovered (see report_sheet.py)
    _reset_price_notes()

    print(f"=== Alif Moliya order {order_number} ===")
    win = get_window()
    focus(win)
    open_zakazy_klientov(win)

    agreement = find_order_smart(win, order_number, agreements, session_state if session_state is not None else {})
    if not agreement:
        print(f"!! {order_number} not found in 1C under any of {agreements} - "
              "this is the rare exceptional case, skipping (needs manual handling).")
        return _status("skipped", "not_found_in_1c")

    confirmed_orders, unconfirmed_orders = read_open_orders_with_dates(win)
    order_date = dict(confirmed_orders).get(order_number)
    if order_date is None:
        if order_number in unconfirmed_orders:
            print(f"!! {order_number} shows only a time (no date) in 1C's grid - not confirmed yet, "
                  "stopping (deal with this one once it has a real date).")
            return _status("skipped", "not_yet_confirmed")
        print(f"!! {order_number} not found among currently open Alif Moliya orders (already closed/"
              "reconciled, or an unexpected Текущее состояние) - can't determine its date, stopping.")
        return _status("skipped", "not_currently_open")

    result = open_and_read_related_documents(win, order_number)
    if result is None:
        print(f"!! Related-documents check failed for {order_number} (see messages above) - "
              "stopping, nothing was changed.")
        return _status("skipped", "related_docs_check_failed")
    if result.get("ekvayring"):
        print(f"=== {order_number} already has an Эквайринговая операция "
              f"({result['ekvayring']}) - reconciled by an earlier run, nothing to do. ===")
        return _status("skipped", "already_has_ekvayring", detail=result["ekvayring"])
    realizatsiya_date, grid = result["date"], result["grid"]
    rep["order_date"] = result.get("order_date") or ""

    open_related_row(win, grid, row_index=0, expect_title_startswith="Заказ клиента")
    guard_no_surprise_dialog(win)
    click_tabitem(win, "Товары")
    time.sleep(0.5)
    focus(win)

    order_items = read_item_rows(win)

    # Report generation is somewhat expensive (a real browser session,
    # several seconds of Yaratish/download) - deliberately done AFTER the
    # cheaper related-docs check above, but BEFORE building the item pairs
    # (multi-item matching needs every report row for this date up front).
    if report_rows is None:
        report_start = order_date - datetime.timedelta(days=1)
        report_end = order_date + datetime.timedelta(days=1)
        print(f"-> {order_number}'s date is {order_date} - generating a Tovarlar report for "
              f"{report_start}..{report_end}")
        with sync_playwright() as p:
            report_path = alif_client.get_tovarlar_report(p, report_start, report_end)
        report_rows = alif_client.parse_tovarlar_report(report_path)
        print(f"   report downloaded and parsed: {len(report_rows)} row(s)")

    try:
        pairs, app_id, delivery_total, gift_total = build_alifmoliya_pairs(order_items, report_rows, order_date)
    except MultiItemMatchError as e:
        print(f"!! {order_number}: {e} - stopping without changes.")
        return _status("skipped", "unhandled_shape", detail=str(e))

    # Report fields: first product line in the primary columns, further
    # product lines' set prices in Item 2..5, a gift line in Bonus. price
    # before/after are the first product line's Tsena before/after this run;
    # price_source is what the report financed (the FULL total for a
    # delivery order, which price_after then has the delivery split out of).
    product_pairs = [(it, pl) for it, pl in pairs if pl["kind"] == "product"]
    gift_pairs = [(it, pl) for it, pl in pairs if pl["kind"] == "gift"]
    rep["id"] = app_id
    rep["name_1c"] = product_pairs[0][0]["nomenklatura"]
    rep["quantity"] = product_pairs[0][0]["kolichestvo"]
    rep["price_before"] = product_pairs[0][0]["tsena"]
    rep["name_source"] = product_pairs[0][1]["source_name"]
    rep["price_source"] = product_pairs[0][1]["report_price"]
    rep["price_after"] = product_pairs[0][1]["target_price"]
    for n, (_, pl) in enumerate(product_pairs[1:5], start=2):
        rep[f"item{n}_price"] = pl["target_price"]
    if gift_pairs:
        gi, gpl = gift_pairs[0]
        rep["bonus"] = f"{gi['nomenklatura']} = {gpl['target_price']}"

    if len(pairs) > 1:
        extras = []
        if delivery_total:
            extras.append(f"delivery {delivery_total} split out")
        if gift_total:
            extras.append(f"{len(gift_pairs)} gift ({gift_total}) folded out of first product")
        print(f"   {order_number} matched to report application {app_id} "
              f"({len(product_pairs)} product line(s){', ' + ', '.join(extras) if extras else ''}):")
        for it, pl in pairs:
            print(f"     [{it['index']}] {it['nomenklatura']!r} <-> {pl['source_name']!r} "
                  f"= {pl['target_price']} ({pl['kind']})")

    for it, pl in pairs:
        if pl["skip_name_check"]:
            continue
        if not names_match_or_confirm(pl["source_name"], it["nomenklatura"]):
            print(f"!! Product name mismatch on {order_number} - stopping without changes.")
            return _status("skipped", "name_mismatch")

    price_summary = ", ".join(str(pl["target_price"]) for _, pl in pairs)
    if not ok(f"Set {order_number}'s order price(s) to [{price_summary}] and post it?"):
        print("Left the order open without changing its price - stopping.")
        return _status("skipped", "user_declined_order_price")
    try:
        for it, pl in pairs:
            if pl["skip_price_fix"]:
                continue  # gift / delivery line - leave its price untouched
            apply_price_fix_for_item(win, it, pl["target_price"])
    except RuntimeError as e:
        print(f"!! {order_number}'s order price fix failed verification: {e} - stopping without posting.")
        return _status("skipped", "order_price_fix_failed", detail=str(e))
    try:
        post_and_close(win, verify_closed_title_contains="Заказ клиента")
    except RuntimeError as e:
        print(f"!! {order_number}'s order did not post: {e} - check by hand.")
        return _status("partial", "order_post_failed", detail=str(e))

    print("-> back to Связанные документы, opening Реализация (row 1)")
    grid = _reopen_related_documents(win, grid, order_number)
    if grid is None:
        print(f"!! Could not reopen Связанные документы for {order_number} after fixing the order price - "
              "stopping, check by hand (order price WAS already changed).")
        return _status("partial", "reopen_failed_after_order_post")

    open_related_row(win, grid, row_index=1, expect_title_startswith="Реализация")
    guard_no_surprise_dialog(win)
    click_tabitem(win, "Товары")
    time.sleep(0.5)
    focus(win)
    if not ok("Set the Реализация's price(s) the same way, and post it?"):
        print("Left the Реализация untouched (order price WAS already changed) - finish this by hand.")
        return _status("partial", "user_declined_realizatsiya_price")
    try:
        realiz_items = read_item_rows(win)
        if len(realiz_items) != len(pairs):
            raise RuntimeError(f"Реализация has {len(realiz_items)} item row(s), expected {len(pairs)}")
        # Re-pair each Реализация row to the order pair with the closest 1C
        # name (its grid's row order isn't guaranteed to match the order's),
        # same approach as reconcile_order's multi-item Реализация fix.
        remaining = list(pairs)
        for ritem in realiz_items:
            remaining.sort(key=lambda p: product_name_coverage(p[0]["nomenklatura"], ritem["nomenklatura"]),
                           reverse=True)
            _, pl = remaining.pop(0)
            if pl["skip_price_fix"]:
                continue
            apply_price_fix_for_item(win, ritem, pl["target_price"])
    except RuntimeError as e:
        print(f"!! {order_number}'s Реализация price fix failed: {e} - "
              "stopping (order price was already fixed, check by hand).")
        return _status("partial", "realizatsiya_price_fix_failed", detail=str(e))
    try:
        post_and_close(win, verify_closed_title_contains="Реализация")
    except RuntimeError as e:
        print(f"!! {order_number}'s Реализация did not post: {e} - order price already fixed, check by hand.")
        return _status("partial", "realizatsiya_post_failed", detail=str(e))

    print("-> back to Связанные документы, opening the order (row 0) for Эквайринговая операция")
    grid = _reopen_related_documents(win, grid, order_number)
    if grid is None:
        print(f"!! Could not reopen Связанные документы for {order_number} to create the Эквайринговая операция - "
              "both prices ARE already fixed, finish this one by hand.")
        return _status("partial", "reopen_failed_before_ekvayring")
    open_related_row(win, grid, row_index=0, expect_title_startswith="Заказ клиента")
    guard_no_surprise_dialog(win)

    try:
        create_ekvayring_operation(win, realizatsiya_date or "", terminal=terminal)
    except RuntimeError as e:
        print(f"!! {order_number}'s Эквайринговая операция setup failed: {e} - "
              "both prices are already fixed and posted; finish this by hand.")
        return _status("partial", "ekvayring_setup_failed", detail=str(e))

    if not ok("Эквайринговая операция form filled in correctly above (date and Терминал set)? Post & close it now?"):
        print("Left the Эквайринговая операция open without posting - "
              "both prices are already fixed and posted; finish this by hand.")
        return _status("partial", "user_declined_ekvayring")
    try:
        post_and_close(win, verify_closed_title_contains="Эквайринговая")
    except RuntimeError as e:
        print(f"!! {order_number}'s Эквайринговая операция did not post: {e} - "
              "both prices are already fixed and posted; finish this by hand.")
        return _status("partial", "ekvayring_post_failed", detail=str(e))

    priced = ", ".join(str(pl["target_price"]) for _, pl in pairs if not pl["skip_price_fix"])
    detail_bits = []
    if delivery_total:
        detail_bits.append(f"delivery {delivery_total} split out of the report total")
    if len(product_pairs) > 1:
        detail_bits.append(f"{len(product_pairs)} product lines")
    if gift_total:
        detail_bits.append(f"{len(gift_pairs)} gift ({gift_total}) folded out of first product")
    print(f"=== Done: {order_number} reconciled at price(s) [{priced}] - no sheet to write back (Alif Moliya "
          "has none, 1C's own 'Все открытые' filter is the source of truth). ===")
    return _status("done", "reconciled", detail=_combine_detail("; ".join(detail_bits), _price_notes_suffix()))


def process_all_alifmoliya_orders(agreements=ALIFMOLIYA_AGREEMENTS, terminal=ALIFMOLIYA_TERMINAL, since_date=None):
    """Batch mode for Alif Moliya (see main()'s --profile alif-moliya
    handling). Unlike the IMAN/Alif Shop batch loops, there's no Google
    Sheet cursor to scan - the "unprocessed" set is simply every order
    currently open under Соглашение=Alif Moliya (see
    read_open_orders_with_dates), optionally narrowed further by
    since_date. Fetches ONE Tovarlar report up front covering every open
    order's own date range (min - 1 day .. max + 1 day) rather than one
    per order (see reconcile_alifmoliya_order's report_rows parameter) -
    much cheaper than regenerating/downloading a report per order.

    since_date is OPTIONAL here (unlike the other profiles' --since, which
    is required for batch mode) - "Все открытые" already bounds this to
    currently-open orders only, there's no unlimited historical backlog to
    worry about looping over."""
    win = get_window()
    focus(win)
    open_zakazy_klientov(win)
    # "Все открытые" passed explicitly: for this profile the filter IS the
    # definition of what to process (there's no sheet saying which orders to
    # look at), unlike an order lookup, where it would hide already-finished
    # orders - see clear_tekushee_sostoyanie.
    set_filters(win, agreements[0], tekushee_sostoyanie="Все открытые")
    # Batch mode wants EVERY open order under this Соглашение - a stale term
    # left in the quick-search box (from a prior manual search or a
    # single-order run) would silently filter the grid down, so clear it
    # first (see clear_search - confirmed live this exact case reported
    # "no open orders" when there really were several).
    clear_search(win)

    orders, unconfirmed_orders = read_open_orders_with_dates(win)
    if unconfirmed_orders:
        print(f"   ({len(unconfirmed_orders)} order(s) show only a time (no date) in 1C's grid - not "
              f"confirmed yet, excluding from this run: {', '.join(unconfirmed_orders)})")
    if since_date is not None:
        orders = [(o, d) for o, d in orders if d >= since_date]
    since_note = f" on/after {since_date}" if since_date else ""
    if not orders:
        print(f"No open (confirmed) Alif Moliya orders{since_note} to process.")
        return {"done": [], "skipped": [], "partial": [], "errored": []}

    dates = [d for _, d in orders]
    report_start = min(dates) - datetime.timedelta(days=1)
    report_end = max(dates) + datetime.timedelta(days=1)
    print(f"=== Batch run: {len(orders)} open Alif Moliya order(s){since_note}, dated {min(dates)}..{max(dates)} - "
          f"generating a Tovarlar report for {report_start}..{report_end} ===\n")

    with sync_playwright() as p:
        report_path = alif_client.get_tovarlar_report(p, report_start, report_end)
    report_rows = alif_client.parse_tovarlar_report(report_path)
    print(f"   report downloaded and parsed: {len(report_rows)} row(s)\n")

    done, skipped, partial, errored = [], [], [], []
    report_ws = _open_report_ws()
    filter_session_state = {}

    try:
        for i, (order_number, order_date) in enumerate(orders, 1):
            print(f"[{i}/{len(orders)}] --- {order_number} ({order_date}) ---")
            rep = {}
            try:
                result = reconcile_alifmoliya_order(order_number, report_rows=report_rows, auto_confirm=True,
                                                     agreements=agreements, terminal=terminal,
                                                     session_state=filter_session_state, report=rep)
            except Exception as e:  # noqa: BLE001 - deliberately broad, see process_all_unchecked_orders's docstring
                print(f"!! Unexpected error on {order_number}: {e!r} - moving on to the next order.")
                result = _status("error", "unexpected_exception", detail=repr(e))

            try:
                close_back_to_zakazy_klientov(win)
            except Exception as e:  # noqa: BLE001 - tab cleanup failing shouldn't kill the batch either
                print(f"   (tab cleanup after {order_number} hit an error: {e!r} - continuing anyway)")

            _log_report_row(report_ws, rep, order_number, terminal, agreements[0], result)
            entry = (order_number, result.get("reason", "?"), result.get("detail", ""))
            status = result.get("status")
            if status == "done":
                done.append(entry)
            elif status == "partial":
                partial.append(entry)
            elif status == "error":
                errored.append(entry)
            else:
                skipped.append(entry)
            print()
            telegram_logger.flush()
    except KeyboardInterrupt:
        print("\n!! Interrupted - stopping the batch early, here's the report for what ran so far.\n")

    with telegram_logger.batch():
        print("=== Batch report ===")
        print(f"Done: {len(done)}")
        print(f"Skipped (safe to ignore, nothing changed): {len(skipped)}")
        print(f"Partial (SOMETHING WAS POSTED - needs a human to finish/check!): {len(partial)}")
        print(f"Errored unexpectedly: {len(errored)}")

        def _print_group(title, entries):
            if not entries:
                return
            print(f"\n{title}:")
            for order_number, reason, detail in entries:
                suffix = f" ({detail})" if detail else ""
                print(f"  {order_number}: {reason}{suffix}")

        _print_group("Skipped", skipped)
        _print_group("!! PARTIAL - needs attention", partial)
        _print_group("!! ERRORED - needs attention", errored)

    return {"done": done, "skipped": skipped, "partial": partial, "errored": errored}


# ----------------------------------------------------------- Uzum Nasiya flow

def reconcile_uzumnasiya_order(order_number, report_rows=None, auto_confirm=False, agreements=UZUM_AGREEMENTS,
                                terminal=UZUM_TERMINAL, session_state=None, report=None):
    """Core reconcile logic for one Uzum Nasiya order - a near-exact sibling
    of reconcile_alifmoliya_order (see its docstring for the shared
    reasoning: no Google Sheet at all, the set of orders to reconcile is
    simply whatever's currently open under this Соглашение in 1C itself,
    see read_open_orders_with_dates). The only real difference is WHERE the
    per-order price comes from: Uzum Nasiya publishes a downloadable CSV
    ("Отчет по филиалам с датами отмененных" - see uzum_client.
    get_cancel_dates_report/parse_report) instead of Alif Moliya's Excel
    Tovarlar report, and its price column ("Стоимость одного товара с НДС
    и комиссией") is already commission-inclusive - no separate commission
    subtraction step like Alif Shop's.

    report_rows, if given, is a PRE-FETCHED CSV report (see
    process_all_uzumnasiya_orders, which fetches ONE report covering every
    open order's date range up front) - if None (the standalone --order
    case, see main()), this fetches its own narrow report covering just
    this order's own date ± 1 day.

    Only single-item, quantity-1 orders are handled so far (same starting
    point every other flow in this file began at) - anything else stops
    rather than guessing.

    session_state, if given, is passed straight to find_order_smart, same
    reasoning as reconcile_order/reconcile_alifmoliya_order - leave it
    None for a one-off single-order run."""
    def ok(prompt):
        return True if auto_confirm else confirm(prompt)

    # report (if the caller passed one - the batch loop does) is filled in
    # as data is discovered, so the batch has whatever's known regardless of
    # which exit point below is hit (see report_sheet.py / the batch loops).
    rep = report if report is not None else {}
    _reset_price_notes()

    print(f"=== Uzum Nasiya order {order_number} ===")
    win = get_window()
    focus(win)
    open_zakazy_klientov(win)

    agreement = find_order_smart(win, order_number, agreements, session_state if session_state is not None else {})
    if not agreement:
        print(f"!! {order_number} not found in 1C under any of {agreements} - "
              "this is the rare exceptional case, skipping (needs manual handling).")
        return _status("skipped", "not_found_in_1c")

    confirmed_orders, unconfirmed_orders = read_open_orders_with_dates(win)
    order_date = dict(confirmed_orders).get(order_number)
    if order_date is None:
        if order_number in unconfirmed_orders:
            print(f"!! {order_number} shows only a time (no date) in 1C's grid - not confirmed yet, "
                  "stopping (deal with this one once it has a real date).")
            return _status("skipped", "not_yet_confirmed")
        print(f"!! {order_number} not found among currently open Uzum Nasiya orders (already closed/"
              "reconciled, or an unexpected Текущее состояние) - can't determine its date, stopping.")
        return _status("skipped", "not_currently_open")

    result = open_and_read_related_documents(win, order_number)
    if result is None:
        print(f"!! Related-documents check failed for {order_number} (see messages above) - "
              "stopping, nothing was changed.")
        return _status("skipped", "related_docs_check_failed")
    if result.get("ekvayring"):
        print(f"=== {order_number} already has an Эквайринговая операция "
              f"({result['ekvayring']}) - reconciled by an earlier run, nothing to do. ===")
        return _status("skipped", "already_has_ekvayring", detail=result["ekvayring"])
    realizatsiya_date, grid = result["date"], result["grid"]
    rep["order_date"] = result.get("order_date") or ""

    open_related_row(win, grid, row_index=0, expect_title_startswith="Заказ клиента")
    guard_no_surprise_dialog(win)
    click_tabitem(win, "Товары")
    time.sleep(0.5)
    focus(win)

    order_items = read_item_rows(win)
    kinds = [_classify_order_item(it["nomenklatura"]) for it in order_items]
    products = [it for it, k in zip(order_items, kinds) if k == "product"]
    delivery = [it for it, k in zip(order_items, kinds) if k == "delivery"]
    gift = [it for it, k in zip(order_items, kinds) if k == "gift"]
    if len(products) != 1:
        print(f"!! {order_number} has {len(products)} product line(s) in 1C (plus {len(delivery)} delivery, "
              f"{len(gift)} gift) - only a single product line is handled for Uzum Nasiya, stopping.")
        return _status("skipped", "multi_item_unhandled")
    order_item = products[0]
    rep["name_1c"] = order_item["nomenklatura"]
    rep["quantity"] = order_item["kolichestvo"]
    rep["price_before"] = order_item["tsena"]
    if order_item["kolichestvo"] != 1:
        print(f"!! {order_number}'s item quantity is {order_item['kolichestvo']}, not 1 - stopping.")
        return _status("skipped", "quantity_not_1")

    # Report generation is somewhat expensive (a real browser session, page
    # navigation, and a download) - deliberately done AFTER the cheaper
    # related-docs/item-shape checks above, so an order that turns out to
    # be unhandled/skippable anyway doesn't pay for a report it'll never use.
    if report_rows is None:
        report_start = order_date - datetime.timedelta(days=1)
        report_end = order_date + datetime.timedelta(days=1)
        print(f"-> {order_number}'s date is {order_date} - generating a cancel-dates report for "
              f"{report_start}..{report_end}")
        with sync_playwright() as p:
            report_path = uzum_client.get_cancel_dates_report(p, report_start, report_end)
        report_rows = uzum_client.parse_report(report_path)
        print(f"   report downloaded and parsed: {len(report_rows)} row(s)")

    best_row, coverage = _match_report_row_for_order(report_rows, order_date, order_item["nomenklatura"],
                                                      name_field="product_name")
    if best_row is None:
        print(f"!! No report row for {order_date} - can't find {order_number}'s price, stopping.")
        return _status("skipped", "no_report_row_for_date")
    print(f"   matched report row: {best_row['product_name']!r} ({best_row['date']}) = {best_row['price']} "
          f"(name coverage={coverage:.2f})")
    rep["name_source"] = best_row["product_name"]
    rep["price_source"] = best_row["price"]
    if not names_match_or_confirm(best_row["product_name"], order_item["nomenklatura"]):
        print(f"!! Product name mismatch on {order_number} - stopping without changes.")
        return _status("skipped", "name_mismatch")

    # Delivery split (per instruction, same rule as Alif Moliya): the report
    # prices the product at the FULL financed total, but 1C splits it into a
    # product line + a Доставка line, so the product line = report total -
    # the delivery line's current price (product + delivery == report total).
    # The note goes in the report's Comment column (NOT a Telegram tag).
    delivery_total = sum((d["tsena"] or 0) for d in delivery)
    product_target = round(best_row["price"] - delivery_total, 2) if delivery else best_row["price"]
    rep["price_after"] = product_target
    split_note = _split_delivery_note(order_number, best_row["price"], delivery, gift)
    if split_note:
        print(f"   {split_note}")

    # Product line gets the (possibly delivery-adjusted) price; delivery and
    # gift lines are left exactly as they are (a no-op re-apply is skipped).
    pairs = [(order_item, {"target_price": product_target, "skip_price_fix": False})]
    pairs += [(x, {"target_price": x["tsena"], "skip_price_fix": True}) for x in delivery + gift]

    price_summary = ", ".join(str(pl["target_price"]) for _, pl in pairs)
    if not ok(f"Set {order_number}'s order price(s) to [{price_summary}] and post it?"):
        print("Left the order open without changing its price - stopping.")
        return _status("skipped", "user_declined_order_price")
    try:
        for it, pl in pairs:
            if pl["skip_price_fix"]:
                continue
            apply_price_fix_for_item(win, it, pl["target_price"])
    except RuntimeError as e:
        print(f"!! {order_number}'s order price fix failed verification: {e} - stopping without posting.")
        return _status("skipped", "order_price_fix_failed", detail=str(e))
    try:
        post_and_close(win, verify_closed_title_contains="Заказ клиента")
    except RuntimeError as e:
        print(f"!! {order_number}'s order did not post: {e} - check by hand.")
        return _status("partial", "order_post_failed", detail=str(e))

    print("-> back to Связанные документы, opening Реализация (row 1)")
    grid = _reopen_related_documents(win, grid, order_number)
    if grid is None:
        print(f"!! Could not reopen Связанные документы for {order_number} after fixing the order price - "
              "stopping, check by hand (order price WAS already changed).")
        return _status("partial", "reopen_failed_after_order_post")

    open_related_row(win, grid, row_index=1, expect_title_startswith="Реализация")
    guard_no_surprise_dialog(win)
    click_tabitem(win, "Товары")
    time.sleep(0.5)
    focus(win)
    if not ok("Set the Реализация's price(s) the same way, and post it?"):
        print("Left the Реализация untouched (order price WAS already changed) - finish this by hand.")
        return _status("partial", "user_declined_realizatsiya_price")
    try:
        realiz_items = read_item_rows(win)
        if len(realiz_items) != len(pairs):
            raise RuntimeError(f"Реализация has {len(realiz_items)} item row(s), expected {len(pairs)}")
        remaining = list(pairs)
        for ritem in realiz_items:
            remaining.sort(key=lambda p: product_name_coverage(p[0]["nomenklatura"], ritem["nomenklatura"]),
                           reverse=True)
            _, pl = remaining.pop(0)
            if pl["skip_price_fix"]:
                continue
            apply_price_fix_for_item(win, ritem, pl["target_price"])
    except RuntimeError as e:
        print(f"!! {order_number}'s Реализация price fix failed: {e} - "
              "stopping (order price was already fixed, check by hand).")
        return _status("partial", "realizatsiya_price_fix_failed", detail=str(e))
    try:
        post_and_close(win, verify_closed_title_contains="Реализация")
    except RuntimeError as e:
        print(f"!! {order_number}'s Реализация did not post: {e} - order price already fixed, check by hand.")
        return _status("partial", "realizatsiya_post_failed", detail=str(e))

    print("-> back to Связанные документы, opening the order (row 0) for Эквайринговая операция")
    grid = _reopen_related_documents(win, grid, order_number)
    if grid is None:
        print(f"!! Could not reopen Связанные документы for {order_number} to create the Эквайринговая операция - "
              "both prices ARE already fixed, finish this one by hand.")
        return _status("partial", "reopen_failed_before_ekvayring")
    open_related_row(win, grid, row_index=0, expect_title_startswith="Заказ клиента")
    guard_no_surprise_dialog(win)

    try:
        create_ekvayring_operation(win, realizatsiya_date or "", terminal=terminal)
    except RuntimeError as e:
        print(f"!! {order_number}'s Эквайринговая операция setup failed: {e} - "
              "both prices are already fixed and posted; finish this by hand.")
        return _status("partial", "ekvayring_setup_failed", detail=str(e))

    if not ok("Эквайринговая операция form filled in correctly above (date and Терминал set)? Post & close it now?"):
        print("Left the Эквайринговая операция open without posting - "
              "both prices are already fixed and posted; finish this by hand.")
        return _status("partial", "user_declined_ekvayring")
    try:
        post_and_close(win, verify_closed_title_contains="Эквайринговая")
    except RuntimeError as e:
        print(f"!! {order_number}'s Эквайринговая операция did not post: {e} - "
              "both prices are already fixed and posted; finish this by hand.")
        return _status("partial", "ekvayring_post_failed", detail=str(e))

    print(f"=== Done: {order_number} reconciled at price(s) [{price_summary}] - no sheet to write back "
          "(Uzum Nasiya has none, 1C's own 'Все открытые' filter is the source of truth). ===")
    return _status("done", "reconciled", detail=_combine_detail(split_note, _price_notes_suffix()))


def process_all_uzumnasiya_orders(agreements=UZUM_AGREEMENTS, terminal=UZUM_TERMINAL, since_date=None):
    """Batch mode for Uzum Nasiya (see main()'s --profile uzum-nasiya
    handling) - a near-exact sibling of process_all_alifmoliya_orders (see
    its docstring). Fetches ONE CSV report up front covering every open
    order's own date range (min - 1 day .. max + 1 day) rather than one per
    order (see reconcile_uzumnasiya_order's report_rows parameter).

    since_date is OPTIONAL here (unlike the IMAN/Alif Shop profiles' --since,
    which is required for batch mode) - "Все открытые" already bounds this
    to currently-open orders only, there's no unlimited historical backlog
    to worry about looping over."""
    win = get_window()
    focus(win)
    open_zakazy_klientov(win)
    # "Все открытые" passed explicitly: for this profile the filter IS the
    # definition of what to process (there's no sheet saying which orders to
    # look at), unlike an order lookup, where it would hide already-finished
    # orders - see clear_tekushee_sostoyanie.
    set_filters(win, agreements[0], tekushee_sostoyanie="Все открытые")
    # Batch mode wants EVERY open order under this Соглашение - a stale term
    # left in the quick-search box (from a prior manual search or a
    # single-order run) would silently filter the grid down, so clear it
    # first (see clear_search - confirmed live this exact case reported
    # "no open orders" when there really were several).
    clear_search(win)

    orders, unconfirmed_orders = read_open_orders_with_dates(win)
    if unconfirmed_orders:
        print(f"   ({len(unconfirmed_orders)} order(s) show only a time (no date) in 1C's grid - not "
              f"confirmed yet, excluding from this run: {', '.join(unconfirmed_orders)})")
    if since_date is not None:
        orders = [(o, d) for o, d in orders if d >= since_date]
    since_note = f" on/after {since_date}" if since_date else ""
    if not orders:
        print(f"No open (confirmed) Uzum Nasiya orders{since_note} to process.")
        return {"done": [], "skipped": [], "partial": [], "errored": []}

    dates = [d for _, d in orders]
    report_start = min(dates) - datetime.timedelta(days=1)
    report_end = max(dates) + datetime.timedelta(days=1)
    print(f"=== Batch run: {len(orders)} open Uzum Nasiya order(s){since_note}, dated {min(dates)}..{max(dates)} - "
          f"generating a cancel-dates report for {report_start}..{report_end} ===\n")

    with sync_playwright() as p:
        report_path = uzum_client.get_cancel_dates_report(p, report_start, report_end)
    report_rows = uzum_client.parse_report(report_path)
    print(f"   report downloaded and parsed: {len(report_rows)} row(s)\n")

    done, skipped, partial, errored = [], [], [], []
    report_ws = _open_report_ws()
    filter_session_state = {}

    try:
        for i, (order_number, order_date) in enumerate(orders, 1):
            print(f"[{i}/{len(orders)}] --- {order_number} ({order_date}) ---")
            rep = {}
            try:
                result = reconcile_uzumnasiya_order(order_number, report_rows=report_rows, auto_confirm=True,
                                                     agreements=agreements, terminal=terminal,
                                                     session_state=filter_session_state, report=rep)
            except Exception as e:  # noqa: BLE001 - deliberately broad, see process_all_unchecked_orders's docstring
                print(f"!! Unexpected error on {order_number}: {e!r} - moving on to the next order.")
                result = _status("error", "unexpected_exception", detail=repr(e))

            try:
                close_back_to_zakazy_klientov(win)
            except Exception as e:  # noqa: BLE001 - tab cleanup failing shouldn't kill the batch either
                print(f"   (tab cleanup after {order_number} hit an error: {e!r} - continuing anyway)")

            _log_report_row(report_ws, rep, order_number, terminal, agreements[0], result)
            entry = (order_number, result.get("reason", "?"), result.get("detail", ""))
            status = result.get("status")
            if status == "done":
                done.append(entry)
            elif status == "partial":
                partial.append(entry)
            elif status == "error":
                errored.append(entry)
            else:
                skipped.append(entry)
            print()
            telegram_logger.flush()
    except KeyboardInterrupt:
        print("\n!! Interrupted - stopping the batch early, here's the report for what ran so far.\n")

    with telegram_logger.batch():
        print("=== Batch report ===")
        print(f"Done: {len(done)}")
        print(f"Skipped (safe to ignore, nothing changed): {len(skipped)}")
        print(f"Partial (SOMETHING WAS POSTED - needs a human to finish/check!): {len(partial)}")
        print(f"Errored unexpectedly: {len(errored)}")

        def _print_group(title, entries):
            if not entries:
                return
            print(f"\n{title}:")
            for order_number, reason, detail in entries:
                suffix = f" ({detail})" if detail else ""
                print(f"  {order_number}: {reason}{suffix}")

        _print_group("Skipped", skipped)
        _print_group("!! PARTIAL - needs attention", partial)
        _print_group("!! ERRORED - needs attention", errored)

    return {"done": done, "skipped": skipped, "partial": partial, "errored": errored}


# -------------------------------------------------------------- manual flow

def main_manual():
    print("=== 1C order -> price fix -> Эквайринговая операция automation (manual mode) ===")
    order_number = input("Order number to search for (e.g. QQ-26/00103): ").strip()
    soglashenie = input("Соглашение filter value (e.g. Alif Moliya): ").strip()
    price = input("New price to set on the Реализация item row(s): ").strip()
    terminal = input("Терминал / tип соглашения for the Эквайринговая операция (leave blank to skip): ").strip()

    win = get_window()
    focus(win)

    open_zakazy_klientov(win)
    set_filters(win, soglashenie)
    search_order(win, order_number)

    result = open_and_read_related_documents(win, order_number)
    if result is None:
        print("Stopping - see messages above.")
        return
    if result.get("ekvayring"):
        print(f"{order_number} already has an Эквайринговая операция ({result['ekvayring']}) - "
              "already reconciled, nothing to do.")
        return
    realizatsiya_date, grid = result["date"], result["grid"]

    if not confirm("Related documents look right above? Continue to open Реализация and fix prices?"):
        return

    open_related_row(win, grid, row_index=1, expect_title_startswith="Реализация")
    guard_no_surprise_dialog(win)
    fix_realizatsiya_prices(win, price)
    post_and_close(win, verify_closed_title_contains="Реализация")

    print("-> back to Связанные документы, opening the order (row 0)")
    grid = _reopen_related_documents(win, grid, order_number)
    if grid is None:
        print("Could not reopen Связанные документы - finish by hand.")
        return
    open_related_row(win, grid, row_index=0, expect_title_startswith="Заказ клиента")
    guard_no_surprise_dialog(win)

    create_ekvayring_operation(win, realizatsiya_date or "", terminal or None)

    if not confirm("Эквайринговая операция form filled in correctly above? Post & close it now?"):
        print("Left the document open without posting - review/close it by hand.")
        return
    post_and_close(win, verify_closed_title_contains="Эквайринговая")

    print("=== done ===")


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument(
        "--manual", action="store_true",
        help="Original manual single-order flow (type in order number/Соглашение/price/terminal by hand) "
             "instead of reading the next unchecked row from the Google Sheet.",
    )
    parser.add_argument(
        "--profile", default=None, choices=sorted(PROFILES),
        help="Named bundle of --worksheet/--agreements/--terminal for a specific sheet/workflow. Required "
             "(no default) for every mode except --manual, which doesn't use it - more profiles are coming "
             "soon (different sheets with different agreements/terminal), and silently defaulting to one "
             "risks running against the wrong one by accident. Add a new entry to PROFILES in this file when "
             "connecting a different sheet instead of overriding the individual flags below by hand every run.",
    )
    parser.add_argument("--worksheet", default=None, help="Google Sheet tab name to read/write "
                                                            "(default: from --profile).")
    parser.add_argument(
        "--yes", action="store_true",
        help="Skip the confirm() checkpoints before posting/writing. Required with --since (see below); "
             "optional elsewhere but not recommended until this has been validated against a few real orders.",
    )
    parser.add_argument(
        "--order", help="Run the reconciliation logic against exactly this order number instead of whatever "
                         "the sheet says is next. Without --amount, looks the order up in the sheet itself "
                         "(same as --since would, just targeting one specific order - this also means the "
                         "result gets written back to Сумма 1С). With --amount, bypasses the sheet entirely "
                         "(useful for an order not yet in the sheet, before credentials\\service_account.json "
                         "exists, or to force a specific amount). With --profile alif-moliya/uzum-nasiya: no "
                         "sheet involved at all - looks the order up directly among currently-open orders in "
                         "1C under that profile's Соглашение and generates a small report just for its date "
                         "to find its price (--amount/--product are ignored for these profiles).",
    )
    parser.add_argument("--amount", type=float, help="With --order: force this amount instead of looking it "
                                                       "up in the sheet (also disables writing back to the sheet).")
    parser.add_argument(
        "--application", help="Alif Shop application id (Alif Nasiya sheet's column B) to check/reconcile - "
                               "requires --profile alif-shop. Looks up its status on business.alif.uz: "
                               "records 'Yetkazib berishga' as waiting and 'Bekor qilindi' as Cancelled "
                               "(row colored red) with nothing further to do; once 'Rasmiylashtirildi', "
                               "computes the post-commission price and reconciles the matching 1C order "
                               "(already named in the sheet's column K) the same way --order does for IMAN.",
    )
    parser.add_argument("--product", default="", help="With --order --amount: expected product name to "
                                                        "fuzzy-match against 1C (skip to always ask you to "
                                                        "confirm it). Ignored otherwise - the sheet already has it.")
    parser.add_argument(
        "--since", metavar="DD.MM.YYYY",
        help="Batch mode (requires --yes): loop over and reconcile EVERY unreconciled row on/after this date, "
             "instead of just the next one, printing a running [i/N] progress line and a final summary report "
             "(done/skipped/needs-attention counts + reasons). Rows are never retried within a run - each is "
             "moved past (whether it succeeded, was skipped, or hit a partial failure) so one bad row can't "
             "loop forever. With --profile imanum: rows whose Дата Реализации is on/after this date, "
             "REQUIRED for batch mode. With --profile alif-shop: rows whose Дата is on/after this date AND "
             "not already success/Cancelled (so a 'waiting' application gets rechecked on the next --since "
             "run, not just a never-checked one), REQUIRED for batch mode. With --profile alif-moliya/"
             "uzum-nasiya: OPTIONAL extra filter narrowing the batch to currently-open orders dated on/after "
             "this date (batch mode itself doesn't need a date - see main()'s alif-moliya/uzum-nasiya "
             "handling). Ignored with --order/--application/--manual.",
    )
    parser.add_argument(
        "--agreements", default=None,
        help="Comma-separated Соглашение values to search 1C under, in order (default: from --profile). "
             "Override only if this specific run needs something different from the profile.",
    )
    parser.add_argument(
        "--terminal", default=None,
        help="Терминал to set on the Эквайринговая операция (default: from --profile). Override only if "
             "this specific run needs something different from the profile.",
    )
    args = parser.parse_args()
    # Only from here on - --help (or a bad flag) exits during parse_args()
    # itself, above, so it never reaches this and never gets relayed.
    telegram_logger.install()
    launch_1c.ensure_1c_open()

    if args.manual:
        main_manual()
        return

    if not args.profile:
        parser.error(f"--profile is required (choices: {', '.join(sorted(PROFILES))}) - no default, since "
                     "more profiles are coming soon and defaulting silently risks running against the wrong "
                     "sheet's agreements/terminal by accident. (Not needed for --manual.)")

    profile = PROFILES[args.profile]
    agreements = tuple(a.strip() for a in args.agreements.split(",") if a.strip()) if args.agreements \
        else profile["agreements"]
    terminal = args.terminal or profile["terminal"]

    if args.application:
        if args.profile != "alif-shop":
            parser.error(f"--application requires --profile alif-shop (got --profile {args.profile!r})")
        report_ws = _open_report_ws()
        rep = {}
        result = reconcile_alifshop_application(args.application, auto_confirm=args.yes, agreements=agreements,
                                                 terminal=terminal, report=rep)
        _log_report_row(report_ws, rep, rep.get("contract", ""), terminal, agreements[0], result)
        return

    if args.profile == "alif-shop":
        # alif-shop has no "worksheet" flag/profile key (see PROFILES) - it
        # always reads/writes the Alif Nasiya sheet directly (see
        # alif_sheets_client.py), so it only makes sense with --application
        # (handled above) or --since (batch mode, handled here) - not the
        # IMAN-only --order/default-next-row flow below.
        if not args.since:
            parser.error("--profile alif-shop needs either --application <id> (one application) or "
                         "--since DD.MM.YYYY --yes (batch mode over every unresolved application).")
        since_date = parse_sheet_date(args.since)
        if since_date is None:
            parser.error(f"--since {args.since!r} doesn't look like DD.MM.YYYY")
        if not args.yes:
            parser.error("--since processes every unresolved application in a loop (batch mode) - pass --yes "
                         "too, since there's no human to click through confirm() prompts for each one.")
        process_all_alifshop_applications(agreements=agreements, since_date=since_date, terminal=terminal)
        return

    if args.profile == "alif-moliya":
        # alif-moliya has no "worksheet" flag/profile key either (see
        # PROFILES) - there's no Google Sheet at all for it (see
        # reconcile_alifmoliya_order's docstring). --order here means "just
        # this one order" (unlike --order for imanum below, this
        # never looks anything up in a sheet - there isn't one); omitting
        # it means batch mode over every currently-open Alif Moliya order,
        # optionally narrowed by --since (optional here, unlike the other
        # profiles, since "Все открытые" already bounds the batch to
        # currently-open orders only - see process_all_alifmoliya_orders).
        if args.order:
            report_ws = _open_report_ws()
            rep = {}
            result = reconcile_alifmoliya_order(args.order, auto_confirm=args.yes, agreements=agreements,
                                                 terminal=terminal, report=rep)
            _log_report_row(report_ws, rep, args.order, terminal, agreements[0], result)
            return
        if not args.yes:
            parser.error("--profile alif-moliya batch mode processes every currently-open order under Alif "
                         "Moliya in one run - pass --yes, since there's no human to click through confirm() "
                         "prompts for each one. Pass --order <number> instead to process just one interactively.")
        since_date = parse_sheet_date(args.since) if args.since else None
        if args.since and since_date is None:
            parser.error(f"--since {args.since!r} doesn't look like DD.MM.YYYY")
        process_all_alifmoliya_orders(agreements=agreements, terminal=terminal, since_date=since_date)
        return

    if args.profile == "uzum-nasiya":
        # Same shape as alif-moliya just above (see its comment) - no
        # Google Sheet, --order for one order, batch mode (optional
        # --since) otherwise.
        if args.order:
            report_ws = _open_report_ws()
            rep = {}
            result = reconcile_uzumnasiya_order(args.order, auto_confirm=args.yes, agreements=agreements,
                                                 terminal=terminal, report=rep)
            _log_report_row(report_ws, rep, args.order, terminal, agreements[0], result)
            return
        if not args.yes:
            parser.error("--profile uzum-nasiya batch mode processes every currently-open order under Uzum "
                         "Nasiya in one run - pass --yes, since there's no human to click through confirm() "
                         "prompts for each one. Pass --order <number> instead to process just one interactively.")
        since_date = parse_sheet_date(args.since) if args.since else None
        if args.since and since_date is None:
            parser.error(f"--since {args.since!r} doesn't look like DD.MM.YYYY")
        process_all_uzumnasiya_orders(agreements=agreements, terminal=terminal, since_date=since_date)
        return

    # imanum-only from here on - needs profile["worksheet"], which
    # alif-shop/alif-moliya/uzum-nasiya's PROFILES entries deliberately
    # don't have (see PROFILES).
    worksheet = args.worksheet or profile["worksheet"]

    if args.order:
        report_ws = _open_report_ws()
        rep = {}
        if args.amount is not None:
            # Standalone override - bypasses the sheet entirely (useful for
            # an order not yet in the sheet, or to force a specific amount).
            row = {"order_number": args.order, "expected_amount": args.amount, "product_name": args.product,
                   "row": None}
            result = reconcile_order(row, agreements=agreements, auto_confirm=args.yes, ws=None,
                                     terminal=terminal, report=rep)
            _log_report_row(report_ws, rep, args.order, terminal, "/".join(agreements), result)
            return

        # No --amount given - look the order up in the sheet instead of
        # requiring it to be typed in by hand every time. Passing ws=ws
        # (not None) also means determine_item_price_pairs can do its own
        # multi-item/bonus-item sheet lookups and the result gets written
        # back to Сумма 1С same as the --since flow.
        ws = get_worksheet(worksheet)
        sheet_rows = find_all_rows_for_order(ws, args.order)
        if not sheet_rows:
            parser.error(f"--order {args.order!r} not found in the {worksheet!r} sheet and no --amount was "
                         "given - pass --amount (and optionally --product) to run this order standalone instead.")
        first = sheet_rows[0]
        row = {"order_number": args.order, "expected_amount": first["expected_amount"],
               "product_name": first["product_name"], "row": first["row"]}
        result = reconcile_order(row, agreements=agreements, auto_confirm=args.yes, ws=ws,
                                 terminal=terminal, report=rep)
        _log_report_row(report_ws, rep, args.order, terminal, "/".join(agreements), result)
        return

    since_date = parse_sheet_date(args.since) if args.since else None
    if args.since and since_date is None:
        parser.error(f"--since {args.since!r} doesn't look like DD.MM.YYYY")

    ws = get_worksheet(worksheet)

    if args.since:
        # --since means "process everything from this date forward", a
        # loop over potentially many orders - there's no human available to
        # click through a confirm() prompt for each one, so this requires
        # --yes rather than silently assuming it.
        if not args.yes:
            parser.error("--since processes every unreconciled row in a loop (batch mode) - pass --yes too, "
                         "since there's no human to click through confirm() prompts for each one. Omit --since "
                         "(just run with no date) to process a single row interactively instead.")
        process_all_unchecked_orders(ws, agreements=agreements, since_date=since_date, terminal=terminal)
        return

    process_next_unchecked_order(ws, agreements=agreements, auto_confirm=args.yes, since_date=since_date,
                                  terminal=terminal)


if __name__ == "__main__":
    main()

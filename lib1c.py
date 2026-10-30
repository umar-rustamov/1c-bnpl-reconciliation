# -*- coding: utf-8 -*-
"""
Reusable helpers for driving the 1C:Enterprise thick/thin client ("<Company>
/ Управление торговлей") window via UI Automation (pywinauto) + raw Win32 input.

Key facts learned while building this (kept here so future edits don't have to
re-discover them the hard way):

- The app exposes a rich, named UIA tree (buttons, tab items, checkboxes, combo
  boxes all have real accessible names) EXCEPT for grid/list contents, which
  are custom-drawn:
    * Dynamic list grids (e.g. "Заказы клиентов", document line grids) DO expose
      each visible cell as a `Custom` control whose window_text() is
      "<value> <column name>" (e.g. "ON-26/00106 Номер"). Only the currently
      *visible/scrolled-into-view* rows are realized as elements - it's a
      virtualized grid.
    * The "Связанные документы" (related documents) tree grid exposes NO
      children via UIA at all. The only reliable way to read its rows is to
      click into it, Ctrl+A / Ctrl+C, and read tab-separated text off the
      clipboard.
- Any interaction (click, keyboard) must happen with the 1C window in the real
  OS foreground - `SetForegroundWindow` on our own window. Do this right before
  every interaction because whatever process is running this script (e.g. this
  terminal/Claude) will keep stealing focus back between script invocations.
- Modal "1С:Предприятие" dialogs (save confirmation, etc.) are SEPARATE top
  level HWNDs, not children of the main window, and they immediately grab
  foreground. Detect them via `win32gui.GetForegroundWindow()` and drive them
  directly through their own UIA window.
- Reference/catalog combo fields (Соглашение, Терминал, ...) support
  incremental search: click the field, Ctrl+A, paste replacement text, wait for
  the suggestion popup, click the exact-text match. Cyrillic text must go
  through the clipboard (SendKeys-style key injection does not reliably type
  Cyrillic); pure ASCII/digits can be typed directly with `keyboard.send_keys`.
- Fixed-enum picklists (e.g. "Вид цены") do NOT support incremental search -
  typing just edits raw text. To pick an item there: click near the right edge
  of the cell to open the dropdown (a plain single click only *selects* the
  cell; the dropdown opens on the second interaction), then scroll the popup
  with `mouse.scroll` and click the item's on-screen position.
- Opening a row in a virtualized grid or in the (non-accessible) related-docs
  tree requires a genuine OS-level double click (`pywinauto.mouse.double_click`
  at absolute screen coords). Doing two separate `element.click_input()` calls
  does NOT reliably register as a double-click.
- Ctrl+Enter on a document form == click its highlighted default button
  ("Провести и закрыть" - post & close). It does NOT just "confirm the current
  cell" - it saves/posts and closes the whole document. Use plain Enter/Tab to
  commit a single field without closing the document.
"""

import ctypes
import re
import time

import win32api
import win32clipboard
import win32con
import win32gui
from pywinauto import Desktop, keyboard, mouse
from PIL import ImageGrab

from env_config import load_env

# The 1C main window's title is "<configuration name> / Управление торговлей,
# редакция 11" - the configuration name is company-specific, so it comes from
# credentials\.env (1c_window_title_hint); empty hints are dropped.
WINDOW_TITLE_HINTS = tuple(
    h for h in (load_env().get("1c_window_title_hint", ""), "Управление торговлей") if h
)

_user32 = ctypes.windll.user32
_kernel32 = ctypes.windll.kernel32
SW_RESTORE = 9


def get_window():
    """Find the 1C main window among all top-level desktop windows.

    Confirmed live this needs more care than "any window whose title
    contains any one hint": the production infobase is also web-published
    (see launch_1c.py), and a plain browser tab left open on that URL gets
    a window titled just "<Company> - Google Chrome" - loosely matching
    WINDOW_TITLE_HINTS on "<Company>" alone just as well as the real 1C
    window ("<Company> / Управление торговлей, редакция 11", which
    contains BOTH hints). Desktop(backend="uia").windows() has no
    guaranteed ordering, so "return the first loose match" could pick
    either one depending on what else happens to be open - confirmed live,
    it picked the Chrome tab on one run.

    Prefers a window matching ALL hints (disambiguates exactly the case
    above) - only falls back to a loose single-hint match if there's
    exactly one such window (e.g. a differently-worded 1C title after some
    future update), and raises rather than guessing if it's still
    ambiguous after that."""
    windows = list(Desktop(backend="uia").windows())
    all_hint_matches = [w for w in windows if all(h in w.window_text() for h in WINDOW_TITLE_HINTS)]
    if len(all_hint_matches) == 1:
        return all_hint_matches[0]
    if len(all_hint_matches) > 1:
        raise RuntimeError(
            f"Multiple windows match ALL of 1C's title hints ({[w.window_text() for w in all_hint_matches]!r}) "
            "- can't tell which is the real 1C window, check by hand."
        )

    loose_matches = [w for w in windows if any(h in w.window_text() for h in WINDOW_TITLE_HINTS)]
    if len(loose_matches) == 1:
        return loose_matches[0]
    if len(loose_matches) > 1:
        raise RuntimeError(
            f"Multiple windows loosely match 1C's title hints ({[w.window_text() for w in loose_matches]!r}) "
            "but none matches all of them - can't tell which is the real 1C window, check by hand."
        )
    raise RuntimeError("1C main window not found - is it running?")


def _try_set_foreground_once(hwnd):
    """Plain SetForegroundWindow is often silently refused by Windows when
    the calling process isn't the one that currently owns 'input'
    attribution (e.g. this script running in a terminal the user just typed
    into to answer an input() prompt). The documented workaround is to
    temporarily attach this thread's input queue to the current foreground
    thread's, which grants permission - see Win32 AttachThreadInput docs."""
    if _user32.IsIconic(hwnd):
        _user32.ShowWindow(hwnd, SW_RESTORE)

    fg = _user32.GetForegroundWindow()
    if fg == hwnd:
        return True

    our_tid = _kernel32.GetCurrentThreadId()
    fg_tid = _user32.GetWindowThreadProcessId(fg, None) if fg else 0
    target_tid = _user32.GetWindowThreadProcessId(hwnd, None)

    attached_fg = attached_target = False
    try:
        if fg_tid and fg_tid != our_tid:
            attached_fg = bool(_user32.AttachThreadInput(our_tid, fg_tid, True))
        if target_tid and target_tid != our_tid and target_tid != fg_tid:
            attached_target = bool(_user32.AttachThreadInput(our_tid, target_tid, True))

        _user32.SetForegroundWindow(hwnd)
        _user32.BringWindowToTop(hwnd)
    finally:
        if attached_fg:
            _user32.AttachThreadInput(our_tid, fg_tid, False)
        if attached_target:
            _user32.AttachThreadInput(our_tid, target_tid, False)

    return _user32.GetForegroundWindow() == hwnd


def force_foreground(win, retries=10):
    hwnd = win.handle
    for _ in range(retries):
        try:
            if _try_set_foreground_once(hwnd):
                return True
        except Exception:
            pass
        time.sleep(0.15)
    return win32gui.GetForegroundWindow() == hwnd


def focus(win):
    win.set_focus()
    time.sleep(0.15)
    force_foreground(win)
    time.sleep(0.15)


def foreground_hwnd():
    return win32gui.GetForegroundWindow()


def get_dialog(hwnd=None):
    """Return the UIA window wrapper for whatever HWND currently owns
    foreground (used for modal '1С:Предприятие' confirmation dialogs, which
    are separate top-level windows, not children of the main window)."""
    hwnd = hwnd or foreground_hwnd()
    return Desktop(backend="uia").window(handle=hwnd)


def dialog_present(win):
    """True if the foreground window is a modal dialog, not the main window."""
    return foreground_hwnd() != win.handle


def get_stable_dialog(win, settle_checks=2, settle_delay=0.15):
    """Like dialog_present()+get_dialog(), but atomic and debounced: captures
    the foreground HWND once and re-checks it a couple of times before
    trusting it. A single dialog_present() read can be a one-frame flicker
    (e.g. while a new document tab is opening) that's gone by the time a
    separate get_dialog() call re-reads foreground - that race is exactly
    what caused a false-positive 'dialog' that turned out to just be the main
    window's own tab-close buttons. Returns (hwnd, uia_window) or (None, None)."""
    hwnd = foreground_hwnd()
    if hwnd == win.handle:
        return None, None
    for _ in range(settle_checks):
        time.sleep(settle_delay)
        if foreground_hwnd() != hwnd:
            return None, None  # it moved on its own - was a flicker, not a real dialog

    # Confirmed live: when this script runs from an IDE's own integrated
    # terminal, the IDE itself can legitimately grab OS foreground between
    # our own force_foreground() calls (e.g. right after a keypress) with no
    # real 1C dialog involved - and it can sit there stably long enough to
    # survive the settle-check above, misreading the IDE window itself as an
    # "unexpected dialog". A genuine modal 1C dialog refuses to give up
    # foreground even when asked; a stray sibling app doesn't. Ask, and only
    # treat this as a real dialog if it refuses to yield.
    if force_foreground(win):
        return None, None
    return hwnd, Desktop(backend="uia").window(handle=hwnd)


def click_dialog_button(button_text_options):
    """Click a button by exact text on whatever dialog currently has
    foreground. button_text_options: iterable of acceptable texts (Да/Нет/Отмена...)."""
    dlg = get_dialog()
    for b in dlg.descendants(control_type="Button"):
        if b.window_text().strip() in button_text_options:
            b.click_input()
            time.sleep(0.5)
            return True
    return False


def click_at(win, x, y, double=False):
    if not force_foreground(win):
        raise RuntimeError("could not bring 1C window to foreground before click")
    mouse.move(coords=(x, y))
    time.sleep(0.08)
    if double:
        mouse.double_click(button="left", coords=(x, y))
    else:
        mouse.click(button="left", coords=(x, y))
    time.sleep(0.3)


def set_clipboard_text(text, retries=20, delay=0.25):
    last_exc = None
    for _ in range(retries):
        try:
            win32clipboard.OpenClipboard()
            try:
                win32clipboard.EmptyClipboard()
                # SetClipboardText (not SetClipboardData) for a Python str:
                # SetClipboardData(CF_UNICODETEXT, str) expects a global
                # memory HANDLE and fails with "invalid handle" (error 6) on
                # this pywin32 build; SetClipboardText does the allocation.
                win32clipboard.SetClipboardText(text, win32con.CF_UNICODETEXT)
            finally:
                win32clipboard.CloseClipboard()
            return
        except Exception as e:
            last_exc = e
            time.sleep(delay)
    raise last_exc


def get_clipboard_text(retries=10, delay=0.2):
    last_exc = None
    for _ in range(retries):
        try:
            win32clipboard.OpenClipboard()
            try:
                return win32clipboard.GetClipboardData(win32clipboard.CF_UNICODETEXT)
            except Exception:
                return ""
            finally:
                win32clipboard.CloseClipboard()
        except Exception as e:
            last_exc = e
            time.sleep(delay)
    raise last_exc


# Virtual-key codes for every key this project ever sends as part of a
# modifier combination (single keys still go through pywinauto).
_VK = {
    "a": 0x41, "c": 0x43, "s": 0x53, "v": 0x56,
    "enter": 0x0D, "insert": 0x2D,
    "ctrl": 0x11, "shift": 0x10, "alt": 0x12,
}


def _key_event(vk, down):
    scan = win32api.MapVirtualKey(vk, 0)
    win32api.keybd_event(vk, scan, 0 if down else win32con.KEYEVENTF_KEYUP, 0)


def send_hotkey(*keys, gap=0.06):
    """Send a modifier combination - send_hotkey("ctrl", "a"),
    send_hotkey("ctrl", "shift", "s") - as individual keybd_event calls with
    real scancodes and a pause between each, rather than pywinauto's
    keyboard.send_keys("^a").

    Confirmed live that send_keys does not deliver Ctrl combinations to the
    1C thin client AT ALL: Ctrl+A, Ctrl+V, Ctrl+C, Ctrl+F and Ctrl+Shift+S
    were every one of them silent no-ops, while single keys ({ENTER}), a
    Shift combo ({VK_SHIFT down}{INSERT}) and plain typing all went through
    fine - and so did those same Ctrl combinations sent the way this function
    sends them. The difference is delivery, not the key codes: send_keys
    packs the whole modifier-down / key / modifier-up sequence into ONE
    SendInput batch with no scancodes, and 1C's input handling doesn't see
    the modifier; separate keybd_event calls carrying scancodes, spaced by a
    few tens of milliseconds, look like a real keypress and land every time.

    This is much bigger than one broken hotkey. With Ctrl+V dead, every
    paste_text() into a filter field silently did nothing, so the script kept
    running against whatever values the form already happened to have on it -
    the filters it "set" were simply whatever was left there from before."""
    seq = [k.lower() for k in keys]
    for k in seq:
        if k not in _VK:
            raise ValueError(f"send_hotkey: unknown key {k!r}")
    for k in seq:                 # modifiers first, then the key itself
        _key_event(_VK[k], down=True)
        time.sleep(gap)
    for k in reversed(seq):       # released in reverse order, as a real one would be
        _key_event(_VK[k], down=False)
        time.sleep(gap)


def paste_text(ctrl, text, select_all=True):
    """Click a field and replace its contents via clipboard paste (safe for
    Cyrillic, unlike keyboard.send_keys)."""
    ctrl.click_input()
    time.sleep(0.15)
    if select_all:
        send_hotkey("ctrl", "a")
        time.sleep(0.1)
    set_clipboard_text(text)
    time.sleep(0.1)
    send_hotkey("ctrl", "v")
    time.sleep(0.4)


def type_digits(ctrl, text):
    """Click a field and type plain ASCII (digits/dates) directly - no
    clipboard needed, and clipboard is sometimes transiently locked by 1C
    right after a grid cell enters edit mode."""
    ctrl.click_input()
    time.sleep(0.2)
    send_hotkey("ctrl", "a")
    time.sleep(0.1)
    keyboard.send_keys(text)
    time.sleep(0.2)


def pick_suggestion(win, exact_text, timeout=2.0):
    """After paste_text() into a searchable reference combo, click the
    exact-text suggestion in the popup list. Returns True if found+clicked."""
    end = time.time() + timeout
    while time.time() < end:
        items = [c for c in win.descendants() if c.window_text().strip() == exact_text]
        if items:
            items[-1].click_input()
            time.sleep(0.4)
            return True
        time.sleep(0.15)
    return False


def nearest_control_right_of_checkbox(win, checkbox_text, control_type="ComboBox", max_gap=80):
    """Find a filter row's value control by locating the labeled CheckBox and
    then the nearest control of `control_type` immediately to its right on the
    same row. Needed because several filter rows share the same Y coordinate
    (two side-by-side filters), so matching purely by vertical position is
    ambiguous."""
    checks = [c for c in win.descendants(control_type="CheckBox") if c.window_text().strip() == checkbox_text]
    if not checks:
        return None, None
    cb = checks[0]
    cbr = cb.rectangle()
    cy = (cbr.top + cbr.bottom) / 2
    best, best_gap = None, None
    for c in win.descendants(control_type=control_type):
        r = c.rectangle()
        rcy = (r.top + r.bottom) / 2
        gap = r.left - cbr.right
        if abs(rcy - cy) < 15 and 0 <= gap < max_gap:
            if best_gap is None or gap < best_gap:
                best_gap, best = gap, c
    return best, cb


def nearest_control_right_of_text(win, label_text, control_type="Edit", max_gap=40):
    """Same idea as nearest_control_right_of_checkbox, but anchored on a Text
    label instead of a CheckBox (used for plain 'Label: <field>' rows that
    have no checkbox, e.g. the date field on the Эквайринговая операция form)."""
    labels = [t for t in win.descendants(control_type="Text") if t.window_text().strip() == label_text]
    if not labels:
        return None
    lr = labels[0].rectangle()
    cy = (lr.top + lr.bottom) / 2
    best, best_gap = None, None
    for c in win.descendants(control_type=control_type):
        r = c.rectangle()
        rcy = (r.top + r.bottom) / 2
        gap = r.left - lr.right
        if abs(rcy - cy) < 15 and 0 <= gap < max_gap:
            if best_gap is None or gap < best_gap:
                best_gap, best = gap, c
    return best


def ensure_checkbox_checked(cb):
    try:
        if hasattr(cb, "get_toggle_state") and cb.get_toggle_state() == 1:
            return
    except Exception:
        pass
    cb.click_input()
    time.sleep(0.2)


def read_grid_via_clipboard(win, grid):
    """For grids with no UIA-accessible rows (e.g. Related Documents tree):
    click it, Select All, Copy, and read the resulting tab-separated text off
    the clipboard. Returns list of (indent_level, text) tuples, one per row."""
    grid.click_input()
    time.sleep(0.2)
    send_hotkey("ctrl", "a")
    time.sleep(0.15)
    send_hotkey("ctrl", "c")
    time.sleep(0.3)
    raw = get_clipboard_text()
    rows = []
    for line in raw.split("\r\n"):
        if not line.strip():
            continue
        stripped = line.lstrip("\t")
        indent = len(line) - len(stripped)
        cell = stripped.split("\t")[0].strip()
        if cell:
            rows.append((indent, cell))
    return rows


DATE_RE = re.compile(r"от (\d{2}\.\d{2}\.\d{4})")


def extract_date(text):
    m = DATE_RE.search(text)
    return m.group(1) if m else None


def wait(seconds):
    time.sleep(seconds)


# ----------------------------------------------------- posted-status (visual)
#
# The "Связанные документы" tree exposes no UIA children at all (see module
# docstring), so the little green-checkmark overlay 1C draws on a posted
# document's icon can't be read as text/state the way everything else in this
# file is read. The only way to see it is to literally look at the pixels.
# Empirically confirmed against a real posted order+realizatsiya pair (both
# rows showed a solid green blob, RGB roughly (18-40, 150-195, 45-75), inside
# roughly the first 150px of the row - the icon shifts right per tree-nesting
# level, hence scanning the whole row width rather than a fixed x position).

def capture_element_image(element):
    """Grab a screenshot of exactly the screen region an element occupies.
    Returns a PIL Image in the same coordinate space as element.rectangle()
    (i.e. subtract rect.left/rect.top from any point before indexing pixels).

    all_screens=True is NOT optional here. Without it ImageGrab.grab captures
    the PRIMARY monitor only and clamps the bbox to it, so on a multi-monitor
    setup every capture of a window living on a secondary screen comes back as
    the wrong region entirely. Confirmed live with 1C open on a monitor to the
    LEFT of the primary one (its window rect starts at x=-1927, i.e. negative
    screen coordinates): row_has_green_checkmark then found no green pixels in
    that garbage image and every single order was rejected as "missing the
    posted (green) checkmark" - a document-state failure reported for what was
    really a screenshot pointed at the wrong screen."""
    r = element.rectangle()
    return ImageGrab.grab(bbox=(r.left, r.top, r.right, r.bottom), all_screens=True)


def _is_greenish(pixel):
    r, g, b = pixel[:3]
    return g > 100 and g > r + 25 and g > b + 25


def row_has_green_checkmark(image, row_index, row_height=31, x_max=150, min_green_px=15):
    """Best-effort visual check for the posted/green-checkmark icon on a given
    row of a grid image captured with capture_element_image(). NOT verified
    against a real *unposted* document (none was available while building
    this) - if this starts misfiring, capture and inspect a real unposted row
    the same way diag_checkmark.py did before trusting it further."""
    w, h = image.size
    y0 = min(row_index * row_height, h)
    y1 = min(y0 + row_height, h)
    x1 = min(x_max, w)
    count = 0
    for y in range(y0, y1):
        for x in range(0, x1):
            if _is_greenish(image.getpixel((x, y))):
                count += 1
                if count >= min_green_px:
                    return True
    return False


# --------------------------------------------------------- number/text utils

_MARKUP_TAG_RE = re.compile(r"<[^>]*>")


def strip_ui_markup(text):
    """1C highlights search-matched substrings within a grid cell's displayed
    text using embedded pseudo-HTML tags, e.g.
    '<b><colorstyle -46>ON</></>-<b><colorstyle -46>26</></>/...' - and
    window_text() returns that markup literally instead of the plain
    rendered text. Confirmed live: this silently broke a plain
    text.startswith(order_number) check right after searching for that same
    order_number (the highlighted cell no longer "starts with" the plain
    string once the tags are spliced in). Strip it before comparing text
    read from any grid cell against a known plain value."""
    return _MARKUP_TAG_RE.sub("", text or "")

def parse_ru_number(text):
    """Parse a 1C grid cell's raw numeric text (e.g. '3049000,00' or '1,000')
    into a float. 1C renders these with a comma decimal separator and no
    thousands separator (unlike the Google Sheet's '3 049 000,00' display -
    see sheets_client.parse_sheet_amount for that variant)."""
    s = (text or "").strip().replace("\xa0", "").replace(" ", "")
    if not s:
        return None
    return float(s.replace(",", "."))


class MultipleItemRowsError(Exception):
    """Raised by read_single_item_row when a document's Товары grid has other
    than exactly one item row. The multi-item-row case is a known gap - not
    yet supported, see automate_order_ekvayring.py's module docstring."""
    def __init__(self, n_rows):
        self.n_rows = n_rows
        super().__init__(f"expected exactly 1 item row, found {n_rows}")


def read_item_rows(win):
    """Read EVERY item row on whatever document's Товары tab is currently
    showing (works for both Заказ клиента and Реализация - both render the
    same grid shape, confirmed live). Unlike read_single_item_row, does not
    raise for a multi-item document - used for the multi-item-order flow
    (see automate_order_ekvayring.determine_item_price_pairs). Filtering
    table.children() by a column-suffix and indexing by row position [i]
    reliably lines up cells for the same row across columns - confirmed
    live and already relied on by the older fix_realizatsiya_prices."""
    tables = win.descendants(control_type="Table")
    if not tables:
        raise RuntimeError("Товары grid (Table) not found")
    kids = tables[0].children()

    def ending_with(suffix):
        # A cell with no value at all (e.g. Цена right after switching Вид
        # цены to <произвольная> resets it blank - confirmed live) renders as
        # just the bare column label with no leading space, e.g. "Цена" not
        # " Цена" - match that bare form too, not only "<value> Цена".
        bare = suffix.strip()
        return [k for k in kids if (t := k.window_text().strip()) == bare or t.endswith(suffix)]

    def value_of(cell, suffix):
        text = cell.window_text().strip()
        if text == suffix.strip():
            return ""
        return text[: -len(suffix)].strip() if text.endswith(suffix) else text

    price_cells = ending_with(" Цена")
    nom_cells = ending_with(" Номенклатура")
    kol_cells = ending_with(" Количество")
    vid_cells = [k for k in kids if "Вид цены" in k.window_text()]

    rows = []
    for i in range(len(price_cells)):
        kolichestvo_raw = value_of(kol_cells[i], " Количество") if i < len(kol_cells) else ""
        tsena_raw = value_of(price_cells[i], " Цена")
        rows.append({
            "index": i,
            "nomenklatura": value_of(nom_cells[i], " Номенклатура") if i < len(nom_cells) else "",
            "kolichestvo_raw": kolichestvo_raw,
            "kolichestvo": parse_ru_number(kolichestvo_raw),
            "vid_tseny_cell": vid_cells[i] if i < len(vid_cells) else None,
            "vid_tseny": vid_cells[i].window_text().strip() if i < len(vid_cells) else "",
            "price_cell": price_cells[i],
            "tsena_raw": tsena_raw,
            "tsena": parse_ru_number(tsena_raw),
        })
    return rows


def read_single_item_row(win):
    """Read the single item row on whatever document's Товары tab is
    currently showing. Raises MultipleItemRowsError if there isn't exactly
    one row - callers should treat that as a stop condition, not silently
    loop over rows (see module docstring)."""
    rows = read_item_rows(win)
    if len(rows) != 1:
        raise MultipleItemRowsError(len(rows))
    return rows[0]


# ------------------------------------------------- fuzzy product-name match

_NAME_TOKEN_RE = re.compile(r"[^a-z0-9а-яё]+")

# A trailing " — <descriptor>" marketing suffix: a SPACED dash (em-dash,
# en-dash, or plain hyphen with spaces around it) followed by trailing
# text. Confirmed live these appear on the source-side name (Google Sheet /
# downloaded report) but never in 1C's Номенклатура, and are pure variant
# descriptors ("— standard", "— Black") that dragged genuine matches below
# the auto-match threshold, forcing needless human confirms:
#   "Samsung ME88SUB/UZ — standard"  vs  "Samsung ME88SUB/UZ"        -> 0.83
#   "Ninebot KickScooter D28E — Black" vs "Ninebot KickScooter D28E" -> 0.80
#   "Premier PRM-460BFNF/DI — Black" vs "...PRM-460BFNF/DI (чёрный)" -> 0.83
# The real identity (brand, product type, model number) is always BEFORE
# the dash, so dropping the suffix makes all of these score 1.0. The dash
# must be SPACED (" - "/" — "/" – ") so model numbers with internal hyphens
# and no surrounding spaces (e.g. "PRM-460BFNF") are never split.
_VARIANT_SUFFIX_RE = re.compile(r"\s[—–-]\s.*$")


def _strip_variant_suffix(name):
    return _VARIANT_SUFFIX_RE.sub("", name or "")


# Cyrillic letters that are glyph-IDENTICAL to a Latin lowercase letter -
# mapped Cyrillic -> Latin so a token typed in mixed/either script still
# compares equal (confirmed real risk: a dimension like "305х76" can have a
# Cyrillic 'х' in one system and a Latin 'x' in the other, which would
# otherwise read as two different tokens). Only the visually-identical
# pairs are listed, so this can only unify true look-alikes, never merge
# genuinely distinct characters; and it never changes a token's digit
# content, so the letters-AND-digits model-token test below is unaffected.
_HOMOGLYPHS = str.maketrans({
    "а": "a", "е": "e", "о": "o", "р": "p", "с": "c", "у": "y", "х": "x",
})


def normalize_product_name(name):
    s = (name or "").lower().translate(_HOMOGLYPHS)
    s = _NAME_TOKEN_RE.sub(" ", s)
    return re.sub(r"\s+", " ", s).strip()


def _is_model_token(token):
    """A token that contains BOTH a letter and a digit (e.g. 'a56',
    'ms23f301taw', '256gb', '460bfnf') - a strong product identifier, as
    opposed to plain words ('samsung', 'смартфон') or bare numbers ('8')."""
    return any(c.isalpha() for c in token) and any(c.isdigit() for c in token)


def product_name_coverage(sheet_name, c1_name):
    """Fraction of sheet_name's tokens that appear (as a substring either way,
    so '256' matches '256gb') somewhere in c1_name's tokens. Both names have
    any trailing " — <descriptor>" variant suffix stripped first (see
    _VARIANT_SUFFIX_RE).

    Model-number guard: a token containing both letters and digits (see
    _is_model_token) is a strong identifier - "Galaxy A56" vs "A17",
    "8/256GB" vs "8/512GB". If ANY such token in sheet_name has no
    counterpart in c1_name, the two are almost certainly different products,
    so the score is capped WELL below the auto-match threshold - otherwise
    "A56" vs "A17" scores 0.86 (just over 0.85) purely on shared
    "5g/8/256gb" boilerplate and would auto-match the wrong product. This
    also keeps a wrong-model candidate ranked below a true match (which has
    no such mismatch) in the greedy multi-item pairing. Plain words that
    differ (a color, "standard") are NOT model tokens, so genuine variants
    still separate by ordinary coverage, not this hard cap.

    Even with all this, coverage is a heuristic prefilter for confident
    auto-matches, not a verdict - see names_match_or_confirm's caller,
    which asks a human whenever coverage isn't high."""
    a_tokens = normalize_product_name(_strip_variant_suffix(sheet_name)).split()
    b_tokens = normalize_product_name(_strip_variant_suffix(c1_name)).split()
    if not a_tokens:
        return 0.0

    def matches(t):
        return any(t in bt or bt in t for bt in b_tokens)

    coverage = sum(1 for t in a_tokens if matches(t)) / len(a_tokens)
    model_tokens = [t for t in a_tokens if _is_model_token(t)]
    if model_tokens and not all(matches(t) for t in model_tokens):
        return min(coverage, 0.5)
    return coverage

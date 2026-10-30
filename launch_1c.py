# -*- coding: utf-8 -*-
"""Launches the 1C thin client and logs in automatically, so
automate_order_ekvayring.py doesn't require the user to have already
opened 1C and signed in by hand.

Copied from emit_transactions/launch_1c.py (that project has its own
self-contained copy - same "own copy, not a cross-import" pattern already
used for lib1c.py between the two - see alif_sheets_client.py's docstring
for the same reasoning applied elsewhere in THIS project) and adapted to
read credentials via this project's env_config.py instead of its own
ad-hoc regex .env parsing.



The production infobase is web-published - it's registered in
the user's 1C list (%APPDATA%\\1C\\1CEStart\\ibases.v8i) as
    ws="https://<your-1c-server>/<infobase>";
(set as 1c_connection_string in credentials\\.env).
We launch 1cv8c.exe (the thin client) straight at that connection string
and pass the 1C user/password on the command line (/N /P), which 1C uses
to fill and submit its own login form - no clicking through the login
window needed.

Credentials come from credentials\\.env (1c_login / 1c_password), the same
folder/format the rest of this project already uses. Nothing here is
destructive: launching and signing in is exactly what the user does
manually every day; the reconciliation edits happen later and are gated
separately (confirm() prompts unless --yes).
"""
import ctypes
import subprocess
import time
from pathlib import Path

from env_config import load_env
from lib1c import WINDOW_TITLE_HINTS
from pywinauto import Desktop

_user32 = ctypes.windll.user32

# The 1C thin client this project was built and confirmed against. Still the
# first choice when it's installed, but no longer the ONLY one - see
# find_thin_client: the same project now ships as .exe builds meant to run on
# other laptops, which can have a different 1C version (or only a different
# one) installed.
THIN_CLIENT_EXE = r"C:\Program Files\1cv8\8.3.13.1644\bin\1cv8c.exe"


def _version_key(path):
    r"""Sort key for a ...\1cv8\<version>\bin\1cv8c.exe path by that version
    segment - numerically, so 8.3.22 ranks above 8.3.9."""
    try:
        return tuple(int(p) for p in Path(path).parents[1].name.split("."))
    except ValueError:
        return ()


def find_thin_client():
    r"""Path of the 1C thin client to launch, or None if none is installed.

    In order: an explicit 1c_thin_client=<full path> in credentials\.env (for
    a machine with 1C somewhere non-standard); then THIN_CLIENT_EXE, the
    version this was confirmed against; then the NEWEST 1cv8c.exe found under
    either Program Files tree. Picking a different version than the pinned one
    is a fallback, not a silent swap - the caller prints which one it used."""
    override = (load_env().get("1c_thin_client") or "").strip().strip('"')
    if override:
        return override if Path(override).exists() else None
    if Path(THIN_CLIENT_EXE).exists():
        return THIN_CLIENT_EXE
    candidates = []
    for root in (r"C:\Program Files\1cv8", r"C:\Program Files (x86)\1cv8"):
        candidates.extend(str(p) for p in Path(root).glob(r"*\bin\1cv8c.exe"))
    return max(candidates, key=_version_key) if candidates else None

# Production infobase web-publish URL (from ibases.v8i), e.g.
# 1c_connection_string=ws=https://<your-1c-server>/<infobase>;
# Unquoted in the connection string on purpose: the URL has no spaces or
# semicolons, so quoting it would only create command-line quote-escaping
# headaches for no benefit.
ONEC_CONNECTION_STRING = (load_env().get("1c_connection_string") or "").strip()


def find_main_window():
    """Return the 1C main-window wrapper if it's open, else None (unlike
    lib1c.get_window, which raises). Same disambiguation as lib1c.
    get_window (see its docstring) - confirmed live that a browser tab left
    open on the web-published infobase can loosely title-match just as
    well as the real 1C window, so a plain "first loose match wins" isn't
    reliable."""
    windows = []
    for w in Desktop(backend="uia").windows():
        try:
            w.window_text()
        except Exception:
            continue
        windows.append(w)

    all_hint_matches = [w for w in windows if all(h in w.window_text() for h in WINDOW_TITLE_HINTS)]
    if len(all_hint_matches) == 1:
        return all_hint_matches[0]
    if len(all_hint_matches) > 1:
        return None  # genuinely ambiguous - treat as "not confirmed open" rather than guessing

    loose_matches = [w for w in windows if any(h in w.window_text() for h in WINDOW_TITLE_HINTS)]
    return loose_matches[0] if len(loose_matches) == 1 else None


def ensure_1c_open(timeout=120, poll=1.5):
    """Make sure the 1C main window is open and logged in, launching+signing
    in if needed. Idempotent: if 1C is already up, returns immediately
    without starting a second instance. Returns the main-window wrapper."""
    existing = find_main_window()
    if existing is not None:
        print("-> 1C is already open.")
        return existing

    env = load_env()
    login = env.get("1c_login")
    password = env.get("1c_password")
    if not login or not password:
        raise RuntimeError(
            "1c_login / 1c_password not found in credentials\\.env - can't auto-launch 1C."
        )
    if not ONEC_CONNECTION_STRING:
        raise RuntimeError(
            "1c_connection_string not found in credentials\\.env - can't auto-launch 1C."
        )
    thin_client = find_thin_client()
    if not thin_client:
        raise RuntimeError(
            "1C thin client (1cv8c.exe) not found - install 1C, or set 1c_thin_client=<full path to "
            "1cv8c.exe> in credentials\\.env, or just open 1C and log in by hand before running this."
        )

    print("-> 1C is not open - launching and logging in...")
    if thin_client != THIN_CLIENT_EXE:
        print(f"   (using {thin_client} - not the {THIN_CLIENT_EXE} this was confirmed against)")
    subprocess.Popen([
        thin_client,
        "ENTERPRISE",
        "/IBConnectionString", ONEC_CONNECTION_STRING,
        "/N", login,
        "/P", password,
        "/DisableStartupDialogs",   # suppress "what's new" / update nags
        "/DisableStartupMessages",
    ])

    end = time.time() + timeout
    while time.time() < end:
        win = find_main_window()
        if win is not None:
            print("-> 1C's main window appeared - waiting for it to actually become responsive...")
            _wait_until_responsive(win.handle)
            time.sleep(3)  # small extra buffer once responsive, before driving the UI
            print("-> 1C opened, logged in, and responsive.")
            return win
        time.sleep(poll)

    raise RuntimeError(
        f"1C did not reach its main window within {timeout}s after launch - "
        "the login may have failed (check credentials) or the server is slow."
    )


def _wait_until_responsive(hwnd, timeout=90, poll=1.5):
    """The main window can exist well before 1C is actually ready to
    receive input - confirmed live: driving it (Продажи -> Заказы
    клиентов etc.) right after a fresh cold-start launch, with only a
    fixed 2s sleep before starting, hung for a very long time inside what
    turned out to be a blocked UIA/COM call (a real 1C client can spend
    10s of seconds after its window first appears loading
    config/caches/sync on first login of the day) - and interrupting a
    hung COM call with Ctrl+C corrupted the interpreter's shutdown instead
    of raising a clean KeyboardInterrupt. IsHungAppWindow is the right
    Win32-level check for this: unlike a UIA call, it can't itself hang -
    it just asks the OS whether the window's message loop is currently
    responding, so this is safe to poll even while 1C is busy."""
    end = time.time() + timeout
    while time.time() < end:
        if not _user32.IsHungAppWindow(hwnd):
            return
        time.sleep(poll)
    print(f"   (1C's window still isn't responding after {timeout}s - proceeding anyway, "
          "the next step may be slow or need a retry)")


if __name__ == "__main__":
    from lib1c import get_window
    ensure_1c_open()
    print(get_window().window_text())

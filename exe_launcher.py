# -*- coding: utf-8 -*-
"""Double-click front end shared by the two .exe builds (IMAN.exe,
AlifMoliya.exe - see IMAN.py / AlifMoliya.py and build_exe.bat).

automate_order_ekvayring.py itself is a command-line tool: it expects its
flags (--profile, --since, --yes) on the command line and, once finished,
simply exits. Double-clicked as an .exe that would mean no way to pass a date,
and a console window that vanishes the instant the batch report is printed -
taking the report with it. So this:

  - checks up front that what the run needs is actually next to the .exe
    (the credentials folder, and for Alif Moliya the browser), and says so in
    plain words if not, instead of failing half-way into a run;
  - asks for the date interactively, with a sensible default on Enter;
  - runs the real main() with the equivalent command line;
  - always waits for Enter before the window closes, whatever happened.

Running an .exe with --selftest checks the same things without touching 1C or
any Google Sheet (and for Alif Moliya opens and closes the bundled browser
once) - the quick way to confirm a copy on another laptop is complete.
"""
import datetime
import os
import sys
import traceback
from pathlib import Path

# Must be importable without automate_order_ekvayring: that module imports
# playwright at the top, and PLAYWRIGHT_BROWSERS_PATH (set below) only takes
# effect if it's in the environment BEFORE playwright is first imported.
from env_config import CREDENTIALS_DIR, app_dir

DATE_FORMAT = "%d.%m.%Y"

# The browser ships as a plain folder beside AlifMoliya.exe rather than inside
# it: ~400 MB unpacked into a temp folder on EVERY launch would add tens of
# seconds to each start, and as a folder it can also be refreshed without a
# rebuild. Playwright finds it through PLAYWRIGHT_BROWSERS_PATH.
BROWSERS_DIR_NAME = "ms-playwright"


def _setup_console(title):
    # A real console renders Cyrillic fine; errors="replace" is only there so
    # that one unprintable character can never crash a run half-way through.
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(errors="replace")
        except Exception:
            pass
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.kernel32.SetConsoleTitleW(title)
        except Exception:
            pass


def _point_playwright_at_bundled_browser():
    browsers = app_dir() / BROWSERS_DIR_NAME
    if browsers.is_dir():
        os.environ["PLAYWRIGHT_BROWSERS_PATH"] = str(browsers)
        return browsers
    return None


def _missing_requirements(needs_browser):
    """Plain-language list of what's missing next to the .exe ([] if none)."""
    problems = []
    if not CREDENTIALS_DIR.is_dir():
        problems.append(f"нет папки credentials рядом с программой (ожидается {CREDENTIALS_DIR})")
    else:
        for name in (".env", "service_account.json"):
            if not (CREDENTIALS_DIR / name).is_file():
                problems.append(f"в папке credentials нет файла {name}")
    if needs_browser and not (app_dir() / BROWSERS_DIR_NAME).is_dir():
        problems.append(f"нет папки {BROWSERS_DIR_NAME} рядом с программой (браузер для сайта Alif)")
    return problems


def _first_of_previous_month(today):
    first_this = today.replace(day=1)
    return (first_this - datetime.timedelta(days=1)).replace(day=1)


def _ask_date(prompt, default):
    """DD.MM.YYYY from the keyboard. Enter alone takes `default` (which may
    be None, meaning "no date"). Re-asks on anything unparseable rather than
    guessing what was meant."""
    while True:
        hint = default.strftime(DATE_FORMAT) if default else "без даты"
        raw = input(f"{prompt} [Enter = {hint}]: ").strip()
        if not raw:
            return default
        try:
            return datetime.datetime.strptime(raw, DATE_FORMAT).date()
        except ValueError:
            print("   Не понял дату - нужно в виде ДД.ММ.ГГГГ, например 01.09.2026.")


def _selftest(needs_browser):
    print("Проверка установки (1С и таблицы не трогаются)\n")
    ok = True
    print(f"  папка программы : {app_dir()}")
    print(f"  credentials     : {CREDENTIALS_DIR}")
    for p in _missing_requirements(needs_browser):
        print(f"  !! {p}")
        ok = False

    try:
        import automate_order_ekvayring  # noqa: F401 - the import itself is the check
        print("  модули          : OK")
    except Exception as e:
        print(f"  !! модули не загрузились: {e!r}")
        return False

    import launch_1c
    client = launch_1c.find_thin_client()
    if client:
        print(f"  клиент 1С       : {client}")
    else:
        print("  !! клиент 1С (1cv8c.exe) не найден - либо установите 1С, либо откройте её "
              "вручную перед запуском")
        ok = False

    # Read-only on purpose: get_report_worksheet() would (re)write the header
    # row, so this only fetches the spreadsheet's title - enough to prove the
    # service-account key, the network and gspread all work in this build.
    try:
        import report_sheet
        if report_sheet.is_configured():
            book = report_sheet._client().open_by_key(report_sheet.REPORT_SPREADSHEET_ID)
            print(f"  Google-таблицы  : OK (доступ к «{book.title}»)")
        else:
            print("  Google-таблицы  : REPORT_SPREADSHEET_URL не задан - отчёт прогона писаться не будет")
    except Exception as e:
        print(f"  !! нет доступа к Google-таблицам: {e!r}")
        ok = False

    # Exercises UI Automation through comtypes - the part of a frozen build
    # most likely to be missing something, and the one every 1C step needs.
    try:
        window = launch_1c.find_main_window()
        print(f"  окно 1С         : {'открыто' if window else 'не открыто (запустится само)'}")
    except Exception as e:
        print(f"  !! не работает доступ к окнам (UI Automation): {e!r}")
        ok = False

    if needs_browser:
        try:
            from playwright.sync_api import sync_playwright
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=False)
                browser.new_page().goto("about:blank")
                browser.close()
            print("  браузер         : OK (открылся и закрылся)")
        except Exception as e:
            print(f"  !! браузер не запустился: {e!r}")
            ok = False

    print("\nВсё готово к работе." if ok else "\nЕсть проблемы - см. строки с !! выше.")
    return ok


def run(title, profile, since_mode, needs_browser):
    """since_mode: "required" (IMAN - its batch mode needs a start date) or
    "optional" (Alif Moliya - Enter alone means every currently-open order)."""
    _setup_console(title)
    _point_playwright_at_bundled_browser()

    try:
        if "--selftest" in sys.argv[1:]:
            _selftest(needs_browser)
            return

        print("=" * 64)
        print(f"  {title}")
        print("=" * 64)

        problems = _missing_requirements(needs_browser)
        if problems:
            print("\nНе могу запуститься:")
            for p in problems:
                print(f"  - {p}")
            print("\nСкопируйте эти папки рядом с программой и запустите снова.")
            return

        today = datetime.date.today()
        if since_mode == "required":
            since = _ask_date("С какой даты сверять", _first_of_previous_month(today))
        else:
            since = _ask_date("С какой даты брать заказы", None)

        argv = ["--profile", profile, "--yes"]
        if since:
            argv += ["--since", since.strftime(DATE_FORMAT)]
        print()
        print("Запускаю. Пока идёт работа - НЕ трогайте мышь и клавиатуру:")
        print("программа сама управляет окном 1С. Остановить - Ctrl+C.")
        print()

        import automate_order_ekvayring
        sys.argv = [sys.argv[0]] + argv
        automate_order_ekvayring.main()
    except SystemExit as e:
        # argparse's own errors, and main()'s deliberate early exits, land
        # here - the message (if any) has already been printed.
        if e.code not in (None, 0):
            print(f"\n(остановлено, код {e.code})")
    except KeyboardInterrupt:
        print("\nОстановлено (Ctrl+C).")
    except Exception:
        print("\n!! Программа упала с ошибкой:\n")
        traceback.print_exc()
    finally:
        try:
            input("\nНажмите Enter, чтобы закрыть окно...")
        except (EOFError, KeyboardInterrupt):
            pass

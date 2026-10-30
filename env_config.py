# -*- coding: utf-8 -*-
"""Shared credentials\\.env loading for every module in this project
(sheets_client.py, alif_client.py, alif_sheets_client.py) - one file, one
parser, so IDs/URLs/logins/passwords all live in the SAME place instead of
being hardcoded across the source files. credentials\\.env (and this whole
folder) is never shared/committed anywhere (see credentials\\README.txt).

A missing .env or a missing individual key never breaks the import: each
module that reads a value falls back to an empty/neutral default (or a
public site URL, see alif_client.ALIF_BUSINESS_BASE_URL) and only fails
when that value is actually needed. credentials\\.env is the SOURCE OF
TRUTH - see .env.example in the repo root for every key.
"""
import sys
from pathlib import Path


def app_dir():
    """The folder this project runs FROM: next to the .exe when built with
    PyInstaller (IMAN.exe / AlifMoliya.exe), next to these .py files
    otherwise.

    The distinction matters because a PyInstaller --onefile build unpacks
    itself into a fresh temp folder on every launch and __file__ points in
    there - so a Path(__file__)-relative credentials folder would be looked
    for inside a throwaway directory that never contains it, and anything
    written there (the saved Alif / Uzum browser sessions) would be thrown
    away when the run ends. sys.executable is the .exe the user actually
    double-clicked."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


# Deliberately NOT bundled into the .exe builds: this folder holds logins,
# passwords and the Google service-account key, so it lives beside the .exe
# where it can be updated without a rebuild and never travels inside a
# binary that might get passed around.
CREDENTIALS_DIR = app_dir() / "credentials"
ENV_PATH = CREDENTIALS_DIR / ".env"


def load_env(path=ENV_PATH):
    """Parse credentials\\.env into a dict - plain KEY=VALUE lines, '#'
    comments and blank lines ignored. Returns {} if the file doesn't exist
    yet (callers fall back to their own defaults in that case - see this
    module's docstring) rather than raising, since not every caller
    strictly requires it (e.g. sheets_client.py's IMAN_SPREADSHEET_ID has
    always worked with zero .env setup)."""
    if not path.exists():
        return {}
    env = {}
    for line in Path(path).read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        env[k.strip()] = v.strip()
    return env

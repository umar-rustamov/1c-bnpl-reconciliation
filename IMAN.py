# -*- coding: utf-8 -*-
"""Entry point for IMAN.exe - the IMANUM STORE sverka, batch mode.
See exe_launcher.py for what the .exe front end does, build_exe.bat to build."""
import exe_launcher

if __name__ == "__main__":
    exe_launcher.run(
        title="IMAN - эквайринг (IMANUM STORE)",
        profile="imanum",
        since_mode="required",
        needs_browser=False,   # this profile reads a Google Sheet, never a website
    )

# -*- coding: utf-8 -*-
"""Entry point for AlifMoliya.exe - every currently-open Alif Moliya order,
batch mode. See exe_launcher.py for what the .exe front end does,
build_exe.bat to build."""
import exe_launcher

if __name__ == "__main__":
    exe_launcher.run(
        title="Alif Moliya - эквайринг",
        profile="alif-moliya",
        since_mode="optional",
        needs_browser=True,    # prices come from business.alif.uz's own report
    )

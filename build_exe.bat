@echo off
rem Builds IMAN.exe and AlifMoliya.exe and assembles a folder ready to copy
rem to another laptop:  dist\1C-Ekvayring\
rem
rem   IMAN.exe, AlifMoliya.exe   the two programs (see exe_launcher.py)
rem   credentials\               copied from here - logins, passwords, Google key
rem   ms-playwright\             the browser AlifMoliya.exe drives on business.alif.uz
rem
rem Needs Python with this project's packages AND pyinstaller installed, on
rem THIS machine only - the laptop the folder is copied to needs no Python.
rem
rem --collect-all comtypes bundles the UI Automation wrappers comtypes has
rem already generated here (comtypes\gen\UIAutomationClient.py ...). Without
rem them a frozen build has to generate them on its first run, which is the
rem usual way pywinauto breaks inside an .exe.

setlocal
cd /d "%~dp0"

set OUT=dist\1C-Ekvayring
set PYI=python -m PyInstaller --noconfirm --onefile --console --distpath "%OUT%" --workpath build --specpath build --collect-all comtypes --collect-submodules pywinauto

echo === building IMAN.exe ===
%PYI% --name IMAN IMAN.py || goto :fail
echo === building AlifMoliya.exe ===
%PYI% --name AlifMoliya AlifMoliya.py || goto :fail

echo === copying credentials ===
robocopy credentials "%OUT%\credentials" /E /NFL /NDL /NJH /NJS >nul
if errorlevel 8 goto :fail

echo === copying the browser for AlifMoliya.exe ===
rem Only chromium-NNNN: every launch here is headed (headless=False), so the
rem separate chromium_headless_shell-NNNN download is never used.
for /d %%D in ("%LOCALAPPDATA%\ms-playwright\chromium-*") do (
    robocopy "%%D" "%OUT%\ms-playwright\%%~nxD" /E /NFL /NDL /NJH /NJS >nul
    if errorlevel 8 goto :fail
)
if not exist "%OUT%\ms-playwright" (
    echo !! no Playwright browser found in %LOCALAPPDATA%\ms-playwright
    echo    run:  python -m playwright install chromium   and build again
    goto :fail
)

copy /y README_EXE.txt "%OUT%\README.txt" >nul

echo.
echo Done: %OUT%
echo Copy that whole folder to the other laptop. It contains PASSWORDS - don't share it.
exit /b 0

:fail
echo.
echo !! build failed - see the messages above
exit /b 1

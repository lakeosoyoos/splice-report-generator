@echo off
REM ===========================================================================
REM  Splice Report — local Windows one-click build
REM ===========================================================================
REM  Run this from a Windows machine with Python 3.11 installed (NOT 3.12).
REM  Produces dist\SpliceReport\SpliceReport.exe + a zip alongside it.
REM
REM  CI does the same steps via .github/workflows/build-windows.yml plus a
REM  boot self-test.  This local script skips the boot test — for an
REM  authoritative answer on whether the build launches, push and let CI run.
REM ===========================================================================

setlocal ENABLEDELAYEDEXPANSION
cd /d "%~dp0"

REM ── 1. Confirm Python 3.11 ────────────────────────────────────────────────
set "PY=py -3.11"
%PY% --version >nul 2>&1
if errorlevel 1 (
    set "PY=python"
    python --version 2>nul | findstr /R "3\.11\." >nul
    if errorlevel 1 (
        echo.
        echo [build.bat] ERROR — Python 3.11 was not found.
        echo             Install Python 3.11 from python.org and re-run.
        echo             3.12+ will compile green but the exe will crash at launch.
        exit /b 1
    )
)
echo [build.bat] Using: %PY%
%PY% --version

REM ── 2. Fresh venv ─────────────────────────────────────────────────────────
if exist .venv (
    echo [build.bat] Removing old .venv ...
    rmdir /s /q .venv
)
%PY% -m venv .venv
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip wheel

REM ── 3. Install deps + re-pin setuptools LAST (so it wins) ─────────────────
pip install -r requirements-desktop.txt
pip install --force-reinstall setuptools==65.5.1

REM ── 4. PyInstaller build ──────────────────────────────────────────────────
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist
pyinstaller SpliceReport.spec --noconfirm --clean
if errorlevel 1 (
    echo.
    echo [build.bat] ERROR — PyInstaller failed.
    exit /b 1
)

REM ── 5. Zip up the dist folder ─────────────────────────────────────────────
powershell -NoProfile -Command ^
    "Compress-Archive -Path 'dist\SpliceReport\*' -DestinationPath 'dist\SpliceReport-Windows.zip' -Force"

echo.
echo [build.bat] ===========================================================
echo [build.bat]  Build OK.
echo [build.bat]    EXE : %CD%\dist\SpliceReport\SpliceReport.exe
echo [build.bat]    ZIP : %CD%\dist\SpliceReport-Windows.zip
echo [build.bat] ===========================================================
echo [build.bat]  REMEMBER — local build is NOT verified to launch.
echo [build.bat]  Push to GitHub and let CI run the boot self-test before
echo [build.bat]  shipping the zip to a tech.
echo [build.bat] ===========================================================

endlocal

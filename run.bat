@echo off
title Word Mail Merge - Setup and Run

cd /d "%~dp0"

echo.
echo ============================================================
echo   Word Mail Merge - Setup and Run
echo ============================================================
echo.

REM ---------- 1. Check if Python is already installed ----------
set "PYTHON_CMD="
where python >nul 2>&1
if %errorlevel%==0 set "PYTHON_CMD=python"

if not defined PYTHON_CMD (
    where py >nul 2>&1
    if %errorlevel%==0 set "PYTHON_CMD=py"
)

if defined PYTHON_CMD goto python_found

REM ---------- 2. Python not found - try to install ----------
echo [1/5] Python not found. Trying to install...
echo.

REM Try winget first
where winget >nul 2>&1
if %errorlevel%==0 goto use_winget
goto use_direct

:use_winget
echo      Using winget to install Python...
winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
call :refresh_path
goto verify_python

:use_direct
set "PY_INSTALLER=%TEMP%\python_installer.exe"
echo      Downloading Python installer...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Invoke-WebRequest -Uri 'https://www.python.org/ftp/python/3.12.7/python-3.12.7-amd64.exe' -OutFile '%PY_INSTALLER%' -UseBasicParsing; exit 0 } catch { exit 1 }"

if not exist "%PY_INSTALLER%" (
    echo.
    echo [ERROR] Could not download Python installer.
    echo         Please install Python manually from https://www.python.org/downloads/
    echo         Make sure to check "Add Python to PATH" during installation.
    echo.
    pause
    exit /b 1
)

echo      Running Python installer...
"%PY_INSTALLER%" /quiet InstallAllUsers=0 PrependPath=1 Include_test=0 Include_pip=1
call :refresh_path

:verify_python
timeout /t 3 /nobreak >nul

set "PYTHON_CMD="
where python >nul 2>&1
if %errorlevel%==0 set "PYTHON_CMD=python"

if not defined PYTHON_CMD (
    where py >nul 2>&1
    if %errorlevel%==0 set "PYTHON_CMD=py"
)

if not defined PYTHON_CMD (
    echo.
    echo [ERROR] Python installed but not found in PATH.
    echo         Close this window and run run.bat again.
    echo.
    pause
    exit /b 1
)

:python_found
echo [OK] Python found: %PYTHON_CMD%
%PYTHON_CMD% --version
echo.

REM ---------- 3. Bypass PowerShell execution policy ----------
echo [2/5] Allowing PowerShell scripts for current user...
powershell -NoProfile -ExecutionPolicy Bypass -Command "try { Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force; Write-Host '      Execution policy set.' } catch { Write-Host '      Skipped.' }"
echo.

REM ---------- 4. Create virtual environment ----------
echo [3/5] Setting up virtual environment...
if not exist "venv\Scripts\python.exe" (
    %PYTHON_CMD% -m venv venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        pause
        exit /b 1
    )
    echo      venv created.
) else (
    echo      venv already exists.
)
echo.

REM ---------- 5. Install requirements ----------
echo [4/5] Installing requirements...
if exist "requirements.txt" (
    "venv\Scripts\python.exe" -m pip install --upgrade pip --quiet
    "venv\Scripts\python.exe" -m pip install -r requirements.txt
) else (
    echo      No requirements.txt found. Installing defaults...
    "venv\Scripts\python.exe" -m pip install --upgrade pip --quiet
    "venv\Scripts\python.exe" -m pip install python-docx pandas openpyxl xlrd
)
echo      Requirements installed.
echo.

REM ---------- 6. Run the app ----------
echo [5/5] Starting Word Mail Merge...
echo.
echo ============================================================
echo.

if not exist "word_placement.py" (
    echo [ERROR] word_placement.py not found in this folder.
    pause
    exit /b 1
)

"venv\Scripts\python.exe" word_placement.py

if errorlevel 1 (
    echo.
    echo [ERROR] The app exited with an error.
    pause
    exit /b 1
)

echo.
echo App closed.
pause
exit /b 0

REM ================================================================
REM  Refresh PATH from registry
REM ================================================================
:refresh_path
for /f "tokens=2*" %%A in ('reg query "HKCU\Environment" /v Path 2^>nul') do set "USER_PATH=%%B"
for /f "tokens=2*" %%A in ('reg query "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment" /v Path 2^>nul') do set "SYS_PATH=%%B"
set "PATH=%SYS_PATH%;%USER_PATH%"
exit /b 0
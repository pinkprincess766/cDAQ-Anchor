@echo off
setlocal
cd /d "%~dp0"

set "PY_CMD="
where py >nul 2>nul
if not errorlevel 1 set "PY_CMD=py"

if "%PY_CMD%"=="" (
    where python >nul 2>nul
    if not errorlevel 1 set "PY_CMD=python"
)

if "%PY_CMD%"=="" (
    echo Python was not found. Install Python 3.10+ and enable "Add python.exe to PATH".
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment...
    %PY_CMD% -m venv .venv
    if errorlevel 1 exit /b %errorlevel%

    echo Installing dependencies...
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 exit /b %errorlevel%
)

".venv\Scripts\python.exe" main.py %*

@echo off
REM ============================================================================
REM SCHEDULED-JOB LAUNCHER - Headless index scan (app-launcher chain hop)
REM ============================================================================
REM Runs main_scan.py --no-ui in the foreground and exits with its exit code:
REM   0 ok, 1 crash (or folder-explorer's index run failed), 2 config missing / no roots,
REM   3 an archive root is missing, 4 folder-explorer busy (its lock held past the wait).
REM Stream Deck keeps launch_scan.bat (windowed, detached).

setlocal
set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "PYTHON=%SCRIPT_DIR%\.venv\Scripts\python.exe"
set "PYTHONUTF8=1"
set "PYTHONUNBUFFERED=1"

cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change to directory: %SCRIPT_DIR%
    exit /b 1
)

if not exist "%PYTHON%" (
    echo [ERROR] Python not found: %PYTHON%
    exit /b 1
)

"%PYTHON%" main_scan.py --no-ui
exit /b %ERRORLEVEL%

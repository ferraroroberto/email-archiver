@echo off
REM ============================================================================
REM STREAM DECK LAUNCHER – Archive selected Outlook email (no console window)
REM ============================================================================

set "SCRIPT_DIR=%~dp0"
set "SCRIPT_DIR=%SCRIPT_DIR:~0,-1%"
set "VENV_DIR=%SCRIPT_DIR%\.venv"

REM Ensure we run from project root so config and imports resolve
cd /d "%SCRIPT_DIR%"
if errorlevel 1 (
    echo [ERROR] Failed to change to directory: %SCRIPT_DIR%
    pause
    exit /b 1
)

REM Start Python without console; use /d so the child process has correct cwd
start "" /d "%SCRIPT_DIR%" "%VENV_DIR%\Scripts\pythonw.exe" main_archive.py

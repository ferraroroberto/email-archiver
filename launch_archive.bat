@echo off
:: Stream Deck launcher – Archive selected Outlook email (no console window)
cd /d "%~dp0"
start "" pythonw main_archive.py

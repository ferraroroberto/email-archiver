@echo off
:: Stream Deck launcher – Scan and index the archive (no console window)
cd /d "%~dp0"
start "" pythonw main_scan.py

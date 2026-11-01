"""Stream Deck entry point: archive the selected Outlook email.

    pythonw main_archive.py   # no console window (recommended for Stream Deck)

Kept at this path for the launcher; the code is `email_archiver.cli.archive`.
"""
import sys
from pathlib import Path

# The package lives in src/; put it first on the path regardless of cwd
sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from email_archiver.cli.archive import main

if __name__ == "__main__":
    main()

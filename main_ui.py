"""Full launcher window: both 'Scan Archive' and 'Archive Email' buttons.

    pythonw main_ui.py   # no console window

Kept at this path for the launcher; the code is `email_archiver.cli.launcher`.
"""
import sys
from pathlib import Path

# The package lives in src/; put it first on the path regardless of cwd
sys.path.insert(0, str(Path(__file__).resolve().parent / "src"))

from email_archiver.cli.launcher import main

if __name__ == "__main__":
    main()

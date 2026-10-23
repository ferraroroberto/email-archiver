"""Scan and index the email archive; `--no-ui` runs headless.

    python main_scan.py --no-ui

Kept at this path for the nightly job and the launcher; the code, with its exit
codes, is `email_archiver.cli.scan`.
"""
import sys
from pathlib import Path

# Ensure project root is on the path regardless of cwd
sys.path.insert(0, str(Path(__file__).resolve().parent))

from email_archiver.cli.scan import main

if __name__ == "__main__":
    sys.exit(main())

"""
Stream Deck entry point – Archive selected Outlook email.

Optimised for instant perceived startup:
  1. Tkinter window is shown within ~100ms.
  2. Outlook query and DB lookup run in a background thread.
  3. Total time to first suggestion display: typically 1–3s.

Usage:
    python main_archive.py
    pythonw main_archive.py   # no console window (recommended for Stream Deck)
"""
import sys
from pathlib import Path

# Ensure project root is on the path regardless of cwd
sys.path.insert(0, str(Path(__file__).resolve().parent))

from email_archiver.config import load_config, setup_logging
from email_archiver.ui.app import ArchiveDialog


def main() -> None:
    cfg = load_config()
    setup_logging(cfg)
    ArchiveDialog(cfg).run()


if __name__ == "__main__":
    main()

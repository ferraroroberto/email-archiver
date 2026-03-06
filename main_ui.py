"""
Full launcher UI – both 'Scan Archive' and 'Archive Email' buttons.

Useful if you want a single Stream Deck button for the full app,
or for desktop use.

Usage:
    python main_ui.py
    pythonw main_ui.py   # no console window
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from email_archiver.config import load_config, setup_logging
from email_archiver.ui.app import LauncherApp


def main() -> None:
    cfg = load_config()
    setup_logging(cfg)
    LauncherApp(cfg).run()


if __name__ == "__main__":
    main()

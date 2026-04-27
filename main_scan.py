"""
Stream Deck entry point – Scan and index the email archive.

Can also be run from the command line for headless operation:
    python main_scan.py           # opens Tkinter progress window
    python main_scan.py --no-ui   # headless, prints progress to stdout
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from email_archiver.config import load_config, setup_logging


def main() -> None:
    parser = argparse.ArgumentParser(description="Scan and index the email archive.")
    parser.add_argument(
        "--no-ui", action="store_true",
        help="Run headless (no Tkinter window); progress printed to stdout."
    )
    args = parser.parse_args()

    cfg = load_config()
    setup_logging(cfg)

    if args.no_ui:
        _run_headless(cfg)
    else:
        from email_archiver.ui.app import ScanWindow
        ScanWindow(cfg).run()


def _run_headless(cfg: dict) -> None:
    import logging
    from email_archiver.scanner.scanner import FolderScanner

    logger = logging.getLogger(__name__)
    scanner = FolderScanner(cfg)

    def on_progress(current: int, total: int, path: str) -> None:
        if current % 500 == 0 or current == total:
            pct = f"{int(current / total * 100)}%" if total else ""
            print(f"\r  {current:>6,} / {total:,}  {pct}  ", end="", flush=True)

    logger.info("Scanning: %s", cfg['archive']['root_path'])
    stats = scanner.scan(progress_callback=on_progress)
    print()  # clear the \r progress line
    logger.info(
        "Done — %s new, %s updated, %s skipped, %s errors in %.1fs",
        stats.newly_indexed, stats.updated, stats.skipped, stats.errors, stats.duration_seconds,
    )


if __name__ == "__main__":
    main()

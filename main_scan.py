"""
Stream Deck entry point – Scan and index the email archive.

Can also be run from the command line for headless operation:
    python main_scan.py           # opens Tkinter progress window
    python main_scan.py --no-ui   # headless, prints progress to stdout

Exit codes (headless):
    0  scan complete (per-file read errors are counted, not a failure)
    1  unexpected crash (uncaught exception)
    2  config missing/unloadable, or no archive roots configured
    3  a configured archive root is missing; its index rows were left untouched
"""
import argparse
import logging
import sys
from pathlib import Path
from typing import Optional

import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent))

from email_archiver.config import get_archive_roots, load_config, setup_logging

logger = logging.getLogger(__name__)

EXIT_OK = 0
EXIT_CONFIG_ERROR = 2
EXIT_ROOT_MISSING = 3

_PROGRESS_EVERY = 500


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Scan and index the email archive.")
    parser.add_argument(
        "--no-ui", action="store_true",
        help="Run headless (no Tkinter window); progress printed to stdout."
    )
    args = parser.parse_args(argv)

    try:
        cfg = load_config()
    except (OSError, yaml.YAMLError, KeyError, TypeError) as exc:
        # Logging is configured from the config, so there is none yet.
        logging.basicConfig(level=logging.INFO)
        logger.error("Cannot load config: %s", exc)
        return EXIT_CONFIG_ERROR
    setup_logging(cfg)

    if args.no_ui:
        return _run_headless(cfg)

    from email_archiver.ui.app import ScanWindow
    ScanWindow(cfg).run()
    return EXIT_OK


def _run_headless(cfg: dict) -> int:
    from email_archiver.scanner.scanner import FolderScanner

    try:
        roots = get_archive_roots(cfg)
        scanner = FolderScanner(cfg)
    except (KeyError, TypeError) as exc:
        logger.error("Config is missing a required key: %s", exc)
        return EXIT_CONFIG_ERROR
    if not roots:
        logger.error("No archive roots configured (archive.root_paths).")
        return EXIT_CONFIG_ERROR

    # A captured job log is not a TTY: there `\r` rewrites pile up into one
    # unreadable line, so print one line per tick instead.
    interactive = sys.stdout.isatty()

    def on_progress(current: int, total: int, path: str) -> None:
        if current % _PROGRESS_EVERY == 0 or current == total:
            pct = f"{int(current / total * 100)}%" if total else ""
            line = f"  {current:>6,} / {total:,}  {pct}  "
            if interactive:
                print(f"\r{line}", end="", flush=True)
            else:
                print(line.rstrip(), flush=True)

    logger.info("Scanning: %s", roots)
    stats = scanner.scan(progress_callback=on_progress)
    if interactive:
        print()  # end the \r progress line
    logger.info(
        "Done — %s new, %s updated, %s skipped, %s deleted, %s errors in %.1fs",
        stats.newly_indexed, stats.updated, stats.skipped, stats.deleted,
        stats.errors, stats.duration_seconds,
    )

    if stats.missing_roots:
        for root in stats.missing_roots:
            logger.error(
                "Archive root missing, index rows under it left untouched: %s", root
            )
        return EXIT_ROOT_MISSING
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())

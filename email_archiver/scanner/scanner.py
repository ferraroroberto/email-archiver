"""
Folder scanner: walks the archive directory, indexes .msg files into SQLite.

Design decisions:
- Incremental: compares os.stat().st_mtime with the stored value; unchanged
  files are skipped entirely, making re-scans fast even with 50k+ emails.
- Batch commits: writes to DB every `batch_size` records to balance memory
  usage vs. write overhead.
- Errors in individual .msg files are logged and skipped; the scan continues.
- progress_callback(current, total, current_file) is called after each file
  so the UI can update a progress bar without polling.
"""
from __future__ import annotations

import logging
import os
import re
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable

import extract_msg
from extract_msg.exceptions import InvalidFileFormatError, UnrecognizedMSGTypeError

from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRecord, EmailRepository

logger = logging.getLogger(__name__)

ProgressCallback = Callable[[int, int, str], None]


# ---------------------------------------------------------------- helpers ---

_RE_REPLY_PREFIX = re.compile(
    r"^\s*(re|rv|fwd?)\s*:?\s*", re.IGNORECASE
)
_RE_MSG_SUFFIX = re.compile(r"\s*\.msg$", re.IGNORECASE)


def _clean_subject(raw: str | None) -> str:
    """Strip Re:/Rv:/Fwd: prefixes and .msg suffix for cleaner indexing."""
    if not raw:
        return ""
    s = _RE_REPLY_PREFIX.sub("", raw.strip())
    s = _RE_MSG_SUFFIX.sub("", s)
    return s.strip()


def _safe_str(value: object) -> str:
    return str(value).strip() if value else ""


def _extract_msg_metadata(
    file_path: str, body_preview_len: int
) -> dict[str, str] | None:
    """
    Open a .msg file and extract metadata.
    Returns None on any read error (error is logged).
    """
    try:
        with extract_msg.Message(file_path) as msg:
            subject = _clean_subject(msg.subject)
            sender = _safe_str(msg.sender)
            recipients = _safe_str(msg.to)
            body = _safe_str(msg.body)[:body_preview_len]

            date_sent = ""
            if msg.date:
                try:
                    if isinstance(msg.date, datetime):
                        date_sent = msg.date.isoformat()
                    else:
                        date_sent = str(msg.date)
                except Exception:
                    pass

            return {
                "subject": subject,
                "sender": sender,
                "recipients": recipients,
                "body_preview": body,
                "date_sent": date_sent,
            }

    except (InvalidFileFormatError, UnrecognizedMSGTypeError) as exc:
        logger.debug("Skipping unreadable .msg %s: %s", file_path, exc)
    except AttributeError as exc:
        logger.debug("AttributeError in %s: %s", file_path, exc)
    except Exception as exc:
        logger.warning("Unexpected error reading %s: %s", file_path, exc)
    return None


# -------------------------------------------------------------- scanner -----

@dataclass
class ScanStats:
    total_found: int = 0
    newly_indexed: int = 0
    updated: int = 0
    skipped: int = 0
    errors: int = 0
    deleted: int = 0
    duration_seconds: float = 0.0


class FolderScanner:
    """
    Recursively scans a root directory for .msg files and indexes them.

    Usage:
        scanner = FolderScanner(config)
        stats = scanner.scan(progress_callback=my_callback)
    """

    def __init__(self, cfg: dict) -> None:
        self._root = Path(cfg["archive"]["root_path"])
        self._db_path = cfg["database"]["path"]
        self._batch_size: int = cfg["scanning"]["batch_size"]
        self._preview_len: int = cfg["scanning"]["body_preview_length"]
        self._max_path: int = cfg["scanning"]["max_path_length"]

    def scan(
        self,
        progress_callback: ProgressCallback | None = None,
        stop_flag: list[bool] | None = None,
    ) -> ScanStats:
        """
        Run a full incremental scan of the archive root.

        Args:
            progress_callback: called as (processed, total_estimate, current_path).
                               total_estimate may be 0 if unknown.
            stop_flag: a single-element list [False]; set to [True] to abort.
        """
        if not self._root.exists():
            raise FileNotFoundError(f"Archive root not found: {self._root}")

        stats = ScanStats()
        start = datetime.now()

        conn = init_db(self._db_path)
        repo = EmailRepository(conn)

        # Collect all .msg paths upfront for progress reporting and for
        # the post-scan purge (delete DB entries whose files were removed).
        logger.info("Discovering .msg files under %s …", self._root)
        all_msg_paths: list[str] = []
        for dirpath, _dirs, files in os.walk(self._root):
            for fname in files:
                if fname.lower().endswith(".msg"):
                    fp = os.path.join(dirpath, fname)
                    if len(fp) <= self._max_path:
                        all_msg_paths.append(fp)

        stats.total_found = len(all_msg_paths)
        logger.info("Found %d .msg files. Starting indexing …", stats.total_found)

        batch_counter = 0
        indexed_folders: set[str] = set()

        for idx, file_path in enumerate(all_msg_paths):
            if stop_flag and stop_flag[0]:
                logger.info("Scan aborted by user after %d files.", idx)
                break

            if progress_callback:
                progress_callback(idx + 1, stats.total_found, file_path)

            folder_path = str(Path(file_path).parent)
            filename = Path(file_path).name

            # --- incremental check ---
            try:
                disk_mtime = os.stat(file_path).st_mtime
            except OSError:
                stats.errors += 1
                continue

            stored_mtime = repo.get_mtime(file_path)
            if stored_mtime is not None and abs(stored_mtime - disk_mtime) < 1.0:
                stats.skipped += 1
                indexed_folders.add(folder_path)
                continue

            # --- read metadata ---
            meta = _extract_msg_metadata(file_path, self._preview_len)
            if meta is None:
                stats.errors += 1
                continue

            was_new = stored_mtime is None
            repo.upsert_email(
                EmailRecord(
                    file_path=file_path,
                    folder_path=folder_path,
                    filename=filename,
                    file_mtime=disk_mtime,
                    **meta,
                )
            )
            repo.upsert_folder(folder_path)
            indexed_folders.add(folder_path)

            if was_new:
                stats.newly_indexed += 1
            else:
                stats.updated += 1

            batch_counter += 1
            if batch_counter >= self._batch_size:
                conn.commit()
                batch_counter = 0
                logger.debug("Committed batch at %d files.", idx + 1)

        # Final commit
        conn.commit()

        # Remove DB entries for files that no longer exist on disk.
        # Only run the purge when the scan was not aborted mid-way, to avoid
        # deleting entries for files we simply haven't visited yet.
        aborted = stop_flag and stop_flag[0]
        if not aborted:
            stats.deleted = repo.delete_missing_emails(all_msg_paths)
            if stats.deleted:
                logger.info("Purged %d deleted email(s) from index.", stats.deleted)
            conn.commit()

        repo.refresh_folder_counts()
        conn.commit()

        stats.duration_seconds = (datetime.now() - start).total_seconds()
        logger.info(
            "Scan complete: %d new, %d updated, %d skipped, %d errors in %.1fs",
            stats.newly_indexed, stats.updated, stats.skipped,
            stats.errors, stats.duration_seconds,
        )
        conn.close()
        return stats

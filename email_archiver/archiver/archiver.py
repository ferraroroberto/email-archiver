"""
Email archiver: saves .msg files and extracts attachments to disk.

Naming convention (matches user spec):
    Email:       NNN - sanitized_subject.msg
    Attachments: NNN - filename.ext

NNN is a zero-padded 3-digit sequence (000-999) shared by an email and its
attachments. With ``naming.date_prefix: true`` in the config, the email's sent
date in local time is prefixed as well:

    Email:       YYYY-MM-DD - NNN - sanitized_subject.msg
    Attachments: YYYY-MM-DD - NNN - filename.ext

The toggle is off by default. When it is on but the sent date cannot be
resolved, the undated form is used for that email instead.

Design decisions:
- Sequence number is derived from the MAXIMUM existing numeric prefix in the
  target folder, not a count, so it is safe even if files were deleted. The
  scan recognises both prefix forms regardless of the toggle, so a folder
  holding a mix of dated and undated files always allocates the next number
  correctly and the toggle stays safe to flip at any time.
- Embedded images (ContentId set) are skipped; only real attachments are saved.
- Subject sanitisation removes characters illegal on Windows file systems.
- SaveAs uses olMSG format constant (3) to produce a proper .msg file.
"""
from __future__ import annotations

import logging
import os
import re
import string
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from email_archiver.config import (
    DEFAULT_DATE_PREFIX_ENABLED,
    DEFAULT_MAX_PATH_LENGTH,
    get_date_prefix_enabled,
    get_max_path_length,
)

logger = logging.getLogger(__name__)

# Outlook SaveAs format constant for .msg
_OL_MSG_FORMAT = 3

# Regexes to find the leading sequence number, in either filename form.
# Dated is tried first since it is the more specific match.
_RE_DATED_PREFIX = re.compile(r"^\d{4}-\d{2}-\d{2} - (\d{3}) - ")
_RE_UNDATED_PREFIX = re.compile(r"^(\d{3}) - ")

# Maximum path length (in chars) the fitted filename must stay within. This is
# the SAME budget the scanner uses — sourced from config via
# ``get_max_path_length`` and threaded in through ``EmailArchiver``; see
# ``config.DEFAULT_MAX_PATH_LENGTH`` for the rationale behind the value.
_ELLIPSIS = "..."

# Bound on the attachment de-duplication loop in ``_save_attachments`` — a
# fitter regression (or a pathological folder with dozens of same-name
# attachments already saved) degrades to a logged, skipped attachment instead
# of hanging the caller's thread forever.
_MAX_DEDUPE_ATTEMPTS = 100


# ----------------------------------------------------------------- types ----

@dataclass
class ArchiveResult:
    email_path: str = ""
    attachment_paths: list[str] = field(default_factory=list)
    sequence_number: str = ""


# ---------------------------------------------------------------- helpers ---

def _sanitize_filename(text: str, max_len: int = 80) -> str:
    """
    Remove characters illegal on Windows NTFS and truncate.
    Keeps ASCII letters/digits + a small set of safe punctuation.
    """
    safe = set(string.ascii_letters + string.digits + " -_.()")
    cleaned = "".join(c if c in safe else "_" for c in text)
    # Collapse multiple underscores/spaces
    cleaned = re.sub(r"[_ ]{2,}", "_", cleaned).strip("_. ")
    return cleaned[:max_len] if cleaned else "email"


def _fit_filename_to_path(
    folder_path: str,
    seq: str,
    stem: str,
    suffix: str,
    *,
    date_prefix: str | None = None,
    max_path: int = DEFAULT_MAX_PATH_LENGTH,
) -> str:
    """
    Build ``"{seq} - {stem}{suffix}"`` — or, when ``date_prefix`` is given,
    ``"{date_prefix} - {seq} - {stem}{suffix}"`` — such that the full path
    ``os.path.join(folder_path, filename)`` stays within ``max_path`` chars.
    The date prefix costs 13 chars and comes out of the stem's budget, never
    out of the path limit.

    If the stem has to be cut, an ellipsis is appended so the resulting
    filename still hints at the original subject. ``suffix`` is preserved.
    """
    prefix = f"{date_prefix} - {seq} - " if date_prefix else f"{seq} - "
    sep_len = 1  # path separator inserted by os.path.join
    filename_budget = max_path - len(folder_path) - sep_len
    fixed = len(prefix) + len(suffix)

    if filename_budget >= fixed + len(stem):
        return f"{prefix}{stem}{suffix}"

    stem_budget = filename_budget - fixed - len(_ELLIPSIS)
    if stem_budget > 0:
        return f"{prefix}{stem[:stem_budget]}{_ELLIPSIS}{suffix}"

    # Pathological: even a one-char stem + ellipsis won't fit.
    # Cut as tightly as possible without the ellipsis; if the folder itself
    # already exceeds the budget there is nothing we can do except let the
    # underlying call surface a clear OSError.
    stem_budget = max(1, filename_budget - fixed)
    return f"{prefix}{stem[:stem_budget]}{suffix}"


def get_next_sequence_number(folder_path: str) -> str:
    """
    Scan the folder for files starting with either the undated ``NNN - `` or
    the dated ``YYYY-MM-DD - NNN - `` prefix and return (max + 1).
    Both forms are always recognised, independently of which form this run
    writes, so a folder holding a mix never gets a colliding number.
    Returns '001' if the folder is empty or has no numbered files.
    """
    try:
        files = os.listdir(folder_path)
    except OSError as exc:
        logger.error("Cannot list folder %s: %s", folder_path, exc)
        return "001"

    numbers: list[int] = []
    for fname in files:
        m = _RE_DATED_PREFIX.match(fname) or _RE_UNDATED_PREFIX.match(fname)
        if m:
            numbers.append(int(m.group(1)))

    next_num = (max(numbers) + 1) if numbers else 1
    if next_num > 999:
        logger.warning("Sequence number exceeds 999 in %s", folder_path)
    return f"{next_num:03d}"


def _get_sent_date_prefix(mail_item: Any) -> str | None:
    """
    Resolve the email's sent date as a ``YYYY-MM-DD`` string in local time.

    Only called when ``naming.date_prefix`` is on. Tries ``SentOn`` (the
    actual send time) first, falling back to ``ReceivedTime``. Returns
    ``None`` — and logs why — when neither resolves, signalling the caller to
    fall back to the undated filename form rather than invent a placeholder
    date.
    """
    last_exc: Exception | None = None
    for attr in ("SentOn", "ReceivedTime"):
        try:
            value = getattr(mail_item, attr)
        except Exception as exc:
            last_exc = exc
            continue
        if value is None:
            continue
        try:
            return f"{value.year:04d}-{value.month:02d}-{value.day:02d}"
        except AttributeError as exc:
            last_exc = exc
            continue

    logger.warning(
        "Could not resolve a sent date for this email (%s); "
        "falling back to the undated filename form",
        last_exc if last_exc is not None else "SentOn and ReceivedTime both empty",
    )
    return None


# ------------------------------------------------------------- archiver -----

class EmailArchiver:
    """Saves an Outlook MailItem (COM object) to a target folder on disk."""

    def __init__(
        self,
        cfg: dict[str, Any] | None = None,
        *,
        date_prefix: bool | None = None,
    ) -> None:
        """Build an archiver bound to a path-length budget and a naming form.

        ``cfg`` is the loaded config dict; the path budget is read from it via
        the shared ``get_max_path_length`` accessor so the archiver and scanner
        honour the same ``path.max_length`` knob, and the sent-date prefix
        toggle via ``get_date_prefix_enabled`` (``naming.date_prefix``). When
        ``cfg`` is omitted the ``DEFAULT_*`` fallbacks are used.

        ``date_prefix`` overrides the configured toggle for this archiver only
        — the archive dialog passes its checkbox state through here, so the
        config value is the checkbox's *starting* position rather than the last
        word. ``None`` (the default) means "use the config".
        """
        self._max_path: int = (
            get_max_path_length(cfg) if cfg is not None else DEFAULT_MAX_PATH_LENGTH
        )
        self._date_prefix_enabled: bool
        if date_prefix is not None:
            self._date_prefix_enabled = date_prefix
        elif cfg is not None:
            self._date_prefix_enabled = get_date_prefix_enabled(cfg)
        else:
            self._date_prefix_enabled = DEFAULT_DATE_PREFIX_ENABLED

    def archive(
        self,
        mail_item: Any,
        folder_path: str,
        subject: str,
    ) -> ArchiveResult:
        """
        Save the email and its attachments to folder_path.

        Args:
            mail_item: Outlook COM MailItem (from OutlookClient.raw_item).
            folder_path: Absolute path to the destination folder.
            subject: Clean subject string (used for filename).

        Returns:
            ArchiveResult with paths of all saved files.
        """
        dest = Path(folder_path)
        if not dest.exists():
            logger.info("Creating destination folder: %s", dest)
            dest.mkdir(parents=True, exist_ok=True)

        seq = get_next_sequence_number(folder_path)
        date_prefix = (
            _get_sent_date_prefix(mail_item) if self._date_prefix_enabled else None
        )
        result = ArchiveResult(sequence_number=seq)

        # ---- save .msg ----
        result.email_path = self._save_msg(
            mail_item, folder_path, seq, subject, date_prefix
        )

        # ---- save attachments ----
        result.attachment_paths = self._save_attachments(
            mail_item, folder_path, seq, date_prefix
        )

        logger.info(
            "Archived email %s → %s (%d attachment(s))",
            seq, result.email_path, len(result.attachment_paths),
        )
        return result

    # -------------------------------------------------- private helpers ----

    def _save_msg(
        self,
        mail_item: Any,
        folder_path: str,
        seq: str,
        subject: str,
        date_prefix: str | None,
    ) -> str:
        safe_subject = _sanitize_filename(subject)
        filename = _fit_filename_to_path(
            folder_path, seq, safe_subject, ".msg",
            date_prefix=date_prefix, max_path=self._max_path,
        )
        file_path = os.path.join(folder_path, filename)

        try:
            mail_item.SaveAs(file_path, _OL_MSG_FORMAT)
            logger.debug("Saved email: %s", file_path)
        except Exception as exc:
            logger.error("Failed to save .msg to %s: %s", file_path, exc)
            raise

        return file_path

    def _save_attachments(
        self,
        mail_item: Any,
        folder_path: str,
        seq: str,
        date_prefix: str | None,
    ) -> list[str]:
        saved: list[str] = []
        att_index = 1

        try:
            attachments = mail_item.Attachments
        except Exception as exc:
            logger.warning("Cannot access attachments: %s", exc)
            return saved

        for attachment in attachments:
            # Determine whether this attachment is inline (embedded in HTML body).
            #
            # ContentId alone is NOT reliable — many email clients (including
            # Outlook) set a ContentId on every attachment, not just embedded ones.
            # The correct signal is PR_ATTACH_FLAGS bit 4 (ATT_MHTML_REF = 0x4),
            # which is only set on true inline/embedded objects.
            # Fall back to ContentId only when the flags property is unavailable.
            try:
                flags = int(attachment.PropertyAccessor.GetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x37140003"
                ) or 0)
                if flags & 4:
                    logger.debug("Skipping inline attachment: %s", attachment.FileName)
                    continue
            except Exception:
                # Flags property not present — fall back to ContentId check
                try:
                    cid = attachment.PropertyAccessor.GetProperty(
                        "http://schemas.microsoft.com/mapi/proptag/0x3712001E"
                    )
                    # Only skip if ContentId looks like a generated image CID
                    # (i.e. attachment is an image type AND has a CID)
                    if cid and Path(attachment.FileName or "").suffix.lower() in {
                        ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp",
                    }:
                        logger.debug("Skipping inline image: %s", attachment.FileName)
                        continue
                except Exception:
                    pass

            try:
                original_name = attachment.FileName or f"attachment_{att_index}"
                stem = _sanitize_filename(Path(original_name).stem, max_len=60)
                suffix = Path(original_name).suffix.lower()

                att_filename = _fit_filename_to_path(
                    folder_path, seq, stem, suffix,
                    date_prefix=date_prefix, max_path=self._max_path,
                )
                att_path = os.path.join(folder_path, att_filename)
                # Avoid overwrite if multiple attachments share the same name.
                # The disambiguating counter is appended to the *suffix* (not
                # folded into the stem) so its width comes out of the stem's
                # fit budget on every iteration via the normal `fixed` term —
                # folding it into the stem instead left `stem_budget` (which
                # depends only on folder/prefix/suffix length, never the
                # stem) unchanged across iterations, so once truncation
                # kicked in every attempt truncated to the exact same
                # characters and `os.path.exists` never went False (#43).
                # Bounded so a fitter regression degrades to a logged, skipped
                # attachment instead of hanging the caller's thread.
                counter = 2
                while os.path.exists(att_path):
                    if counter > _MAX_DEDUPE_ATTEMPTS:
                        logger.error(
                            "Could not find a non-colliding filename for "
                            "attachment %r in %s after %d attempts; skipping",
                            original_name, folder_path, _MAX_DEDUPE_ATTEMPTS,
                        )
                        att_path = None
                        break
                    att_filename = _fit_filename_to_path(
                        folder_path, seq, stem, f"_{counter}{suffix}",
                        date_prefix=date_prefix, max_path=self._max_path,
                    )
                    att_path = os.path.join(folder_path, att_filename)
                    counter += 1

                if att_path is None:
                    continue

                attachment.SaveAsFile(att_path)
                saved.append(att_path)
                logger.debug("Saved attachment: %s", att_path)
                att_index += 1

            except Exception as exc:
                logger.warning(
                    "Failed to save attachment %s: %s",
                    getattr(attachment, "FileName", "?"), exc,
                )

        return saved

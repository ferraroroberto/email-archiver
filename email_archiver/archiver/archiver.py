"""
Email archiver: saves .msg files and extracts attachments to disk.

Naming convention (matches user spec):
    Email:       NNN_sanitized_subject.msg
    Attachments: NNN_01_filename.ext
                 NNN_02_filename.ext

Where NNN is zero-padded to 3 digits (000–999).

Design decisions:
- Sequence number is derived from the MAXIMUM existing numeric prefix in the
  target folder, not a count, so it is safe even if files were deleted.
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

logger = logging.getLogger(__name__)

# Outlook SaveAs format constant for .msg
_OL_MSG_FORMAT = 3

# Regex to find the leading NNN_ prefix in filenames
_RE_PREFIX = re.compile(r"^(\d+)")


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


def get_next_sequence_number(folder_path: str) -> str:
    """
    Scan the folder for files starting with NNN_ and return (max + 1).
    Returns '001' if the folder is empty or has no numbered files.
    """
    try:
        files = os.listdir(folder_path)
    except OSError as exc:
        logger.error("Cannot list folder %s: %s", folder_path, exc)
        return "001"

    numbers: list[int] = []
    for fname in files:
        m = _RE_PREFIX.match(fname)
        if m:
            numbers.append(int(m.group(1)))

    next_num = (max(numbers) + 1) if numbers else 1
    if next_num > 999:
        logger.warning("Sequence number exceeds 999 in %s", folder_path)
    return f"{next_num:03d}"


# ------------------------------------------------------------- archiver -----

class EmailArchiver:
    """Saves an Outlook MailItem (COM object) to a target folder on disk."""

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
        result = ArchiveResult(sequence_number=seq)

        # ---- save .msg ----
        result.email_path = self._save_msg(mail_item, folder_path, seq, subject)

        # ---- save attachments ----
        result.attachment_paths = self._save_attachments(
            mail_item, folder_path, seq
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
    ) -> str:
        safe_subject = _sanitize_filename(subject)
        filename = f"{seq} - {safe_subject}.msg"
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
                safe_name = stem + suffix

                att_filename = f"{seq} - {safe_name}"
                att_path = os.path.join(folder_path, att_filename)
                # Avoid overwrite if multiple attachments share the same name
                counter = 2
                while os.path.exists(att_path):
                    safe_name = f"{stem}_{counter}{suffix}"
                    att_filename = f"{seq} - {safe_name}"
                    att_path = os.path.join(folder_path, att_filename)
                    counter += 1

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

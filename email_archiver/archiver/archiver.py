"""
Email archiver: saves .msg files and extracts attachments to disk.

Naming convention (matches user spec):
    Email:       NNN - sanitized_subject.msg
    Attachments: NNN - filename.ext

NNN is a zero-padded 3-digit sequence (000-999) shared by an email and its
attachments. With the sent-date prefix in effect, the email's sent date in
local time is prefixed as well:

    Email:       YYYY-MM-DD - NNN - sanitized_subject.msg
    Attachments: YYYY-MM-DD - NNN - filename.ext

``naming.date_prefix`` in the config controls this per folder:
    false (default) → never prefix
    true            → always prefix
    "auto"          → infer the form from what the destination folder
                       already holds (see ``infer_date_prefix``), falling
                       back to the undated form when inference is ambiguous

Whichever form is chosen, if the sent date cannot be resolved the undated
form is used for that email instead.

Design decisions:
- Sequence number is derived from the MAXIMUM existing numeric prefix in the
  target folder, not a count, so it is safe even if files were deleted. The
  scan recognises both prefix forms regardless of the toggle, so a folder
  holding a mix of dated and undated files always allocates the next number
  correctly and the toggle stays safe to flip at any time.
- ``infer_date_prefix`` shares that same listing pass (see ``_scan_folder``):
  the archiver never lists a destination folder twice for one archive call.
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
    DATE_PREFIX_AUTO,
    DEFAULT_DATE_PREFIX_ENABLED,
    DEFAULT_MAX_PATH_LENGTH,
    get_date_prefix_mode,
    get_max_path_length,
)
from email_archiver.outlook.client import _sent_datetime

logger = logging.getLogger(__name__)

# Outlook SaveAs format constant for .msg
_OL_MSG_FORMAT = 3

# Regexes to find the leading sequence number, in either filename form.
# Dated is tried first since it is the more specific match. Applied in exactly
# one place — ``split_sequence_prefix`` below; every caller reads a filename
# through that, so the naming rules stay owned by this module.
_RE_DATED_PREFIX = re.compile(r"^(\d{4}-\d{2}-\d{2}) - (\d{3}) - ")
_RE_UNDATED_PREFIX = re.compile(r"^(\d{3}) - ")

# Appended to a filename's stem when it has to be cut to fit the path budget
# (see ``_fit_filename_to_path``).
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

@dataclass(frozen=True)
class SequencePrefix:
    """A filename's leading sequence field, split into its parts.

    ``"023 - report.pdf"`` → ``SequencePrefix("", "023", "report.pdf")`` and
    ``"2026-03-14 - 023 - report.pdf"`` →
    ``SequencePrefix("2026-03-14", "023", "report.pdf")``. Keeping the date
    prefix as written (rather than a bool) is what lets a caller rewrite the
    number without touching the form or re-deriving the date.
    """

    date_prefix: str  # "" for the undated form, "YYYY-MM-DD" otherwise
    number: str       # the three digits exactly as written
    rest: str         # everything after the prefix — the name itself

    @property
    def dated(self) -> bool:
        return bool(self.date_prefix)

    def with_number(self, number: str) -> str:
        """The same filename carrying ``number`` instead, in the same form."""
        head = f"{self.date_prefix} - " if self.date_prefix else ""
        return f"{head}{number} - {self.rest}"


def split_sequence_prefix(filename: str) -> SequencePrefix | None:
    """Split ``filename``'s leading sequence field, or ``None`` when it has none.

    Both forms are always recognised, dated first since it is the more specific
    match — the same rule sequence allocation has always used, now shared with
    :mod:`email_archiver.renumber` so the two can never drift apart on what
    counts as a numbered file.
    """
    m = _RE_DATED_PREFIX.match(filename)
    if m:
        return SequencePrefix(m.group(1), m.group(2), filename[m.end():])
    m = _RE_UNDATED_PREFIX.match(filename)
    if m:
        return SequencePrefix("", m.group(1), filename[m.end():])
    return None


def sanitize_filename(text: str, max_len: int = 80) -> str:
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


@dataclass
class _FolderScan:
    """The result of one ``os.listdir`` pass over a destination folder."""

    next_sequence: str
    inferred_date_prefix: bool | None


def _scan_folder(folder_path: str) -> _FolderScan:
    """
    List ``folder_path`` once and derive both the next sequence number and
    the per-folder date-prefix inference from that single pass.

    This is the one place that actually calls ``os.listdir`` on a
    destination folder — ``get_next_sequence_number`` and
    ``infer_date_prefix`` are thin, independently-testable wrappers around
    it, but ``EmailArchiver.archive`` calls this directly so a real archive
    never lists a OneDrive-backed folder twice.
    """
    try:
        files = os.listdir(folder_path)
    except OSError as exc:
        logger.error("Cannot list folder %s: %s", folder_path, exc)
        return _FolderScan(next_sequence="001", inferred_date_prefix=None)

    numbers: list[int] = []
    dated = 0
    undated = 0
    for fname in files:
        parsed = split_sequence_prefix(fname)
        if parsed is None:
            continue
        numbers.append(int(parsed.number))
        if parsed.dated:
            dated += 1
        else:
            undated += 1

    next_num = (max(numbers) + 1) if numbers else 1
    if next_num > 999:
        logger.warning("Sequence number exceeds 999 in %s", folder_path)
    next_sequence = f"{next_num:03d}"

    # Empty/unnumbered (0 == 0) and a genuine tie both mean "no clear form" —
    # neither is worth forcing a guess on, so both come out as None.
    inferred: bool | None = None if dated == undated else dated > undated

    return _FolderScan(next_sequence=next_sequence, inferred_date_prefix=inferred)


def get_next_sequence_number(folder_path: str) -> str:
    """
    Scan the folder for files starting with either the undated ``NNN - `` or
    the dated ``YYYY-MM-DD - NNN - `` prefix and return (max + 1).
    Both forms are always recognised, independently of which form this run
    writes, so a folder holding a mix never gets a colliding number.
    Returns '001' if the folder is empty or has no numbered files.
    """
    return _scan_folder(folder_path).next_sequence


def infer_date_prefix(folder_path: str) -> bool | None:
    """
    Infer which filename form ``folder_path`` already uses, from the same
    dated-vs-undated counts ``get_next_sequence_number`` derives.

    Majority wins: ``True`` when more of the folder's existing numbered
    files use the dated ``YYYY-MM-DD - NNN - `` form than the undated
    ``NNN - `` form, ``False`` for the reverse. Returns ``None`` when the
    folder is empty, has no numbered files, or the two forms tie — callers
    fall back to the undated form in that case rather than guess.
    """
    return _scan_folder(folder_path).inferred_date_prefix


def resolve_date_prefix_for_folder(cfg: dict[str, Any], folder_path: str) -> bool:
    """
    The date-prefix form ``EmailArchiver.archive`` would pick for
    ``folder_path`` under ``cfg``, absent an explicit per-call override.

    ``naming.date_prefix: auto`` infers from the folder's own contents,
    falling back to ``False`` when inference is ambiguous; ``true``/``false``
    apply uniformly regardless of the folder. Used by batch ``plan`` to
    report each candidate's form so the caller can hand it straight back to
    an ``apply`` decision.
    """
    mode = get_date_prefix_mode(cfg)
    inferred = infer_date_prefix(folder_path) if mode == DATE_PREFIX_AUTO else None
    return _resolve_date_prefix_mode(mode, inferred)


def _resolve_date_prefix_mode(mode: bool | str, inferred: bool | None) -> bool:
    """Shared by ``EmailArchiver`` and ``resolve_date_prefix_for_folder``:
    ``"auto"`` uses ``inferred`` (``False`` when it is ``None``), otherwise
    ``mode`` is already the plain boolean to use."""
    if mode == DATE_PREFIX_AUTO:
        return bool(inferred) if inferred is not None else False
    return bool(mode)


def _get_sent_date_prefix(mail_item: Any) -> str | None:
    """
    Resolve the email's sent date as a ``YYYY-MM-DD`` string in local time.

    Only called when ``naming.date_prefix`` is on. Formats whatever
    ``outlook.client._sent_datetime`` resolves (``SentOn`` falling back to
    ``ReceivedTime``) — the single implementation of that rule, so the date a
    batch ``plan`` reports and the date a ``YYYY-MM-DD -`` prefix carries can
    never disagree. Returns ``None`` — and logs why — when neither resolves,
    signalling the caller to fall back to the undated filename form rather
    than invent a placeholder date.
    """
    value = _sent_datetime(mail_item)
    if value is None:
        logger.warning(
            "Could not resolve a sent date for this email "
            "(SentOn and ReceivedTime both empty or unreadable); "
            "falling back to the undated filename form"
        )
        return None
    return f"{value.year:04d}-{value.month:02d}-{value.day:02d}"


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
        honour the same ``path.max_length`` knob, and the sent-date prefix mode
        via ``get_date_prefix_mode`` (``naming.date_prefix``: ``True``,
        ``False`` or ``"auto"``). When ``cfg`` is omitted the ``DEFAULT_*``
        fallbacks are used.

        ``date_prefix`` overrides the configured mode for this archiver only
        — the archive dialog passes its checkbox state through here, and batch
        mode passes the caller's per-mail decision, so the config is only the
        *starting* position, never the last word. ``None`` (the default) means
        "use the config": ``auto`` then infers the form per destination folder
        at ``archive()`` time, falling back to the undated form when
        inference is ambiguous.
        """
        self._max_path: int = (
            get_max_path_length(cfg) if cfg is not None else DEFAULT_MAX_PATH_LENGTH
        )
        # Explicit override always wins over the config; None defers to it.
        self._date_prefix_override: bool | None = date_prefix
        self._date_prefix_mode: bool | str = (
            get_date_prefix_mode(cfg) if cfg is not None else DEFAULT_DATE_PREFIX_ENABLED
        )

    def archive(
        self,
        mail_item: Any,
        folder_path: str,
        subject: str,
    ) -> ArchiveResult:
        """
        Save the email and its attachments to folder_path.

        Args:
            mail_item: Outlook COM MailItem, freshly acquired by the caller
                (COM objects are STA — never cache one across threads).
            folder_path: Absolute path to the destination folder.
            subject: Clean subject string (used for filename).

        Returns:
            ArchiveResult with paths of all saved files.
        """
        dest = Path(folder_path)
        if not dest.exists():
            logger.info("Creating destination folder: %s", dest)
            dest.mkdir(parents=True, exist_ok=True)

        # One listing pass covers both the next sequence number and (when the
        # config mode is "auto") the per-folder date-prefix inference.
        scan = _scan_folder(folder_path)
        seq = scan.next_sequence
        date_prefix_enabled = self._resolve_date_prefix_enabled(
            scan.inferred_date_prefix
        )
        date_prefix = (
            _get_sent_date_prefix(mail_item) if date_prefix_enabled else None
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

    def _resolve_date_prefix_enabled(self, inferred: bool | None) -> bool:
        """Precedence for this call: explicit constructor override beats the
        config outright; otherwise the config decides — ``"auto"`` via
        ``inferred`` (already derived from this same destination folder by
        ``archive``), a plain boolean applied as-is."""
        if self._date_prefix_override is not None:
            return self._date_prefix_override
        return _resolve_date_prefix_mode(self._date_prefix_mode, inferred)

    # -------------------------------------------------- private helpers ----

    def _save_msg(
        self,
        mail_item: Any,
        folder_path: str,
        seq: str,
        subject: str,
        date_prefix: str | None,
    ) -> str:
        safe_subject = sanitize_filename(subject)
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
                stem = sanitize_filename(Path(original_name).stem, max_len=60)
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

"""
Headless batch mode: plan / apply / revert the whole Inbox with JSON I/O.

This module is what ``main_batch.py`` runs and what the tests exercise. It is
pure orchestration over :class:`~email_archiver.outlook.client.OutlookClient`,
:class:`~email_archiver.engine.suggester.SuggestionEngine`,
:class:`~email_archiver.archiver.archiver.EmailArchiver` and
:class:`~email_archiver.database.repository.EmailRepository` — no COM, no
tkinter — so a fake client drives every verb end to end in a unit test.

Design decisions:

- **Identity is the Internet Message-ID, never EntryID.** Outlook rewrites
  ``EntryID`` when a mail moves between folders, which is exactly what ``apply``
  does to every mail it touches; the Message-ID survives. A mail carrying no
  Message-ID (rare — drafts, some system mail) is *reported and skipped*, never
  guessed at.
- **The archiver stays the sole owner of the naming rules.** Batch mode chooses
  nothing about filenames; it passes the caller's per-mail ``date_prefix``
  decision into ``EmailArchiver`` and lets it do what the dialog does.
- **One failing mail never aborts the run.** Every verb returns a document with
  a per-mail result carrying its own ``error``, and the process still exits 0.
  A non-zero exit means the run could not *start* at all (see ``main_batch.py``).
- **``files`` is filled in before the move**, so a mail that was written to disk
  but failed to move is still fully revertible from the ``apply`` output. That
  is why an entry can carry ``ok: false`` and a non-empty ``files`` at once.
- **``revert`` deletes only what it is given, and only inside the archive
  roots.** Anything resolving outside ``archive.root_paths`` is refused per
  file with a reason rather than deleted, so a malformed or hostile items file
  cannot reach the rest of the disk.
"""
from __future__ import annotations

import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Any

from email_archiver.archiver.archiver import EmailArchiver
from email_archiver.config import (
    get_archive_roots,
    get_outlook_archive_folder,
    get_outlook_category,
)
from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRepository
from email_archiver.engine.suggester import SuggestionEngine
from email_archiver.outlook.client import EmailData
from email_archiver.text import normalize_message_id

logger = logging.getLogger(__name__)

# Bumped when the shape of a document below changes incompatibly, so a consumer
# spawning this process can refuse a version it does not understand instead of
# reading a field that silently moved.
SCHEMA_VERSION = 1

# Default number of ranked folder candidates `plan` reports per mail. Larger
# than the dialog's three: the caller re-ranks (an LLM, in task-os's case) and
# needs more than the engine's own top pick to choose from.
DEFAULT_CANDIDATES = 10

# Why a mail in the Inbox got no plan entry.
SKIP_NO_MESSAGE_ID = "no_message_id"

# `error.code` values, and the reason each one is worth telling apart.
ERROR_NOT_IN_INBOX = "not_in_inbox"            # already moved, or never there
ERROR_NOT_IN_ARCHIVE = "not_in_archive_folder"  # revert cannot find it back
ERROR_BAD_DECISION = "bad_decision"             # caller sent an unusable entry
ERROR_ARCHIVE_FAILED = "archive_failed"         # disk write / Outlook SaveAs
ERROR_MOVE_FAILED = "move_failed"               # files are on disk, mail is not
# The move landed and only the tag did not. Its own code because the two are
# genuinely different states to recover from: after a move_failed the mail is
# still in the Inbox, after a category_failed it is already filed and only
# looks untouched in Outlook.
ERROR_CATEGORY_FAILED = "category_failed"

# Per-file outcomes in a revert result.
REFUSED_OUTSIDE_ROOTS = "outside_archive_roots"
REFUSED_UNRESOLVABLE = "unresolvable_path"


# ---------------------------------------------------------------- helpers ---

def _now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def _iso_or_empty(value: datetime | None) -> str:
    return value.isoformat() if value is not None else ""


def _engine_for(cfg: dict[str, Any], candidates: int) -> SuggestionEngine:
    """A suggestion engine capped at ``candidates`` instead of the config's 3.

    The cap is the only knob batch mode changes, and it changes it on a copy —
    the loaded config is shared process-wide and the dialog must keep seeing
    its own ``suggestion.max_suggestions``.
    """
    tuned = dict(cfg)
    tuned["suggestion"] = {**(cfg.get("suggestion") or {}), "max_suggestions": candidates}
    return SuggestionEngine(tuned)


def _candidate_dict(suggestion: Any) -> dict[str, Any]:
    return {
        "folder_path": suggestion.folder_path,
        "display_name": suggestion.display_name,
        "score": round(suggestion.score, 4),
        "match_count": suggestion.match_count,
        "sample_subjects": list(suggestion.sample_subjects),
    }


def _mail_dict(mail: Any) -> dict[str, Any]:
    """The fields of a live mail every verb reports, in one shape."""
    return {
        "message_id": mail.message_id,
        "entry_id": mail.entry_id,
        "subject": mail.subject,
        "sender": mail.sender,
        "recipients": mail.recipients,
        "date_sent": _iso_or_empty(mail.date_sent),
        "body_preview": mail.body_preview,
        "attachment_count": mail.attachment_count,
        "flag_status": mail.flag_status,
    }


def _is_within(child: Path, root: Path) -> bool:
    """Whether ``child`` is ``root`` or lives underneath it.

    Compared through ``os.path.normcase`` rather than
    ``Path.is_relative_to`` alone: on Windows the two paths routinely differ in
    case (the config spells a root one way, ``resolve()`` another) and a
    case-sensitive comparison would refuse a file that is genuinely inside the
    archive. The trailing separator is what stops ``C:\\Archive`` from matching
    ``C:\\ArchiveOther``.
    """
    child_s = os.path.normcase(str(child))
    root_s = os.path.normcase(str(root))
    if child_s == root_s:
        return True
    return child_s.startswith(root_s.rstrip("\\/") + os.sep)


def _resolve_under_roots(
    raw_path: str, roots: list[Path]
) -> tuple[Path | None, str | None]:
    """Resolve a path and require it to sit under one of the archive roots.

    Returns ``(path, None)`` when it does, ``(None, reason)`` when it does not.
    The refusal is deliberately not an exception: ``revert`` reports it per file
    and carries on with the rest of the bundle.
    """
    try:
        resolved = Path(raw_path).resolve()
    except (OSError, ValueError):
        return None, REFUSED_UNRESOLVABLE
    for root in roots:
        if _is_within(resolved, root):
            return resolved, None
    return None, REFUSED_OUTSIDE_ROOTS


def _archive_roots(cfg: dict[str, Any]) -> list[Path]:
    resolved: list[Path] = []
    for raw in get_archive_roots(cfg):
        try:
            resolved.append(Path(raw).resolve())
        except (OSError, ValueError):
            logger.warning("Ignoring unresolvable archive root: %r", raw)
    return resolved


def _envelope(verb: str, cfg: dict[str, Any]) -> dict[str, Any]:
    return {
        "verb": verb,
        "schema_version": SCHEMA_VERSION,
        "generated_at": _now_iso(),
        "archive_folder": get_outlook_archive_folder(cfg),
        "category": get_outlook_category(cfg),
    }


def error_document(verb: str, code: str, message: str) -> dict[str, Any]:
    """The document printed when the run could not start at all.

    Deliberately the same envelope shape minus the results, so a consumer
    parses one JSON document either way and branches on ``"error" in doc``
    rather than on the exit code alone.
    """
    return {
        "verb": verb,
        "schema_version": SCHEMA_VERSION,
        "generated_at": _now_iso(),
        "error": {"code": code, "message": message},
    }


# ------------------------------------------------------------------- plan ---

def plan(
    client: Any,
    cfg: dict[str, Any],
    *,
    candidates: int = DEFAULT_CANDIDATES,
) -> dict[str, Any]:
    """Enumerate the Inbox and rank archive folders for every mail in it.

    Every Inbox mail is offered: nothing is filtered on age, sender or size.
    A mail whose Message-ID is already in the index is reported with
    ``already_archived`` set to the file it was archived as, and no candidates —
    ranking a mail that is already filed would only invite filing it twice.
    """
    preview_len = int((cfg.get("scanning") or {}).get("body_preview_length", 500))
    engine = _engine_for(cfg, candidates)

    conn = init_db(cfg["database"]["path"])
    repo = EmailRepository(conn)

    mails: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []
    already = 0

    try:
        for mail in client.iter_inbox(preview_len):
            if not mail.message_id:
                skipped.append({
                    "entry_id": mail.entry_id,
                    "subject": mail.subject,
                    "reason": SKIP_NO_MESSAGE_ID,
                })
                continue

            entry = _mail_dict(mail)
            archived_as = repo.find_path_by_message_id(mail.message_id)
            entry["already_archived"] = archived_as
            if archived_as:
                already += 1
                entry["candidates"] = []
            else:
                entry["candidates"] = [
                    _candidate_dict(s)
                    for s in engine.suggest(EmailData(
                        subject=mail.subject,
                        sender=mail.sender,
                        recipients=mail.recipients,
                        date_sent=mail.date_sent,
                    ))
                ]
            mails.append(entry)
    finally:
        conn.close()

    doc = _envelope("plan", cfg)
    doc["counts"] = {
        "inbox": len(mails) + len(skipped),
        "planned": len(mails) - already,
        "already_archived": already,
        "skipped": len(skipped),
    }
    doc["mails"] = mails
    doc["skipped"] = skipped
    logger.info(
        "Plan: %d mail(s) in the Inbox, %d already archived, %d skipped.",
        doc["counts"]["inbox"], already, len(skipped),
    )
    return doc


# ------------------------------------------------------------------ apply ---

def _blank_apply_result(message_id: str, folder_path: str) -> dict[str, Any]:
    return {
        "message_id": message_id,
        "folder_path": folder_path,
        "ok": False,
        "sequence_number": "",
        "files": [],
        "entry_id": "",
        "moved": False,
        "categorized": False,
        "error": None,
    }


def apply(
    client: Any,
    cfg: dict[str, Any],
    decisions: list[dict[str, Any]],
) -> dict[str, Any]:
    """Archive each decided mail, move it out of the Inbox and tag it.

    ``decisions`` is a list of ``{message_id, folder_path, date_prefix}``. The
    ``date_prefix`` flag is per mail and comes from the caller — batch mode
    never reads the global ``naming.date_prefix`` toggle, because the caller
    infers the form the destination folder actually uses.

    Whatever happens to one mail, the next is still attempted; the per-mail
    ``error`` says what went wrong and ``files`` says what is already on disk.
    """
    archive_folder = get_outlook_archive_folder(cfg)
    category = get_outlook_category(cfg)
    results: list[dict[str, Any]] = []

    for raw in decisions:
        message_id = normalize_message_id(raw.get("message_id"))
        folder_path = str(raw.get("folder_path") or "")
        result = _blank_apply_result(message_id, folder_path)

        if not message_id or not folder_path:
            result["error"] = {
                "code": ERROR_BAD_DECISION,
                "message": "a decision needs both a message_id and a folder_path",
            }
            results.append(result)
            continue

        item = None
        try:
            item = client.find_by_message_id(message_id, None)
        except Exception as exc:  # a COM failure on one lookup, not the run
            result["error"] = {
                "code": ERROR_NOT_IN_INBOX,
                "message": f"Inbox lookup failed: {type(exc).__name__}: {exc}",
            }
            results.append(result)
            continue

        if item is None:
            result["error"] = {
                "code": ERROR_NOT_IN_INBOX,
                "message": "no mail with this Message-ID is in the Inbox",
            }
            results.append(result)
            continue

        try:
            mail = client.read_mail(item)
            archiver = EmailArchiver(
                cfg, date_prefix=bool(raw.get("date_prefix", False))
            )
            archived = archiver.archive(item, folder_path, mail.subject)
        except Exception as exc:
            logger.exception("Archiving %s failed", message_id)
            result["error"] = {
                "code": ERROR_ARCHIVE_FAILED,
                "message": f"{type(exc).__name__}: {exc}",
            }
            results.append(result)
            continue

        result["sequence_number"] = archived.sequence_number
        # Filled in before the move on purpose: if the move fails these files
        # exist and the caller must be able to revert them.
        result["files"] = [archived.email_path, *archived.attachment_paths]

        try:
            moved = client.move_to(item, archive_folder)
            result["moved"] = True
            result["entry_id"] = client.entry_id(moved)
        except Exception as exc:
            logger.exception("Moving %s to %r failed", message_id, archive_folder)
            result["error"] = {
                "code": ERROR_MOVE_FAILED,
                "message": f"{type(exc).__name__}: {exc}",
            }
            results.append(result)
            continue

        # Tagged in its own step: a category that would not stick is a
        # different state from a move that did not happen, and the caller
        # recovers from the two differently.
        try:
            client.set_category(moved, category)
            result["categorized"] = True
        except Exception as exc:
            logger.exception("Tagging %s with %r failed", message_id, category)
            result["error"] = {
                "code": ERROR_CATEGORY_FAILED,
                "message": f"the mail was filed and moved, but not tagged: "
                           f"{type(exc).__name__}: {exc}",
            }
            results.append(result)
            continue

        result["ok"] = True
        results.append(result)

    doc = _envelope("apply", cfg)
    applied = sum(1 for r in results if r["ok"])
    doc["counts"] = {
        "requested": len(results),
        "applied": applied,
        "failed": len(results) - applied,
    }
    doc["results"] = results
    logger.info("Apply: %d of %d mail(s) filed.", applied, len(results))
    return doc


# ----------------------------------------------------------------- revert ---

def _blank_revert_result(message_id: str) -> dict[str, Any]:
    return {
        "message_id": message_id,
        "ok": False,
        "deleted": [],
        "missing": [],
        "refused": [],
        "file_errors": [],
        "moved_back": False,
        "category_removed": False,
        "entry_id": "",
        "error": None,
    }


def revert(
    client: Any,
    cfg: dict[str, Any],
    items: list[dict[str, Any]],
) -> dict[str, Any]:
    """Delete the listed archive files and put each mail back in the Inbox.

    ``items`` is a list of ``{message_id, files}`` — normally straight from an
    ``apply`` result. Only the listed files are touched, and only those that
    resolve inside ``archive.root_paths``; anything else is refused with a
    reason and left alone.

    The four per-file buckets are deliberately distinct facts: ``deleted``
    (gone now), ``missing`` (already gone — not an error, a revert run twice),
    ``refused`` (policy said no) and ``file_errors`` (the delete was attempted
    and the OS said no).
    """
    archive_folder = get_outlook_archive_folder(cfg)
    category = get_outlook_category(cfg)
    roots = _archive_roots(cfg)
    results: list[dict[str, Any]] = []

    for raw in items:
        message_id = normalize_message_id(raw.get("message_id"))
        result = _blank_revert_result(message_id)

        for raw_path in raw.get("files") or []:
            resolved, refusal = _resolve_under_roots(str(raw_path), roots)
            if resolved is None:
                result["refused"].append({"path": str(raw_path), "reason": refusal})
                continue
            try:
                resolved.unlink()
                result["deleted"].append(str(resolved))
            except FileNotFoundError:
                result["missing"].append(str(resolved))
            except OSError as exc:
                result["file_errors"].append(
                    {"path": str(resolved), "message": f"{type(exc).__name__}: {exc}"}
                )

        if not message_id:
            result["error"] = {
                "code": ERROR_BAD_DECISION,
                "message": "an item needs a message_id to move the mail back",
            }
            results.append(result)
            continue

        try:
            item = client.find_by_message_id(message_id, archive_folder)
        except Exception as exc:
            item = None
            logger.warning("Archive-folder lookup for %s failed: %s", message_id, exc)

        if item is None:
            result["error"] = {
                "code": ERROR_NOT_IN_ARCHIVE,
                "message": (
                    f"no mail with this Message-ID is in {archive_folder!r}; "
                    "the files listed above were still processed"
                ),
            }
            results.append(result)
            continue

        try:
            client.clear_category(item, category)
            result["category_removed"] = True
            moved = client.move_to(item, None)
            result["moved_back"] = True
            result["entry_id"] = client.entry_id(moved)
        except Exception as exc:
            logger.exception("Moving %s back to the Inbox failed", message_id)
            result["error"] = {
                "code": ERROR_MOVE_FAILED,
                "message": f"{type(exc).__name__}: {exc}",
            }
            results.append(result)
            continue

        result["ok"] = not result["refused"] and not result["file_errors"]
        results.append(result)

    doc = _envelope("revert", cfg)
    reverted = sum(1 for r in results if r["ok"])
    doc["counts"] = {
        "requested": len(results),
        "reverted": reverted,
        "failed": len(results) - reverted,
    }
    doc["results"] = results
    logger.info("Revert: %d of %d mail(s) restored.", reverted, len(results))
    return doc

"""
Headless batch mode: plan / apply / revert / renumber, with JSON I/O.

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
- **A mail is never archived twice.** ``apply`` re-checks the index per
  decision: a mail already filed there (the one ``plan`` reports
  ``already_archived``, still in the Inbox because an earlier run's move
  failed) is *finished* rather than re-written — ``reused: true``, the existing
  file reported back, nothing new on disk.
- **The reference that is moved is not the one that was archived.**
  ``MailItem.SaveAs`` can leave an item flagged as modified and ``Move`` then
  refuses it with MAPI_E_OBJECT_CHANGED, which is how a mail ends up with its
  files written and its place in the Inbox kept. ``apply`` moves a reference
  re-acquired by EntryID, and on that refusal saves and retries once; which
  path finished it is logged and reported as ``move_via``.
- **``revert`` deletes only what it is given, and only inside the archive
  roots.** Anything resolving outside ``archive.root_paths`` is refused per
  file with a reason rather than deleted, so a malformed or hostile items file
  cannot reach the rest of the disk.
- **Renumbering is opt-in and reported, never implied.** ``apply`` and
  ``revert`` leave the sequence exactly as they always have unless
  ``--renumber`` is passed; with it, each touched folder is renumbered once and
  the document carries the old → new map under ``renumbered`` so the consumer
  can heal the ``.msg`` paths it stored. A folder that could *not* be
  renumbered is listed under ``renumber_refused`` with its reason — never an
  empty map, which means something else entirely ("already in order").
"""
from __future__ import annotations

import logging
import os
from datetime import datetime
from pathlib import Path
from typing import Any

from email_archiver.archiver.archiver import EmailArchiver, resolve_date_prefix_for_folder
from email_archiver.config import (
    get_outlook_archive_folder,
    get_outlook_category,
)
from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRepository
from email_archiver.engine.suggester import SuggestionEngine
from email_archiver.outlook.client import EmailData, is_message_changed_error
from email_archiver.paths import (
    REFUSED_OUTSIDE_ROOTS,
    REFUSED_UNRESOLVABLE,
    archive_roots as _archive_roots,
    resolve_under_roots as _resolve_under_roots,
)
from email_archiver.renumber import RenumberResult, renumber_folder
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

# The per-file outcomes in a revert result — REFUSED_OUTSIDE_ROOTS and
# REFUSED_UNRESOLVABLE — are imported above from `paths`, which owns the
# archive-root guard now that `renumber` applies it too. They stay readable as
# `batch.REFUSED_*`, which is where consumers and tests look for them.

# Which reference finished an `apply` move, reported as the result's
# `move_via`. Three genuinely different stories about the same mail, and the
# one place that says whether the re-acquire is earning its keep on this
# mailbox — see `_move_out_of_inbox`.
MOVE_VIA_REFETCHED = "refetched"     # a reference read back out of the store
MOVE_VIA_ORIGINAL = "original"       # the re-acquire failed; the original moved
MOVE_VIA_SAVED_RETRY = "saved_retry"  # refused with 0x80040109, saved, retried


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


def _candidate_dict(
    cfg: dict[str, Any], suggestion: Any, date_prefix_cache: dict[str, bool]
) -> dict[str, Any]:
    """A ranked candidate, plus the ``date_prefix`` form its own folder would
    get under ``cfg`` — inferred per folder when ``naming.date_prefix`` is
    ``"auto"``, otherwise the config's fixed boolean. The caller (task-os)
    hands this straight back as the ``date_prefix`` on its ``apply``
    decision for the folder it picks; see ``resolve_date_prefix_for_folder``.

    ``date_prefix_cache`` is ``plan``'s own dict, keyed by folder path and
    shared across every mail in the run: a handful of folders tend to be
    everyone's top suggestion, so without it a full-Inbox plan would re-list
    the same OneDrive-backed folder once per mail that suggests it — exactly
    the redundant-listing cost the per-``archive()`` single-pass guarantee is
    there to avoid, just reappearing one layer up.
    """
    folder_path = suggestion.folder_path
    if folder_path not in date_prefix_cache:
        date_prefix_cache[folder_path] = resolve_date_prefix_for_folder(cfg, folder_path)
    return {
        "folder_path": folder_path,
        "display_name": suggestion.display_name,
        "score": round(suggestion.score, 4),
        "match_count": suggestion.match_count,
        "sample_subjects": list(suggestion.sample_subjects),
        "date_prefix": date_prefix_cache[folder_path],
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
        # Always true today — ``plan`` enumerates the Inbox, so a mail it
        # reports is in it by construction. Stated anyway because it is the
        # fact that makes an ``already_archived`` mail *retryable*: its files
        # are on disk but it never left the Inbox, and a consumer reading only
        # ``already_archived`` cannot tell that from a mail that is properly
        # filed and gone (issue #59).
        "in_inbox": True,
    }


def _envelope(verb: str, cfg: dict[str, Any]) -> dict[str, Any]:
    return {
        "verb": verb,
        "schema_version": SCHEMA_VERSION,
        "generated_at": _now_iso(),
        "archive_folder": get_outlook_archive_folder(cfg),
        "category": get_outlook_category(cfg),
    }


def _unique_folders(paths: list[str]) -> list[str]:
    """The distinct folders in ``paths``, first spelling wins.

    Deduplicated case-insensitively because Windows hands the same folder back
    spelled several ways, and renumbering one folder twice in a run would
    reorder what the first pass just fixed.
    """
    seen: set[str] = set()
    folders: list[str] = []
    for raw in paths:
        folder = os.path.normpath(raw)
        key = os.path.normcase(folder)
        if key in seen:
            continue
        seen.add(key)
        folders.append(folder)
    return folders


def _renumber_folders(
    cfg: dict[str, Any], repo: EmailRepository, folders: list[str]
) -> tuple[dict[str, list[dict[str, Any]]], list[dict[str, Any]]]:
    """Renumber each touched folder once; return ``(renumbered, refused)``.

    ``renumbered`` is the map a consumer heals its stored paths from, keyed by
    folder. A folder that could not be renumbered at all is **not** an empty
    map in there — an empty map means "already in order", which is a different
    fact — it goes into ``refused`` with its reason.
    """
    roots = _archive_roots(cfg)
    renumbered: dict[str, list[dict[str, Any]]] = {}
    refused: list[dict[str, Any]] = []

    for folder in folders:
        try:
            result = renumber_folder(folder, repo, roots)
        except Exception as exc:  # one folder's failure is not the run's
            logger.exception("Renumbering %r failed", folder)
            refused.append({
                "folder_path": folder,
                "reason": f"{type(exc).__name__}: {exc}",
            })
            continue
        if result.refused:
            refused.append({"folder_path": folder, "reason": result.refused})
            continue
        renumbered[result.folder_path] = result.renamed

    return renumbered, refused


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
    # Shared across every mail in this run — see _candidate_dict.
    date_prefix_cache: dict[str, bool] = {}

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
                    _candidate_dict(cfg, s, date_prefix_cache)
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
        # True when the mail was already on disk and this run only finished it
        # in Outlook — nothing was written, and `files` is what was found, not
        # what was created. `sequence_number` stays empty: none was allocated.
        "reused": False,
        # Which reference the move went through, "" when it never happened.
        "move_via": "",
        "error": None,
    }


def apply(
    client: Any,
    cfg: dict[str, Any],
    decisions: list[dict[str, Any]],
    *,
    renumber: bool = False,
) -> dict[str, Any]:
    """Archive each decided mail, move it out of the Inbox and tag it.

    ``decisions`` is a list of ``{message_id, folder_path, date_prefix}``. The
    ``date_prefix`` flag is per mail and comes from the caller — batch mode
    never reads the global ``naming.date_prefix`` toggle, because the caller
    infers the form the destination folder actually uses.

    A decision for a mail the index already holds — the one ``plan`` reported
    ``already_archived``, still sitting in the Inbox because an earlier run
    wrote its files and then failed to move it — writes **nothing**: the
    existing ``.msg`` is reported back as ``files``, ``reused`` is true, and
    the mail is moved and tagged like any other. That is what makes such a mail
    finishable at all; re-archiving it would file a second copy beside the
    first (issue #59).

    Whatever happens to one mail, the next is still attempted; the per-mail
    ``error`` says what went wrong and ``files`` says what is already on disk.

    With ``renumber`` on, every destination folder this run put a bundle in is
    renumbered once afterwards — a mail older than the folder's newest file
    took ``max + 1`` on the way in, and this is what puts it back in date order
    — and the document carries the old → new map per folder under
    ``renumbered``. Off (the default), the verb behaves exactly as before and
    neither key appears.
    """
    archive_folder = get_outlook_archive_folder(cfg)
    category = get_outlook_category(cfg)
    results: list[dict[str, Any]] = []

    conn = init_db(cfg["database"]["path"])
    repo = EmailRepository(conn)
    try:
        for raw in decisions:
            results.append(
                _apply_one(client, cfg, repo, raw, archive_folder, category)
            )
        doc = _envelope("apply", cfg)
        if renumber:
            # Only a folder this run actually wrote into, and only once each.
            doc["renumbered"], doc["renumber_refused"] = _renumber_folders(
                cfg, repo,
                _unique_folders([r["folder_path"] for r in results if r["files"]]),
            )
    finally:
        conn.close()

    applied = sum(1 for r in results if r["ok"])
    doc["counts"] = {
        "requested": len(results),
        "applied": applied,
        "failed": len(results) - applied,
    }
    doc["results"] = results
    logger.info("Apply: %d of %d mail(s) filed.", applied, len(results))
    return doc


def _archived_file_to_reuse(repo: EmailRepository, message_id: str) -> str | None:
    """The ``.msg`` this mail is already filed as, when it really is on disk.

    An index row is not proof the file is still there: a row can outlive its
    file (a manual delete, a move in Explorer). Reporting ``reused`` over a
    path that no longer exists would leave the mail tagged and filed in Outlook
    with nothing on disk to show for it, so an unbacked row falls through to a
    normal archive instead.
    """
    path = repo.find_path_by_message_id(message_id)
    if not path:
        return None
    if Path(path).exists():
        return path
    logger.info(
        "The index has %s at %r but the file is gone; archiving it again.",
        message_id, path,
    )
    return None


def _apply_one(
    client: Any,
    cfg: dict[str, Any],
    repo: EmailRepository,
    raw: dict[str, Any],
    archive_folder: str,
    category: str,
) -> dict[str, Any]:
    """One decision, start to finish. Never raises: every failure is a result."""
    message_id = normalize_message_id(raw.get("message_id"))
    folder_path = str(raw.get("folder_path") or "")
    result = _blank_apply_result(message_id, folder_path)

    if not message_id or not folder_path:
        result["error"] = {
            "code": ERROR_BAD_DECISION,
            "message": "a decision needs both a message_id and a folder_path",
        }
        return result

    try:
        item = client.find_by_message_id(message_id, None)
    except Exception as exc:  # a COM failure on one lookup, not the run
        result["error"] = {
            "code": ERROR_NOT_IN_INBOX,
            "message": f"Inbox lookup failed: {type(exc).__name__}: {exc}",
        }
        return result

    if item is None:
        result["error"] = {
            "code": ERROR_NOT_IN_INBOX,
            "message": "no mail with this Message-ID is in the Inbox",
        }
        return result

    reused = _archived_file_to_reuse(repo, message_id)
    if reused:
        # Nothing to write: an earlier run already wrote this bundle and only
        # the Outlook half of it is unfinished. Writing again would file a
        # second copy of the same mail beside the first.
        logger.info(
            "Mail %s is already archived as %r; finishing it in Outlook "
            "without writing anything.", message_id, reused,
        )
        result["reused"] = True
        result["files"] = [reused]
    else:
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
            return result

        result["sequence_number"] = archived.sequence_number
        # Filled in before the move on purpose: if the move fails these files
        # exist and the caller must be able to revert them.
        result["files"] = [archived.email_path, *archived.attachment_paths]

    try:
        moved, via = _move_out_of_inbox(client, item, archive_folder, message_id)
        result["moved"] = True
        result["move_via"] = via
        result["entry_id"] = client.entry_id(moved)
    except Exception as exc:
        logger.exception("Moving %s to %r failed", message_id, archive_folder)
        result["error"] = {
            "code": ERROR_MOVE_FAILED,
            "message": f"{type(exc).__name__}: {exc}{_move_remedy(exc)}",
        }
        return result

    # Tagged in its own step: a category that would not stick is a different
    # state from a move that did not happen, and the caller recovers from the
    # two differently.
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
        return result

    result["ok"] = True
    return result


def _move_remedy(exc: Exception) -> str:
    """The sentence appended to a ``move_failed`` message, when there is one.

    A move refused as *the message has been changed* even after the re-acquire
    and the save-and-retry is its own condition, and a recoverable one: the
    files are on disk, the mail is still in the Inbox, and the state that
    refuses the write is held by the running Outlook process itself — a
    restart clears it and re-applying the same decision then finishes the mail
    without writing anything (observed on the mail in issue #59: refused on a
    freshly re-acquired reference *and* on ``Save()``, moved on the first
    attempt after Outlook was restarted). Saying so here is what makes the
    next occurrence diagnosable from the run's own output.
    """
    if not is_message_changed_error(exc):
        return ""
    return (
        " — Outlook refused the move as a changed message twice, on a "
        "re-acquired reference and after saving it. The files are on disk and "
        "the mail is still in the Inbox: restarting Outlook and applying this "
        "same decision again finishes it without writing anything."
    )


def _move_out_of_inbox(
    client: Any, item: Any, archive_folder: str, message_id: str
) -> tuple[Any, str]:
    """Move a just-archived mail out of the Inbox; return it and how it went.

    ``MailItem.SaveAs`` can leave the in-memory item flagged as modified, and
    ``Move`` on that same reference is then refused with MAPI_E_OBJECT_CHANGED
    — the mail keeps its place in the Inbox with its files already on disk, and
    the next ``plan`` reports it ``already_archived``, so batch mode could
    never file it (issue #59). Two defences, in order:

    1. move a reference re-acquired from the store by EntryID, which carries
       none of the original's modified state;
    2. on a 0x80040109 anyway, ``Save()`` that reference — which is what clears
       the flag — and retry the move exactly once.

    Any other failure is raised untouched: a store that refuses a move for a
    different reason is a different problem, and retrying it would only hide
    it. Which of the three paths ran is logged and reported as ``move_via``,
    because "the re-acquire was enough" and "it took a save and a retry" are
    the two facts that say whether this fix is holding on a real mailbox.
    """
    target, via = item, MOVE_VIA_ORIGINAL
    try:
        fresh = client.refetch(item)
    except Exception as exc:  # a refetch that raises is still just a fallback
        fresh = None
        logger.warning("Re-acquiring %s by EntryID raised %s: %s",
                       message_id, type(exc).__name__, exc)
    if fresh is not None:
        target, via = fresh, MOVE_VIA_REFETCHED
    else:
        logger.info(
            "Could not re-acquire %s from the store; moving the original "
            "reference.", message_id,
        )

    try:
        moved = client.move_to(target, archive_folder)
    except Exception as exc:
        if not is_message_changed_error(exc):
            raise
        logger.info(
            "Outlook refused to move %s (the message has been changed) on the "
            "%s reference; saving it and retrying once.", message_id, via,
        )
        client.save_item(target)
        moved = client.move_to(target, archive_folder)
        via = MOVE_VIA_SAVED_RETRY

    logger.info("Moved %s to %r via the %s reference.", message_id, archive_folder, via)
    return moved, via


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
        "index_rows_removed": 0,
        "error": None,
    }


def revert(
    client: Any,
    cfg: dict[str, Any],
    items: list[dict[str, Any]],
    *,
    renumber: bool = False,
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

    Deleting a ``.msg`` also removes its index row in the same step (via
    ``EmailRepository.delete_by_path``, next to ``delete_missing_emails``'s
    end-of-scan sweep): otherwise the row survives until the next full scan,
    and ``plan`` in between reports the mail ``already_archived`` at a path
    that no longer exists instead of offering it again. A file whose delete
    failed leaves its row alone — the file is still there, so the row is
    still correct. Reported per result as ``index_rows_removed``.

    With ``renumber`` on, every folder this run deleted a file from is
    renumbered once afterwards, closing the gap the delete left, and the
    document carries the old → new map per folder under ``renumbered``. Off
    (the default), the verb behaves exactly as before and neither key appears.
    """
    archive_folder = get_outlook_archive_folder(cfg)
    category = get_outlook_category(cfg)
    roots = _archive_roots(cfg)
    results: list[dict[str, Any]] = []

    conn = init_db(cfg["database"]["path"])
    repo = EmailRepository(conn)
    try:
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
                    continue
                if resolved.suffix.lower() == ".msg":
                    result["index_rows_removed"] += repo.delete_by_path(str(resolved))
            conn.commit()

            results.append(_finish_revert_item(client, archive_folder, category, result, message_id))
        doc = _envelope("revert", cfg)
        if renumber:
            # The source folders: a gap only exists where a file really went.
            doc["renumbered"], doc["renumber_refused"] = _renumber_folders(
                cfg, repo,
                _unique_folders([
                    os.path.dirname(path)
                    for r in results for path in r["deleted"]
                ]),
            )
    finally:
        conn.close()

    reverted = sum(1 for r in results if r["ok"])
    doc["counts"] = {
        "requested": len(results),
        "reverted": reverted,
        "failed": len(results) - reverted,
    }
    doc["results"] = results
    logger.info("Revert: %d of %d mail(s) restored.", reverted, len(results))
    return doc


def _finish_revert_item(
    client: Any,
    archive_folder: str,
    category: str,
    result: dict[str, Any],
    message_id: str,
) -> dict[str, Any]:
    """The Outlook side of one ``revert`` item, after its files are handled."""
    if not message_id:
        result["error"] = {
            "code": ERROR_BAD_DECISION,
            "message": "an item needs a message_id to move the mail back",
        }
        return result

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
        return result

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
        return result

    result["ok"] = not result["refused"] and not result["file_errors"]
    return result


# --------------------------------------------------------------- renumber ---

def renumber(
    cfg: dict[str, Any], folder_path: str, *, dry_run: bool = False
) -> dict[str, Any]:
    """Renumber one folder and return the old → new map.

    The only verb that touches neither Outlook nor COM: it reads a folder
    listing and the index, and renames files. ``dry_run`` reports exactly the
    same map without changing anything on disk.

    The map is reported under the same ``renumbered`` key ``apply --renumber``
    and ``revert --renumber`` use, so a consumer healing its stored ``.msg``
    paths reads one shape whichever verb produced it.
    """
    normalised = os.path.normpath(folder_path)
    conn = init_db(cfg["database"]["path"])
    repo = EmailRepository(conn)
    try:
        result = renumber_folder(
            folder_path, repo, _archive_roots(cfg), dry_run=dry_run
        )
    except Exception as exc:
        # Never a traceback over an empty stdout: this process's contract is one
        # JSON document, in one shape, whatever happened. A run that stopped
        # part-way is refused with its reason and reports no map — a map the
        # disk may not match is worse to hand a consumer than none at all.
        logger.exception("Renumbering %r failed", normalised)
        result = RenumberResult(
            folder_path=normalised,
            dry_run=dry_run,
            refused=f"{type(exc).__name__}: {exc}",
        )
    finally:
        conn.close()

    doc = _envelope("renumber", cfg)
    doc["dry_run"] = result.dry_run
    doc["folder_path"] = result.folder_path
    doc["first_number"] = result.first_number
    doc["last_number"] = result.last_number
    doc["counts"] = result.counts()
    doc["skipped"] = result.skipped
    doc["renumbered"] = {} if result.refused else {result.folder_path: result.renamed}
    doc["renumber_refused"] = (
        [{"folder_path": result.folder_path, "reason": result.refused}]
        if result.refused else []
    )
    return doc

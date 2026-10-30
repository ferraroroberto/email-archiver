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
- **The reference that is moved is not the one that was archived.** The move,
  its retries and the ``message_changed`` error are ``email_archiver.batch_move``.
- **Nothing is written outside the archive roots.** ``revert`` refuses per
  file, and ``apply`` refuses per decision, any path that does not resolve
  inside ``archive.root_paths`` — so a malformed, hostile or LLM-invented
  decisions file cannot reach the rest of the disk. A new folder *inside* a
  root is still created on demand.
- **``plan`` is targetable, and untargeted by default.** ``--message-id``,
  ``--since``, ``--search`` and ``--ref`` narrow it for a caller that already
  knows which mail it wants to file; with none of them the document is exactly
  what a full-Inbox plan always was, with no ``filters`` key at all.
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
from collections.abc import Callable, Iterator, Sequence
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from email_archiver.archiver.archiver import EmailArchiver, resolve_date_prefix_for_folder
from email_archiver.config import (
    DATE_PREFIX_AUTO,
    Mailbox,
    MailboxRegistry,
    get_mailbox_archive_folder,
    get_outlook_category,
)
from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRepository
from email_archiver.engine.suggester import SuggestionEngine
from email_archiver.outlook.client import EmailData
from email_archiver.batch_common import (
    SCHEMA_VERSION as SCHEMA_VERSION,
    envelope,
    error_document as error_document,
    now_iso as now_iso,
    renumber_folders,
    unique_folders,
)
from email_archiver.batch_move import (
    ERROR_MESSAGE_CHANGED as ERROR_MESSAGE_CHANGED,
    ERROR_MOVE_FAILED,
    MOVE_VIA_FOLDER_SWITCH as MOVE_VIA_FOLDER_SWITCH,
    MOVE_VIA_KEPT_IN_SENT,
    MOVE_VIA_ORIGINAL as MOVE_VIA_ORIGINAL,
    MOVE_VIA_REFETCHED as MOVE_VIA_REFETCHED,
    MOVE_VIA_RELEASED_RETRY as MOVE_VIA_RELEASED_RETRY,
    move_failure_error,
    move_out_of_inbox,
)
from email_archiver.outlook.mapi import MailboxNotInOutlookError
from email_archiver.paths import (
    REFUSED_OUTSIDE_ROOTS as REFUSED_OUTSIDE_ROOTS,
    REFUSED_UNRESOLVABLE as REFUSED_UNRESOLVABLE,
    archive_roots as _archive_roots,
    resolve_under_roots as _resolve_under_roots,
)
from email_archiver.renumber import RenumberResult, renumber_folder
from email_archiver.text import normalize_message_id

logger = logging.getLogger(__name__)

# Default number of ranked folder candidates `plan` reports per mail. Larger
# than the dialog's three: the caller re-ranks (an LLM, in task-os's case) and
# needs more than the engine's own top pick to choose from.
DEFAULT_CANDIDATES = 10

# Why a mail got no plan entry.
SKIP_NO_MESSAGE_ID = "no_message_id"
SKIP_NOT_IN_INBOX = "not_in_inbox"  # a --message-id that matched nothing
SKIP_NOT_IN_SENT = "not_in_sent"    # the same, under `--folder sent`

# Where a plan reads from and an apply decision files from.
# The Inbox is the default and the only source that moves the mail out; Sent
# Items is opt-in, for a sent mail whose self-BCC copy never reached the Inbox.
SOURCE_INBOX = "inbox"
SOURCE_SENT = "sent"
SOURCES = (SOURCE_INBOX, SOURCE_SENT)

# `error.code` values, and the reason each one is worth telling apart.
ERROR_NOT_IN_INBOX = "not_in_inbox"            # already moved, or never there
ERROR_NOT_IN_SENT = "not_in_sent"              # a `source: sent` mail not in Sent Items
ERROR_NOT_IN_ARCHIVE = "not_in_archive_folder"  # revert cannot find it back
ERROR_BAD_DECISION = "bad_decision"             # caller sent an unusable entry
ERROR_ARCHIVE_FAILED = "archive_failed"         # disk write / Outlook SaveAs
# The move landed and only the tag did not. Its own code because the two are
# genuinely different states to recover from: after a move_failed the mail is
# still in the Inbox, after a category_failed it is already filed and only
# looks untouched in Outlook.
ERROR_CATEGORY_FAILED = "category_failed"

# The per-file outcomes in a revert result — REFUSED_OUTSIDE_ROOTS and
# REFUSED_UNRESOLVABLE — are imported above from `paths`, which owns the
# archive-root guard now that `renumber` applies it too. They stay readable as
# `batch.REFUSED_*`, which is where consumers and tests look for them.

# A mailbox's `outlook_status` in the `mailboxes` document.
MAILBOX_PRESENT = "present"   # its store is in the running Outlook
MAILBOX_MISSING = "missing"   # Outlook answered; the mailbox is not in it
MAILBOX_UNKNOWN = "unknown"   # Outlook not running, or the check could not run


# ---------------------------------------------------------------- helpers ---

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


def _mail_dict(mail: Any, *, in_inbox: bool = True) -> dict[str, Any]:
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
        # True for every Inbox plan — it enumerates the Inbox, so a mail it
        # reports is in it by construction (false only under `--folder sent`,
        # where the mail is in Sent Items). Stated anyway because it is the
        # fact that makes an ``already_archived`` mail *retryable*: its files
        # are on disk but it never left the Inbox, and a consumer reading only
        # ``already_archived`` cannot tell that from a mail that is properly
        # filed and gone.
        "in_inbox": in_inbox,
    }


# ------------------------------------------------------------------- plan ---

@dataclass(frozen=True)
class PlanFilters:
    """What narrows a ``plan``. All optional and combinable (AND); the default,
    nothing set, is the full-Inbox plan.

    - ``message_ids``: look each one up instead of enumerating the Inbox.
    - ``since``: only mail received at/after it (naive local time).
    - ``search``: every term must appear, case-insensitively, in the subject,
      sender, recipients or body preview.
    - ``ref``: the mail's ``X-Archive-Ref`` header equals it exactly.
    """

    message_ids: tuple[str, ...] = ()
    since: datetime | None = None
    search: tuple[str, ...] = ()
    ref: str | None = None
    folder: str = SOURCE_INBOX

    @property
    def active(self) -> bool:
        return bool(
            self.message_ids or self.since or self.search or self.ref
            or self.folder != SOURCE_INBOX
        )

    def as_dict(self) -> dict[str, Any]:
        doc = {
            "message_ids": list(self.message_ids),
            "since": self.since.isoformat(timespec="minutes") if self.since else None,
            "search": list(self.search),
            "ref": self.ref,
        }
        # Only when it is not the default, so an Inbox plan's document stays
        # exactly what it always was.
        if self.folder != SOURCE_INBOX:
            doc["folder"] = self.folder
        return doc


def _plan_source(
    client: Any, preview_len: int, filters: PlanFilters, not_found: list[str]
) -> Iterator[Any]:
    """The mails a plan looks at: looked up by id, restricted by date, or the
    whole Inbox — cheapest first. A ``--message-id`` matching nothing is
    appended to ``not_found``."""
    sent = filters.folder == SOURCE_SENT
    if filters.message_ids:
        for message_id in filters.message_ids:
            item = (
                client.find_sent_by_message_id(message_id) if sent
                else client.find_by_message_id(message_id, None)
            )
            if item is None:
                not_found.append(message_id)
                continue
            yield client.read_mail(item, preview_len)
    elif filters.since is not None:
        if sent:
            yield from client.iter_sent_since(filters.since, preview_len)
        else:
            yield from client.iter_inbox_received_since(filters.since, preview_len)
    elif sent:
        yield from client.iter_sent(preview_len)
    else:
        yield from client.iter_inbox(preview_len)


def mail_matches(
    mail: Any, terms: Sequence[str], since: datetime | None, sent: bool
) -> bool:
    """Whether ``mail`` carries every term (casefolded, in the subject, sender,
    recipients or body preview) and is dated at or after ``since``.

    The one per-mail rule shared by ``plan --search``/``--since`` and the
    ``find`` folder-walk fallback. ``since`` is checked here even after a
    server-side Restrict: this is the exact comparison, the Restrict only
    narrows the walk."""
    if since is not None:
        # A Sent Items mail's own clock is when it was sent; its received time
        # is whatever the store stamped and means nothing here.
        when = (
            (mail.date_sent or mail.date_received)
            if sent
            else (mail.date_received or mail.date_sent)
        )
        if when is None or when < since:
            return False
    if terms:
        haystack = "\n".join(
            (mail.subject, mail.sender, mail.recipients, mail.body_preview)
        ).casefold()
        if not all(term.casefold() in haystack for term in terms):
            return False
    return True


def _matches(client: Any, mail: Any, filters: PlanFilters) -> bool:
    """Whether one mail passes every filter, the cheap checks before the
    header read."""
    if not mail_matches(mail, filters.search, filters.since, filters.folder == SOURCE_SENT):
        return False
    if filters.ref is not None and client.archive_ref(mail.item) != filters.ref:
        return False
    return True


def plan(
    client: Any,
    cfg: dict[str, Any],
    *,
    candidates: int = DEFAULT_CANDIDATES,
    filters: PlanFilters | None = None,
    mailbox: Mailbox | None = None,
) -> dict[str, Any]:
    """Rank archive folders for every Inbox mail, or for the ones ``filters``
    select.

    Unfiltered, every Inbox mail is offered: nothing is filtered on age,
    sender or size. A mail whose Message-ID is already in the index is reported
    with ``already_archived`` set to the file it was archived as, and no
    candidates — ranking a mail that is already filed would only invite filing
    it twice.

    ``candidates=0`` reports the matching mails with no ranking at all, for a
    caller that already knows the destination. With any filter set the
    document gains ``filters`` (what was applied), ``counts.inbox`` counts the
    Inbox mails that matched, and a ``--message-id`` found nowhere in the Inbox
    is listed under ``skipped`` with ``reason: "not_in_inbox"`` and its
    ``message_id``.
    """
    filters = filters or PlanFilters()
    preview_len = int((cfg.get("scanning") or {}).get("body_preview_length", 500))
    engine = _engine_for(cfg, candidates) if candidates > 0 else None

    conn = init_db(cfg["database"]["path"])
    repo = EmailRepository(conn)

    mails: list[dict[str, Any]] = []
    skipped: list[dict[str, Any]] = []
    not_found: list[str] = []
    already = 0
    # Shared across every mail in this run — see _candidate_dict.
    date_prefix_cache: dict[str, bool] = {}

    try:
        for mail in _plan_source(client, preview_len, filters, not_found):
            if filters.active and not _matches(client, mail, filters):
                continue
            if not mail.message_id:
                skipped.append({
                    "entry_id": mail.entry_id,
                    "subject": mail.subject,
                    "reason": SKIP_NO_MESSAGE_ID,
                })
                continue

            entry = _mail_dict(mail, in_inbox=filters.folder == SOURCE_INBOX)
            archived_as = repo.find_path_by_message_id(mail.message_id)
            entry["already_archived"] = archived_as
            if archived_as:
                already += 1
                entry["candidates"] = []
            elif engine is None:
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

    listed = len(mails) + len(skipped)
    not_found_reason = SKIP_NOT_IN_SENT if filters.folder == SOURCE_SENT else SKIP_NOT_IN_INBOX
    skipped.extend(
        {"entry_id": "", "subject": "", "reason": not_found_reason, "message_id": mid}
        for mid in not_found
    )

    doc = envelope("plan", cfg, mailbox=mailbox)
    if filters.active:
        doc["filters"] = filters.as_dict()
    doc["counts"] = {
        # The mails the plan read: "inbox" as it always was, "sent" under
        # `--folder sent` so the key never says Inbox about Sent Items.
        filters.folder: listed,
        "planned": len(mails) - already,
        "already_archived": already,
        "skipped": len(skipped),
    }
    doc["mails"] = mails
    doc["skipped"] = skipped
    logger.info(
        "Plan: %d mail(s) in %s%s, %d already archived, %d skipped.",
        listed, "Sent Items" if filters.folder == SOURCE_SENT else "the Inbox",
        f" matching {filters.as_dict()}" if filters.active else "",
        already, len(skipped),
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
    category: str | None = None,
    mailbox: Mailbox | None = None,
) -> dict[str, Any]:
    """Archive each decided mail, move it out of the Inbox and tag it.

    ``decisions`` is a list of ``{message_id, folder_path, date_prefix}``. The
    ``date_prefix`` flag is per mail and comes from the caller — batch mode
    never reads the global ``naming.date_prefix`` toggle for a boolean, because
    the caller infers the form the destination folder actually uses. A caller
    that picked a folder ``plan`` never suggested passes ``"auto"`` instead,
    and gets exactly the form ``plan`` would have reported for that folder
    (``resolve_date_prefix_for_folder``).

    A ``folder_path`` that does not resolve inside ``archive.root_paths`` is a
    ``bad_decision`` for that mail: nothing is looked up, written or moved, and
    the next decision is still attempted.

    ``category`` overrides ``outlook.category`` for this run; the document's
    ``category`` reports the one used. ``mailbox`` supplies the archive
    folder (its own, else ``outlook.archive_folder``); the client it came with
    is already bound to that mailbox's store.

    A decision may carry ``"source": "sent"`` (default ``"inbox"``, anything
    else is a ``bad_decision``): the mail is then looked up in Sent Items,
    archived and tagged like any other, but **not moved** — ``moved`` stays
    false and ``move_via`` is ``kept_in_sent``. ``revert`` does not
    undo that move, since there was none: it deletes the files and reports the
    mail ``not_in_archive_folder``.

    A decision for a mail the index already holds — the one ``plan`` reported
    ``already_archived``, still sitting in the Inbox because an earlier run
    wrote its files and then failed to move it — writes **nothing**: the
    existing ``.msg`` is reported back as ``files``, ``reused`` is true, and
    the mail is moved and tagged like any other. That is what makes such a mail
    finishable at all; re-archiving it would file a second copy beside the
    first.

    Whatever happens to one mail, the next is still attempted; the per-mail
    ``error`` says what went wrong and ``files`` says what is already on disk.

    With ``renumber`` on, every destination folder this run put a bundle in is
    renumbered once afterwards — a mail older than the folder's newest file
    took ``max + 1`` on the way in, and this is what puts it back in date order
    — and the document carries the old → new map per folder under
    ``renumbered``. Off (the default), the verb behaves exactly as before and
    neither key appears.
    """
    archive_folder = get_mailbox_archive_folder(cfg, mailbox)
    category = category or get_outlook_category(cfg)
    roots = _archive_roots(cfg)
    results: list[dict[str, Any]] = []

    conn = init_db(cfg["database"]["path"])
    repo = EmailRepository(conn)
    try:
        for raw in decisions:
            results.append(
                _apply_one(client, cfg, repo, raw, archive_folder, category, roots)
            )
        doc = envelope("apply", cfg, category, mailbox)
        if renumber:
            # Only a folder this run actually wrote into, and only once each.
            doc["renumbered"], doc["renumber_refused"] = renumber_folders(
                cfg, repo,
                unique_folders([r["folder_path"] for r in results if r["files"]]),
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


def _decision_refusal(
    raw: dict[str, Any], message_id: str, folder_path: str, roots: list[Path]
) -> dict[str, str] | None:
    """The ``bad_decision`` error for a decision that cannot be acted on, else ``None``."""
    if not message_id or not folder_path:
        return {
            "code": ERROR_BAD_DECISION,
            "message": "a decision needs both a message_id and a folder_path",
        }

    # An unknown source is refused, never read as the Inbox: filing the wrong
    # copy of a mail is not recoverable the way not filing one is.
    source = raw.get("source", SOURCE_INBOX)
    if source not in SOURCES:
        return {
            "code": ERROR_BAD_DECISION,
            "message": f"source {source!r} is not one of {', '.join(SOURCES)}",
        }

    # Before Outlook is asked anything: the archiver creates a missing folder,
    # so an unchecked path is a write anywhere on the disk.
    _, refusal = _resolve_under_roots(folder_path, roots)
    if refusal is not None:
        logger.warning(
            "Refusing the decision for %s: folder_path %r is %s.",
            message_id, folder_path, refusal,
        )
        return {
            "code": ERROR_BAD_DECISION,
            "message": f"{refusal}: folder_path must resolve inside archive.root_paths",
        }
    return None


def _archive_decision(
    client: Any,
    cfg: dict[str, Any],
    repo: EmailRepository,
    raw: dict[str, Any],
    item: Any,
    result: dict[str, Any],
) -> bool:
    """Write the decided mail's files, or find the ones an earlier run wrote.

    Fills ``files`` (and ``sequence_number`` / ``reused``) in ``result``; on a
    failure sets ``error`` instead and returns False.
    """
    message_id, folder_path = result["message_id"], result["folder_path"]
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
        return True

    try:
        mail = client.read_mail(item)
        archiver = EmailArchiver(
            cfg, date_prefix=_decision_date_prefix(cfg, raw, folder_path, message_id)
        )
        archived = archiver.archive(item, folder_path, mail.subject)
    except Exception as exc:
        logger.exception("Archiving %s failed", message_id)
        result["error"] = {
            "code": ERROR_ARCHIVE_FAILED,
            "message": f"{type(exc).__name__}: {exc}",
        }
        return False

    result["sequence_number"] = archived.sequence_number
    # Filled in before the move on purpose: if the move fails these files
    # exist and the caller must be able to revert them.
    result["files"] = [archived.email_path, *archived.attachment_paths]
    return True


def _apply_one(
    client: Any,
    cfg: dict[str, Any],
    repo: EmailRepository,
    raw: dict[str, Any],
    archive_folder: str,
    category: str,
    roots: list[Path],
) -> dict[str, Any]:
    """One decision, start to finish. Never raises: every failure is a result."""
    message_id = normalize_message_id(raw.get("message_id"))
    folder_path = str(raw.get("folder_path") or "")
    result = _blank_apply_result(message_id, folder_path)

    result["error"] = _decision_refusal(raw, message_id, folder_path, roots)
    if result["error"] is not None:
        return result
    from_sent = raw.get("source", SOURCE_INBOX) == SOURCE_SENT

    where = "Sent Items" if from_sent else "the Inbox"
    not_found_code = ERROR_NOT_IN_SENT if from_sent else ERROR_NOT_IN_INBOX
    try:
        item = (
            client.find_sent_by_message_id(message_id) if from_sent
            else client.find_by_message_id(message_id, None)
        )
    except Exception as exc:  # a COM failure on one lookup, not the run
        result["error"] = {
            "code": not_found_code,
            "message": f"{where} lookup failed: {type(exc).__name__}: {exc}",
        }
        return result

    if item is None:
        result["error"] = {
            "code": not_found_code,
            "message": f"no mail with this Message-ID is in {where}",
        }
        return result

    if not _archive_decision(client, cfg, repo, raw, item, result):
        return result

    try:
        if from_sent:
            # Filed from Sent Items: the files are written, the mail stays.
            moved, via = item, MOVE_VIA_KEPT_IN_SENT
            logger.info("Archived %s from Sent Items; leaving it there.", message_id)
        else:
            # The move takes over the only reference: nothing in this frame
            # may keep the mail alive while Outlook is asked to let go of it.
            held, item = [item], None
            moved, via = move_out_of_inbox(client, held, archive_folder, message_id)
            result["moved"] = True
        result["move_via"] = via
        result["entry_id"] = client.entry_id(moved)
    except Exception as exc:
        logger.exception("Moving %s to %r failed", message_id, archive_folder)
        result["error"] = move_failure_error(client, exc, message_id)
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
            "message": f"the mail was filed{'' if from_sent else ' and moved'}, but not tagged: "
                       f"{type(exc).__name__}: {exc}",
        }
        return result

    result["ok"] = True
    return result


def _decision_date_prefix(
    cfg: dict[str, Any], raw: dict[str, Any], folder_path: str, message_id: str
) -> bool:
    """A decision's ``date_prefix`` as the boolean the archiver takes.

    ``"auto"`` resolves against the destination folder exactly as ``plan``
    resolves a candidate's; anything else keeps its long-standing boolean
    reading, so an existing caller's decisions file means what it always did.
    """
    value = raw.get("date_prefix", False)
    if isinstance(value, str) and value.strip().casefold() == DATE_PREFIX_AUTO:
        resolved = resolve_date_prefix_for_folder(cfg, folder_path)
        logger.info(
            "date_prefix auto for %s resolved to %s for its folder.",
            message_id, resolved,
        )
        return resolved
    return bool(value)


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
    category: str | None = None,
    mailbox: Mailbox | None = None,
) -> dict[str, Any]:
    """Delete the listed archive files and put each mail back in the Inbox.

    ``category`` is the one to remove, overriding ``outlook.category`` — pass
    the same one the ``apply`` used.

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
    archive_folder = get_mailbox_archive_folder(cfg, mailbox)
    category = category or get_outlook_category(cfg)
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
        doc = envelope("revert", cfg, category, mailbox)
        if renumber:
            # The source folders: a gap only exists where a file really went.
            doc["renumbered"], doc["renumber_refused"] = renumber_folders(
                cfg, repo,
                unique_folders([
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


# -------------------------------------------------------------- mailboxes ---

def mailboxes(
    registry: MailboxRegistry, client_for: Callable[[Mailbox], Any] | None
) -> dict[str, Any]:
    """The registry as a document, with each mailbox's ``outlook_status``.

    ``client_for`` builds a client bound to one mailbox; ``None`` means
    Outlook is not running or cannot be reached, which makes every status
    ``unknown`` — never ``present``, which would claim a check that did not
    run. ``missing`` means Outlook answered and the mailbox has no store in it;
    any other failure of the check is ``unknown`` too, with its reason. This is
    how a consumer learns the mailbox list: it never reads the file.
    """
    entries: list[dict[str, Any]] = []
    for mailbox in registry.mailboxes.values():
        address = mailbox.address
        status, reason = MAILBOX_UNKNOWN, "Outlook is not running"
        if client_for is not None:
            client = client_for(mailbox)
            try:
                client.store()
                status, reason = MAILBOX_PRESENT, ""
                if mailbox.synthesized:
                    address = client.default_account_smtp()
            except MailboxNotInOutlookError as exc:
                status, reason = MAILBOX_MISSING, str(exc)
            except Exception as exc:  # a check that could not run is not "missing"
                status, reason = MAILBOX_UNKNOWN, f"{type(exc).__name__}: {exc}"
        entries.append({
            "alias": mailbox.alias,
            "address": address,
            "display_name": mailbox.display_name,
            "aliases": list(mailbox.aliases),
            "archive_folder": mailbox.archive_folder,
            "default": mailbox.alias == registry.default,
            "outlook_status": status,
            "outlook_status_reason": reason,
        })
    logger.info(
        "Mailboxes: %s.", ", ".join(f"{e['alias']} {e['outlook_status']}" for e in entries)
    )
    return {
        "verb": "mailboxes",
        "schema_version": SCHEMA_VERSION,
        "generated_at": now_iso(),
        "registry": "synthesized" if registry.synthesized else "file",
        "default": registry.default,
        "mailboxes": entries,
    }


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

    doc = envelope("renumber", cfg)
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

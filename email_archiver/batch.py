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
  re-acquired by EntryID and StoreID, with the original let go first. On that
  refusal anyway, every reference to the mail is let go, Outlook's view is
  parked off the Inbox when it is showing it (the Explorer holds its selected
  mail too, issue #111), and the mail is fetched again and moved while the view
  is still away (issue #130); it is never ``Save()``d. Which path finished it is
  logged and reported as ``move_via``.
  A refusal that survives all of that is reported as ``message_changed`` rather
  than ``move_failed`` — usually a mail left open in an Outlook window, and the
  message says so.
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

import gc
import hashlib
import logging
import os
import time
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
from email_archiver.outlook.mapi import MailboxNotInOutlookError, is_message_changed_error
from email_archiver.paths import (
    REFUSED_OUTSIDE_ROOTS as REFUSED_OUTSIDE_ROOTS,
    REFUSED_UNRESOLVABLE as REFUSED_UNRESOLVABLE,
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
ERROR_MOVE_FAILED = "move_failed"               # files are on disk, mail is not
# A move refused as MAPI_E_OBJECT_CHANGED after both retries — almost always a
# mail left open in an Outlook window. Its own code because it is the one move
# failure an operator clears in seconds without anything being wrong: close the
# window, apply the same decision again. A caller that cannot tell it from a
# plain `move_failed` has to treat both as needing a human.
ERROR_MESSAGE_CHANGED = "message_changed"
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
# `move_via`. Genuinely different stories about the same mail, and the one
# place that says which defence is earning its keep on this mailbox — see
# `_move_out_of_inbox`.
MOVE_VIA_REFETCHED = "refetched"     # a reference read back out of the store
MOVE_VIA_ORIGINAL = "original"       # the re-acquire failed; the original moved
# Refused with 0x80040109, then moved on a reference fetched again once every
# other one was let go; Outlook's view was not on the Inbox, so it was left
# alone. Replaces `saved_retry`: a refused mail is never saved.
MOVE_VIA_RELEASED_RETRY = "released_retry"
# The same, with Outlook's Explorer showing the Inbox and so holding the mail as
# its selection: the view was parked on another folder while the move was
# retried, then put back. The one path that disturbs the
# user's view.
MOVE_VIA_FOLDER_SWITCH = "folder_switch"
# A `source: sent` mail is archived and tagged but never moved: Sent Items is
# the user's record of what went out, so the mail stays where it is.
MOVE_VIA_KEPT_IN_SENT = "kept_in_sent"

# The pauses before each fresh-reference retry of a refused move, all made
# while the view is away: one straight after it settles, one more after a
# pause. Two tries, not more: the view is the user's, and a refusal that
# outlasts both is reported as `message_changed` with its remedy.
MOVE_RETRY_DELAYS_SECONDS = (0.0, 2.0)

# A mailbox's `outlook_status` in the `mailboxes` document.
MAILBOX_PRESENT = "present"   # its store is in the running Outlook
MAILBOX_MISSING = "missing"   # Outlook answered; the mailbox is not in it
MAILBOX_UNKNOWN = "unknown"   # Outlook not running, or the check could not run


# ---------------------------------------------------------------- helpers ---

def now_iso() -> str:
    """The local-time timestamp every batch-family document carries."""
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


def _envelope(
    verb: str, cfg: dict[str, Any], category: str | None = None,
    mailbox: Mailbox | None = None,
) -> dict[str, Any]:
    """The fields every verb's document opens with. ``category`` is the one
    this run actually used, when the caller overrode the config's;
    ``archive_folder`` is the one of ``mailbox`` (``main_batch`` adds the
    ``mailbox`` key itself, to every Outlook verb's document)."""
    return {
        "verb": verb,
        "schema_version": SCHEMA_VERSION,
        "generated_at": now_iso(),
        "archive_folder": get_mailbox_archive_folder(cfg, mailbox),
        "category": category or get_outlook_category(cfg),
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


def error_document(verb: str, code: str, message: str, **details: Any) -> dict[str, Any]:
    """The document printed when the run could not start at all.

    Deliberately the same envelope shape minus the results, so a consumer
    parses one JSON document either way and branches on ``"error" in doc``
    rather than on the exit code alone. ``details`` add keys to ``error``
    (``send``'s refusal names the parts that differ under ``differs``).
    """
    return {
        "verb": verb,
        "schema_version": SCHEMA_VERSION,
        "generated_at": now_iso(),
        "error": {"code": code, "message": message, **details},
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

    doc = _envelope("plan", cfg, mailbox=mailbox)
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
    false and ``move_via`` is ``kept_in_sent`` (issue #103). ``revert`` does not
    undo that move, since there was none: it deletes the files and reports the
    mail ``not_in_archive_folder``.

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
        doc = _envelope("apply", cfg, category, mailbox)
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
    roots: list[Path],
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

    # An unknown source is refused, never read as the Inbox: filing the wrong
    # copy of a mail is not recoverable the way not filing one is.
    source = raw.get("source", SOURCE_INBOX)
    if source not in SOURCES:
        result["error"] = {
            "code": ERROR_BAD_DECISION,
            "message": f"source {source!r} is not one of {', '.join(SOURCES)}",
        }
        return result
    from_sent = source == SOURCE_SENT

    # Before Outlook is asked anything: the archiver creates a missing folder,
    # so an unchecked path is a write anywhere on the disk.
    _, refusal = _resolve_under_roots(folder_path, roots)
    if refusal is not None:
        logger.warning(
            "Refusing the decision for %s: folder_path %r is %s.",
            message_id, folder_path, refusal,
        )
        result["error"] = {
            "code": ERROR_BAD_DECISION,
            "message": f"{refusal}: folder_path must resolve inside archive.root_paths",
        }
        return result

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
                cfg, date_prefix=_decision_date_prefix(cfg, raw, folder_path, message_id)
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
        if from_sent:
            # Filed from Sent Items: the files are written, the mail stays.
            moved, via = item, MOVE_VIA_KEPT_IN_SENT
            logger.info("Archived %s from Sent Items; leaving it there.", message_id)
        else:
            # The move takes over the only reference: nothing in this frame
            # may keep the mail alive while Outlook is asked to let go of it
            # `mail` goes too — an InboxMail carries the item.
            held, item, mail = [item], None, None
            moved, via = _move_out_of_inbox(client, held, archive_folder, message_id)
            result["moved"] = True
        result["move_via"] = via
        result["entry_id"] = client.entry_id(moved)
    except Exception as exc:
        logger.exception("Moving %s to %r failed", message_id, archive_folder)
        result["error"] = _move_failure_error(client, exc, message_id)
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


def _move_failure_error(
    client: Any, exc: Exception, message_id: str
) -> dict[str, str]:
    """The ``error`` a failed ``apply`` move reports: its code and its message.

    Anything other than MAPI_E_OBJECT_CHANGED after both defences is a plain
    ``move_failed`` — a store that refused the write for its own reasons is a
    different problem, and dressing it up as this one would hide it.

    A move refused as *the message has been changed* on a re-acquired
    reference *and* on fresh ones fetched after every reference was let go is
    its own condition, and a recoverable one: the files are on disk, the mail
    is still in the Inbox, and something in the running Outlook is holding the
    item. Almost always that something is an open mail window — an inspector
    holds its item for the life of the window, so both defences fail by
    construction (issue #76: a live filing run refused this way finished on
    the first attempt the moment the window was closed, with no restart).
    Issue #59 saw the same refusal clear on an Outlook restart, which is why
    that stays as the fallback rather than the headline.

    When no window holds the mail the cause is still unresolved (issue #84).
    What the 2026-09-16 occurrence ruled out: an inspector (the probe read
    every window and none held it); the modified flag ``SaveAs`` leaves behind
    (the two re-applies wrote nothing, and the first ``Move`` of a freshly
    re-acquired reference was still refused); and anything transient (the
    ``Save()`` meant to clear the flag was refused too, three times across
    separate processes over 40 seconds). So it is state held on the item in
    the running Outlook, not a race, and a bounded retry-with-delay would not
    have cleared it. The message therefore says outright that re-applying
    before a restart is not expected to help, rather than inviting a caller to
    burn retries on it; restarting stays the user's call, never this tool's.

    So the message leads with the cheap remedy, and says which of the two it
    is on the evidence rather than on the guess: Outlook is asked whether this
    mail is actually open. A check that could not be completed is reported as
    *not determined* and never as "no window", because pointing an operator at
    a restart they do not need is exactly the failure this replaces.

    Issue #111 found a third holder: the Explorer showing the mail as its
    selected / reading-pane item, freed by selecting another folder and the
    message again with no restart. ``_move_out_of_inbox`` does that switch
    itself, retrying the move while the view is away (issue #130), and a
    refusal it wraps in ``_MoveRefused`` says whether it ran. The
    restart advice is kept only for one that ran and did not help; otherwise
    the message gives the by-hand switch as the remedy.
    """
    if not is_message_changed_error(getattr(exc, "cause", exc)):
        return {
            "code": ERROR_MOVE_FAILED,
            "message": f"{type(exc).__name__}: {exc}",
        }

    folder_switched = getattr(exc, "folder_switched", None)
    entry_id = getattr(exc, "entry_id", "")
    exc = getattr(exc, "cause", exc)

    try:
        open_in_window = client.is_open_in_inspector(entry_id)
    except Exception as probe_exc:  # the probe is a courtesy, never the run
        logger.warning(
            "Could not check whether %s is open in a window: %s: %s",
            message_id, type(probe_exc).__name__, probe_exc,
        )
        open_in_window = None
    logger.info(
        "Move of %s was refused as a changed message; open in a window: %s; "
        "Outlook's view parked on another folder for the retries: %s.",
        message_id, {True: "yes", False: "no", None: "not determined"}[open_in_window],
        {True: "yes", False: "not applicable", None: "not determined"}[folder_switched],
    )

    # Only a switch that actually happened and did not help is evidence the
    # hold outlasts one. One that was not applicable or could not run says
    # nothing, so it points the user at the same switch by hand, not a restart.
    tried_switch = (
        "Outlook's view was moved to another folder automatically and the move "
        "retried there, and that did not free it. "
    )
    by_hand = (
        "Select another folder in Outlook, then this message again, and apply "
        "this same decision again; restart Outlook only if that does not clear it."
    )
    if open_in_window is True:
        cause = (
            "This mail is open in an Outlook window, which holds it and "
            "refuses every write to it. Close that window and apply this same "
            "decision again."
        )
    elif open_in_window is False and folder_switched is True:
        cause = (
            "This mail is not open in any Outlook window, so something else "
            f"in the running Outlook is holding it. {tried_switch}"
            "That hold outlasts re-applies, so applying again before a restart "
            "is not expected to help. Restart Outlook, then apply this same "
            "decision again. Deleting the mail from the Inbox by hand instead "
            "leaves the files already written as its only copy, with nothing in "
            "Outlook pointing at them."
        )
    elif open_in_window is False:
        cause = (
            "This mail is not open in any Outlook window, but Outlook may "
            f"still be holding it as the message selected in its folder view. {by_hand}"
        )
    elif folder_switched is True:
        cause = (
            "Whether this mail is open in an Outlook window could not be "
            f"determined. {tried_switch}A mail left open is the usual cause, so "
            "close any window showing it and apply this same decision again; "
            "restart Outlook if that does not clear it."
        )
    else:
        cause = (
            "Whether this mail is open in an Outlook window could not be "
            "determined. A mail left open is the usual cause, so close any "
            "window showing it and apply this same decision again. Failing "
            "that, Outlook may be holding it as the message selected in its "
            "folder view: select another folder in Outlook, then this message "
            "again, and apply again; restart Outlook if that does not clear it."
        )

    return {
        "code": ERROR_MESSAGE_CHANGED,
        "message": (
            f"{type(exc).__name__}: {exc} — Outlook refused the move as a "
            f"changed message on a re-acquired reference, and again on a "
            f"fresh one after every reference to it was released. {cause} "
            f"The files are on disk and the mail is still "
            f"in the Inbox, so re-applying writes nothing."
        ),
    }


class _MoveRefused(Exception):
    """A move still refused as a changed message after every retry.

    Carries the refusal itself (``cause``), what the folder switch did, and the
    mail's EntryID as a plain string (``entry_id``, ``""`` when unreadable) —
    the caller let go of the mail itself before the move (issue #130), so the
    window check behind ``_move_failure_error`` asks by id. A switch that
    happened and did not help is the only evidence a restart is what is left.
    ``folder_switched`` is ``True`` (it ran and did not help), ``False`` (the
    Explorer was not on the Inbox, nothing to switch) or ``None`` (it could not
    be determined or could not run).
    """

    def __init__(
        self, cause: Exception, folder_switched: bool | None, entry_id: str
    ) -> None:
        super().__init__(str(cause))
        self.cause = cause
        self.folder_switched = folder_switched
        self.entry_id = entry_id


def _store_note(store_id: str) -> str:
    """How a log line names the StoreID a re-fetch used: a short fingerprint,
    never the id itself — a StoreID can embed the store's file path."""
    if not store_id:
        return "without a StoreID (none was readable)"
    return f"with StoreID #{hashlib.sha1(store_id.encode()).hexdigest()[:8]}"


def _yes_no(answer: bool | None) -> str:
    return {True: "yes", False: "no", None: "not determined"}[answer]


def _refusal_without_frames(exc: Exception) -> Exception:
    """The refusal with its traceback dropped, so keeping it keeps no frame.

    A traceback keeps every frame it passed through alive, and with them their
    locals — ``move_to``'s own ``item`` among them. Holding the refusal for
    the error message would otherwise hold the very reference the retry
    depends on having let go of (issue #130).
    """
    exc.__traceback__ = None
    return exc


def _move_out_of_inbox(
    client: Any, held: list[Any], archive_folder: str, message_id: str
) -> tuple[Any, str]:
    """Move a just-archived mail out of the Inbox; return it and how it went.

    ``held`` is a one-item list carrying the caller's reference to the mail;
    it is emptied here, so the caller keeps none of its own.

    ``MailItem.SaveAs`` can leave the in-memory item flagged as modified, and
    ``Move`` on that same reference is then refused with MAPI_E_OBJECT_CHANGED
    — the mail keeps its place in the Inbox with its files already on disk
    (issue #59). Outlook documents the same refusal as intermittent on IMAP
    stores, cleared by viewing another folder and retrying — the owner's own
    manual remedy (issue #130). Two defences, in order:

    1. move a reference re-acquired from the store by EntryID and StoreID,
       with the original let go before the move;
    2. on a 0x80040109 anyway, let go of every reference to the mail, park
       Outlook's view on another folder when it is showing the Inbox (the
       Explorer holds its selected mail too, issue #111), then fetch the mail
       again from the two ids and move it *while the view is still away*,
       retrying after a pause (``MOVE_RETRY_DELAYS_SECONDS``). The view is put
       back afterwards whatever happened.

    No ``Save()`` is ever called on a refused mail: on a refused object it is
    refused too (the batch log from 2026-10-14 on), and a write to a mail
    Outlook will not let go of is not one to force. Any other failure is
    raised untouched: a store that refuses a move for a different reason is a
    different problem, and retrying it would only hide it. Which path finished
    the move is logged and reported as ``move_via``, because those are the
    facts that say whether this fix is holding on a real mailbox. A 0x80040109
    that survives every retry is raised as ``_MoveRefused`` so the message can
    say whether the switch was tried.
    """
    original = held.pop()
    entry_id, store_id = client.item_ids(original)
    try:
        fresh = client.refetch(entry_id, store_id) if entry_id else None
    except Exception as exc:  # a refetch that raises is still just a fallback
        fresh = None
        logger.warning("Re-acquiring %s by EntryID raised %s: %s",
                       message_id, type(exc).__name__, exc)
    if fresh is not None:
        logger.info(
            "Re-acquired %s by EntryID %s; the re-fetched object is the "
            "original: %s.", message_id, _store_note(store_id),
            _yes_no(client.is_same_object(original, fresh)),
        )
        target, via = fresh, MOVE_VIA_REFETCHED
    else:
        logger.info(
            "Could not re-acquire %s from the store; moving the original "
            "reference.", message_id,
        )
        target, via = original, MOVE_VIA_ORIGINAL
    original = fresh = None  # `target` is now the only reference to the mail

    try:
        moved = client.move_to(target, archive_folder)
    except Exception as exc:
        if not is_message_changed_error(exc):
            raise
        refusal = _refusal_without_frames(exc)
    else:
        logger.info("Moved %s to %r via the %s reference.", message_id, archive_folder, via)
        return moved, via

    target = None
    if not entry_id:
        # Nothing to fetch the mail again by; the original was the only way in.
        raise _MoveRefused(refusal, None, entry_id) from refusal
    gc.collect()  # a reference caught in a cycle is only let go here
    logger.info(
        "Outlook refused to move %s (the message has been changed) on the %s "
        "reference. References to it held now: none (all released); fetching "
        "it again %s while Outlook's view is moved off the Inbox.",
        message_id, via, _store_note(store_id),
    )

    view_away: bool | None = None
    with client.view_away_from_inbox() as view_away:
        logger.info(
            "Outlook's view while %s is moved: %s.", message_id,
            {True: "parked on another folder",
             False: "not on the Inbox, left where it was",
             None: "not determined"}[view_away],
        )
        for attempt, delay in enumerate(MOVE_RETRY_DELAYS_SECONDS, start=2):
            if delay:
                time.sleep(delay)
            try:
                target = client.refetch(entry_id, store_id)
            except Exception as exc:
                logger.warning("Re-acquiring %s by EntryID raised %s: %s",
                               message_id, type(exc).__name__, exc)
                target = None
            if target is None:
                logger.info("Attempt %d: %s could not be fetched again.", attempt, message_id)
                continue
            try:
                moved = client.move_to(target, archive_folder)
            except Exception as exc:
                if not is_message_changed_error(exc):
                    raise
                refusal = _refusal_without_frames(exc)
                target = None
                gc.collect()
                logger.info("Attempt %d to move %s was refused again.", attempt, message_id)
                continue
            via = MOVE_VIA_FOLDER_SWITCH if view_away is True else MOVE_VIA_RELEASED_RETRY
            logger.info(
                "Moved %s to %r on attempt %d, %.1f s after the one before, on "
                "a fresh reference with no other held (move_via %s).",
                message_id, archive_folder, attempt, delay, via,
            )
            return moved, via

    raise _MoveRefused(refusal, view_away, entry_id) from refusal


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
        doc = _envelope("revert", cfg, category, mailbox)
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

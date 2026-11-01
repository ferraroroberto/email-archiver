"""
Moving an archived mail out of the Inbox, and explaining a move Outlook refused.

``apply`` writes a mail's files first and moves the mail second, so a refused
move leaves files on disk and the mail in the Inbox. Everything that happens
between "the files are written" and "the mail is filed" lives here:

- **The reference that is moved is not the one that was archived.**
  ``MailItem.SaveAs`` can leave an item flagged as modified, and ``Move`` then
  refuses it with MAPI_E_OBJECT_CHANGED. ``move_out_of_inbox`` moves a
  reference re-acquired by EntryID and StoreID, with the original let go first.
  On that refusal anyway, every reference to the mail is let go, Outlook's view
  is parked off the Inbox when it is showing it (the Explorer holds its
  selected mail too), and the mail is fetched again and moved while the view is
  still away. It is never ``Save()``d. Which path finished the move is logged
  and reported as ``move_via``.
- **A refusal that survives all of that is its own error.** It is reported as
  ``message_changed`` rather than ``move_failed``: usually a mail left open in
  an Outlook window, and the message says so.
"""
from __future__ import annotations

import gc
import hashlib
import logging
import time
from typing import Any

from email_archiver.outlook.mapi import is_message_changed_error

logger = logging.getLogger(__name__)

ERROR_MOVE_FAILED = "move_failed"               # files are on disk, mail is not
# A move refused as MAPI_E_OBJECT_CHANGED after both retries — almost always a
# mail left open in an Outlook window. Its own code because it is the one move
# failure an operator clears in seconds without anything being wrong: close the
# window, apply the same decision again. A caller that cannot tell it from a
# plain `move_failed` has to treat both as needing a human.
ERROR_MESSAGE_CHANGED = "message_changed"

# Which reference finished an `apply` move, reported as the result's
# `move_via`. Genuinely different stories about the same mail, and the one
# place that says which defence is earning its keep on this mailbox — see
# `move_out_of_inbox`.
MOVE_VIA_REFETCHED = "refetched"     # a reference read back out of the store
MOVE_VIA_ORIGINAL = "original"       # the re-acquire failed; the original moved
# Refused with 0x80040109, then moved on a reference fetched again once every
# other one was let go; Outlook's view was not on the Inbox, so it was left
# alone. A refused mail is never saved.
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


class MoveRefused(Exception):
    """A move still refused as a changed message after every retry.

    Carries the refusal itself (``cause``), what the folder switch did, and the
    mail's EntryID as a plain string (``entry_id``, ``""`` when unreadable) —
    the caller let go of the mail itself before the move, so the window check
    behind ``move_failure_error`` asks by id. A switch that happened and did
    not help is the only evidence a restart is what is left.
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


def _changed_message_cause(open_in_window: bool | None, folder_switched: bool | None) -> str:
    """The operator-facing cause and remedy sentences of a ``message_changed`` error.

    Only a folder switch that actually happened and did not help is evidence
    the hold outlasts one. One that was not applicable or could not run says
    nothing, so it points the user at the same switch by hand, not a restart.
    """
    tried_switch = (
        "Outlook's view was moved to another folder automatically and the move "
        "retried there, and that did not free it. "
    )
    by_hand = (
        "Select another folder in Outlook, then this message again, and apply "
        "this same decision again; restart Outlook only if that does not clear it."
    )
    if open_in_window is True:
        return (
            "This mail is open in an Outlook window, which holds it and "
            "refuses every write to it. Close that window and apply this same "
            "decision again."
        )
    if open_in_window is False and folder_switched is True:
        return (
            "This mail is not open in any Outlook window, so something else "
            f"in the running Outlook is holding it. {tried_switch}"
            "That hold outlasts re-applies, so applying again before a restart "
            "is not expected to help. Restart Outlook, then apply this same "
            "decision again. Deleting the mail from the Inbox by hand instead "
            "leaves the files already written as its only copy, with nothing in "
            "Outlook pointing at them."
        )
    if open_in_window is False:
        return (
            "This mail is not open in any Outlook window, but Outlook may "
            f"still be holding it as the message selected in its folder view. {by_hand}"
        )
    if folder_switched is True:
        return (
            "Whether this mail is open in an Outlook window could not be "
            f"determined. {tried_switch}A mail left open is the usual cause, so "
            "close any window showing it and apply this same decision again; "
            "restart Outlook if that does not clear it."
        )
    return (
        "Whether this mail is open in an Outlook window could not be "
        "determined. A mail left open is the usual cause, so close any "
        "window showing it and apply this same decision again. Failing "
        "that, Outlook may be holding it as the message selected in its "
        "folder view: select another folder in Outlook, then this message "
        "again, and apply again; restart Outlook if that does not clear it."
    )


def move_failure_error(client: Any, exc: Exception, message_id: str) -> dict[str, str]:
    """The ``error`` a failed ``apply`` move reports: its code and its message.

    Anything other than MAPI_E_OBJECT_CHANGED after both defences is a plain
    ``move_failed`` — a store that refused the write for its own reasons is a
    different problem, and dressing it up as this one would hide it.

    A move refused as *the message has been changed* on a re-acquired reference
    *and* on fresh ones fetched after every reference was let go is its own,
    recoverable condition: the files are on disk, the mail is still in the
    Inbox, and something in the running Outlook is holding the item. Usually
    that is an open mail window — an inspector holds its item for the life of
    the window, so both defences fail by construction. Otherwise the cause is
    unresolved: it is state held on the item in the running Outlook, not a
    race, so a bounded retry-with-delay does not clear it. The message says
    outright that re-applying before a restart is not expected to help, rather
    than inviting a caller to burn retries on it; restarting stays the user's
    call, never this tool's.

    So the message leads with the cheap remedy and says which of the two it is
    on the evidence rather than on the guess: Outlook is asked whether this
    mail is actually open. A check that could not be completed is reported as
    *not determined* and never as "no window", because pointing an operator at
    a restart they do not need is the failure this avoids. The Explorer
    showing the mail as its selected item is a third holder, freed by selecting
    another folder; ``move_out_of_inbox`` does that switch itself, and a
    refusal it wraps in ``MoveRefused`` says whether it ran.
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
    cause = _changed_message_cause(open_in_window, folder_switched)

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
    depends on having let go of.
    """
    exc.__traceback__ = None
    return exc


def _retry_with_view_away(
    client: Any,
    refusal: Exception,
    via: str,
    entry_id: str,
    store_id: str,
    archive_folder: str,
    message_id: str,
) -> tuple[Any, str]:
    """Retry a refused move on fresh references while Outlook's view is off the Inbox.

    Raises:
        MoveRefused: the refusal outlasted every retry.
    """
    if not entry_id:
        # Nothing to fetch the mail again by; the original was the only way in.
        raise MoveRefused(refusal, None, entry_id) from refusal
    gc.collect()  # a reference caught in a cycle is only let go here
    logger.info(
        "Outlook refused to move %s (the message has been changed) on the %s "
        "reference. References to it held now: none (all released); fetching "
        "it again %s while Outlook's view is moved off the Inbox.",
        message_id, via, _store_note(store_id),
    )

    target = None
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

    raise MoveRefused(refusal, view_away, entry_id) from refusal


def move_out_of_inbox(
    client: Any, held: list[Any], archive_folder: str, message_id: str
) -> tuple[Any, str]:
    """Move a just-archived mail out of the Inbox; return it and how it went.

    ``held`` is a one-item list carrying the caller's reference to the mail;
    it is emptied here, so the caller keeps none of its own.

    ``MailItem.SaveAs`` can leave the in-memory item flagged as modified, and
    ``Move`` on that same reference is then refused with MAPI_E_OBJECT_CHANGED
    — the mail keeps its place in the Inbox with its files already on disk.
    Outlook documents the same refusal as intermittent on IMAP stores, cleared
    by viewing another folder and retrying. Two defences, in order:

    1. move a reference re-acquired from the store by EntryID and StoreID,
       with the original let go before the move;
    2. on a 0x80040109 anyway, let go of every reference to the mail, park
       Outlook's view on another folder when it is showing the Inbox (the
       Explorer holds its selected mail too), then fetch the mail again from
       the two ids and move it *while the view is still away*, retrying after
       a pause (``MOVE_RETRY_DELAYS_SECONDS``). The view is put back afterwards
       whatever happened.

    No ``Save()`` is ever called on a refused mail: on a refused object it is
    refused too, and a write to a mail Outlook will not let go of is not one
    to force. Any other failure is raised untouched: a store that refuses a
    move for a different reason is a different problem, and retrying it would
    only hide it. Which path finished the move is logged and reported as
    ``move_via``, because those are the facts that say whether this fix is
    holding on a real mailbox. A 0x80040109 that survives every retry is
    raised as ``MoveRefused`` so the message can say whether the switch was
    tried.
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
    return _retry_with_view_away(
        client, refusal, via, entry_id, store_id, archive_folder, message_id,
    )

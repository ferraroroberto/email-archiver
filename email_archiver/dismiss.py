"""
Move Inbox mail into the mailbox's archive folder without keeping a copy.

This is what ``main_batch.py dismiss`` runs. ``apply`` writes a ``.msg`` first
and only then moves the mail; ``dismiss`` is for the mail the caller does not
want kept (a notification it has read), so it takes the mail out of the Inbox
and leaves nothing on disk.

Design decisions:

- **Nothing is written to the archive.** No ``.msg``, no index row, no
  archive-roots check: there is no destination path to guard. The mail's only
  new home is the mailbox's own archive folder in Outlook.
- **Inbox only.** The lookup is the Inbox's, so a mail already filed, sent or
  anywhere else is ``not_in_inbox`` and never touched.
- **Tagged apart from filed mail.** The category differs from ``apply``'s, so a
  mail in the archive folder that has no file behind it is recognisable.
- **Each failure is its own state.** A lookup that raised is not "not in the
  Inbox", a refused move is not a failed tag, and a tag that failed after a
  move leaves ``moved`` true.

Like ``batch.py`` it is pure orchestration over the Outlook client, so a fake
client drives it in a unit test.
"""
from __future__ import annotations

import logging
from typing import Any

from email_archiver.batch_common import ERROR_CATEGORY_FAILED, ERROR_NOT_IN_INBOX, envelope
from email_archiver.batch_move import move_failure_error, move_out_of_inbox
from email_archiver.config import (
    Mailbox,
    get_mailbox_archive_folder,
    get_outlook_dismiss_category,
)
from email_archiver.text import normalize_message_id

logger = logging.getLogger(__name__)

VERB = "dismiss"

# The Inbox lookup itself raised, so whether the mail is there is not known.
ERROR_LOOKUP_FAILED = "lookup_failed"


def _blank_result(message_id: str, folder: str) -> dict[str, Any]:
    return {
        "message_id": message_id,
        "folder": folder,
        "ok": False,
        "moved": False,
        "categorized": False,
        # Which reference the move went through, "" when it never happened.
        "move_via": "",
        "entry_id": "",
        "error": None,
    }


def _dismiss_one(
    client: Any, message_id: str, archive_folder: str, category: str
) -> dict[str, Any]:
    """One mail, start to finish. Never raises: every failure is a result."""
    result = _blank_result(message_id, archive_folder)

    try:
        item = client.find_by_message_id(message_id, None)
    except Exception as exc:  # a COM failure on one lookup, not the run
        logger.exception("Looking up %s in the Inbox failed", message_id)
        result["error"] = {
            "code": ERROR_LOOKUP_FAILED,
            "message": f"the Inbox lookup failed: {type(exc).__name__}: {exc}",
        }
        return result
    if item is None:
        result["error"] = {
            "code": ERROR_NOT_IN_INBOX,
            "message": "no mail with this Message-ID is in the Inbox",
        }
        return result

    try:
        # The move takes over the only reference: nothing in this frame may
        # keep the mail alive while Outlook is asked to let go of it.
        held, item = [item], None
        moved, via = move_out_of_inbox(client, held, archive_folder, message_id)
        result["moved"] = True
        result["move_via"] = via
        result["entry_id"] = client.entry_id(moved)
    except Exception as exc:
        logger.exception("Moving %s to %r failed", message_id, archive_folder)
        result["error"] = move_failure_error(client, exc, message_id)
        return result

    # Tagged in its own step: a tag that would not stick leaves the mail moved,
    # which the caller recovers from differently than a move that did not happen.
    try:
        client.set_category(moved, category)
        result["categorized"] = True
    except Exception as exc:
        logger.exception("Tagging %s with %r failed", message_id, category)
        result["error"] = {
            "code": ERROR_CATEGORY_FAILED,
            "message": f"the mail was moved, but not tagged: {type(exc).__name__}: {exc}",
        }
        return result

    result["ok"] = True
    return result


def dismiss(
    client: Any,
    cfg: dict[str, Any],
    message_ids: list[str],
    *,
    category: str | None = None,
    mailbox: Mailbox | None = None,
) -> dict[str, Any]:
    """Move each Inbox mail into ``mailbox``'s archive folder and tag it.

    ``category`` overrides ``outlook.dismiss_category`` for this run; the
    document's ``category`` reports the one used. A Message-ID listed twice is
    dismissed once — a second lookup would find the mail gone and report a
    ``not_in_inbox`` for a mail this run did move. Whatever happens to one mail,
    the next is still attempted.
    """
    archive_folder = get_mailbox_archive_folder(cfg, mailbox)
    category = category or get_outlook_dismiss_category(cfg)

    wanted = list(dict.fromkeys(normalize_message_id(raw) for raw in message_ids))
    results = [_dismiss_one(client, mid, archive_folder, category) for mid in wanted]

    dismissed = sum(1 for r in results if r["ok"])
    doc = envelope(VERB, cfg, category, mailbox)
    doc["counts"] = {
        "requested": len(results),
        "dismissed": dismissed,
        "failed": len(results) - dismissed,
    }
    doc["results"] = results
    logger.info("Dismiss: %d of %d mail(s) moved to %r.", dismissed, len(results), archive_folder)
    return doc

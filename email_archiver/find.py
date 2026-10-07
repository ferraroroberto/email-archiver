"""
Read-only live search over one mailbox's Outlook folders (issue #110).

This is what ``main_batch.py find`` runs. The index (``data/emails.db``) only
knows mail already archived to disk, and ``plan --search`` reads only the Inbox
or Sent; a mailbox's history lives in its Outlook folders. ``find`` searches
them live.

Like ``batch.py`` it is pure orchestration over
:class:`~email_archiver.outlook.client.OutlookClient` — no COM here — so a fake
client drives it in a unit test.

Design decisions:

- **Read-only.** Nothing is moved, saved, deleted, drafted or sent; a missing
  archive folder is reported ``absent``, never created.
- **Every word must match**, in the subject, sender, recipients or body: the
  same semantics as the index search. Narrowed server-side with
  ``Items.Restrict`` (synchronous; no ``AdvancedSearch`` callback).
- **Coverage is stated, not implied.** With Gmail over IMAP and All Mail not
  synced, mail archived from Gmail web or phone without the Archive label is
  invisible to Outlook; the document says ``coverage: "outlook_folders"``.
"""
from __future__ import annotations

import logging
from datetime import datetime
from typing import Any

from email_archiver.batch import SCHEMA_VERSION, mail_matches, now_iso
from email_archiver.config import Mailbox, get_mailbox_archive_folder
from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRepository
from email_archiver.outlook.client import FIND_ARCHIVE, FIND_INBOX, FIND_SENT
from email_archiver.outlook.mapi import words_filter

logger = logging.getLogger(__name__)

VERB = "find"
FOLDERS = (FIND_INBOX, FIND_ARCHIVE, FIND_SENT)
DEFAULT_LIMIT = 25
COVERAGE = "outlook_folders"
COVERAGE_NOTE = (
    "Searches this mailbox's Outlook folders only. Mail archived from Gmail web "
    "or phone without the Archive label lives only in All Mail, which is not "
    "synced to Outlook, so it is not found here."
)


def _iso_or_empty(value: datetime | None) -> str:
    return value.isoformat() if value is not None else ""


def find(
    client: Any,
    cfg: dict[str, Any],
    mailbox: Mailbox | None,
    words: list[str],
    *,
    since: datetime | None = None,
    folders: tuple[str, ...] = FOLDERS,
    limit: int = DEFAULT_LIMIT,
) -> dict[str, Any]:
    """Search ``folders`` of the mailbox ``client`` is bound to; the ``find``
    document, newest ``limit`` hits first across all folders."""
    preview_len = int((cfg.get("scanning") or {}).get("body_preview_length", 500))
    archive_folder = get_mailbox_archive_folder(cfg, mailbox)

    searched: list[dict[str, Any]] = []
    found: list[tuple[str, Any]] = []
    for folder in folders:
        sent = folder == FIND_SENT
        result = client.search_folder(
            folder, archive_folder, words_filter(words, since, sent=sent),
            lambda mail, sent=sent: mail_matches(mail, words, since, sent),
            preview_len=preview_len, limit=limit,
        )
        searched.append({
            "folder": result.folder, "name": result.name,
            "via": result.via, "matched": result.matched,
        })
        found.extend((result.name, mail) for mail in result.mails)

    found.sort(key=lambda pair: pair[1].date_sent or datetime.min, reverse=True)
    kept = found[:limit]

    conn = init_db(cfg["database"]["path"])
    repo = EmailRepository(conn)
    try:
        hits = [
            {
                "message_id": mail.message_id,
                "subject": mail.subject,
                "sender": mail.sender,
                "recipients": mail.recipients,
                "date": _iso_or_empty(mail.date_sent),
                "folder": name,
                "entry_id": mail.entry_id,
                "body_preview": mail.body_preview,
                "already_archived": (
                    repo.find_path_by_message_id(mail.message_id) if mail.message_id else None
                ),
            }
            for name, mail in kept
        ]
    finally:
        conn.close()

    matched = sum(f["matched"] for f in searched)
    logger.info(
        "Find: %d match(es) across %s, %d returned.",
        matched, ", ".join(f["folder"] for f in searched), len(hits),
    )
    return {
        "verb": VERB,
        "schema_version": SCHEMA_VERSION,
        "generated_at": now_iso(),
        "query": {
            "words": list(words),
            "since": since.isoformat(timespec="minutes") if since else None,
            "folders": list(folders),
            "limit": limit,
        },
        "coverage": COVERAGE,
        "coverage_note": COVERAGE_NOTE,
        "folders": searched,
        "counts": {"matched": matched, "returned": len(hits)},
        "truncated": matched > len(hits),
        "hits": hits,
    }

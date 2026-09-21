"""
Guarded send: read a draft back, and send it only while it is still what was approved.

This is what ``main_batch.py read`` and ``main_batch.py send`` run. Like
``draft.py`` it is pure orchestration — the COM reads live in
:class:`~email_archiver.outlook.client.OutlookClient` and the one ``Send`` call
in :mod:`email_archiver.outlook.sending` — so a fake client drives it in tests.

Design decisions:

- **Approval is a fingerprint of the stored item, not of what a caller wrote.**
  ``read`` reports a sha256 per part (``to``, ``cc``, ``bcc``, ``subject``,
  ``body``, ``attachments``) and one ``hash`` over them. ``body`` is the whole
  ``HTMLBody``, signature included, and ``attachments`` are their bytes, so an
  edit anywhere in the item — made in an Outlook window, by hand — changes it.
- **The comparison runs on the live item immediately before ``Send()``**, on
  the same reference that is then sent. Nothing is compared to a copy.
- **A refusal names what differs** (``differs``) when the caller passed the
  per-part hashes, and sends nothing. There is no partial or forced send.
- **One item, by EntryID.** No lookup by subject or search, and no recipient
  is resolved or added; a recipient with no readable address refuses the send
  rather than being skipped.
- **``draft`` keeps no send path.** Sending is this separate verb, so a caller
  has to choose it deliberately.
"""
from __future__ import annotations

import hashlib
import json
import logging
from collections.abc import Mapping
from typing import Any

from email_archiver.batch import SCHEMA_VERSION, now_iso
from email_archiver.draft import html_to_text
from email_archiver.outlook.drafts import DraftSnapshot
from email_archiver.outlook.sending import send_if_approved

logger = logging.getLogger(__name__)

READ_VERB = "read"
SEND_VERB = "send"

PARTS = ("to", "cc", "bcc", "subject", "body", "attachments")

APPROVAL_MISMATCH = "approval_mismatch"
RECIPIENT_UNREADABLE = "recipient_unreadable"


class SendRefused(Exception):
    """The draft was not sent; ``code`` says why, ``differs`` names the parts."""

    def __init__(self, code: str, message: str, differs: list[str] | None = None) -> None:
        super().__init__(message)
        self.code = code
        self.differs = differs or []


def _addresses(values: list[str]) -> list[str]:
    """Who receives it: order, case and repeats do not change that."""
    return sorted({value.strip().casefold() for value in values if value.strip()})


def _digest(value: Any) -> str:
    canonical = json.dumps(value, sort_keys=True, ensure_ascii=False, separators=(",", ":"))
    return hashlib.sha256(canonical.encode("utf-8")).hexdigest()


def part_values(snapshot: DraftSnapshot) -> dict[str, Any]:
    """Each fingerprinted part in the form it is hashed in."""
    return {
        "to": _addresses(snapshot.to),
        "cc": _addresses(snapshot.cc),
        "bcc": _addresses(snapshot.bcc),
        "subject": snapshot.subject,
        "body": snapshot.html_body,
        "attachments": sorted([a.name, a.size_bytes, a.sha256] for a in snapshot.attachments),
    }


def fingerprint(snapshot: DraftSnapshot) -> dict[str, Any]:
    """``{"hash", "parts"}``: a sha256 per part, and one over those."""
    parts = {name: _digest(value) for name, value in part_values(snapshot).items()}
    return {"hash": _digest(parts), "parts": parts}


def read_document(snapshot: DraftSnapshot) -> dict[str, Any]:
    """The ``read`` document: the draft as stored, and its fingerprint."""
    region = snapshot.body_region_html
    return {
        "verb": READ_VERB,
        "schema_version": SCHEMA_VERSION,
        "generated_at": now_iso(),
        "entry_id": snapshot.entry_id,
        "subject": snapshot.subject,
        "to": list(snapshot.to),
        "cc": list(snapshot.cc),
        "bcc": list(snapshot.bcc),
        "unreadable_recipients": snapshot.unreadable_recipients,
        "body": {
            "html": snapshot.html_body,
            "region_html": region,
            "region_text": None if region is None else html_to_text(region),
        },
        "attachments": [
            {"name": a.name, "size_bytes": a.size_bytes, "sha256": a.sha256}
            for a in snapshot.attachments
        ],
        "fingerprint": fingerprint(snapshot),
    }


def check_approval(
    snapshot: DraftSnapshot,
    expect_hash: str,
    expect_parts: Mapping[str, str],
    expect_to: list[str] | None,
) -> None:
    """Raise :class:`SendRefused` unless ``snapshot`` is the approved item.

    ``expect_parts`` only names the parts in a refusal; the whole ``hash`` is
    what decides. ``expect_to``, when given, must be exactly the To line.
    """
    if snapshot.unreadable_recipients:
        raise SendRefused(
            RECIPIENT_UNREADABLE,
            f"{snapshot.unreadable_recipients} recipient(s) on the draft have no readable "
            "address; refusing to send to someone the approval cannot name",
        )
    live = fingerprint(snapshot)
    differs = [name for name in PARTS if name in expect_parts and expect_parts[name] != live["parts"][name]]
    if expect_to is not None and _addresses(expect_to) != _addresses(snapshot.to) and "to" not in differs:
        differs.append("to")
        differs.sort(key=PARTS.index)
    if live["hash"] == expect_hash and not differs:
        return
    named = ", ".join(differs) if differs else "a part not named (pass --expect-part to name it)"
    raise SendRefused(
        APPROVAL_MISMATCH,
        f"the draft is not what was approved; it differs in: {named}. Nothing was sent.",
        differs,
    )


def send(
    client: Any,
    entry_id: str,
    expect_hash: str,
    expect_parts: Mapping[str, str],
    expect_to: list[str] | None,
) -> dict[str, Any]:
    """Send the draft ``entry_id`` if it still matches; the ``send`` document.

    Raises:
        SendRefused: the item differs from the approval; nothing was sent.
        DraftUpdateError: the item is gone, sent, or not in Drafts.
    """
    snapshot = send_if_approved(
        client, entry_id, lambda live: check_approval(live, expect_hash, expect_parts, expect_to),
    )
    logger.info("Sent draft %s: its fingerprint matched the approval.", entry_id)
    return {
        "verb": SEND_VERB,
        "schema_version": SCHEMA_VERSION,
        "generated_at": now_iso(),
        "entry_id": entry_id,
        "subject": snapshot.subject,
        "to": list(snapshot.to),
        "cc": list(snapshot.cc),
        "bcc": list(snapshot.bcc),
        "attachments": [a.name for a in snapshot.attachments],
        "fingerprint": expect_hash,
        "sent": True,
        "sent_at": now_iso(),
    }

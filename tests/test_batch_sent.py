"""Tests for filing from Sent Items (issue #103).

life-os files the self-BCC copy of a sent draft from the Inbox; when that copy
never arrives it asks for the sent mail itself. These tests pin what it relies
on: `plan --folder sent` reads Sent Items only, `apply` with `source: "sent"`
archives and tags the mail without moving it, and an unknown folder or source
is refused rather than read as the Inbox. The Inbox contract is pinned by
``test_batch.py`` and ``test_batch_targeted.py``.

Everything here is synthetic: no real folder name, address or subject.
"""
from __future__ import annotations

import argparse
from datetime import datetime, timedelta
from pathlib import Path

import pytest

import main_batch
from email_archiver import batch
from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRecord, EmailRepository
from tests.test_batch import (  # noqa: F401 - fixtures are used by name
    _mail,
    archive_root,
    cfg,
)
from tests.test_batch_targeted import TargetedFakeClient

SENT = "__sent__"
REF = "ref-0001"


class SentFakeClient(TargetedFakeClient):
    """The targeted fake plus a Sent Items folder and its surface."""

    def __init__(self, inbox=(), sent=()) -> None:
        super().__init__(list(inbox))
        self.folders[SENT] = list(sent)

    def iter_sent(self, preview_len: int = 500):
        self.calls.append("iter_sent")
        for item in list(self.folders[SENT]):
            yield self.read_mail(item, preview_len)

    def iter_sent_since(self, since: datetime, preview_len: int = 500):
        # Deliberately the whole folder, the widest a minute-precise Restrict
        # could be, so only the exact check in `batch.plan` can pass the test.
        self.calls.append("iter_sent_since")
        yield from self.iter_sent(preview_len)

    def find_sent_by_message_id(self, message_id: str):
        self.calls.append("find_sent_by_message_id")
        return self.find_by_message_id(message_id, SENT)


def _sent_mail(message_id: str, subject: str, **kw):
    item = _mail(message_id, subject, **kw)
    item.headers = f"X-Archive-Ref: {REF}\r\nSubject: {subject}\r\n"
    return item


def test_plan_folder_sent_reads_sent_items_only(cfg):
    client = SentFakeClient(
        inbox=[_mail("inbox@example.invalid", "Inbox mail")],
        sent=[_sent_mail("sent@example.invalid", "Sent mail")],
    )

    doc = batch.plan(client, cfg, candidates=0, filters=batch.PlanFilters(
        ref=REF, since=datetime(2026, 1, 1), folder=batch.SOURCE_SENT,
    ))

    assert "iter_inbox_received_since" not in client.calls
    assert client.calls[0] == "iter_sent_since"
    assert [m["message_id"] for m in doc["mails"]] == ["sent@example.invalid"]
    assert doc["mails"][0]["in_inbox"] is False
    assert doc["filters"]["folder"] == "sent"
    assert doc["counts"] == {"sent": 1, "planned": 1, "already_archived": 0, "skipped": 0}


def test_plan_without_folder_is_the_inbox_plan_it_always_was(cfg):
    client = SentFakeClient(
        inbox=[_mail("inbox@example.invalid", "Inbox mail")],
        sent=[_sent_mail("sent@example.invalid", "Sent mail")],
    )

    doc = batch.plan(client, cfg, candidates=0, filters=batch.PlanFilters(ref=REF))

    assert doc["mails"] == []
    assert "folder" not in doc["filters"]
    assert "inbox" in doc["counts"] and "sent" not in doc["counts"]
    assert "iter_sent" not in client.calls


def test_plan_sent_since_compares_the_sent_time(cfg):
    boundary = datetime(2026, 9, 30, 11, 13)
    old = _sent_mail("old@example.invalid", "Before", sent=boundary - timedelta(minutes=1))
    new = _sent_mail("new@example.invalid", "After", sent=boundary + timedelta(minutes=4))
    # A received time that would pass or fail the check must not matter.
    old.received = boundary + timedelta(days=1)
    new.received = boundary - timedelta(days=1)
    client = SentFakeClient(sent=[old, new])

    doc = batch.plan(client, cfg, candidates=0, filters=batch.PlanFilters(
        since=boundary, folder=batch.SOURCE_SENT,
    ))

    assert [m["message_id"] for m in doc["mails"]] == ["new@example.invalid"]


def test_plan_sent_message_id_not_found_is_skipped_as_not_in_sent(cfg):
    client = SentFakeClient(inbox=[_mail("inbox@example.invalid", "Inbox mail")])

    doc = batch.plan(client, cfg, candidates=0, filters=batch.PlanFilters(
        message_ids=("inbox@example.invalid",), folder=batch.SOURCE_SENT,
    ))

    assert doc["mails"] == []
    assert doc["skipped"] == [{
        "entry_id": "", "subject": "", "reason": batch.SKIP_NOT_IN_SENT,
        "message_id": "inbox@example.invalid",
    }]


def test_apply_from_sent_archives_and_tags_but_keeps_the_mail_in_sent(cfg, archive_root):
    dest = archive_root / "Project Alpha"
    item = _sent_mail("sent@example.invalid", "Sent mail")
    client = SentFakeClient(sent=[item])

    doc = batch.apply(client, cfg, [{
        "message_id": "sent@example.invalid", "folder_path": str(dest),
        "date_prefix": "auto", "source": "sent",
    }])

    result = doc["results"][0]
    assert (result["ok"], result["error"]) == (True, None)
    assert result["moved"] is False
    assert result["move_via"] == batch.MOVE_VIA_KEPT_IN_SENT
    assert result["categorized"] is True
    assert result["entry_id"] == item.EntryID
    assert result["files"] and all(Path(p).exists() for p in result["files"])
    assert client.folders[SENT] == [item]
    assert client.moves == []
    assert item.Categories == "Filed by batch"


def test_apply_from_sent_after_a_scan_reuses_the_written_file(cfg, archive_root):
    """The mail stays in Sent Items, so unlike an Inbox mail it can be decided
    again; once the scan has indexed the file, that writes nothing."""
    dest = archive_root / "Project Alpha"
    client = SentFakeClient(sent=[_sent_mail("sent@example.invalid", "Sent mail")])
    decision = {"message_id": "sent@example.invalid", "folder_path": str(dest),
                "date_prefix": "auto", "source": "sent"}

    first = batch.apply(client, cfg, [decision])["results"][0]
    conn = init_db(cfg["database"]["path"])
    EmailRepository(conn).upsert_email(EmailRecord(
        file_path=first["files"][0], folder_path=str(dest),
        filename=Path(first["files"][0]).name, subject="Sent mail",
        sender="sender@example.invalid", recipients="me@example.invalid",
        date_sent="2026-03-14T09:30:00", body_preview="", file_mtime=1.0,
        message_id="sent@example.invalid",
    ))
    conn.commit()
    conn.close()
    second = batch.apply(client, cfg, [decision])["results"][0]

    assert second["ok"] is True and second["reused"] is True
    assert second["files"] == first["files"]
    assert len(list(dest.glob("*.msg"))) == 1


def test_apply_from_sent_does_not_file_an_inbox_only_mail(cfg, archive_root):
    dest = archive_root / "Project Alpha"
    client = SentFakeClient(inbox=[_mail("inbox@example.invalid", "Inbox mail")])

    result = batch.apply(client, cfg, [{
        "message_id": "inbox@example.invalid", "folder_path": str(dest),
        "date_prefix": "auto", "source": "sent",
    }])["results"][0]

    assert result["ok"] is False
    assert result["error"]["code"] == batch.ERROR_NOT_IN_SENT
    assert result["files"] == []
    assert not dest.exists() or not list(dest.glob("*.msg"))


def test_apply_refuses_an_unknown_source_instead_of_reading_it_as_the_inbox(cfg, archive_root):
    dest = archive_root / "Project Alpha"
    client = SentFakeClient(inbox=[_mail("a@example.invalid", "Inbox mail")])

    result = batch.apply(client, cfg, [{
        "message_id": "a@example.invalid", "folder_path": str(dest),
        "date_prefix": "auto", "source": "drafts",
    }])["results"][0]

    assert result["ok"] is False
    assert result["error"]["code"] == batch.ERROR_BAD_DECISION
    assert client.folders[None] and client.moves == []


def _args(**kw) -> argparse.Namespace:
    base = dict(message_ids=[], since=None, search=[], ref=None, folder="inbox")
    return argparse.Namespace(**{**base, **kw})


def test_main_batch_accepts_sent_and_refuses_any_other_folder():
    assert main_batch._plan_filters(_args(folder=" Sent ")).folder == batch.SOURCE_SENT
    assert main_batch._plan_filters(_args()).folder == batch.SOURCE_INBOX
    with pytest.raises(ValueError, match="--folder"):
        main_batch._plan_filters(_args(folder="drafts"))

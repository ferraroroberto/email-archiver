"""A move that landed is reported as moved, whatever happens to the id read after it.

Synthetic throughout, against the fake Outlook client; no real mailbox is
touched.
"""
from __future__ import annotations

from email_archiver import batch, dismiss
from tests.batch.batch_support import (  # noqa: F401 - fixtures are used by name
    FakeOutlookClient,
    _mail,
    _no_retry_pauses,
    archive_root,
    cfg,
)
from tests.batch.test_batch_sent import SentFakeClient, _sent_mail


class _EntryIdUnreadable(FakeOutlookClient):
    """The move lands, then Outlook will not hand over the mail's new EntryID."""

    def entry_id(self, item) -> str:
        raise RuntimeError("the EntryID could not be read")


class _SentEntryIdUnreadable(SentFakeClient):
    def entry_id(self, item) -> str:
        raise RuntimeError("the EntryID could not be read")


def _decision(archive_root, **extra):
    return {
        "message_id": "a@example.invalid", "folder_path": str(archive_root / "Project Alpha"),
        "date_prefix": "auto", **extra,
    }


def test_apply_reports_a_landed_move_as_moved_when_the_entry_id_read_raises(cfg, archive_root):
    client = _EntryIdUnreadable([_mail("a@example.invalid", "Synthetic notice")])

    [result] = batch.apply(client, cfg, [_decision(archive_root)])["results"]

    assert client.moves == [("a@example.invalid", "Archive")]
    assert result["error"] is None
    assert (result["ok"], result["moved"], result["categorized"]) == (True, True, True)
    assert result["entry_id"] == ""
    assert result["move_via"]
    assert result["files"]


def test_dismiss_reports_a_landed_move_as_moved_when_the_entry_id_read_raises(cfg):
    client = _EntryIdUnreadable([_mail("a@example.invalid", "Synthetic notice")])

    [result] = dismiss.dismiss(client, cfg, ["a@example.invalid"])["results"]

    assert client.moves == [("a@example.invalid", "Archive")]
    assert result["error"] is None
    assert (result["ok"], result["moved"], result["categorized"]) == (True, True, True)
    assert result["entry_id"] == ""


def test_a_tag_failure_after_an_unreadable_entry_id_is_still_a_tag_failure(cfg, archive_root):
    client = _EntryIdUnreadable([_mail("a@example.invalid", "Synthetic notice")])
    client.category_should_fail = True

    [result] = batch.apply(client, cfg, [_decision(archive_root)])["results"]

    assert result["error"]["code"] == batch.ERROR_CATEGORY_FAILED
    assert (result["ok"], result["moved"], result["categorized"]) == (False, True, False)
    assert result["entry_id"] == ""


def test_apply_from_sent_is_not_a_move_failure_when_the_entry_id_read_raises(cfg, archive_root):
    item = _sent_mail("a@example.invalid", "Sent mail")
    client = _SentEntryIdUnreadable(sent=[item])

    [result] = batch.apply(client, cfg, [_decision(archive_root, source="sent")])["results"]

    assert result["error"] is None
    assert (result["ok"], result["moved"], result["categorized"]) == (True, False, True)
    assert result["move_via"] == batch.MOVE_VIA_KEPT_IN_SENT
    assert result["entry_id"] == ""

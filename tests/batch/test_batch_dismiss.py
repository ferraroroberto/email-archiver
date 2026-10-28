"""``dismiss``: Inbox mail moved to the archive folder with no copy kept.

Synthetic throughout, against the fake Outlook client; no real mailbox is
touched.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from email_archiver import dismiss
from email_archiver.config import DEFAULT_ARCHIVED_CATEGORY, DEFAULT_DISMISSED_CATEGORY
from tests.batch.batch_support import FakeOutlookClient, _mail
from tests.batch.batch_support import (  # noqa: F401 - fixtures are used by name
    _no_retry_pauses,
    archive_root,
    cfg,
)
from tests.test_mailboxes import batch_process  # noqa: F401 - fixture is used by name

REGISTRY = {
    "schema_version": 1,
    "default": "owner",
    "mailboxes": {
        "owner": {"address": "owner@example.invalid"},
        "second": {
            "address": "second@example.invalid",
            "aliases": ["household"],
            "archive_folder": "Filed",
        },
    },
}


class _ProcessClient(FakeOutlookClient):
    """The fake as ``main_batch`` drives it, which starts Outlook first."""

    def ensure_running(self, timeout: float = 0) -> None:
        return None


def _tree(root: Path) -> list[Path]:
    return sorted(root.rglob("*"))


def test_dismiss_moves_the_mail_to_the_archive_folder_and_writes_nothing(cfg, archive_root):
    mail = _mail("m-1", "Synthetic notice")
    client = FakeOutlookClient([mail, _mail("m-2", "Stays put")])

    doc = dismiss.dismiss(client, cfg, ["m-1"])

    [result] = doc["results"]
    assert result["ok"] and result["moved"] and result["categorized"]
    assert result["folder"] == "Archive" and result["move_via"]
    assert result["error"] is None
    assert client.moves == [("m-1", "Archive")]
    assert [m.message_id for m in client.folders[None]] == ["m-2"]
    assert doc["counts"] == {"requested": 1, "dismissed": 1, "failed": 0}
    assert doc["verb"] == "dismiss" and doc["archive_folder"] == "Archive"
    # Kept nothing: no .msg anywhere, no SaveAs, and the index was never opened.
    assert mail.saved_as == [] and _tree(archive_root) == []
    assert not Path(cfg["database"]["path"]).exists()


def test_dismiss_tags_with_its_own_category_not_the_filing_one(cfg):
    mail = _mail("m-1", "Synthetic notice")

    doc = dismiss.dismiss(FakeOutlookClient([mail]), cfg, ["m-1"])

    assert doc["category"] == DEFAULT_DISMISSED_CATEGORY != DEFAULT_ARCHIVED_CATEGORY
    assert mail.Categories == DEFAULT_DISMISSED_CATEGORY


def test_the_configured_and_the_requested_category_win_in_that_order(cfg):
    cfg["outlook"] = {"dismiss_category": "From config"}
    first, second = _mail("m-1", "One"), _mail("m-2", "Two")

    from_config = dismiss.dismiss(FakeOutlookClient([first]), cfg, ["m-1"])
    requested = dismiss.dismiss(FakeOutlookClient([second]), cfg, ["m-2"], category="Asked for")

    assert (from_config["category"], first.Categories) == ("From config", "From config")
    assert (requested["category"], second.Categories) == ("Asked for", "Asked for")


def test_a_message_id_not_in_the_inbox_is_its_own_error_and_the_rest_still_move(cfg):
    client = FakeOutlookClient([_mail("m-1", "Synthetic notice")])
    client.folders["Archive"] = [_mail("m-filed", "Already filed")]

    doc = dismiss.dismiss(client, cfg, ["m-filed", "m-nowhere", "m-1"])

    by_id = {r["message_id"]: r for r in doc["results"]}
    assert by_id["m-filed"]["error"]["code"] == "not_in_inbox"
    assert by_id["m-nowhere"]["error"]["code"] == "not_in_inbox"
    assert by_id["m-1"]["ok"]
    assert not by_id["m-filed"]["moved"]
    assert client.moves == [("m-1", "Archive")]
    assert doc["counts"] == {"requested": 3, "dismissed": 1, "failed": 2}


def test_a_failed_move_leaves_the_mail_in_the_inbox_with_move_failed(cfg):
    client = FakeOutlookClient([_mail("m-1", "Synthetic notice")])
    client.move_should_fail = True

    doc = dismiss.dismiss(client, cfg, ["m-1"])

    [result] = doc["results"]
    assert result["error"]["code"] == "move_failed"
    assert not result["ok"] and not result["moved"]
    assert [m.message_id for m in client.folders[None]] == ["m-1"]


def test_a_move_outlook_keeps_refusing_is_message_changed(cfg):
    client = FakeOutlookClient([_mail("m-1", "Synthetic notice")])
    client.move_message_changed_always = True

    [result] = dismiss.dismiss(client, cfg, ["m-1"])["results"]

    assert result["error"]["code"] == "message_changed"
    assert not result["moved"]


def test_a_tag_that_would_not_stick_is_category_failed_on_a_moved_mail(cfg):
    client = FakeOutlookClient([_mail("m-1", "Synthetic notice")])
    client.category_should_fail = True

    [result] = dismiss.dismiss(client, cfg, ["m-1"])["results"]

    assert result["error"]["code"] == "category_failed"
    assert result["moved"] and not result["categorized"] and not result["ok"]
    assert client.moves == [("m-1", "Archive")]


def test_a_lookup_that_raises_is_not_reported_as_not_in_the_inbox(cfg):
    class _Unreachable(FakeOutlookClient):
        def find_by_message_id(self, message_id, folder_name=None):
            raise RuntimeError("Outlook would not answer")

    [result] = dismiss.dismiss(_Unreachable([_mail("m-1", "Synthetic")]), cfg, ["m-1"])["results"]

    assert result["error"]["code"] == "lookup_failed"
    assert not result["moved"]


def test_a_repeated_message_id_is_moved_once_and_brackets_are_ignored(cfg):
    client = FakeOutlookClient([_mail("m-1", "Synthetic notice")])

    doc = dismiss.dismiss(client, cfg, ["m-1", "<m-1>", "m-1"])

    assert [r["message_id"] for r in doc["results"]] == ["m-1"]
    assert doc["counts"]["failed"] == 0
    assert client.moves == [("m-1", "Archive")]


# ------------------------------------------------------ process contract ---

def test_the_verb_exits_0_when_every_mail_moved(batch_process):
    code, doc, built = batch_process(
        ["dismiss", "--message-id", "m-1", "--message-id", "m-2"], registry=REGISTRY,
        client_factory=lambda mb: _ProcessClient([_mail("m-1", "One"), _mail("m-2", "Two")]),
    )

    assert code == 0, doc
    assert doc["mailbox"] == "owner" and doc["counts"]["dismissed"] == 2
    assert [m for m, _ in built[0].moves] == ["m-1", "m-2"]


@pytest.mark.parametrize("state", [{}, {"move_should_fail": True}])
def test_a_mail_that_did_not_move_exits_1_and_still_prints_the_document(batch_process, state):
    def _factory(mb):
        client = _ProcessClient([_mail("m-1", "One")])
        for name, value in state.items():
            setattr(client, name, value)
        return client

    code, doc, _ = batch_process(
        ["dismiss", "--message-id", "m-1" if state else "m-gone"], registry=REGISTRY,
        client_factory=_factory,
    )

    assert code == 1
    assert doc["counts"]["failed"] == 1
    assert doc["results"][0]["error"]["code"] == ("move_failed" if state else "not_in_inbox")


def test_the_verb_acts_on_the_named_mailbox_only(batch_process):
    seen = []

    def _factory(mailbox):
        seen.append(mailbox.alias)
        return _ProcessClient([_mail("m-1", "One")])

    code, doc, built = batch_process(
        ["dismiss", "--mailbox", "HOUSEHOLD", "--message-id", "m-1"],
        registry=REGISTRY, client_factory=_factory,
    )

    assert code == 0, doc
    assert (doc["mailbox"], doc["archive_folder"]) == ("second", "Filed")
    assert built[0].moves == [("m-1", "Filed")]
    assert seen == ["second"]


def test_an_unknown_mailbox_is_refused_before_outlook(batch_process):
    code, doc, built = batch_process(
        ["dismiss", "--mailbox", "nobody", "--message-id", "m-1"], registry=REGISTRY,
        client_factory=lambda mb: pytest.fail("Outlook touched"),
    )

    assert code == 2 and built == []
    assert doc["error"]["code"] == "mailbox_unknown"


@pytest.mark.parametrize("argv", [
    ["dismiss"],
    ["dismiss", "--message-id", "  "],
    ["dismiss", "--message-id", "m-1", "--category", " "],
])
def test_unusable_input_is_refused_before_outlook(batch_process, argv):
    code, doc, built = batch_process(
        argv, registry=REGISTRY, client_factory=lambda mb: pytest.fail("Outlook touched"),
    )

    assert code == 2 and built == []
    assert doc["error"]["code"] == "bad_input"

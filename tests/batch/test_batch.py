"""Tests for headless batch mode (issue #53).

Batch mode is what another local app spawns to file the whole Inbox in one run,
so its JSON documents are an interface, not an implementation detail: a
consumer parses them without a human reading the output. These tests drive
every verb end to end through a fake Outlook client — real archiver, real
suggestion engine, real SQLite index, fake COM — and pin the parts a consumer
depends on:

- the shapes of the three documents, and that a per-mail failure is *reported*
  rather than fatal;
- already-archived detection, which is the only thing stopping a second run
  from filing the same mail twice;
- the per-decision ``date_prefix``, which must beat the global config toggle
  because the caller infers the form the destination folder actually uses;
- ``revert`` refusing to delete anything outside the configured archive roots.

Everything on screen here is synthetic: no real folder name, address or subject.
"""
from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

import pytest

from email_archiver import batch
from email_archiver.archiver import archiver as archiver_module
from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRecord, EmailRepository
from tests.batch.batch_support import (
    FAKE_STORE_ID,
    FakeOutlookClient,
    _FakeAttachment,
    _mail,
    _seed_index,
    _stuck_move,
)
from tests.batch.batch_support import (  # noqa: F401 - fixtures are used by name
    _no_retry_pauses,
    archive_root,
    cfg,
)


# ------------------------------------------------------------------- plan ---

def test_plan_lists_every_inbox_mail_with_ranked_candidates(cfg, archive_root):
    _seed_index(cfg, archive_root, {
        "Project Alpha": ["Project Alpha kickoff", "Project Alpha budget"],
        "Project Beta": ["Project Beta retro"],
    })
    client = FakeOutlookClient([
        _mail("a@example.invalid", "Project Alpha weekly update"),
        _mail("b@example.invalid", "Project Beta retro follow-up"),
    ])

    doc = batch.plan(client, cfg, candidates=5)

    assert doc["verb"] == "plan"
    assert doc["schema_version"] == batch.SCHEMA_VERSION
    assert doc["archive_folder"] == "Archive"
    assert doc["counts"] == {
        "inbox": 2, "planned": 2, "already_archived": 0, "skipped": 0,
    }
    assert [m["message_id"] for m in doc["mails"]] == [
        "a@example.invalid", "b@example.invalid",
    ]

    first = doc["mails"][0]
    assert set(first) == {
        "message_id", "entry_id", "subject", "sender", "recipients", "date_sent",
        "body_preview", "attachment_count", "flag_status", "already_archived",
        "in_inbox", "candidates",
    }
    assert first["already_archived"] is None
    assert first["candidates"], "a seeded index must produce candidates"
    assert set(first["candidates"][0]) == {
        "folder_path", "display_name", "score", "match_count", "sample_subjects",
        "date_prefix",
    }
    assert first["candidates"][0]["folder_path"] == str(archive_root / "Project Alpha")
    # cfg's naming.date_prefix is False (not "auto"), so every candidate gets
    # the fixed config value regardless of what its folder holds.
    assert first["candidates"][0]["date_prefix"] is False


def test_plan_infers_date_prefix_per_candidate_folder_in_auto_mode(cfg, archive_root):
    """naming.date_prefix: auto — each candidate's date_prefix reflects what
    its own folder already holds, not one fixed value for every candidate."""
    cfg["naming"]["date_prefix"] = "auto"
    _seed_index(cfg, archive_root, {
        "Project Alpha": ["Project Alpha kickoff"],
        "Project Beta": ["Project Beta retro"],
    })
    dated_folder = archive_root / "Project Alpha"
    undated_folder = archive_root / "Project Beta"
    dated_folder.mkdir(parents=True, exist_ok=True)
    undated_folder.mkdir(parents=True, exist_ok=True)
    (dated_folder / "2026-01-01 - 001 - a.msg").write_text("x")
    (dated_folder / "2026-01-02 - 002 - b.msg").write_text("x")
    (undated_folder / "001 - a.msg").write_text("x")
    (undated_folder / "002 - b.msg").write_text("x")

    client = FakeOutlookClient([
        _mail("a@example.invalid", "Project Alpha weekly update"),
        _mail("b@example.invalid", "Project Beta retro follow-up"),
    ])
    doc = batch.plan(client, cfg, candidates=5)

    by_folder = {
        c["folder_path"]: c["date_prefix"]
        for m in doc["mails"] for c in m["candidates"]
    }
    assert by_folder[str(dated_folder)] is True
    assert by_folder[str(undated_folder)] is False


def test_plan_lists_a_shared_candidate_folder_only_once_per_run(
    cfg, archive_root, monkeypatch
):
    """A folder that two different mails both suggest must only be listed
    once for the whole plan() run, not once per mail that suggests it — a
    full-Inbox plan has a handful of popular folders and many mails, so an
    uncached per-candidate os.listdir would re-list the same OneDrive-backed
    folder over and over (the caching in _candidate_dict's date_prefix_cache
    exists to prevent exactly that)."""
    cfg["naming"]["date_prefix"] = "auto"
    _seed_index(cfg, archive_root, {"Project Alpha": ["Project Alpha kickoff"]})
    folder = archive_root / "Project Alpha"
    folder.mkdir(parents=True, exist_ok=True)
    (folder / "2026-01-01 - 001 - a.msg").write_text("x")

    client = FakeOutlookClient([
        _mail("a@example.invalid", "Project Alpha weekly update"),
        _mail("b@example.invalid", "Project Alpha budget question"),
    ])

    real_listdir = os.listdir
    calls: list[str] = []

    def _counting_listdir(path):
        calls.append(str(path))
        return real_listdir(path)

    monkeypatch.setattr(archiver_module.os, "listdir", _counting_listdir)

    doc = batch.plan(client, cfg, candidates=1)

    assert [m["candidates"][0]["folder_path"] for m in doc["mails"]] == [
        str(folder), str(folder),
    ]
    assert calls.count(str(folder)) == 1


def test_plan_honours_the_candidate_cap_not_the_config(cfg, archive_root):
    """--candidates is the caller's, and must beat suggestion.max_suggestions."""
    _seed_index(cfg, archive_root, {
        f"Project {n}": [f"Project {n} status report"] for n in range(1, 6)
    })
    client = FakeOutlookClient([_mail("a@example.invalid", "Project status report")])

    assert cfg["suggestion"]["max_suggestions"] == 3
    doc = batch.plan(client, cfg, candidates=1)
    assert len(doc["mails"][0]["candidates"]) == 1
    # And the shared config was not mutated on the way through.
    assert cfg["suggestion"]["max_suggestions"] == 3


def test_plan_marks_an_already_archived_mail_and_offers_no_candidates(
    cfg, archive_root
):
    _seed_index(cfg, archive_root, {"Project Alpha": ["Project Alpha kickoff"]})
    conn = init_db(cfg["database"]["path"])
    EmailRepository(conn).upsert_email(EmailRecord(
        file_path=str(archive_root / "Project Alpha" / "007 - already.msg"),
        folder_path=str(archive_root / "Project Alpha"),
        filename="007 - already.msg",
        subject="Project Alpha weekly update",
        file_mtime=99.0,
        message_id="a@example.invalid",
    ))
    conn.commit()
    conn.close()

    client = FakeOutlookClient([
        _mail("a@example.invalid", "Project Alpha weekly update"),
        _mail("b@example.invalid", "Project Alpha budget question"),
    ])
    doc = batch.plan(client, cfg, candidates=5)

    archived, fresh = doc["mails"]
    assert archived["already_archived"] == str(
        archive_root / "Project Alpha" / "007 - already.msg"
    )
    assert archived["candidates"] == []
    # Its files are on disk but it is still sitting in the Inbox: the shape a
    # consumer offers a retry for (issue #59).
    assert archived["in_inbox"] is True
    assert fresh["already_archived"] is None
    assert doc["counts"]["already_archived"] == 1
    assert doc["counts"]["planned"] == 1


def test_plan_reports_a_mail_with_no_message_id_as_skipped(cfg, archive_root):
    """A mail with no identity is reported, never guessed at."""
    _seed_index(cfg, archive_root, {"Project Alpha": ["Project Alpha kickoff"]})
    client = FakeOutlookClient([
        _mail("", "A draft with no Message-ID"),
        _mail("b@example.invalid", "Project Alpha budget"),
    ])

    doc = batch.plan(client, cfg, candidates=5)

    assert doc["counts"] == {
        "inbox": 2, "planned": 1, "already_archived": 0, "skipped": 1,
    }
    assert doc["skipped"] == [{
        "entry_id": "entry-none",
        "subject": "A draft with no Message-ID",
        "reason": batch.SKIP_NO_MESSAGE_ID,
    }]
    assert [m["message_id"] for m in doc["mails"]] == ["b@example.invalid"]


def test_plan_on_an_empty_inbox_is_an_empty_document_not_an_error(cfg):
    doc = batch.plan(FakeOutlookClient([]), cfg)
    assert doc["counts"]["inbox"] == 0
    assert doc["mails"] == []


# ------------------------------------------------------------------ apply ---

def test_apply_writes_the_bundle_moves_the_mail_and_tags_it(cfg, archive_root):
    dest = archive_root / "Project Alpha"
    item = _mail(
        "a@example.invalid", "Project Alpha weekly update",
        attachments=[_FakeAttachment("report.pdf")],
    )
    client = FakeOutlookClient([item])

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest),
         "date_prefix": False},
    ])

    assert doc["verb"] == "apply"
    assert doc["counts"] == {"requested": 1, "applied": 1, "failed": 0}
    result = doc["results"][0]
    assert set(result) == {
        "message_id", "folder_path", "ok", "sequence_number", "files",
        "entry_id", "moved", "categorized", "reused", "move_via", "error",
    }
    assert result["reused"] is False
    assert result["ok"] is True
    assert result["error"] is None
    assert result["sequence_number"] == "001"
    assert [Path(p).name for p in result["files"]] == [
        "001 - Project Alpha weekly update.msg", "001 - report.pdf",
    ]
    assert all(Path(p).exists() for p in result["files"])

    # Moved out of the Inbox, into the configured folder, with the category.
    assert client.folders[None] == []
    assert client.folders["Archive"] == [item]
    assert item.Categories == "Filed by batch"
    # The reported EntryID is the post-move one, never the stale Inbox id.
    assert result["entry_id"] == "entry-a@example.invalid-in-Archive"


def test_apply_uses_the_per_decision_date_prefix_over_the_config(cfg, archive_root):
    """The caller owns the naming form; the global toggle is not consulted."""
    dest = archive_root / "Project Alpha"
    assert cfg["naming"]["date_prefix"] is False
    client = FakeOutlookClient([
        _mail("a@example.invalid", "Dated one",
              sent=datetime(2026, 3, 14, 9, 0, 0)),
        _mail("b@example.invalid", "Undated one",
              sent=datetime(2026, 3, 14, 9, 0, 0)),
    ])

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest),
         "date_prefix": True},
        {"message_id": "b@example.invalid", "folder_path": str(dest),
         "date_prefix": False},
    ])

    dated, undated = (Path(r["files"][0]).name for r in doc["results"])
    assert dated == "2026-03-14 - 001 - Dated one.msg"
    assert undated == "002 - Undated one.msg"


def test_apply_defaults_date_prefix_to_off_when_the_decision_omits_it(
    cfg, archive_root
):
    dest = archive_root / "Project Alpha"
    client = FakeOutlookClient([_mail("a@example.invalid", "No opinion")])
    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])
    assert Path(doc["results"][0]["files"][0]).name == "001 - No opinion.msg"


def test_apply_reports_a_missing_mail_without_aborting_the_run(cfg, archive_root):
    dest = archive_root / "Project Alpha"
    client = FakeOutlookClient([_mail("b@example.invalid", "Present")])

    doc = batch.apply(client, cfg, [
        {"message_id": "gone@example.invalid", "folder_path": str(dest)},
        {"message_id": "b@example.invalid", "folder_path": str(dest)},
    ])

    assert doc["counts"] == {"requested": 2, "applied": 1, "failed": 1}
    missing, present = doc["results"]
    assert missing["ok"] is False
    assert missing["error"]["code"] == batch.ERROR_NOT_IN_INBOX
    assert missing["files"] == []
    assert present["ok"] is True, "one bad decision must not stop the next"


def test_apply_rejects_a_decision_with_no_folder(cfg):
    client = FakeOutlookClient([_mail("a@example.invalid", "Somewhere")])
    doc = batch.apply(client, cfg, [{"message_id": "a@example.invalid"}])
    assert doc["results"][0]["error"]["code"] == batch.ERROR_BAD_DECISION
    assert client.folders[None], "a rejected decision must not move the mail"


def test_apply_keeps_the_written_files_when_the_move_fails(cfg, archive_root):
    """The case that makes an apply result revertible even when it failed."""
    dest = archive_root / "Project Alpha"
    client = FakeOutlookClient([_mail("a@example.invalid", "Half done")])
    client.move_should_fail = True

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])

    result = doc["results"][0]
    assert result["ok"] is False
    assert result["moved"] is False
    assert result["error"]["code"] == batch.ERROR_MOVE_FAILED
    assert result["files"] and Path(result["files"][0]).exists(), (
        "the files are on disk, so the caller must be told about them"
    )


def test_apply_tells_a_failed_tag_apart_from_a_failed_move(cfg, archive_root):
    """Two different states to recover from, so two different codes.

    After a move_failed the mail is still in the Inbox; after a
    category_failed it is already filed and only *looks* untouched in Outlook.
    Folding the second into the first would send a caller looking in the wrong
    folder.
    """
    dest = archive_root / "Project Alpha"
    item = _mail("a@example.invalid", "Tagged badly")
    client = FakeOutlookClient([item])
    client.category_should_fail = True

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])

    result = doc["results"][0]
    assert result["ok"] is False
    assert result["error"]["code"] == batch.ERROR_CATEGORY_FAILED
    assert result["moved"] is True, "the move did happen and must be reported so"
    assert result["categorized"] is False
    assert client.folders["Archive"] == [item], "the mail really is filed"
    assert result["files"] and Path(result["files"][0]).exists()


def test_apply_moves_a_reference_re_acquired_from_the_store(cfg, archive_root):
    """The mail archived is not the reference moved (issue #59).

    ``SaveAs`` can leave the in-memory MailItem flagged as modified; moving a
    reference read back out of the store by EntryID is what avoids the
    0x80040109 that flag causes.
    """
    dest = archive_root / "Project Alpha"
    item = _mail("a@example.invalid", "Fresh reference")
    client = FakeOutlookClient([item])

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])

    result = doc["results"][0]
    assert result["ok"] is True
    assert result["move_via"] == batch.MOVE_VIA_REFETCHED
    assert client.refetched == [("entry-a@example.invalid", FAKE_STORE_ID)], (
        "the re-fetch must name the store the mail came from (issue #130)"
    )


def test_apply_retries_on_a_fresh_reference_without_saving_when_refused(
    cfg, archive_root
):
    """The reported failure: the first Move is refused with 0x80040109.

    A store that refuses the fresh reference too gets the mail fetched again
    once every reference is let go, and moved — never ``Save()``d, since a
    save on a refused mail is refused as well (issue #130) — and the result
    says which path finished it.
    """
    dest = archive_root / "Project Alpha"
    item = _mail("a@example.invalid", "Changed underfoot")
    client = FakeOutlookClient([item])
    client.move_message_changed_once = True
    saves_before = item.save_calls

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])

    result = doc["results"][0]
    assert result["ok"] is True, "a changed message must not strand the mail"
    assert result["moved"] is True
    assert result["error"] is None
    assert result["move_via"] == batch.MOVE_VIA_RELEASED_RETRY
    # The one save is the category's, on the mail already moved; the refused
    # reference itself is never saved (issue #130).
    assert item.save_calls == saves_before + 1, "a refused mail is never saved"
    assert client.refetched == [("entry-a@example.invalid", FAKE_STORE_ID)] * 2
    assert client.folders[None] == [], "the Inbox must lose the mail"
    assert client.folders["Archive"] == [item]
    assert result["entry_id"] == "entry-a@example.invalid-in-Archive"


def test_apply_still_moves_when_the_item_cannot_be_re_acquired(cfg, archive_root):
    """A failed re-acquire is a fallback, not a failure: the original
    reference is moved and the result says so."""
    dest = archive_root / "Project Alpha"
    item = _mail("a@example.invalid", "No fresh copy")
    client = FakeOutlookClient([item])
    client.refetch_should_fail = True

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])

    result = doc["results"][0]
    assert result["ok"] is True
    assert result["move_via"] == batch.MOVE_VIA_ORIGINAL
    assert client.folders["Archive"] == [item]


def test_apply_reports_a_move_that_fails_for_another_reason_unchanged(
    cfg, archive_root
):
    """Only 0x80040109 earns the retry — any other refusal is still a
    ``move_failed`` with the files reported for a revert."""
    dest = archive_root / "Project Alpha"
    client = FakeOutlookClient([_mail("a@example.invalid", "Refused outright")])
    client.move_should_fail = True

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])

    result = doc["results"][0]
    assert result["ok"] is False
    assert result["error"]["code"] == batch.ERROR_MOVE_FAILED
    assert result["move_via"] == ""
    assert result["files"] and Path(result["files"][0]).exists()


def test_a_move_refused_twice_as_changed_is_its_own_error_code(cfg, archive_root):
    """The state the real mailbox was in: refused on the re-acquired reference
    and after saving it too, because something in the running Outlook was
    holding the item. It gets its own code so a consumer can tell "close a
    window and retry" apart from a mail that needs a human (issue #76), and the
    rest of the result still describes a revertible half-done mail."""
    client, result = _stuck_move(cfg, archive_root, open_in_inspector=True)

    assert result["ok"] is False
    assert result["error"]["code"] == batch.ERROR_MESSAGE_CHANGED
    assert result["error"]["code"] != batch.ERROR_MOVE_FAILED
    assert "re-acquired reference" in result["error"]["message"]
    assert "re-applying writes nothing" in result["error"]["message"]
    assert result["files"], "the files are on disk and must be reported"
    assert client.folders[None], "the mail is still in the Inbox"


def test_a_mail_open_in_a_window_is_told_to_close_it_not_restart(cfg, archive_root):
    """The 2026-09-16 occurrence: an open inspector held the mail, and closing
    that one window cleared it with no restart. Leading with the restart is
    what this replaces, so the message must not ask for one here."""
    _, result = _stuck_move(cfg, archive_root, open_in_inspector=True)

    message = result["error"]["message"]
    assert "open in an Outlook window" in message
    assert "Close that window" in message
    assert "Restart Outlook" not in message


def test_a_mail_no_window_holds_is_told_to_restart(cfg, archive_root):
    """Issue #59's case, now reached only on evidence rather than by assumption:
    the windows were read and none holds this mail, so a restart is the remedy
    that is actually left."""
    _, result = _stuck_move(cfg, archive_root, open_in_inspector=False)

    message = result["error"]["message"]
    assert "not open in any Outlook window" in message
    assert "Restart Outlook" in message
    assert "moved to another folder automatically" in message


def test_a_mail_no_window_holds_is_not_told_a_re_apply_will_help(cfg, archive_root):
    """Issue #84: with no window holding the mail, three re-applies over 40
    seconds were all refused — the message must not invite a caller to burn
    retries before the restart, and must flag that deleting the Inbox mail by
    hand leaves the written files as its only, unlinked copy."""
    _, result = _stuck_move(cfg, archive_root, open_in_inspector=False)

    message = result["error"]["message"]
    assert "not expected to help" in message
    assert "Deleting the mail from the Inbox by hand" in message
    assert message.index("not expected to help") < message.index("Restart Outlook")


@pytest.mark.parametrize(
    "client_state",
    [
        {"open_in_inspector": None},     # the check ran and could not answer
        {"inspector_check_raises": True},  # the check itself blew up
    ],
    ids=["undetermined", "check_raised"],
)
def test_an_undetermined_window_check_says_so_and_never_claims_no_window(
    cfg, archive_root, client_state
):
    """A check that could not establish the fact reports that as its own state
    (global CLAUDE.md) — folding it into "no window" would point an operator at
    a restart they do not need, which is the whole failure #76 is about."""
    _, result = _stuck_move(cfg, archive_root, **client_state)

    message = result["error"]["message"]
    assert result["error"]["code"] == batch.ERROR_MESSAGE_CHANGED
    assert "could not be determined" in message
    assert "not open in any Outlook window" not in message
    # Both remedies offered, cheap one first — that is all that is honest here.
    assert message.index("close any window") < message.index("restart Outlook")

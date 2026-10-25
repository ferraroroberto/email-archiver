"""The ``apply`` move: Outlook refusing it, and what the retries and the error report.
"""
from __future__ import annotations

from pathlib import Path


from email_archiver import batch
from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRecord, EmailRepository
from email_archiver.outlook.mapi import (
    is_message_changed_error,
)
from tests.batch.batch_support import (
    FAKE_STORE_ID,
    FakeOutlookClient,
    HeldByReferenceClient,
    _FakeComError,
    _mail,
    _message_changed_error,
    _stuck_move,
)
from tests.batch.batch_support import (  # noqa: F401 - fixtures are used by name
    _no_retry_pauses,
    archive_root,
    cfg,
)


# ----------------------------------------------------- selection hold (#111) ---

def _held_by_the_explorer(cfg, archive_root, **client_state):
    """Run an apply whose Move is refused only while Outlook's Explorer holds
    the mail as its selection, and return the client and the one result."""
    dest = archive_root / "Project Alpha"
    client = FakeOutlookClient([_mail("a@example.invalid", "Selected in the pane")])
    client.selection_held = True
    for name, value in client_state.items():
        setattr(client, name, value)

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])
    return client, doc["results"][0]


def test_apply_frees_a_mail_held_by_the_explorer_and_moves_it(cfg, archive_root):
    """Issue #111: the mail was selected in the reading pane, every defence was
    refused, and selecting another folder then the mail again let the very same
    decision file. ``apply`` parks the view itself, moves the mail while it is
    away, and reports it as its own ``move_via`` — no restart, no second run."""
    client, result = _held_by_the_explorer(cfg, archive_root, folder_switch_result=True)

    assert result["ok"] is True, "a selection hold must not strand the mail"
    assert result["moved"] is True
    assert result["error"] is None
    assert result["move_via"] == batch.MOVE_VIA_FOLDER_SWITCH
    assert client.folder_switches == 1, "the switch is tried once, after the refusal"
    assert client.view_away is False, "the view is put back afterwards"
    assert client.folders[None] == [], "the Inbox must lose the mail"
    assert len(client.folders["Archive"]) == 1


def test_the_folder_switch_is_only_tried_after_the_move_was_refused(cfg, archive_root):
    """Switching folders disturbs the user's view, so a move that goes through
    on its first attempt must never touch the Explorer."""
    dest = archive_root / "Project Alpha"
    client = FakeOutlookClient([_mail("a@example.invalid", "Moves fine")])
    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])
    assert doc["results"][0]["ok"] is True
    assert client.folder_switches == 0


def test_a_move_refused_for_another_reason_never_switches_folders(cfg, archive_root):
    dest = archive_root / "Project Alpha"
    client = FakeOutlookClient([_mail("a@example.invalid", "Refused outright")])
    client.move_should_fail = True

    batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])

    assert client.folder_switches == 0


def test_an_explorer_on_another_folder_is_left_alone_and_told_the_remedy(
    cfg, archive_root
):
    """The Explorer is not showing the mail's folder, so there is nothing for
    ``apply`` to switch. The refusal is reported with the folder-switch remedy
    the user can do by hand — not a restart, which nothing here justifies yet."""
    client, result = _held_by_the_explorer(
        cfg, archive_root, folder_switch_result=False, open_in_inspector=False
    )

    message = result["error"]["message"]
    assert result["error"]["code"] == batch.ERROR_MESSAGE_CHANGED
    assert "Select another folder in Outlook, then this message again" in message
    assert "Restart Outlook" not in message
    assert message.index("Select another folder") < message.index("restart Outlook")
    assert client.folders[None], "the mail is still in the Inbox"


def test_a_folder_switch_that_could_not_run_is_not_reported_as_tried(
    cfg, archive_root
):
    """A switch that did not happen is not evidence the hold survives one:
    claiming it did would send the user to restart Outlook for nothing. The
    result is the manual remedy, and the run still ends cleanly."""
    _, result = _held_by_the_explorer(
        cfg, archive_root, open_in_inspector=False, folder_switch_result=None
    )

    message = result["error"]["message"]
    assert result["ok"] is False
    assert result["error"]["code"] == batch.ERROR_MESSAGE_CHANGED
    assert "Select another folder in Outlook" in message
    assert "moved to another folder automatically" not in message
    assert "Restart Outlook" not in message


def test_a_refusal_that_survives_the_folder_switch_keeps_the_restart_advice(
    cfg, archive_root
):
    """Restart advice is for a hold that outlasted the release too — the case
    issue #59 and #84 saw — and it says the switch was already tried."""
    client, result = _stuck_move(cfg, archive_root, open_in_inspector=False)

    message = result["error"]["message"]
    assert client.folder_switches == 1, "the view is parked once, never again"
    assert "Restart Outlook" in message
    assert "moved to another folder automatically" in message
    assert "did not free it" in message


def test_a_mail_open_in_a_window_is_still_told_to_close_it_after_a_switch(
    cfg, archive_root
):
    _, result = _stuck_move(cfg, archive_root, open_in_inspector=True)

    message = result["error"]["message"]
    assert "Close that window" in message
    assert "Restart Outlook" not in message


# ------------------------------------------------- stale reference (#130) ---

def _apply_held_by_reference(cfg, archive_root, **client_state):
    dest = archive_root / "Project Alpha"
    item = _mail("a@example.invalid", "Held twice over")
    client = HeldByReferenceClient([item])
    for name, value in client_state.items():
        setattr(client, name, value)
    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])
    return client, item, doc["results"][0]


def test_a_refusal_held_by_a_reference_and_the_selection_clears_once_both_let_go(
    cfg, archive_root
):
    """Issue #130: the folder switch kept failing because the retry worked on
    the same stale object — the reference ``apply`` still held, saved and moved
    again after coming back to the Inbox. The move goes through only on a fresh
    reference that is the sole one alive, made while the view is still away;
    and nothing refused is ever saved."""
    client, item, result = _apply_held_by_reference(cfg, archive_root)

    assert result["ok"] is True, result["error"]
    assert result["moved"] is True
    assert result["move_via"] == batch.MOVE_VIA_FOLDER_SWITCH
    assert item.save_calls == 1, "only the category's own save, never a refused mail's"
    assert client.folders[None] == [], "the Inbox must lose the mail"
    assert client.folders["Archive"] == [item]
    assert client.view_away is False, "the view is put back afterwards"


def test_the_original_is_let_go_before_the_first_move(cfg, archive_root):
    """With no selection holding it, a mail refused only while ``apply`` still
    held the reference it archived moves on the first attempt: the original is
    let go before the re-fetched reference is moved, and the view is never
    touched."""
    client, _, result = _apply_held_by_reference(cfg, archive_root, selection_held=False)

    assert result["ok"] is True, result["error"]
    assert result["move_via"] == batch.MOVE_VIA_REFETCHED
    assert client.refusals == 0
    assert client.folder_switches == 0


def test_a_refused_move_leaves_breadcrumbs_for_the_next_occurrence(
    cfg, archive_root, caplog
):
    """The next refusal on a real mailbox must be diagnosable from the log
    alone: whether the re-fetch was the original object, that every reference
    was let go first, which store was named (as a fingerprint, never the id),
    where the view was, and which attempt finally moved it."""
    with caplog.at_level("INFO", logger="email_archiver.batch_move"):
        _apply_held_by_reference(cfg, archive_root)

    log = caplog.text
    assert "the re-fetched object is the original: no" in log
    assert "References to it held now: none (all released)" in log
    assert "with StoreID #" in log and FAKE_STORE_ID not in log
    assert "Outlook's view while a@example.invalid is moved: parked on another folder" in log
    assert "on attempt 2, 0.0 s after the one before" in log


def test_apply_finishes_an_already_archived_mail_without_writing_files(
    cfg, archive_root
):
    """The second half of issue #59: a mail whose files are already on disk but
    which never left the Inbox can be finished by applying a decision for it.

    Nothing new is written, the existing ``.msg`` is reported back, and the
    mail is moved and tagged like any other.
    """
    dest = archive_root / "Project Alpha"
    dest.mkdir(parents=True)
    existing = dest / "007 - Already written.msg"
    existing.write_text("synthetic .msg", encoding="utf-8")

    conn = init_db(cfg["database"]["path"])
    EmailRepository(conn).upsert_email(EmailRecord(
        file_path=str(existing),
        folder_path=str(dest),
        filename=existing.name,
        subject="Already written",
        file_mtime=99.0,
        message_id="a@example.invalid",
    ))
    conn.commit()
    conn.close()

    item = _mail("a@example.invalid", "Already written")
    client = FakeOutlookClient([item])

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])

    result = doc["results"][0]
    assert result["ok"] is True
    assert result["reused"] is True
    assert result["files"] == [str(existing)]
    assert result["sequence_number"] == ""
    assert item.saved_as == [], "nothing may be written for a reused mail"
    assert list(dest.iterdir()) == [existing], "no new file in the folder"
    assert client.folders[None] == []
    assert client.folders["Archive"] == [item]
    assert item.Categories == "Filed by batch"
    assert result["moved"] is True and result["categorized"] is True


def test_apply_archives_again_when_the_indexed_file_is_gone(cfg, archive_root):
    """A stale index row is not proof the files are there: the mail is filed
    for real rather than reported as reused over a path that no longer exists.
    """
    dest = archive_root / "Project Alpha"
    conn = init_db(cfg["database"]["path"])
    EmailRepository(conn).upsert_email(EmailRecord(
        file_path=str(dest / "007 - Deleted since.msg"),
        folder_path=str(dest),
        filename="007 - Deleted since.msg",
        subject="Deleted since",
        file_mtime=99.0,
        message_id="a@example.invalid",
    ))
    conn.commit()
    conn.close()

    client = FakeOutlookClient([_mail("a@example.invalid", "Deleted since")])
    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])

    result = doc["results"][0]
    assert result["ok"] is True
    assert result["reused"] is False
    assert Path(result["files"][0]).name == "001 - Deleted since.msg"
    assert Path(result["files"][0]).exists()


def test_a_changed_message_is_told_apart_from_any_other_com_failure():
    """The retry is keyed on the MAPI scode, not on the message text."""
    assert is_message_changed_error(_message_changed_error()) is True
    assert is_message_changed_error(
        _FakeComError(-2147221233, "The item could not be found.", None, None)
    ) is False
    assert is_message_changed_error(RuntimeError("the store refused the move")) is False


def test_apply_accepts_a_bracketed_message_id_from_the_caller(cfg, archive_root):
    """A caller echoing a raw header back must still match."""
    dest = archive_root / "Project Alpha"
    client = FakeOutlookClient([_mail("a@example.invalid", "Bracketed")])
    doc = batch.apply(client, cfg, [
        {"message_id": "<a@example.invalid>", "folder_path": str(dest)},
    ])
    assert doc["results"][0]["ok"] is True

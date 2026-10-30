"""The ``revert`` verb, the category helpers and the error document.
"""
from __future__ import annotations

from pathlib import Path


from email_archiver import batch, batch_common
from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRecord, EmailRepository
from email_archiver.outlook.mapi import (
    with_category,
    without_category,
)
from tests.batch.batch_support import (
    FakeOutlookClient,
    _FakeAttachment,
    _index_archived_file,
    _mail,
)
from tests.batch.batch_support import (  # noqa: F401 - fixtures are used by name
    _no_retry_pauses,
    archive_root,
    cfg,
)


# ----------------------------------------------------------------- revert ---

def _apply_one(cfg, archive_root, client, message_id="a@example.invalid"):
    dest = archive_root / "Project Alpha"
    return batch.apply(client, cfg, [
        {"message_id": message_id, "folder_path": str(dest)},
    ])["results"][0]


def test_revert_deletes_the_files_and_puts_the_mail_back(cfg, archive_root):
    item = _mail(
        "a@example.invalid", "Round trip",
        attachments=[_FakeAttachment("report.pdf")],
    )
    client = FakeOutlookClient([item])
    applied = _apply_one(cfg, archive_root, client)
    assert client.folders["Archive"] == [item]

    doc = batch.revert(client, cfg, [
        {"message_id": "a@example.invalid", "files": applied["files"]},
    ])

    assert doc["verb"] == "revert"
    assert doc["counts"] == {"requested": 1, "reverted": 1, "failed": 0}
    result = doc["results"][0]
    assert set(result) == {
        "message_id", "ok", "deleted", "missing", "refused", "file_errors",
        "moved_back", "category_removed", "entry_id", "error",
        "index_rows_removed",
    }
    assert result["ok"] is True
    assert sorted(result["deleted"]) == sorted(str(Path(p)) for p in applied["files"])
    assert not any(Path(p).exists() for p in applied["files"])
    assert result["moved_back"] is True
    assert result["category_removed"] is True
    assert client.folders[None] == [item]
    assert client.folders["Archive"] == []
    assert item.Categories == ""
    # No row was ever indexed for this mail (no scan ran), so there is
    # nothing to remove -- this is not the stale-index regression below.
    assert result["index_rows_removed"] == 0


def _index_decoy_row(cfg, applied, subject):
    """A second, unrelated row filed in the same folder as ``applied``'s file.

    Gives the suggestion engine's FTS index something to rank for that folder
    once the row for the reverted mail itself is gone -- otherwise "offered
    again with candidates" can't be told apart from "offered again with an
    empty candidate list because the index is now empty".
    """
    conn = init_db(cfg["database"]["path"])
    repo = EmailRepository(conn)
    folder = Path(applied["files"][0]).parent
    decoy_path = str(folder / "decoy.msg")
    repo.upsert_email(EmailRecord(
        file_path=decoy_path,
        folder_path=str(folder),
        filename="decoy.msg",
        subject=subject,
    ))
    conn.commit()
    conn.close()


def test_revert_removes_the_index_row_so_the_next_plan_offers_the_mail_again(
    cfg, archive_root
):
    """Regression for #57: revert used to leave the index row for the deleted
    ``.msg`` in place, so the next `plan` reported the mail `already_archived`
    at a path that no longer existed instead of offering it again."""
    subject = "Stale index row"
    item = _mail("a@example.invalid", subject)
    client = FakeOutlookClient([item])
    applied = _apply_one(cfg, archive_root, client)
    _index_archived_file(cfg, applied, subject=subject)
    _index_decoy_row(cfg, applied, subject=subject)

    doc = batch.revert(client, cfg, [
        {"message_id": "a@example.invalid", "files": applied["files"]},
    ])
    result = doc["results"][0]
    assert result["ok"] is True
    assert result["index_rows_removed"] == 1

    conn = init_db(cfg["database"]["path"])
    repo = EmailRepository(conn)
    assert repo.find_path_by_message_id("a@example.invalid") is None
    conn.close()

    plan_doc = batch.plan(client, cfg)
    mail = plan_doc["mails"][0]
    assert mail["already_archived"] is None
    assert mail["candidates"], "the mail must be offered again with candidates"


def test_revert_leaves_the_index_row_when_the_file_delete_fails(
    cfg, archive_root, monkeypatch
):
    """A revert whose file delete failed must leave the row alone -- the file
    is still there, so the index still describing it is correct, not stale."""
    item = _mail("a@example.invalid", "Delete failure")
    client = FakeOutlookClient([item])
    applied = _apply_one(cfg, archive_root, client)
    _index_archived_file(cfg, applied)

    real_unlink = Path.unlink

    def failing_unlink(self, *args, **kwargs):
        if str(self) == str(Path(applied["files"][0])):
            raise PermissionError("simulated: file is locked")
        return real_unlink(self, *args, **kwargs)

    monkeypatch.setattr(Path, "unlink", failing_unlink)

    doc = batch.revert(client, cfg, [
        {"message_id": "a@example.invalid", "files": applied["files"]},
    ])
    result = doc["results"][0]
    assert result["file_errors"]
    assert result["index_rows_removed"] == 0

    conn = init_db(cfg["database"]["path"])
    repo = EmailRepository(conn)
    assert repo.find_path_by_message_id("a@example.invalid") == applied["files"][0]
    conn.close()


def test_revert_refuses_a_file_outside_the_archive_roots(cfg, tmp_path):
    """The path-safety guard: a decisions file cannot reach the rest of the disk."""
    outsider = tmp_path / "not-the-archive" / "precious.txt"
    outsider.parent.mkdir()
    outsider.write_text("do not delete me", encoding="utf-8")

    client = FakeOutlookClient([])
    doc = batch.revert(client, cfg, [
        {"message_id": "a@example.invalid", "files": [str(outsider)]},
    ])

    result = doc["results"][0]
    assert result["deleted"] == []
    assert result["refused"] == [
        {"path": str(outsider), "reason": batch.REFUSED_OUTSIDE_ROOTS}
    ]
    assert outsider.exists(), "a refused file must still be on disk"
    assert result["ok"] is False


def test_revert_refuses_a_traversal_out_of_an_archive_root(cfg, archive_root, tmp_path):
    outsider = tmp_path / "precious.txt"
    outsider.write_text("do not delete me", encoding="utf-8")
    sneaky = str(archive_root / ".." / "precious.txt")

    doc = batch.revert(FakeOutlookClient([]), cfg, [
        {"message_id": "a@example.invalid", "files": [sneaky]},
    ])

    assert doc["results"][0]["refused"][0]["reason"] == batch.REFUSED_OUTSIDE_ROOTS
    assert outsider.exists()


def test_revert_reports_an_already_deleted_file_as_missing_not_an_error(
    cfg, archive_root
):
    """A revert run twice is not a failure to explain."""
    item = _mail("a@example.invalid", "Twice")
    client = FakeOutlookClient([item])
    applied = _apply_one(cfg, archive_root, client)
    items = [{"message_id": "a@example.invalid", "files": applied["files"]}]

    batch.revert(client, cfg, items)
    doc = batch.revert(client, cfg, items)

    result = doc["results"][0]
    assert result["deleted"] == []
    assert result["missing"] == [str(Path(applied["files"][0]))]
    assert result["file_errors"] == []
    # The mail is back in the Inbox already, so it is no longer in Archive.
    assert result["error"]["code"] == batch.ERROR_NOT_IN_ARCHIVE


def test_revert_still_deletes_the_files_when_the_mail_cannot_be_found(
    cfg, archive_root
):
    """Files first, mail second: a mail the user already moved by hand must not
    strand its files on disk."""
    item = _mail("a@example.invalid", "Moved by hand")
    client = FakeOutlookClient([item])
    applied = _apply_one(cfg, archive_root, client)
    client.folders["Archive"].clear()

    doc = batch.revert(client, cfg, [
        {"message_id": "a@example.invalid", "files": applied["files"]},
    ])

    result = doc["results"][0]
    assert result["deleted"] == [str(Path(applied["files"][0]))]
    assert result["moved_back"] is False
    assert result["error"]["code"] == batch.ERROR_NOT_IN_ARCHIVE


def test_revert_needs_a_message_id_to_move_a_mail_back(cfg):
    doc = batch.revert(FakeOutlookClient([]), cfg, [{"files": []}])
    assert doc["results"][0]["error"]["code"] == batch.ERROR_BAD_DECISION


# ------------------------------------------------------- category handling ---

def test_a_category_is_added_without_losing_the_user_s_own():
    assert with_category("Red; Blue", "Filed") == "Red; Blue; Filed"


def test_adding_a_category_twice_does_not_duplicate_it():
    assert with_category("Red; Filed", "Filed") == "Red; Filed"
    assert with_category("Red; filed", "Filed") == "Red; filed"


def test_removing_a_category_leaves_the_others_alone():
    assert without_category("Red; Filed; Blue", "Filed") == "Red; Blue"


def test_removing_a_category_that_is_not_there_changes_nothing():
    assert without_category("Red; Blue", "Filed") == "Red; Blue"


def test_the_category_helpers_survive_an_empty_string():
    assert with_category("", "Filed") == "Filed"
    assert without_category("", "Filed") == ""
    assert without_category(None, "Filed") == ""


# ------------------------------------------------------- the error envelope --

def test_the_error_document_carries_a_code_and_the_same_envelope():
    doc = batch_common.error_document("plan", "outlook_unavailable", "nope")
    assert doc["verb"] == "plan"
    assert doc["schema_version"] == batch_common.SCHEMA_VERSION
    assert doc["error"] == {"code": "outlook_unavailable", "message": "nope"}

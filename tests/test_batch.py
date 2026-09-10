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
from email_archiver import renumber as renumber_module
from email_archiver.archiver import archiver as archiver_module
from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRecord, EmailRepository
from email_archiver.outlook.client import (
    InboxMail,
    is_message_changed_error,
    with_category,
    without_category,
)
from email_archiver.scanner.scanner import MsgFacts


# --------------------------------------------------------------- fake COM ---

class _FakePropertyAccessor:
    """Raises for an absent property, exactly as Outlook's does."""

    def __init__(self, props: dict[str, object]) -> None:
        self._props = props

    def GetProperty(self, key: str):  # noqa: N802 - COM's spelling
        if key in self._props:
            return self._props[key]
        raise RuntimeError(f"property not found: {key}")


class _FakeComError(Exception):
    """A ``pywintypes.com_error`` look-alike.

    Same ``args`` shape as the real thing — ``(hresult, description, excepinfo,
    argerror)`` with the MAPI scode buried in ``excepinfo[5]`` — so the
    detection helper is exercised exactly as it will be against Outlook, with
    no pywin32 import in the test suite.
    """


def _message_changed_error() -> _FakeComError:
    """The failure this issue is about: MAPI_E_OBJECT_CHANGED (0x80040109).

    Raised by ``MailItem.Move`` on an item Outlook considers modified — which
    ``SaveAs`` can leave it as, on the very same reference the archiver just
    wrote to disk.
    """
    return _FakeComError(
        -2147352567,
        "Exception occurred.",
        (
            4096,
            "Microsoft Outlook",
            "The operation cannot be performed because the message has been changed.",
            None,
            0,
            -2147221239,
        ),
        None,
    )


class _FakeAttachment:
    def __init__(self, filename: str, payload: bytes = b"attachment") -> None:
        self.FileName = filename
        self.PropertyAccessor = _FakePropertyAccessor({})
        self._payload = payload

    def SaveAsFile(self, path: str) -> None:  # noqa: N802 - COM's spelling
        Path(path).write_bytes(self._payload)


class _FakeMailItem:
    """The slice of a COM MailItem the archiver and the client actually touch."""

    def __init__(
        self,
        *,
        message_id: str,
        subject: str,
        sender: str = "sender@example.invalid",
        recipients: str = "me@example.invalid",
        sent: datetime | None = None,
        attachments: list[_FakeAttachment] | None = None,
        flag_status: int = 0,
    ) -> None:
        self.message_id = message_id
        self.Subject = subject
        self.sender = sender
        self.recipients = recipients
        self.SentOn = sent or datetime(2026, 3, 14, 9, 30, 0)
        self.Attachments = attachments or []
        self.Categories = ""
        self.EntryID = f"entry-{message_id or 'none'}"
        self.flag_status = flag_status
        self.saved_as: list[str] = []
        self.save_calls = 0

    def SaveAs(self, path: str, fmt: int) -> None:  # noqa: N802 - COM's spelling
        Path(path).write_text("synthetic .msg", encoding="utf-8")
        self.saved_as.append(path)

    def Save(self) -> None:  # noqa: N802 - COM's spelling
        self.save_calls += 1


class FakeOutlookClient:
    """Stands in for OutlookClient's batch surface.

    Folders are plain lists keyed by name, with ``None`` meaning the Inbox —
    the same convention the real client uses, so the orchestration under test is
    the orchestration that ships.
    """

    def __init__(self, inbox: list[_FakeMailItem]) -> None:
        self.folders: dict[str | None, list[_FakeMailItem]] = {None: list(inbox)}
        self.move_should_fail = False
        # One Move refused with 0x80040109 before the store accepts it — the
        # shape issue #59 hit on a real mailbox.
        self.move_message_changed_once = False
        # Every Move refused with 0x80040109 — the state the real mailbox was
        # in until Outlook itself was restarted (issue #59).
        self.move_message_changed_always = False
        self.refetch_should_fail = False
        self.refetched: list[str] = []
        self.category_should_fail = False
        self.moves: list[tuple[str, str | None]] = []

    # -- the surface batch.py calls -------------------------------------------

    def iter_inbox(self, preview_len: int = 500):
        for item in list(self.folders[None]):
            yield self.read_mail(item, preview_len)

    def read_mail(self, item: _FakeMailItem, preview_len: int = 0) -> InboxMail:
        return InboxMail(
            message_id=item.message_id,
            entry_id=item.EntryID,
            subject=item.Subject,
            sender=item.sender,
            recipients=item.recipients,
            date_sent=item.SentOn,
            body_preview=("synthetic body " * 40)[:preview_len],
            attachment_count=len(item.Attachments),
            flag_status=item.flag_status,
            item=item,
        )

    def find_by_message_id(self, message_id: str, folder_name: str | None = None):
        if not message_id:
            return None
        for item in self.folders.get(folder_name, []):
            if item.message_id == message_id:
                return item
        return None

    def refetch(self, item: _FakeMailItem) -> _FakeMailItem | None:
        """Stand-in for ``Session.GetItemFromID`` — a reference straight from
        the store, which on a real mailbox is a *different* COM object carrying
        none of the modified state ``SaveAs`` left on the original."""
        if self.refetch_should_fail:
            raise _FakeComError(-2147221233, "The item could not be found.", None, None)
        self.refetched.append(item.EntryID)
        return item

    def save_item(self, item: _FakeMailItem) -> None:
        item.Save()

    def move_to(self, item: _FakeMailItem, folder_name: str | None):
        if self.move_should_fail:
            raise RuntimeError("the store refused the move")
        if self.move_message_changed_always:
            raise _message_changed_error()
        if self.move_message_changed_once:
            self.move_message_changed_once = False
            raise _message_changed_error()
        for name, items in self.folders.items():
            if item in items:
                items.remove(item)
                self.moves.append((item.message_id, folder_name))
                break
        self.folders.setdefault(folder_name, []).append(item)
        # Outlook rewrites EntryID on a move; a fake that did not would let a
        # bug reading the pre-move id pass unnoticed.
        item.EntryID = f"entry-{item.message_id}-in-{folder_name or 'inbox'}"
        return item

    def set_category(self, item: _FakeMailItem, name: str) -> None:
        if self.category_should_fail:
            raise RuntimeError("the store refused the category")
        item.Categories = with_category(item.Categories, name)
        item.Save()

    def clear_category(self, item: _FakeMailItem, name: str) -> None:
        item.Categories = without_category(item.Categories, name)
        item.Save()

    def entry_id(self, item: _FakeMailItem) -> str:
        return item.EntryID


# --------------------------------------------------------------- fixtures ---

@pytest.fixture
def archive_root(tmp_path) -> Path:
    root = tmp_path / "archive"
    root.mkdir()
    return root


@pytest.fixture
def cfg(tmp_path, archive_root) -> dict:
    return {
        "archive": {"root_paths": [str(archive_root)]},
        "database": {"path": str(tmp_path / "emails.db")},
        "scanning": {"batch_size": 500, "body_preview_length": 500},
        "suggestion": {"max_suggestions": 3, "min_score": 0.0},
        "naming": {"date_prefix": False},
        "outlook": {"archive_folder": "Archive", "category": "Filed by batch"},
        "path": {"max_length": 255},
    }


def _seed_index(cfg: dict, archive_root: Path, folders: dict[str, list[str]]) -> None:
    """Put synthetic archived mail in the index so the engine has something to
    rank. ``folders`` maps a folder name to the subjects it already holds."""
    conn = init_db(cfg["database"]["path"])
    repo = EmailRepository(conn)
    n = 0
    for folder, subjects in folders.items():
        folder_path = str(archive_root / folder)
        for subject in subjects:
            n += 1
            repo.upsert_email(EmailRecord(
                file_path=f"{folder_path}/{n:03d} - seed.msg",
                folder_path=folder_path,
                filename=f"{n:03d} - seed.msg",
                subject=subject,
                sender="sender@example.invalid",
                recipients="me@example.invalid",
                date_sent="2026-01-01T00:00:00",
                body_preview=subject,
                file_mtime=float(n),
                message_id=f"seed-{n}@example.invalid",
            ))
    conn.commit()
    conn.close()


def _mail(message_id: str, subject: str, **kw) -> _FakeMailItem:
    return _FakeMailItem(message_id=message_id, subject=subject, **kw)


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
    assert client.refetched == ["entry-a@example.invalid"]


def test_apply_saves_and_retries_once_when_the_message_has_been_changed(
    cfg, archive_root
):
    """The reported failure: the first Move is refused with 0x80040109.

    A store that refuses the fresh reference too gets one ``Save()``-then-Move
    before ``apply`` gives up, and the result says which path finished it.
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
    assert result["move_via"] == batch.MOVE_VIA_SAVED_RETRY
    assert item.save_calls > saves_before, "the retry must Save() first"
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


def test_a_move_refused_twice_as_changed_says_how_to_recover(cfg, archive_root):
    """The state the real mailbox was in: refused on the re-acquired reference
    and after saving it too, because the running Outlook process itself was
    holding the item. The message has to say that, or the next occurrence is a
    bare HRESULT again."""
    dest = archive_root / "Project Alpha"
    client = FakeOutlookClient([_mail("a@example.invalid", "Stuck fast")])
    client.move_message_changed_always = True

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])

    result = doc["results"][0]
    assert result["ok"] is False
    assert result["error"]["code"] == batch.ERROR_MOVE_FAILED
    assert "restarting Outlook" in result["error"]["message"]
    assert "without writing anything" in result["error"]["message"]
    assert result["files"], "the files are on disk and must be reported"
    assert client.folders[None], "the mail is still in the Inbox"


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


def _index_archived_file(cfg, applied, message_id="a@example.invalid", subject=""):
    """Simulate the scan that normally indexes a freshly archived file.

    ``apply`` only writes the file to disk -- ``main_scan.py`` is what
    populates the index. The stale-row regression this guards only shows up
    once a row exists, so the test has to write one by hand rather than rely
    on ``apply`` to have done it.
    """
    conn = init_db(cfg["database"]["path"])
    repo = EmailRepository(conn)
    file_path = applied["files"][0]
    repo.upsert_email(EmailRecord(
        file_path=file_path,
        folder_path=str(Path(file_path).parent),
        filename=Path(file_path).name,
        subject=subject,
        message_id=message_id,
    ))
    conn.commit()
    conn.close()


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
    doc = batch.error_document("plan", "outlook_unavailable", "nope")
    assert doc["verb"] == "plan"
    assert doc["schema_version"] == batch.SCHEMA_VERSION
    assert doc["error"] == {"code": "outlook_unavailable", "message": "nope"}


# --------------------------------------------------------------- renumber ---

def test_apply_without_the_flag_reports_no_renumber_keys_at_all(cfg, archive_root):
    """The flag is opt-in: without it the document is byte-for-byte what a
    consumer built against before this shipped."""
    dest = archive_root / "Project Alpha"
    client = FakeOutlookClient([_mail("a@example.invalid", "Untouched")])

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])

    assert "renumbered" not in doc
    assert "renumber_refused" not in doc


def test_apply_with_renumber_puts_an_older_mail_in_its_date_position(
    cfg, archive_root, monkeypatch
):
    """The reported case: a mail filed into a folder after a correction takes
    `max + 1` even though it is older than what is already there.

    The mail `apply` just wrote has no index row yet — nothing has scanned it —
    so its date comes from the file. Outlook writes a real .msg there and
    extract-msg reads it; the fake client writes a placeholder, so the read is
    faked here too rather than silently degrading to "unknown date".
    """
    monkeypatch.setattr(
        renumber_module, "read_msg_facts",
        lambda path: MsgFacts("2026-03-14T09:00:00+01:00", "a@example.invalid", ()),
    )
    dest = archive_root / "Project Alpha"
    dest.mkdir(parents=True)
    newer = dest / "001 - Already filed.msg"
    newer.write_text("synthetic .msg", encoding="utf-8")

    conn = init_db(cfg["database"]["path"])
    EmailRepository(conn).upsert_email(EmailRecord(
        file_path=str(newer),
        folder_path=str(dest),
        filename=newer.name,
        subject="Already filed",
        date_sent="2026-03-20T09:00:00+01:00",
        file_mtime=1.0,
        message_id="already@example.invalid",
    ))
    conn.commit()
    conn.close()

    client = FakeOutlookClient([
        _mail("a@example.invalid", "Older one", sent=datetime(2026, 3, 14, 9, 0)),
    ])
    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ], renumber=True)

    assert doc["results"][0]["ok"] is True
    # It went in as 002 and came out as 001; the mail already there moved up.
    assert sorted(p.name for p in dest.iterdir()) == [
        "001 - Older one.msg", "002 - Already filed.msg",
    ]
    assert doc["renumber_refused"] == []
    moved = {
        Path(e["from"]).name: (Path(e["to"]).name, e["message_id"])
        for e in doc["renumbered"][str(dest)]
    }
    assert moved == {
        "001 - Already filed.msg": (
            "002 - Already filed.msg", "already@example.invalid",
        ),
        "002 - Older one.msg": ("001 - Older one.msg", "a@example.invalid"),
    }


def test_apply_with_renumber_lists_each_destination_folder_once(cfg, archive_root):
    dest = archive_root / "Project Alpha"
    other = archive_root / "Project Beta"
    client = FakeOutlookClient([
        _mail("a@example.invalid", "One", sent=datetime(2026, 3, 14, 9, 0)),
        _mail("b@example.invalid", "Two", sent=datetime(2026, 3, 15, 9, 0)),
        _mail("c@example.invalid", "Three", sent=datetime(2026, 3, 16, 9, 0)),
    ])

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
        {"message_id": "b@example.invalid", "folder_path": str(dest)},
        {"message_id": "c@example.invalid", "folder_path": str(other)},
    ], renumber=True)

    assert sorted(doc["renumbered"]) == sorted([str(dest), str(other)])


def test_apply_with_renumber_refuses_a_folder_outside_the_archive_roots(
    cfg, tmp_path
):
    """`apply` can file anywhere the caller points it — the renumber that
    follows may not."""
    outside = tmp_path / "not-the-archive"
    client = FakeOutlookClient([_mail("a@example.invalid", "Elsewhere")])

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(outside)},
    ], renumber=True)

    assert doc["results"][0]["ok"] is True
    assert doc["renumbered"] == {}
    assert doc["renumber_refused"] == [
        {"folder_path": str(outside), "reason": batch.REFUSED_OUTSIDE_ROOTS}
    ]


def test_revert_with_renumber_closes_the_gap_it_just_made(cfg, archive_root):
    dest = archive_root / "Project Alpha"
    client = FakeOutlookClient([
        _mail("a@example.invalid", "First", sent=datetime(2026, 3, 14, 9, 0)),
        _mail("b@example.invalid", "Second", sent=datetime(2026, 3, 15, 9, 0)),
        _mail("c@example.invalid", "Third", sent=datetime(2026, 3, 16, 9, 0)),
    ])
    applied = batch.apply(client, cfg, [
        {"message_id": mid, "folder_path": str(dest)}
        for mid in ("a@example.invalid", "b@example.invalid", "c@example.invalid")
    ])
    for result in applied["results"]:
        _index_archived_file(
            cfg, result, message_id=result["message_id"],
            subject=Path(result["files"][0]).stem,
        )
    assert sorted(p.name for p in dest.iterdir()) == [
        "001 - First.msg", "002 - Second.msg", "003 - Third.msg",
    ]

    doc = batch.revert(client, cfg, [
        {"message_id": "b@example.invalid", "files": applied["results"][1]["files"]},
    ], renumber=True)

    assert doc["results"][0]["ok"] is True
    assert sorted(p.name for p in dest.iterdir()) == [
        "001 - First.msg", "002 - Third.msg",
    ]
    entry = doc["renumbered"][str(dest)][0]
    assert Path(entry["from"]).name == "003 - Third.msg"
    assert Path(entry["to"]).name == "002 - Third.msg"
    assert doc["renumber_refused"] == []


def test_revert_with_renumber_ignores_a_folder_it_deleted_nothing_from(
    cfg, archive_root
):
    """No delete, no gap: a refused or already-missing file must not drag a
    whole folder through a renumber it did not need."""
    dest = archive_root / "Project Alpha"
    dest.mkdir(parents=True)
    (dest / "003 - Untouched.msg").write_text("synthetic .msg", encoding="utf-8")

    doc = batch.revert(FakeOutlookClient([]), cfg, [
        {"message_id": "a@example.invalid",
         "files": [str(dest / "009 - never existed.msg")]},
    ], renumber=True)

    assert doc["results"][0]["missing"]
    assert doc["renumbered"] == {}
    assert sorted(p.name for p in dest.iterdir()) == ["003 - Untouched.msg"]


def test_the_renumber_verb_reports_the_map_and_a_dry_run_changes_nothing(
    cfg, archive_root
):
    dest = archive_root / "Project Alpha"
    dest.mkdir(parents=True)
    for name in ("001 - one.msg", "003 - three.msg"):
        (dest / name).write_text("synthetic .msg", encoding="utf-8")

    conn = init_db(cfg["database"]["path"])
    repo = EmailRepository(conn)
    for name, date_sent in (
        ("001 - one.msg", "2026-01-01T09:00:00+01:00"),
        ("003 - three.msg", "2026-01-03T09:00:00+01:00"),
    ):
        repo.upsert_email(EmailRecord(
            file_path=str(dest / name), folder_path=str(dest), filename=name,
            subject=name, date_sent=date_sent, file_mtime=1.0,
            message_id=f"{name}@example.invalid",
        ))
    conn.commit()
    conn.close()

    planned = batch.renumber(cfg, str(dest), dry_run=True)
    assert planned["verb"] == "renumber"
    assert planned["schema_version"] == batch.SCHEMA_VERSION
    assert planned["dry_run"] is True
    assert planned["counts"]["bundles"] == 2
    assert planned["counts"]["renamed"] == 1
    assert sorted(p.name for p in dest.iterdir()) == [
        "001 - one.msg", "003 - three.msg",
    ]

    done = batch.renumber(cfg, str(dest))
    assert done["renumbered"] == planned["renumbered"]
    assert done["counts"]["index_rows_updated"] == 1
    assert sorted(p.name for p in dest.iterdir()) == [
        "001 - one.msg", "002 - three.msg",
    ]


def test_the_renumber_verb_refuses_a_folder_outside_the_archive_roots(cfg, tmp_path):
    outside = tmp_path / "not-the-archive"
    outside.mkdir()

    doc = batch.renumber(cfg, str(outside))

    assert doc["renumbered"] == {}
    assert doc["renumber_refused"] == [
        {"folder_path": str(outside), "reason": batch.REFUSED_OUTSIDE_ROOTS}
    ]

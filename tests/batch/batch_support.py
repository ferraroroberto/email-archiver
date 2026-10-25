"""The fake Outlook client, fixtures and helpers the batch tests share.

Synthetic throughout: no real folder name, address or subject.
"""
from __future__ import annotations

import weakref
from contextlib import contextmanager
from datetime import datetime
from pathlib import Path

import pytest

from email_archiver import batch, batch_move
from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRecord, EmailRepository
from email_archiver.outlook.client import InboxMail
from email_archiver.outlook.mapi import (
    with_category,
    without_category,
)


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
        # Every (EntryID, StoreID) the mail was fetched again by.
        self.refetched: list[tuple[str, str]] = []
        # What Outlook answers when asked whether the mail is open in a window:
        # True, False, or None for "the check could not be completed" — the
        # three states `message_changed` reports separately (issue #76).
        self.open_in_inspector: bool | None = None
        self.inspector_check_raises = False
        self.category_should_fail = False
        self.moves: list[tuple[str, str | None]] = []
        # The Explorer holding the item as the selected / reading-pane mail
        # (issue #111): every Move is refused while the view is on the Inbox.
        # Coming back to the Inbox selects it again (issue #130), so only a
        # move made while the view is parked elsewhere gets through.
        self.selection_held = False
        # What the real client tells the block when asked to park the Explorer
        # off the Inbox: True (parked), False (the Explorer is not showing the
        # Inbox), None (could not be determined).
        self.folder_switch_result: bool | None = False
        self.folder_switches = 0
        self.view_away = False

    def store(self):
        # The real client resolves its mailbox's store; nothing to resolve here.
        return self

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

    def item_ids(self, item: _FakeMailItem) -> tuple[str, str]:
        return item.EntryID, FAKE_STORE_ID

    def refetch(self, entry_id: str, store_id: str) -> _FakeMailItem | None:
        """Stand-in for ``Session.GetItemFromID`` — a reference straight from
        the store, which on a real mailbox is a *different* COM object carrying
        none of the modified state ``SaveAs`` left on the original."""
        if self.refetch_should_fail:
            raise _FakeComError(-2147221233, "The item could not be found.", None, None)
        self.refetched.append((entry_id, store_id))
        for items in self.folders.values():
            for item in items:
                if item.EntryID == entry_id:
                    return item
        return None

    def is_same_object(self, first, second) -> bool | None:
        return first is second

    def move_to(self, item: _FakeMailItem, folder_name: str | None):
        if self.move_should_fail:
            raise RuntimeError("the store refused the move")
        if self.move_message_changed_always or (self.selection_held and not self.view_away):
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

    @contextmanager
    def view_away_from_inbox(self):
        self.folder_switches += 1
        self.view_away = self.folder_switch_result is True
        try:
            yield self.folder_switch_result
        finally:
            self.view_away = False

    def is_open_in_inspector(self, entry_id: str) -> bool | None:
        if self.inspector_check_raises:
            raise RuntimeError("Outlook would not list its windows")
        return self.open_in_inspector


class _Ref:
    """One COM reference to a stored mail. ``GetItemFromID`` hands out a new
    object on every call, never the one already held (the issue #130 probe),
    so each lookup here is a new ``_Ref`` over the same ``_FakeMailItem``."""

    def __init__(self, mail: _FakeMailItem) -> None:
        self.mail = mail

    def __getattr__(self, name: str):
        return getattr(self.mail, name)


class HeldByReferenceClient(FakeOutlookClient):
    """Outlook refusing a move with 0x80040109 while the mail is still
    referenced anywhere else, or while the Explorer shows the Inbox and so
    holds it as its selection — the refusal issue #130 found surviving a
    save, a re-fetch, and a switch away and back. It clears only for a fresh
    reference that is the sole one alive, moved while the view is away."""

    def __init__(self, inbox: list[_FakeMailItem]) -> None:
        super().__init__(inbox)
        self.selection_held = True
        self.folder_switch_result = True
        self.refusals = 0
        self._refs: weakref.WeakSet[_Ref] = weakref.WeakSet()

    def _ref(self, mail: _FakeMailItem | None) -> _Ref | None:
        if mail is None:
            return None
        ref = _Ref(mail)
        self._refs.add(ref)
        return ref

    def find_by_message_id(self, message_id: str, folder_name: str | None = None):
        return self._ref(super().find_by_message_id(message_id, folder_name))

    def refetch(self, entry_id: str, store_id: str) -> _Ref | None:
        return self._ref(super().refetch(entry_id, store_id))

    def item_ids(self, item: _Ref) -> tuple[str, str]:
        return super().item_ids(item.mail)

    def is_same_object(self, first: _Ref, second: _Ref) -> bool | None:
        return first is second

    def move_to(self, item: _Ref, folder_name: str | None):
        if self._others_alive(item) or (self.selection_held and not self.view_away):
            self.refusals += 1
            raise _message_changed_error()
        return super().move_to(item.mail, folder_name)

    def _others_alive(self, item: _Ref) -> bool:
        return any(ref is not item and ref.mail is item.mail for ref in self._refs)


# --------------------------------------------------------------- fixtures ---

FAKE_STORE_ID = "store-0001"


@pytest.fixture(autouse=True)
def _no_retry_pauses(monkeypatch):
    """The real pauses between move retries are seconds of a test run each."""
    monkeypatch.setattr(batch_move, "MOVE_RETRY_DELAYS_SECONDS", (0.0, 0.0))

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


def _stuck_move(cfg, archive_root, **client_state):
    """Run an apply whose every Move is refused as 0x80040109, and return the
    one result — the shape all the `message_changed` cases share. The Explorer
    is switched away and back by default (issue #111), so a refusal here is one
    that survived the folder switch; pass ``folder_switch_result`` to change it."""
    dest = archive_root / "Project Alpha"
    client = FakeOutlookClient([_mail("a@example.invalid", "Stuck fast")])
    client.move_message_changed_always = True
    client.folder_switch_result = True
    for name, value in client_state.items():
        setattr(client, name, value)

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])
    return client, doc["results"][0]


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

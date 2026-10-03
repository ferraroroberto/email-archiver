"""Tests for the read-only live ``find`` verb (issue #110).

``find`` searches one mailbox's Inbox, archive folder and Sent with a DASL
``Items.Restrict`` (every word must match), returns the newest hits with
whether each is already in the index, says its coverage is Outlook's folders
only, and never moves, saves, deletes, drafts or sends. Everything here is
synthetic.
"""
from __future__ import annotations

import types
from datetime import datetime

import pytest

from email_archiver import find
from email_archiver.config import Mailbox
from email_archiver.outlook import mapi
from email_archiver.outlook.client import (
    FIND_VIA_ABSENT,
    FIND_VIA_RESTRICT,
    FIND_VIA_WALK,
    FolderHits,
    InboxMail,
    OutlookClient,
)
from tests.test_batch import _seed_index, archive_root, cfg  # noqa: F401 - fixtures are used by name
from tests.test_batch_targeted import TargetedFakeClient
from tests.test_mailboxes import batch_process  # noqa: F401 - fixture is used by name

MUTATIONS = ("Move", "Save", "Delete", "Send", "Display", "Close", "Add")


# ------------------------------------------------------------- DASL filter ---

def test_every_word_must_match_in_one_of_the_search_fields():
    flt = mapi.words_filter(["alpha", "beta"])

    assert flt.startswith("@SQL=")
    alpha, beta = flt[len("@SQL="):].split(" AND ")
    for field in mapi.DASL_FIND_FIELDS:
        assert f"\"{field}\" LIKE '%alpha%'" in alpha
        assert f"\"{field}\" LIKE '%beta%'" in beta


def test_a_quote_in_a_word_cannot_break_out_of_the_literal():
    assert "LIKE '%o''brien%'" in mapi.words_filter(["o'brien"])


def test_since_compares_received_time_or_sent_time_in_utc():
    since = datetime(2026, 9, 15, 8, 0)
    assert mapi.DASL_DATE_RECEIVED in mapi.words_filter(["x"], since)
    assert mapi.DASL_DATE_SENT in mapi.words_filter(["x"], since, sent=True)
    assert mapi.received_since_filter(since).split("= ")[1] in mapi.words_filter(["x"], since)


# ------------------------------------------------- client, read-only search ---

class _ComMail:
    """A MailItem that records any call that would change something."""

    Class = mapi.OL_CLASS_MAIL_ITEM
    SenderEmailType = "SMTP"

    def __init__(self, message_id: str, sent_on: datetime, calls: list) -> None:
        self.message_id = message_id
        self.EntryID = f"entry-{message_id}"
        self.Subject = f"Synthetic {message_id}"
        self.SenderEmailAddress = "sender@example.invalid"
        self.Recipients = []
        self.SentOn = sent_on
        self.ReceivedTime = sent_on
        self.Body = "synthetic body"
        self.Attachments = types.SimpleNamespace(Count=0)
        self.PropertyAccessor = types.SimpleNamespace(
            GetProperty=lambda prop: f"<{message_id}>" if prop == mapi.DASL_INTERNET_MESSAGE_ID else 0,
        )
        for name in MUTATIONS:
            setattr(self, name, lambda *a, _n=name: calls.append(_n))


class _Items:
    def __init__(self, mails: list, *, reject: bool = False) -> None:
        self._mails = list(mails)
        self._reject = reject
        self.filters: list[str] = []

    @property
    def Count(self) -> int:  # noqa: N802 - COM's spelling
        return len(self._mails)

    def Item(self, i: int):  # noqa: N802 - COM's spelling
        return self._mails[i - 1]

    def Restrict(self, flt: str) -> "_Items":  # noqa: N802 - COM's spelling
        if self._reject:
            raise RuntimeError("the store rejected the filter")
        self.filters.append(flt)
        return _Items(self._mails)

    def Sort(self, prop: str, descending: bool) -> None:  # noqa: N802 - COM's spelling
        assert prop == "[SentOn]"
        self._mails.sort(key=lambda m: m.SentOn, reverse=descending)


class _Folder:
    def __init__(self, name: str, items: _Items) -> None:
        self.Name = name
        self.Items = items


class _Root:
    def __init__(self, folders: list[_Folder], calls: list) -> None:
        self._folders = folders
        self.Name = "root"
        self.Folders = types.SimpleNamespace(
            Count=len(folders), Item=lambda i: folders[i - 1],
            Add=lambda name: calls.append("Add"),
        )


def _client(monkeypatch, *, inbox=(), archive=None, sent=(), reject=False):
    import win32com.client

    monkeypatch.setattr(win32com.client, "Dispatch", lambda obj: obj)
    calls: list = []
    folders = {
        mapi.OL_FOLDER_INBOX: _Folder("Inbox", _Items([_ComMail(m, d, calls) for m, d in inbox], reject=reject)),
        mapi.OL_FOLDER_SENT_MAIL: _Folder("Sent Mail", _Items([_ComMail(m, d, calls) for m, d in sent])),
    }
    root_folders = [] if archive is None else [
        _Folder("Archive", _Items([_ComMail(m, d, calls) for m, d in archive])),
    ]
    store = types.SimpleNamespace(
        GetDefaultFolder=lambda kind: folders[kind],
        GetRootFolder=lambda: _Root(root_folders, calls),
    )
    client = OutlookClient(Mailbox(alias="second", address="second@example.invalid"))
    monkeypatch.setattr(client, "store", lambda: store)
    return client, calls, folders


def _search(client, folder: str, limit: int = 10, matches=lambda m: True) -> FolderHits:
    return client.search_folder(folder, "Archive", mapi.words_filter(["synthetic"]), matches,
                                preview_len=100, limit=limit)


def test_find_reads_and_never_moves_saves_deletes_or_sends(monkeypatch):
    mails = [("m-1", datetime(2026, 9, 1)), ("m-2", datetime(2026, 9, 3)), ("m-3", datetime(2026, 9, 2))]
    client, calls, folders = _client(monkeypatch, inbox=mails, archive=mails, sent=mails)

    for folder in find.FOLDERS:
        hits = _search(client, folder, limit=2)
        assert hits.via == FIND_VIA_RESTRICT and hits.matched == 3
        assert [m.message_id for m in hits.mails] == ["m-2", "m-3"]  # newest first, limited

    assert calls == []
    assert folders[mapi.OL_FOLDER_INBOX].Items.filters == [mapi.words_filter(["synthetic"])]


def test_a_missing_archive_folder_is_reported_absent_and_never_created(monkeypatch):
    client, calls, _ = _client(monkeypatch, archive=None)

    hits = _search(client, find.FIND_ARCHIVE)

    assert (hits.via, hits.matched, hits.mails) == (FIND_VIA_ABSENT, 0, [])
    assert "Add" not in calls


def test_a_store_that_rejects_the_filter_is_walked_with_the_same_rule(monkeypatch):
    mails = [("keep", datetime(2026, 9, 1)), ("drop", datetime(2026, 9, 2))]
    client, calls, _ = _client(monkeypatch, inbox=mails, reject=True)

    hits = _search(client, find.FIND_INBOX, matches=lambda m: m.message_id == "keep")

    assert hits.via == FIND_VIA_WALK
    assert [m.message_id for m in hits.mails] == ["keep"]
    assert calls == []


# --------------------------------------------------------- orchestration ---

class _FindClient:
    """Stands in for ``search_folder``, one canned result per folder."""

    def __init__(self, per_folder: dict[str, list[InboxMail]]) -> None:
        self.per_folder = per_folder
        self.searched: list[tuple[str, str]] = []

    def search_folder(self, folder, archive_folder, dasl, matches, *, preview_len, limit):
        self.searched.append((folder, archive_folder))
        mails = self.per_folder.get(folder, [])
        return FolderHits(folder=folder, name=folder.title(), via=FIND_VIA_RESTRICT,
                          matched=len(mails), mails=mails[:limit])


def _hit(message_id: str, day: int) -> InboxMail:
    return InboxMail(message_id=message_id, entry_id=f"e-{message_id}", subject="S",
                     sender="a@example.invalid", recipients="b@example.invalid",
                     date_sent=datetime(2026, 9, day), body_preview="p")


def test_find_merges_folders_newest_first_and_marks_what_is_archived(cfg, archive_root):
    _seed_index(cfg, archive_root, {"Folder": ["Synthetic"]})  # Message-ID seed-1@example.invalid
    client = _FindClient({
        "inbox": [_hit("in-inbox", 5)],
        "archive": [_hit("seed-1@example.invalid", 9), _hit("old", 1)],
        "sent": [_hit("sent", 7)],
    })

    doc = find.find(client, cfg, Mailbox(alias="second", address="s@example.invalid",
                                         archive_folder="Filed"), ["word"], limit=3)

    assert [h["message_id"] for h in doc["hits"]] == ["seed-1@example.invalid", "sent", "in-inbox"]
    assert doc["hits"][0]["folder"] == "Archive" and doc["hits"][0]["already_archived"]
    assert doc["hits"][1]["already_archived"] is None
    assert [f for f, _ in client.searched] == ["inbox", "archive", "sent"]
    assert {a for _, a in client.searched} == {"Filed"}
    assert doc["counts"] == {"matched": 4, "returned": 3} and doc["truncated"] is True
    assert doc["coverage"] == "outlook_folders"


def test_folders_narrow_the_search(cfg):
    client = _FindClient({})
    doc = find.find(client, cfg, None, ["word"], folders=("archive",))
    assert [f for f, _ in client.searched] == ["archive"]
    assert doc["query"]["folders"] == ["archive"]


# ------------------------------------------------------ process contract ---

@pytest.mark.parametrize("argv, message", [
    (["find", "--query", "  "], "--query"),
    (["find", "--query", "x", "--folders", "inbox,trash"], "--folders"),
    (["find", "--query", "x", "--limit", "0"], "--limit"),
    (["find", "--query", "x", "--since", "yesterday"], "--since"),
])
def test_bad_find_flags_exit_2_before_outlook(batch_process, argv, message):
    code, doc, built = batch_process(argv, client_factory=lambda mb: pytest.fail("Outlook touched"))

    assert code == 2 and doc["error"]["code"] == "bad_input" and message in doc["error"]["message"]
    assert built == []


def test_find_runs_on_the_named_mailbox(batch_process):
    registry = {
        "schema_version": 1, "default": "owner",
        "mailboxes": {"owner": {"address": "o@example.invalid"}, "second": {"address": "s@example.invalid"}},
    }

    def _factory(mailbox):
        fake = TargetedFakeClient([])
        fake.search_folder = _FindClient({"archive": [_hit("m-1", 2)]}).search_folder
        return fake

    code, doc, _ = batch_process(["find", "--mailbox", "second", "--query", "word"],
                                 registry=registry, client_factory=_factory)

    assert code == 0, doc
    assert doc["mailbox"] == "second" and [h["message_id"] for h in doc["hits"]] == ["m-1"]

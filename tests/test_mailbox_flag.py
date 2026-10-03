"""Tests for acting on a named mailbox: --mailbox, drafting and sending as it (issue #109).

Filing a second mailbox reads its Inbox and moves into its own archive folder;
a draft made for it sends as its account, blind-copies its address and sits in
its own Drafts; ``send`` refuses a draft whose sending account is not the
mailbox's. Drafting and sending as the second mailbox are exercised on COM
fakes only. Everything here is synthetic.
"""
from __future__ import annotations

import json
import types

import pytest

from email_archiver import draft, send
from email_archiver.config import Mailbox, parse_mailbox_registry
from email_archiver.outlook import mapi, process
from email_archiver.outlook.client import OutlookClient
from email_archiver.outlook.drafts import (
    SENDING_ACCOUNT_MISMATCH,
    CreatedDraft,
    DraftAccountError,
)
from email_archiver.outlook.mapi import AccountNotInOutlookError
from tests.test_batch import _mail, archive_root, cfg  # noqa: F401 - fixtures are used by name
from tests.test_batch_targeted import TargetedFakeClient
from tests.test_draft import FakeDraftClient, _FakeApplication, _FakeComMail
from tests.test_mailboxes import batch_process  # noqa: F401 - fixture is used by name
from tests.test_send import _FakeDraft

OWNER = "owner@example.invalid"
SECOND = "second@example.invalid"
SECOND_MAILBOX = Mailbox(alias="second", address=SECOND)
REGISTRY = {
    "schema_version": 1,
    "default": "owner",
    "mailboxes": {
        "owner": {"address": OWNER, "aliases": ["me"]},
        "second": {"address": SECOND, "aliases": ["household"], "archive_folder": "Filed"},
    },
}


class _Account:
    def __init__(self, smtp: str) -> None:
        self.SmtpAddress = smtp


class _Accounts:
    def __init__(self, *smtps: str) -> None:
        self._items = [_Account(s) for s in smtps]
        self.Count = len(self._items)

    def Item(self, i: int) -> _Account:  # noqa: N802 - COM's spelling
        return self._items[i - 1]


class _AccountMail(_FakeComMail):
    """A new MailItem that records its sending account and where it is saved.

    ``accept_account=False`` is an Outlook that takes the assignment without
    error and changes nothing, the silent failure the read-back exists for.
    """

    def __init__(self, *, accept_account: bool = True, parent: str = "mailbox-drafts") -> None:
        super().__init__()
        self._account = None
        self._accept = accept_account
        self.Parent = types.SimpleNamespace(EntryID=parent)

    @property
    def SendUsingAccount(self):  # noqa: N802 - COM's spelling
        return self._account

    @SendUsingAccount.setter
    def SendUsingAccount(self, account) -> None:  # noqa: N802 - COM's spelling
        self.calls.append(("SendUsingAccount=", account.SmtpAddress))
        if self._accept:
            self._account = account

    def Close(self, mode: int) -> None:  # noqa: N802 - COM's spelling
        self.calls.append(("Close", mode))

    def Move(self, folder):  # noqa: N802 - COM's spelling
        self.calls.append(("Move", folder.EntryID))
        self.Parent = folder
        return self


def _drafting_client(monkeypatch, mail, mailbox=SECOND_MAILBOX, accounts=(OWNER, SECOND)):
    """A client bound to ``mailbox`` whose Outlook is entirely fake."""
    app = _FakeApplication(mail)
    app.Inspectors = types.SimpleNamespace(Count=0)
    monkeypatch.setattr(process, "get_active_application", lambda: app)
    client = OutlookClient(mailbox)
    monkeypatch.setattr(client, "_namespace", lambda: types.SimpleNamespace(Accounts=_Accounts(*accounts)))
    store = types.SimpleNamespace(
        GetDefaultFolder=lambda kind: types.SimpleNamespace(EntryID="mailbox-drafts"),
    )
    monkeypatch.setattr(client, "store", lambda: store)
    return client, app


def _create(client: OutlookClient, **overrides) -> CreatedDraft:
    kwargs = dict(
        to=["a@example.invalid"], cc=None, bcc=[SECOND], subject="Synthetic",
        body_html="<p>Hello</p>", attachments=[], ref=None, display=True,
    )
    kwargs.update(overrides)
    return client.create_draft(**kwargs)


# --------------------------------------------------------- drafting as X ---

def test_a_draft_for_a_mailbox_sends_as_its_account_from_the_start(monkeypatch):
    mail = _AccountMail()
    client, _ = _drafting_client(monkeypatch, mail)

    created = _create(client)

    names = [c[0] for c in mail.calls]
    # Before Display, so the signature Outlook inserts is that account's.
    assert names.index("SendUsingAccount=") < names.index("Display") < names.index("Save")
    assert mail.SendUsingAccount.SmtpAddress == SECOND
    assert created.from_address == SECOND
    assert "Move" not in names  # already in the mailbox's own Drafts


def test_a_draft_saved_outside_the_mailbox_drafts_is_moved_into_them(monkeypatch):
    mail = _AccountMail(parent="default-store-drafts")
    client, _ = _drafting_client(monkeypatch, mail)

    _create(client, display=False)

    assert mail.calls[-1] == ("Move", "mailbox-drafts")
    assert mail.Parent.EntryID == "mailbox-drafts"


def test_an_account_outlook_will_not_set_discards_the_draft_unsaved(monkeypatch):
    mail = _AccountMail(accept_account=False)
    client, _ = _drafting_client(monkeypatch, mail)

    with pytest.raises(DraftAccountError) as refused:
        _create(client)

    assert refused.value.code == SENDING_ACCOUNT_MISMATCH
    names = [c[0] for c in mail.calls]
    assert ("Close", mapi.OL_DISCARD) in mail.calls
    assert "Save" not in names and "Display" not in names


def test_a_mailbox_with_no_account_creates_nothing(monkeypatch):
    mail = _AccountMail()
    client, app = _drafting_client(monkeypatch, mail, accounts=(OWNER,))

    with pytest.raises(AccountNotInOutlookError):
        _create(client)

    assert app.created == []
    assert mail.calls == []


def test_with_no_registry_the_sending_account_is_left_to_outlook(monkeypatch):
    mail = _AccountMail()
    client, _ = _drafting_client(monkeypatch, mail, mailbox=None)

    _create(client)

    assert "SendUsingAccount=" not in [c[0] for c in mail.calls]


def test_a_registry_mailbox_blind_copies_its_own_address():
    cfg = {"outlook": {"self_address": "configured@example.invalid"}}

    assert draft.resolve_self_address(cfg, FakeDraftClient(), SECOND_MAILBOX) == SECOND
    synthesized = Mailbox(alias="default", synthesized=True)
    assert draft.resolve_self_address(cfg, FakeDraftClient(), synthesized) == "configured@example.invalid"


# --------------------------------------------------------- sending as X ---

def _sending_client(monkeypatch, mail: _FakeDraft) -> OutlookClient:
    client = OutlookClient(SECOND_MAILBOX)
    monkeypatch.setattr(client, "_namespace", lambda: types.SimpleNamespace(
        GetItemFromID=lambda entry_id: mail,
    ))
    monkeypatch.setattr(client, "store", lambda: types.SimpleNamespace(
        GetDefaultFolder=lambda kind: types.SimpleNamespace(EntryID="drafts"),
    ))
    app = types.SimpleNamespace(Inspectors=types.SimpleNamespace(Count=0))
    monkeypatch.setattr(process, "get_active_application", lambda: app)
    return client


def _approval(monkeypatch, mail: _FakeDraft) -> dict:
    return send.fingerprint(_sending_client(monkeypatch, mail).read_draft(mail.EntryID))


def test_a_draft_in_the_mailbox_drafts_sending_as_it_is_read_and_sent(monkeypatch):
    mail = _FakeDraft()
    mail.SendUsingAccount = _Account(SECOND)
    approval = _approval(monkeypatch, mail)
    client = _sending_client(monkeypatch, mail)

    assert send.read_document(client.read_draft(mail.EntryID))["from_address"] == SECOND
    document = send.send(client, mail.EntryID, approval["hash"], approval["parts"], None,
                         expect_account=SECOND)

    assert mail.calls == [("Send",)]
    assert document["from_address"] == SECOND


@pytest.mark.parametrize("account", [_Account(OWNER), None], ids=["changed", "unset"])
def test_a_draft_not_sending_as_the_mailbox_is_refused_and_not_sent(monkeypatch, account):
    mail = _FakeDraft()
    mail.SendUsingAccount = _Account(SECOND)
    approval = _approval(monkeypatch, mail)
    mail.SendUsingAccount = account  # changed by hand in Outlook after approval
    client = _sending_client(monkeypatch, mail)

    with pytest.raises(send.SendRefused) as refused:
        send.send(client, mail.EntryID, approval["hash"], approval["parts"], None,
                  expect_account=SECOND)

    assert refused.value.code == SENDING_ACCOUNT_MISMATCH
    assert mail.calls == []


# ------------------------------------------------------ process contract ---

def test_an_unknown_mailbox_exits_2_before_outlook(batch_process):
    code, doc, built = batch_process(
        ["plan", "--mailbox", "nobody"], registry=REGISTRY,
        client_factory=lambda mb: pytest.fail("Outlook touched"),
    )

    assert doc["error"]["code"] == "mailbox_unknown"
    assert code == 2 and built == []


def test_plan_and_apply_act_on_the_named_mailbox(batch_process, archive_root, tmp_path):
    seen: list[Mailbox] = []

    def _factory(mailbox):
        seen.append(mailbox)
        return TargetedFakeClient([_mail("m-1", "Synthetic subject")])

    code, doc, _ = batch_process(["plan", "--mailbox", "HOUSEHOLD", "--candidates", "0"],
                                 registry=REGISTRY, client_factory=_factory)
    assert code == 0 and doc["mailbox"] == "second"

    decisions = tmp_path / "decisions.json"
    decisions.write_text(json.dumps([
        {"message_id": "m-1", "folder_path": str(archive_root), "date_prefix": False},
    ]), encoding="utf-8")
    code, doc, built = batch_process(["apply", "--mailbox", "second", "--decisions", str(decisions)],
                                     client_factory=_factory)
    assert code == 0, doc
    assert (doc["mailbox"], doc["archive_folder"]) == ("second", "Filed")
    assert built[-1].moves == [("m-1", "Filed")]
    assert [m.alias for m in seen] == ["second", "second"]


def test_no_mailbox_flag_is_the_registry_default(batch_process):
    code, doc, _ = batch_process(
        ["plan", "--candidates", "0"], registry=REGISTRY,
        client_factory=lambda mb: TargetedFakeClient([]),
    )
    assert code == 0 and doc["mailbox"] == "owner"


def test_draft_as_a_mailbox_blind_copies_it_and_reports_it(batch_process, tmp_path):
    spec = tmp_path / "spec.json"
    spec.write_text('{"to": ["a@example.invalid"], "subject": "S", "body_text": "Hi"}', encoding="utf-8")
    fakes: list[FakeDraftClient] = []

    def _factory(mailbox):
        fakes.append(FakeDraftClient(account_smtp=OWNER))
        return fakes[-1]

    code, doc, _ = batch_process(["draft", "--mailbox", "second", "--spec", str(spec)],
                                 registry=REGISTRY, client_factory=_factory)

    assert code == 0, doc
    assert doc["mailbox"] == "second" and SECOND in doc["bcc"] and OWNER not in doc["bcc"]


def test_a_mailbox_with_no_account_exits_2_with_its_own_code(batch_process, tmp_path):
    spec = tmp_path / "spec.json"
    spec.write_text('{"to": ["a@example.invalid"], "subject": "S", "body_text": "Hi"}', encoding="utf-8")

    def _factory(mailbox):
        fake = FakeDraftClient()
        fake.refuse_create = AccountNotInOutlookError("no account for this mailbox")
        return fake

    code, doc, _ = batch_process(["draft", "--mailbox", "second", "--spec", str(spec)],
                                 registry=REGISTRY, client_factory=_factory)

    assert code == 2 and doc["error"]["code"] == "account_not_in_outlook"


def test_the_registry_names_resolve_case_insensitively():
    registry = parse_mailbox_registry(REGISTRY)
    assert registry.select("Household") is registry.select("second")

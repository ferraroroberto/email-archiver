"""Tests for the headless `draft` verb (issue #70).

`main_batch.py draft` is spawned by another local app to turn a written email
into an Outlook draft the user reviews and sends by hand. These tests pin what
that caller depends on:

- a bad spec is refused as ``bad_input`` before Outlook is ever started;
- the sender's own address is always on the BCC line, exactly once, and an
  address that cannot be resolved stops the run before a draft exists;
- the document's shape, including an honest ``ref_header``;
- nothing ever sends: the fake COM mail item records a ``Send`` call, and the
  draft modules carry no call to ``Send`` at all.

Every address, path and subject here is synthetic.
"""
from __future__ import annotations

import json
import re
import sys
import types
from pathlib import Path

import pytest

import main_batch
from email_archiver import draft
from email_archiver.outlook import client as client_mod
from email_archiver.outlook.client import (
    DASL_X_ARCHIVE_REF,
    CreatedDraft,
    OutlookClient,
    insert_body_html,
)

REPO_ROOT = Path(__file__).resolve().parents[1]
SELF = "me@example.invalid"


@pytest.fixture
def attachment(tmp_path) -> Path:
    path = tmp_path / "note.txt"
    path.write_text("synthetic attachment", encoding="utf-8")
    return path


def _spec(**overrides) -> dict:
    spec = {"to": ["someone@example.invalid"], "subject": "A subject", "body_text": "Hello"}
    spec.update(overrides)
    return spec


# --------------------------------------------------------------- fake COM ---

class _FakePropertyAccessor:
    def __init__(self, calls: list, refuse: bool) -> None:
        self._calls = calls
        self._refuse = refuse

    def SetProperty(self, name: str, value: str) -> None:  # noqa: N802 - COM's spelling
        if self._refuse:
            raise RuntimeError("the store refused the property")
        self._calls.append(("SetProperty", name, value))


class _FakeComAttachments:
    def __init__(self, calls: list) -> None:
        self._calls = calls

    def Add(self, path: str) -> None:  # noqa: N802 - COM's spelling
        self._calls.append(("Attachments.Add", path))


class _FakeComMail:
    """A new MailItem that logs every call, in order, and every body write.

    ``Display`` puts a signature into ``HTMLBody``, the way Outlook does when it
    opens the compose window for an account with a default signature.
    """

    SIGNATURE_HTML = '<html><body lang=EN><div id="_MailAutoSig">-- sig</div></body></html>'

    def __init__(self, refuse_property: bool = False) -> None:
        self.calls: list = []
        self.To = self.CC = self.BCC = self.Subject = ""
        self._html = ""
        self.EntryID = "draft-entry-1"
        self.Attachments = _FakeComAttachments(self.calls)
        self.PropertyAccessor = _FakePropertyAccessor(self.calls, refuse_property)

    @property
    def HTMLBody(self) -> str:  # noqa: N802 - COM's spelling
        return self._html

    @HTMLBody.setter
    def HTMLBody(self, value: str) -> None:  # noqa: N802 - COM's spelling
        self.calls.append(("HTMLBody=",))
        self._html = value

    def Display(self, modal: bool) -> None:  # noqa: N802 - COM's spelling
        self.calls.append(("Display", modal))
        self._html = self.SIGNATURE_HTML

    def Save(self) -> None:  # noqa: N802 - COM's spelling
        self.calls.append(("Save",))

    def Send(self) -> None:  # noqa: N802 - COM's spelling
        self.calls.append(("Send",))


class _FakeApplication:
    def __init__(self, mail: _FakeComMail) -> None:
        self.mail = mail
        self.created: list[int] = []

    def CreateItem(self, kind: int) -> _FakeComMail:  # noqa: N802 - COM's spelling
        self.created.append(kind)
        return self.mail


class FakeDraftClient:
    """Stands in for OutlookClient's draft surface."""

    def __init__(self, account_smtp: str = SELF, ref_stamped: bool = True) -> None:
        self.account_smtp = account_smtp
        self.ref_stamped = ref_stamped
        self.account_lookups = 0
        self.drafts: list[dict] = []

    def ensure_running(self, timeout: float = 60.0) -> None:
        return None

    def default_account_smtp(self) -> str:
        self.account_lookups += 1
        return self.account_smtp

    def create_draft(self, **kwargs) -> CreatedDraft:
        self.drafts.append(kwargs)
        return CreatedDraft(
            entry_id="draft-entry-1",
            ref_stamped=self.ref_stamped and bool(kwargs["ref"]),
            ref_reason="" if self.ref_stamped else "SetProperty refused: synthetic",
            displayed=kwargs["display"],
        )


# ------------------------------------------------------------------- spec ---

def test_a_valid_spec_resolves_attachments_to_absolute_paths(attachment, monkeypatch):
    monkeypatch.chdir(attachment.parent)
    spec = draft.parse_spec(_spec(attachments=[attachment.name], ref=" tok ", cc=["c@example.invalid"]))

    assert spec.attachments == [str(attachment.resolve())]
    assert spec.cc == ["c@example.invalid"]
    assert spec.ref == "tok"
    assert spec.display is True


@pytest.mark.parametrize(
    ("overrides", "message"),
    [
        ({"to": []}, "`to` needs at least one address"),
        ({"to": "someone@example.invalid"}, "`to` must be a list"),
        ({"cc": [" "]}, "`cc` contains a blank address"),
        ({"subject": "  "}, "`subject` must be a non-empty string"),
        ({"body_html": "<p>x</p>"}, "exactly one of `body_text` or `body_html`"),
        ({"attachments": ["Z:/definitely/not/here.pdf"]}, "attachment is not an existing file"),
        ({"ref": ""}, "`ref` must be a non-empty string"),
        ({"display": "yes"}, "`display` must be true or false"),
        ({"atachments": []}, "unknown spec keys: atachments"),
    ],
)
def test_an_unusable_spec_is_refused_with_the_field_named(overrides, message):
    with pytest.raises(draft.SpecError, match=re.escape(message)):
        draft.parse_spec(_spec(**overrides))


def test_a_spec_with_no_body_is_refused():
    data = _spec()
    del data["body_text"]
    with pytest.raises(draft.SpecError, match="exactly one of"):
        draft.parse_spec(data)


def test_a_directory_is_not_an_attachment(tmp_path):
    with pytest.raises(draft.SpecError, match="not an existing file"):
        draft.parse_spec(_spec(attachments=[str(tmp_path)]))


def test_a_plain_text_body_becomes_escaped_paragraphs():
    html = draft.text_to_html("Dear <team>,\r\nline two\n\n\nA & B\n")

    assert html == "<p>Dear &lt;team&gt;,<br>line two</p><p>A &amp; B</p>"


def test_an_html_body_is_passed_through_untouched():
    data = _spec()
    del data["body_text"]
    data["body_html"] = "<b>as written</b>"
    assert draft.parse_spec(data).body_html == "<b>as written</b>"


# ----------------------------------------------------------- self address ---

def test_the_configured_self_address_wins_and_outlook_is_not_asked():
    client = FakeDraftClient(account_smtp="other@example.invalid")
    cfg = {"outlook": {"self_address": " configured@example.invalid "}}

    assert draft.resolve_self_address(cfg, client) == "configured@example.invalid"
    assert client.account_lookups == 0


def test_the_default_account_is_the_fallback():
    assert draft.resolve_self_address({"outlook": {"self_address": ""}}, FakeDraftClient()) == SELF


@pytest.mark.parametrize("reported", ["", "/o=Exchange/cn=Recipients/cn=someone"])
def test_something_that_is_not_an_smtp_address_does_not_resolve(reported):
    assert draft.resolve_self_address({}, FakeDraftClient(account_smtp=reported)) is None


def test_self_is_blind_copied_exactly_once_whatever_the_spelling():
    assert draft.with_self_bcc([], SELF) == [SELF]
    assert draft.with_self_bcc(["x@example.invalid", "ME@example.invalid"], SELF) == [
        "x@example.invalid", "ME@example.invalid",
    ]


# ------------------------------------------------------------- document ---

def test_the_document_carries_the_draft_with_self_in_bcc(attachment):
    client = FakeDraftClient()
    spec = draft.parse_spec(_spec(
        bcc=["hidden@example.invalid"], attachments=[str(attachment)], ref="tok-1",
    ))

    doc = draft.create(client, spec, SELF)

    assert doc["verb"] == "draft"
    assert doc["schema_version"] == 1
    assert doc["entry_id"] == "draft-entry-1"
    assert doc["subject"] == "A subject"
    assert doc["to"] == ["someone@example.invalid"]
    assert doc["cc"] == []
    assert doc["bcc"] == ["hidden@example.invalid", SELF]
    assert doc["attachments"] == [str(attachment.resolve())]
    assert doc["ref"] == "tok-1"
    assert doc["ref_header"] == "stamped"
    assert doc["ref_header_reason"] == ""
    assert doc["displayed"] is True
    assert doc["created_at"]
    assert client.drafts[0]["bcc"] == ["hidden@example.invalid", SELF]
    assert client.drafts[0]["body_html"] == "<p>Hello</p>"


def test_a_ref_that_could_not_be_stamped_is_reported_with_its_reason():
    doc = draft.create(FakeDraftClient(ref_stamped=False), draft.parse_spec(_spec(ref="tok")), SELF)

    assert doc["ref_header"] == "not_stamped"
    assert doc["ref_header_reason"].startswith("SetProperty refused")


def test_no_ref_is_reported_as_not_stamped():
    doc = draft.create(FakeDraftClient(), draft.parse_spec(_spec()), SELF)

    assert (doc["ref"], doc["ref_header"], doc["ref_header_reason"]) == (
        None, "not_stamped", "no ref given",
    )


# ------------------------------------------------------ client, fake COM ---

def _client_with(monkeypatch, mail: _FakeComMail) -> tuple[OutlookClient, _FakeApplication]:
    app = _FakeApplication(mail)
    monkeypatch.setattr(client_mod, "_get_active_application", lambda: app)
    return OutlookClient(), app


def _create(client: OutlookClient, **overrides) -> CreatedDraft:
    kwargs = dict(
        to=["a@example.invalid"], cc=["c@example.invalid"], bcc=[SELF, "b@example.invalid"],
        subject="S", body_html="<p>Body</p>", attachments=["C:/synthetic/file.pdf"],
        ref="tok", display=True,
    )
    kwargs.update(overrides)
    return client.create_draft(**kwargs)


def test_create_draft_fills_saves_and_shows_the_mail_but_never_sends(monkeypatch):
    mail = _FakeComMail()
    client, app = _client_with(monkeypatch, mail)

    created = _create(client)

    assert app.created == [client_mod.OL_MAIL_ITEM]
    assert (mail.To, mail.CC, mail.BCC, mail.Subject) == (
        "a@example.invalid", "c@example.invalid", f"{SELF}; b@example.invalid", "S",
    )
    assert mail.calls == [
        ("Attachments.Add", "C:/synthetic/file.pdf"),
        ("SetProperty", DASL_X_ARCHIVE_REF, "tok"),
        ("Display", False),
        ("HTMLBody=",),
        ("Save",),
    ]
    assert ("Send",) not in mail.calls
    assert created == CreatedDraft(
        entry_id="draft-entry-1", ref_stamped=True, ref_reason="", displayed=True,
    )


def test_the_body_goes_above_the_signature_outlook_inserted(monkeypatch):
    mail = _FakeComMail()
    client, _ = _client_with(monkeypatch, mail)

    _create(client)

    assert mail.HTMLBody == (
        '<html><body lang=EN><p>Body</p><div id="_MailAutoSig">-- sig</div></body></html>'
    )


def test_an_undisplayed_draft_is_still_saved_with_its_body(monkeypatch):
    mail = _FakeComMail()
    client, _ = _client_with(monkeypatch, mail)

    created = _create(client, display=False, ref=None)

    assert mail.calls == [("Attachments.Add", "C:/synthetic/file.pdf"), ("HTMLBody=",), ("Save",)]
    assert mail.HTMLBody == "<html><body><p>Body</p></body></html>"
    assert (created.displayed, created.ref_stamped, created.ref_reason) == (False, False, "")


def test_a_refused_ref_header_does_not_stop_the_draft(monkeypatch):
    mail = _FakeComMail(refuse_property=True)
    client, _ = _client_with(monkeypatch, mail)

    created = _create(client)

    assert created.ref_stamped is False
    assert "the store refused the property" in created.ref_reason
    assert ("Save",) in mail.calls


def test_the_body_insert_is_case_insensitive_and_keeps_body_attributes():
    assert insert_body_html("<HTML><BODY class=x>sig</BODY></HTML>", "<p>b</p>") == (
        "<HTML><BODY class=x><p>b</p>sig</BODY></HTML>"
    )


def test_no_draft_code_path_calls_send():
    # Spelled in two halves so a grep of the repo for the call finds nothing,
    # this test included.
    needle = "." + "Send("
    for relative in ("email_archiver/draft.py", "email_archiver/outlook/client.py", "main_batch.py"):
        assert needle not in (REPO_ROOT / relative).read_text(encoding="utf-8"), relative


# ------------------------------------------- client, default account smtp ---

class _FakeStore:
    def __init__(self, store_id: str) -> None:
        self.StoreID = store_id


class _FakeAccount:
    def __init__(self, smtp: str, store_id: str) -> None:
        self.SmtpAddress = smtp
        self.DeliveryStore = _FakeStore(store_id)


class _FakeAccounts:
    def __init__(self, accounts: list[_FakeAccount]) -> None:
        self._accounts = accounts
        self.Count = len(accounts)

    def Item(self, i: int) -> _FakeAccount:  # noqa: N802 - COM's spelling
        return self._accounts[i - 1]


def _client_with_accounts(monkeypatch, accounts, default_store="store-b") -> OutlookClient:
    namespace = types.SimpleNamespace(
        Accounts=_FakeAccounts(accounts), DefaultStore=_FakeStore(default_store),
    )
    client = OutlookClient()
    monkeypatch.setattr(client, "_namespace", lambda: namespace)
    return client


def test_the_account_owning_the_default_store_is_the_sender(monkeypatch):
    client = _client_with_accounts(monkeypatch, [
        _FakeAccount("a@example.invalid", "store-a"), _FakeAccount("b@example.invalid", "store-b"),
    ])
    assert client.default_account_smtp() == "b@example.invalid"


def test_a_single_account_needs_no_store_match(monkeypatch):
    client = _client_with_accounts(monkeypatch, [_FakeAccount("a@example.invalid", "x")])
    assert client.default_account_smtp() == "a@example.invalid"


def test_several_accounts_with_no_default_match_resolve_to_nothing(monkeypatch):
    client = _client_with_accounts(monkeypatch, [
        _FakeAccount("a@example.invalid", "x"), _FakeAccount("b@example.invalid", "y"),
    ])
    assert client.default_account_smtp() == ""


# ------------------------------------------------------ process contract ---

@pytest.fixture
def batch_process(tmp_path, monkeypatch, capsys):
    """``main_batch.main`` with a temp config, no log file, a stub pythoncom
    and an Outlook client that records whether it was ever constructed."""
    cfg = {"outlook": {}}
    monkeypatch.setattr(main_batch, "load_config", lambda: cfg)
    monkeypatch.setattr(main_batch, "setup_logging", lambda _cfg: None)
    monkeypatch.setitem(sys.modules, "pythoncom", types.SimpleNamespace(
        CoInitialize=lambda: None, CoUninitialize=lambda: None,
    ))
    clients: list[FakeDraftClient] = []

    def run(spec: dict, account_smtp: str = SELF) -> tuple[int, dict, list[FakeDraftClient]]:
        def _factory() -> FakeDraftClient:
            fake = FakeDraftClient(account_smtp=account_smtp)
            clients.append(fake)
            return fake

        monkeypatch.setattr(main_batch, "OutlookClient", _factory)
        spec_path = tmp_path / "spec.json"
        spec_path.write_text(json.dumps(spec), encoding="utf-8")
        code = main_batch.main(["draft", "--spec", str(spec_path)])
        return code, json.loads(capsys.readouterr().out), clients

    return run


@pytest.mark.parametrize(
    "overrides",
    [
        {"attachments": ["Z:/definitely/not/here.pdf"]},
        {"to": []},
        {"body_html": "<p>both</p>"},
    ],
)
def test_a_bad_spec_exits_2_without_starting_outlook(batch_process, overrides):
    code, doc, clients = batch_process(_spec(**overrides))

    assert code == main_batch.EXIT_CANNOT_START
    assert doc["error"]["code"] == "bad_input"
    assert clients == []


def test_an_unresolved_self_address_exits_2_and_creates_no_draft(batch_process):
    code, doc, clients = batch_process(_spec(), account_smtp="")

    assert code == main_batch.EXIT_CANNOT_START
    assert doc["error"]["code"] == "self_address_unresolved"
    assert len(clients) == 1 and clients[0].drafts == []


def test_a_good_spec_prints_the_draft_document(batch_process):
    code, doc, clients = batch_process(_spec(ref="tok"))

    assert code == main_batch.EXIT_OK
    assert doc["verb"] == "draft" and doc["bcc"] == [SELF]
    assert doc["ref_header"] == "stamped"
    assert len(clients[0].drafts) == 1

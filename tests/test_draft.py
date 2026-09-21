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
from email_archiver.outlook import mapi, process
from email_archiver.outlook.client import OutlookClient
from email_archiver.outlook.drafts import CreatedDraft, DraftUpdateError
from email_archiver.outlook.mapi import (
    DASL_X_ARCHIVE_REF,
    insert_body_html,
    mark_body_html,
    replace_marked_body_html,
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
        self.paths: list[str] = []

    @property
    def Count(self) -> int:  # noqa: N802 - COM's spelling
        return len(self.paths)

    def Add(self, path: str) -> None:  # noqa: N802 - COM's spelling
        self._calls.append(("Attachments.Add", path))
        self.paths.append(path)

    def Remove(self, index: int) -> None:  # noqa: N802 - COM's spelling
        self._calls.append(("Attachments.Remove", index))
        del self.paths[index - 1]


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
        self.updates: list[dict] = []

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

    refuse_update: DraftUpdateError | None = None

    def update_draft(self, **kwargs) -> CreatedDraft:
        if self.refuse_update is not None:
            raise self.refuse_update
        self.updates.append(kwargs)
        return CreatedDraft(entry_id=kwargs["entry_id"], ref_stamped=bool(kwargs["ref"]),
                            displayed=kwargs["display"])


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
    assert doc["updated"] is False
    assert doc["created_at"] and "updated_at" not in doc
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
    monkeypatch.setattr(process, "get_active_application", lambda: app)
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

    assert app.created == [mapi.OL_MAIL_ITEM]
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
        '<html><body lang=EN><div id="archive-draft-body"><p>Body</p></div><!--/archive-draft-body-->'
        '<div id="_MailAutoSig">-- sig</div></body></html>'
    )


def test_an_undisplayed_draft_is_still_saved_with_its_body(monkeypatch):
    mail = _FakeComMail()
    client, _ = _client_with(monkeypatch, mail)

    created = _create(client, display=False, ref=None)

    assert mail.calls == [("Attachments.Add", "C:/synthetic/file.pdf"), ("HTMLBody=",), ("Save",)]
    assert mail.HTMLBody == f"<html><body>{mark_body_html('<p>Body</p>')}</body></html>"
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
    for relative in (
        "email_archiver/draft.py",
        "email_archiver/outlook/client.py",
        "email_archiver/outlook/drafts.py",
        "main_batch.py",
    ):
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

    def run(
        spec: dict, account_smtp: str = SELF, extra: tuple[str, ...] = (),
        refuse_update: DraftUpdateError | None = None,
    ) -> tuple[int, dict, list[FakeDraftClient]]:
        def _factory() -> FakeDraftClient:
            fake = FakeDraftClient(account_smtp=account_smtp)
            fake.refuse_update = refuse_update
            clients.append(fake)
            return fake

        monkeypatch.setattr(main_batch, "OutlookClient", _factory)
        spec_path = tmp_path / "spec.json"
        spec_path.write_text(json.dumps(spec), encoding="utf-8")
        code = main_batch.main(["draft", "--spec", str(spec_path), *extra])
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


# ------------------------------------------------------------ update (#80) ---

SIG = '<div id="_MailAutoSig">-- sig</div>'


def _marked(body: str) -> str:
    return f"<html><body lang=EN>{mark_body_html(body)}{SIG}</body></html>"


def test_the_marked_body_is_replaced_and_the_signature_kept():
    assert replace_marked_body_html(_marked("<p>old</p>"), "<p>new</p>") == _marked("<p>new</p>")


def test_a_body_with_its_own_nested_divs_is_replaced_whole():
    existing = _marked("<div><div>deep</div></div><p>tail of old</p>")

    assert replace_marked_body_html(existing, "<p>new</p>") == _marked("<p>new</p>")


def test_the_marker_is_found_after_outlook_requotes_it():
    existing = "<body><DIV ID=archive-draft-body><p>old</p></DIV>\n<!-- /archive-draft-body -->sig</body>"

    assert replace_marked_body_html(existing, "<p>new</p>") == f"<body>{mark_body_html('<p>new</p>')}sig</body>"


@pytest.mark.parametrize("existing", [
    "", None, f"<html><body><p>no marker</p>{SIG}</body></html>",
    '<html><body><div id="archive-draft-body"><p>opened, never closed</p></body></html>',
])
def test_an_unmarked_body_has_no_replacement(existing):
    assert replace_marked_body_html(existing, "<p>new</p>") is None


class _FakeFolder:
    def __init__(self, entry_id: str) -> None:
        self.EntryID = entry_id


class _FakeExistingMail(_FakeComMail):
    def __init__(self, html: str, sent: bool = False, parent: str = "drafts") -> None:
        super().__init__()
        self._html = html
        self.Sent = sent
        self.Parent = _FakeFolder(parent)
        self.Attachments.paths = ["C:/synthetic/old-1.pdf", "C:/synthetic/old-2.pdf"]
        self.To, self.Subject = "old@example.invalid", "Old subject"

    def Display(self, modal: bool) -> None:  # noqa: N802 - COM's spelling
        # Reopening a saved draft shows it as it is; the signature is inserted
        # only when a new item is first displayed.
        self.calls.append(("Display", modal))


class _FakeNamespace:
    def __init__(self, mail: _FakeExistingMail | None) -> None:
        self.mail = mail

    def GetItemFromID(self, entry_id: str):  # noqa: N802 - COM's spelling
        if self.mail is None or entry_id != self.mail.EntryID:
            raise RuntimeError("The operation failed. An object could not be found.")
        return self.mail

    def GetDefaultFolder(self, kind: int) -> _FakeFolder:  # noqa: N802 - COM's spelling
        assert kind == mapi.OL_FOLDER_DRAFTS
        return _FakeFolder("drafts")


class _FakeInspector:
    def __init__(self, item, on_close=None) -> None:
        self.CurrentItem = item
        self.closed_with: list[int] = []
        self._on_close = on_close

    def Close(self, mode: int) -> None:  # noqa: N802 - COM's spelling
        self.closed_with.append(mode)
        if self._on_close:
            self._on_close()


class _FakeInspectors:
    def __init__(self, inspectors: list[_FakeInspector]) -> None:
        self._inspectors = inspectors

    @property
    def Count(self) -> int:  # noqa: N802 - COM's spelling
        return len(self._inspectors)

    def Item(self, i: int) -> _FakeInspector:  # noqa: N802 - COM's spelling
        return self._inspectors[i - 1]


def _update(
    monkeypatch, mail: _FakeExistingMail | None, inspectors: list[_FakeInspector] | None = None, **overrides,
) -> CreatedDraft:
    client = OutlookClient()
    monkeypatch.setattr(client, "_namespace", lambda: _FakeNamespace(mail))
    # Never the real Outlook: the only one on this machine is the user's own.
    app = types.SimpleNamespace(Inspectors=_FakeInspectors(inspectors or []))
    monkeypatch.setattr(process, "get_active_application", lambda: app)
    kwargs = dict(
        entry_id="draft-entry-1", to=["a@example.invalid"], cc=[], bcc=[SELF],
        subject="New subject", body_html="<p>New</p>", attachments=["C:/synthetic/new.pdf"],
        ref="tok", display=False,
    )
    kwargs.update(overrides)
    return client.update_draft(**kwargs)


def test_update_refills_the_same_item_keeps_the_signature_and_never_sends(monkeypatch):
    mail = _FakeExistingMail(_marked("<p>Old</p>"))

    updated = _update(monkeypatch, mail)

    assert (mail.To, mail.BCC, mail.Subject) == ("a@example.invalid", SELF, "New subject")
    assert mail.HTMLBody == _marked("<p>New</p>")
    assert mail.Attachments.paths == ["C:/synthetic/new.pdf"]
    assert mail.calls == [
        ("Attachments.Remove", 1), ("Attachments.Remove", 1),
        ("Attachments.Add", "C:/synthetic/new.pdf"),
        ("SetProperty", DASL_X_ARCHIVE_REF, "tok"),
        ("HTMLBody=",), ("Save",),
    ]
    assert ("Send",) not in mail.calls
    assert updated == CreatedDraft(entry_id="draft-entry-1", ref_stamped=True, displayed=False)


def test_an_updated_draft_can_be_shown_again(monkeypatch):
    mail = _FakeExistingMail(_marked("<p>Old</p>"))

    updated = _update(monkeypatch, mail, display=True)

    assert mail.calls[-2:] == [("Save",), ("Display", False)]
    assert mail.HTMLBody == _marked("<p>New</p>"), "Display after Save must not re-insert anything"
    assert updated.displayed is True


@pytest.mark.parametrize(("state", "entry_id", "code"), [
    ("missing", "draft-entry-1", "draft_not_found"),
    ("marked", "some-other-id", "draft_not_found"),
    ("sent", "draft-entry-1", "draft_not_editable"),
    ("inbox", "draft-entry-1", "draft_not_editable"),
    ("unmarked", "draft-entry-1", "draft_body_unmarked"),
])
def test_a_refused_update_names_why_and_leaves_the_item_untouched(monkeypatch, state, entry_id, code):
    mail = {
        "missing": None,
        "marked": _FakeExistingMail(_marked("<p>Old</p>")),
        "sent": _FakeExistingMail(_marked("<p>Old</p>"), sent=True),
        "inbox": _FakeExistingMail(_marked("<p>Old</p>"), parent="inbox"),
        "unmarked": _FakeExistingMail(f"<html><body><p>Old</p>{SIG}</body></html>"),
    }[state]
    before = None if mail is None else (mail.To, mail.Subject, mail.HTMLBody, list(mail.Attachments.paths))

    with pytest.raises(DraftUpdateError) as caught:
        _update(monkeypatch, mail, entry_id=entry_id)

    assert caught.value.code == code
    if mail is not None:
        assert mail.calls == []
        assert (mail.To, mail.Subject, mail.HTMLBody, list(mail.Attachments.paths)) == before


def test_the_update_document_says_it_updated():
    client = FakeDraftClient()
    spec = draft.parse_spec(_spec(ref="tok", bcc=["x@example.invalid"]))

    doc = draft.update(client, "draft-entry-1", spec, SELF)

    assert doc["verb"] == "draft" and doc["updated"] is True
    assert doc["entry_id"] == "draft-entry-1"
    assert doc["updated_at"] and "created_at" not in doc
    assert doc["bcc"] == ["x@example.invalid", SELF]
    assert client.drafts == [] and client.updates[0]["entry_id"] == "draft-entry-1"


def test_update_on_the_command_line_edits_rather_than_creates(batch_process):
    code, doc, clients = batch_process(_spec(ref="tok"), extra=("--update", " draft-entry-1 "))

    assert code == main_batch.EXIT_OK
    assert doc["updated"] is True
    assert clients[0].drafts == [] and clients[0].updates[0]["entry_id"] == "draft-entry-1"


def test_a_blank_update_id_is_bad_input_before_outlook(batch_process):
    code, doc, clients = batch_process(_spec(), extra=("--update", "  "))

    assert code == main_batch.EXIT_CANNOT_START
    assert doc["error"]["code"] == "bad_input"
    assert clients == []


def test_a_refused_update_exits_2_with_its_own_code(batch_process):
    refusal = DraftUpdateError("draft_body_unmarked", "no marked body region")

    code, doc, _ = batch_process(_spec(), extra=("--update", "draft-entry-1"), refuse_update=refusal)

    assert code == main_batch.EXIT_CANNOT_START
    assert doc["error"] == {"code": "draft_body_unmarked", "message": "no marked body region"}


def test_an_open_window_on_the_draft_is_saved_and_closed_before_the_update(monkeypatch):
    # The window's copy is saved first (here: the user typed a line), and the
    # update replaces the body of what it saved, not of the stale reference.
    mail = _FakeExistingMail(_marked("<p>Old</p>"))
    other = _FakeInspector(types.SimpleNamespace(EntryID="another-draft"))

    def _user_edit_saved() -> None:
        mail._html = _marked("<p>Old</p><p>typed by hand</p>")

    own = _FakeInspector(mail, on_close=_user_edit_saved)

    _update(monkeypatch, mail, inspectors=[other, own])

    assert own.closed_with == [mapi.OL_SAVE]
    assert other.closed_with == []
    assert mail.HTMLBody == _marked("<p>New</p>")


def test_an_unmarked_draft_is_refused_without_closing_its_open_window(monkeypatch):
    mail = _FakeExistingMail(f"<html><body><p>Old</p>{SIG}</body></html>")
    window = _FakeInspector(mail)

    with pytest.raises(DraftUpdateError, match="no marked body region"):
        _update(monkeypatch, mail, inspectors=[window])

    assert window.closed_with == []

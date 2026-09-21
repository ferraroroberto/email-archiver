"""Tests for the draft read-back and the guarded `send` verb (issue #82).

A caller shows the user a preview, the user approves it in a later message,
and only then is `main_batch.py send` spawned. These tests pin what that
approval depends on:

- ``read`` reports the item as stored — recipients by line, the whole body,
  attachment bytes — and a fingerprint that is stable while nothing changes;
- ``send`` sends exactly one item, the right one, when the live fingerprint
  matches, and in every refusal path sends nothing and names what differs;
- the one ``Send`` call lives in ``outlook/sending.py`` and nowhere else.

Every address, path and subject here is synthetic.
"""
from __future__ import annotations

import hashlib
import json
import sys
import types
from pathlib import Path

import pytest

import main_batch
from email_archiver import draft, send
from email_archiver.outlook import mapi, process
from email_archiver.outlook.client import OutlookClient
from email_archiver.outlook.drafts import DraftUpdateError
from email_archiver.outlook.mapi import (
    mark_body_html,
    marked_body_region,
    recipient_address,
)

REPO_ROOT = Path(__file__).resolve().parents[1]
SELF = "me@example.invalid"
OTHER = "someone@example.invalid"
REGION = draft.text_to_html("Hello,\n\nSecond paragraph & more.")
HTML = f"<html><body>{mark_body_html(REGION)}<div id=\"_MailAutoSig\">-- sig</div></body></html>"


# --------------------------------------------------------------- fake COM ---

class _FakeRecipient:
    def __init__(self, kind: int, address: str = "", name: str = "", entry_address: str = "") -> None:
        self.Type = kind
        self.Address = address
        self.Name = name
        self.AddressEntry = types.SimpleNamespace(Address=entry_address)


class _FakeCollection:
    def __init__(self, items: list) -> None:
        self.items = items

    @property
    def Count(self) -> int:  # noqa: N802 - COM's spelling
        return len(self.items)

    def Item(self, i: int):  # noqa: N802 - COM's spelling
        return self.items[i - 1]


class _FakeAttachment:
    def __init__(self, name: str, data: bytes) -> None:
        self.FileName = name
        self.data = data

    def SaveAsFile(self, path: str) -> None:  # noqa: N802 - COM's spelling
        Path(path).write_bytes(self.data)


class _FakeDraft:
    def __init__(self, entry_id: str = "draft-1") -> None:
        self.EntryID = entry_id
        self.Subject = "A subject"
        self.HTMLBody = HTML
        self.Sent = False
        self.Parent = types.SimpleNamespace(EntryID="drafts")
        # Unresolved, as on a fresh draft: the address is in Name, Address is empty.
        self.Recipients = _FakeCollection([
            _FakeRecipient(mapi.OL_TO, name=OTHER),
            _FakeRecipient(mapi.OL_BCC, address=SELF),
        ])
        self.Attachments = _FakeCollection([_FakeAttachment("note.txt", b"synthetic bytes")])
        self.calls: list[tuple] = []

    def Send(self) -> None:  # noqa: N802 - COM's spelling
        self.calls.append(("Send",))
        self.Sent = True


class _FakeNamespace:
    def __init__(self, items: dict[str, _FakeDraft]) -> None:
        self.items = items

    def GetItemFromID(self, entry_id: str):  # noqa: N802 - COM's spelling
        if entry_id not in self.items:
            raise RuntimeError("The operation failed. An object could not be found.")
        return self.items[entry_id]

    def GetDefaultFolder(self, kind: int):  # noqa: N802 - COM's spelling
        assert kind == mapi.OL_FOLDER_DRAFTS
        return types.SimpleNamespace(EntryID="drafts")


class _FakeInspector:
    def __init__(self, item: _FakeDraft, on_close=None) -> None:
        self.CurrentItem = item
        self.closed_with: list[int] = []
        self._on_close = on_close

    def Close(self, mode: int) -> None:  # noqa: N802 - COM's spelling
        self.closed_with.append(mode)
        if self._on_close:
            self._on_close()


def _client(monkeypatch, *items: _FakeDraft, inspectors: list[_FakeInspector] | None = None) -> OutlookClient:
    client = OutlookClient()
    namespace = _FakeNamespace({item.EntryID: item for item in items})
    monkeypatch.setattr(client, "_namespace", lambda: namespace)
    # Never the real Outlook: the only one on this machine is the user's own.
    app = types.SimpleNamespace(Inspectors=_FakeCollection(inspectors or []))
    monkeypatch.setattr(process, "get_active_application", lambda: app)
    return client


def _approved(monkeypatch, mail: _FakeDraft) -> dict:
    """The fingerprint `read` reports for ``mail`` as it is now."""
    return send.fingerprint(_client(monkeypatch, mail).read_draft(mail.EntryID))


def _send(monkeypatch, mail: _FakeDraft, approval: dict, *, parts: bool = True,
          expect_to: list[str] | None = None, inspectors=None) -> dict:
    client = _client(monkeypatch, mail, inspectors=inspectors)
    return send.send(client, mail.EntryID, approval["hash"], approval["parts"] if parts else {}, expect_to)


# ------------------------------------------------------------------ read ---

def test_read_reports_the_item_as_stored_and_changes_nothing(monkeypatch):
    mail = _FakeDraft()
    doc = send.read_document(_client(monkeypatch, mail).read_draft("draft-1"))

    assert (doc["verb"], doc["entry_id"], doc["subject"]) == ("read", "draft-1", "A subject")
    assert (doc["to"], doc["cc"], doc["bcc"], doc["unreadable_recipients"]) == ([OTHER], [], [SELF], 0)
    assert doc["body"] == {
        "html": HTML, "region_html": REGION, "region_text": "Hello,\n\nSecond paragraph & more.",
    }
    assert doc["attachments"] == [{
        "name": "note.txt", "size_bytes": len(b"synthetic bytes"),
        "sha256": hashlib.sha256(b"synthetic bytes").hexdigest(),
    }]
    assert set(doc["fingerprint"]["parts"]) == set(send.PARTS)
    assert mail.calls == [] and mail.HTMLBody == HTML


def test_the_fingerprint_is_stable_while_nothing_changes(monkeypatch):
    mail = _FakeDraft()
    assert _approved(monkeypatch, mail) == _approved(monkeypatch, mail)


def test_recipient_order_and_case_do_not_change_the_fingerprint(monkeypatch):
    mail = _FakeDraft()
    before = _approved(monkeypatch, mail)
    mail.Recipients.items.reverse()
    mail.Recipients.items[0].Address = SELF.upper()
    assert _approved(monkeypatch, mail) == before


@pytest.mark.parametrize("recipient, expected", [
    (_FakeRecipient(1, address=OTHER), OTHER),
    (_FakeRecipient(1, name=OTHER), OTHER),
    (_FakeRecipient(1, address="/o=ExchangeLabs/cn=x", entry_address=OTHER, name="Some One"), OTHER),
    (_FakeRecipient(1, name="Some One"), ""),
])
def test_a_recipient_address_is_read_from_wherever_outlook_left_it(recipient, expected):
    assert recipient_address(recipient) == expected


def test_the_region_text_is_only_given_for_a_body_text_to_html_wrote():
    assert draft.html_to_text(draft.text_to_html("a  b\nc &lt; d\n\ne")) == "a  b\nc &lt; d\n\ne"
    assert draft.html_to_text("<p>Outlook <b>rewrote</b> this</p>") is None
    assert draft.html_to_text("<div>not ours</div>") is None
    assert marked_body_region("<body>no markers</body>") is None


# ------------------------------------------------------------ send: match ---

def test_a_matching_draft_is_sent_exactly_once(monkeypatch):
    mail, bystander = _FakeDraft(), _FakeDraft("draft-2")
    approval = _approved(monkeypatch, mail)
    client = _client(monkeypatch, mail, bystander)

    doc = send.send(client, "draft-1", approval["hash"], approval["parts"], [OTHER])

    assert mail.calls == [("Send",)]
    assert bystander.calls == []
    assert (doc["verb"], doc["sent"], doc["entry_id"], doc["to"], doc["bcc"]) == ("send", True, "draft-1", [OTHER], [SELF])


def test_a_window_open_on_the_draft_is_closed_saving_before_the_read(monkeypatch):
    mail = _FakeDraft()
    approval = _approved(monkeypatch, mail)
    inspector = _FakeInspector(mail)

    _send(monkeypatch, mail, approval, inspectors=[inspector])

    assert inspector.closed_with == [mapi.OL_SAVE]
    assert mail.calls == [("Send",)]


# --------------------------------------------------------- send: refusals ---

def _tamper_body(mail):
    mail.HTMLBody = HTML.replace("Second paragraph", "Second paragraph!")


def _tamper_outside_region(mail):
    mail.HTMLBody = HTML.replace("-- sig", "-- sig, and a line typed by hand")


def _tamper_subject(mail):
    mail.Subject = "A subject."


def _tamper_recipient(mail):
    mail.Recipients.items[0].Name = "other@example.invalid"


def _add_recipient(mail):
    mail.Recipients.items.append(_FakeRecipient(mapi.OL_CC, address="extra@example.invalid"))


def _tamper_attachment(mail):
    mail.Attachments.items[0].data = b"synthetic bytez"


@pytest.mark.parametrize("tamper, differs", [
    (_tamper_body, ["body"]),
    (_tamper_outside_region, ["body"]),
    (_tamper_subject, ["subject"]),
    (_tamper_recipient, ["to"]),
    (_add_recipient, ["cc"]),
    (_tamper_attachment, ["attachments"]),
])
def test_any_change_after_approval_is_refused_naming_the_part(monkeypatch, tamper, differs):
    mail = _FakeDraft()
    approval = _approved(monkeypatch, mail)
    tamper(mail)

    with pytest.raises(send.SendRefused) as refused:
        _send(monkeypatch, mail, approval)

    assert (refused.value.code, refused.value.differs) == ("approval_mismatch", differs)
    assert "Nothing was sent" in str(refused.value)
    assert mail.calls == []


def test_an_edit_saved_from_an_open_window_is_caught(monkeypatch):
    mail = _FakeDraft()
    approval = _approved(monkeypatch, mail)
    inspector = _FakeInspector(mail, on_close=lambda: _tamper_body(mail))

    with pytest.raises(send.SendRefused) as refused:
        _send(monkeypatch, mail, approval, inspectors=[inspector])

    assert refused.value.differs == ["body"]
    assert mail.calls == []


def test_a_hash_mismatch_without_parts_is_still_refused(monkeypatch):
    mail = _FakeDraft()
    approval = _approved(monkeypatch, mail)
    _tamper_body(mail)

    with pytest.raises(send.SendRefused) as refused:
        _send(monkeypatch, mail, approval, parts=False)

    assert (refused.value.code, refused.value.differs) == ("approval_mismatch", [])
    assert mail.calls == []


def test_a_wrong_expected_to_is_refused_even_when_the_hash_matches(monkeypatch):
    mail = _FakeDraft()
    approval = _approved(monkeypatch, mail)

    with pytest.raises(send.SendRefused) as refused:
        _send(monkeypatch, mail, approval, expect_to=[OTHER, "extra@example.invalid"])

    assert refused.value.differs == ["to"]
    assert mail.calls == []


def test_a_recipient_with_no_readable_address_is_refused(monkeypatch):
    mail = _FakeDraft()
    mail.Recipients.items.append(_FakeRecipient(mapi.OL_CC, name="Nobody Resolvable"))
    approval = _approved(monkeypatch, mail)

    with pytest.raises(send.SendRefused) as refused:
        _send(monkeypatch, mail, approval)

    assert refused.value.code == "recipient_unreadable"
    assert mail.calls == []


@pytest.mark.parametrize("setup, code", [
    (lambda mail: setattr(mail, "Sent", True), "draft_not_editable"),
    (lambda mail: setattr(mail, "Parent", types.SimpleNamespace(EntryID="inbox")), "draft_not_editable"),
])
def test_a_sent_or_moved_item_is_refused(monkeypatch, setup, code):
    mail = _FakeDraft()
    approval = _approved(monkeypatch, mail)
    setup(mail)

    with pytest.raises(DraftUpdateError) as refused:
        _send(monkeypatch, mail, approval)

    assert refused.value.code == code
    assert mail.calls == []


def test_a_missing_item_is_refused(monkeypatch):
    mail = _FakeDraft()
    approval = _approved(monkeypatch, mail)
    client = _client(monkeypatch)

    with pytest.raises(DraftUpdateError) as refused:
        send.send(client, "draft-1", approval["hash"], approval["parts"], None)

    assert refused.value.code == "draft_not_found"
    assert mail.calls == []


def test_send_is_called_only_from_the_sending_module():
    # Spelled in two halves so a grep of the repo for the call finds only the one.
    needle = "." + "Send("
    holders = sorted(
        str(path.relative_to(REPO_ROOT)).replace("\\", "/")
        for path in [*REPO_ROOT.glob("*.py"), *(REPO_ROOT / "email_archiver").rglob("*.py")]
        if needle in path.read_text(encoding="utf-8")
    )
    assert holders == ["email_archiver/outlook/sending.py"]


# ------------------------------------------------------ process contract ---

@pytest.fixture
def batch_process(monkeypatch, capsys):
    """``main_batch.main`` with a temp config, a stub pythoncom and a fake
    Outlook client over the given drafts."""
    monkeypatch.setattr(main_batch, "load_config", lambda: {"outlook": {}})
    monkeypatch.setattr(main_batch, "setup_logging", lambda _cfg: None)
    monkeypatch.setitem(sys.modules, "pythoncom", types.SimpleNamespace(
        CoInitialize=lambda: None, CoUninitialize=lambda: None,
    ))
    started: list[OutlookClient] = []

    def run(argv: list[str], *items: _FakeDraft) -> tuple[int, dict, list]:
        client = _client(monkeypatch, *items)
        monkeypatch.setattr(client, "ensure_running", lambda timeout: None)

        def _factory() -> OutlookClient:
            started.append(client)
            return client

        monkeypatch.setattr(main_batch, "OutlookClient", _factory)
        code = main_batch.main(argv)
        return code, json.loads(capsys.readouterr().out), started

    return run


def test_read_prints_the_snapshot_document(batch_process):
    code, doc, _ = batch_process(["read", "--entry-id", "draft-1"], _FakeDraft())

    assert code == main_batch.EXIT_OK
    assert doc["verb"] == "read" and len(doc["fingerprint"]["hash"]) == 64


def test_a_matching_send_prints_its_document(batch_process, monkeypatch):
    mail = _FakeDraft()
    approval = _approved(monkeypatch, mail)
    argv = ["send", "--entry-id", "draft-1", "--expect-hash", approval["hash"], "--expect-to", OTHER]
    argv += [f"--expect-part={name}={digest}" for name, digest in approval["parts"].items()]

    code, doc, _ = batch_process(argv, mail)

    assert (code, doc["sent"]) == (main_batch.EXIT_OK, True)
    assert mail.calls == [("Send",)]


def test_a_refused_send_exits_2_naming_the_part(batch_process, monkeypatch):
    mail = _FakeDraft()
    approval = _approved(monkeypatch, mail)
    _tamper_body(mail)
    argv = ["send", "--entry-id", "draft-1", "--expect-hash", approval["hash"],
            f"--expect-part=body={approval['parts']['body']}"]

    code, doc, _ = batch_process(argv, mail)

    assert code == main_batch.EXIT_CANNOT_START
    assert (doc["error"]["code"], doc["error"]["differs"]) == ("approval_mismatch", ["body"])
    assert mail.calls == []


@pytest.mark.parametrize("argv", [
    ["send", "--entry-id", "draft-1", "--expect-hash", "not-a-hash"],
    ["send", "--entry-id", "draft-1", "--expect-hash", "a" * 64, "--expect-part", "signature=" + "b" * 64],
    ["send", "--entry-id", "draft-1", "--expect-hash", "a" * 64, "--expect-to", " "],
    ["send", "--entry-id", " ", "--expect-hash", "a" * 64],
    ["read", "--entry-id", " "],
])
def test_bad_send_input_exits_2_without_starting_outlook(batch_process, argv):
    code, doc, started = batch_process(argv, _FakeDraft())

    assert code == main_batch.EXIT_CANNOT_START
    assert doc["error"]["code"] == "bad_input"
    assert started == []

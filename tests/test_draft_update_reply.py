"""``draft --update`` (re-filling an unsent draft in place) and ``draft`` with
``reply_to`` (a threaded reply built from the original mail). The two share their
fake namespace, signature and quote fixtures.

Every address, path and subject here is synthetic.
"""
from __future__ import annotations

import re
import types
from pathlib import Path

import pytest

from email_archiver.cli import batch as main_batch
from email_archiver import draft
from email_archiver.outlook import mapi, process
from email_archiver.outlook.client import OutlookClient
from email_archiver.outlook.drafts import (
    REPLY_BY_MESSAGE_ID,
    REPLY_BY_MSG_PATH,
    CreatedDraft,
    DraftUpdateError,
    ReplySourceError,
    ReplyTarget,
)
from email_archiver.outlook.mapi import (
    DASL_X_ARCHIVE_REF,
    mark_body_html,
    replace_marked_body_html,
)
from tests.draft_support import (
    FakeDraftClient,
    SELF,
    _FakeApplication,
    _FakeComMail,
    _FakePropertyAccessor,
    _spec,
)
from tests.draft_support import batch_process  # noqa: F401 - fixtures are used by name


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


# What Outlook's editor stores for a reply drafted with the compose window open
# (issue #105, read back live): the opening div survives, unquoted, the closing
# comment does not, and each paragraph is padded with an empty <o:p>.
WORD_REPLY = (
    "<html><body lang=EN><div class=WordSection1>"
    "<div id=archive-draft-body><p>New<o:p></o:p></p></div>"
    '<p class=MsoNormal><o:p>&nbsp;</o:p></p><div id="_MailAutoSig">-- sig</div>'
    '<div id="quote">On a day, they wrote:<blockquote>original text</blockquote></div>'
    "</div></body></html>"
)


def test_a_region_whose_closing_comment_outlook_dropped_is_still_found():
    assert mapi.marked_body_region(WORD_REPLY) == "<p>New<o:p></o:p></p>"
    after = mapi.html_after_marked_region(WORD_REPLY)
    assert after is not None
    assert after.startswith('<p class=MsoNormal>') and "original text" in after and "archive-draft-body" not in after


def test_a_commentless_region_with_its_own_nested_divs_ends_at_its_own_div():
    existing = (
        "<body><div id=archive-draft-body><div><div>deep</div></div><p>tail</p></div>"
        f"{SIG}{QUOTE}</body>"
    )

    assert mapi.marked_body_region(existing) == "<div><div>deep</div></div><p>tail</p>"
    assert mapi.html_after_marked_region(existing) == f"{SIG}{QUOTE}</body>"


def test_the_body_of_a_commentless_reply_is_replaced_and_the_quote_kept():
    updated = replace_marked_body_html(WORD_REPLY, "<p>Newer</p>")

    assert updated is not None
    assert mapi.marked_body_region(updated) == "<p>Newer</p>"
    assert updated.endswith(WORD_REPLY[WORD_REPLY.index('<p class=MsoNormal>'):])
    assert updated.count("original text") == 1


def test_the_text_after_a_region_with_its_comment_is_the_same_text():
    existing = f"<body>{mark_body_html('<p>x</p>')}{SIG}{QUOTE}</body>"

    assert mapi.html_after_marked_region(existing) == f"{SIG}{QUOTE}</body>"
    assert mapi.html_after_marked_region("<body>no markers</body>") is None


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

    @property
    def DefaultStore(self):  # noqa: N802 - COM's spelling
        # The default store's folders are this namespace's (issue #108).
        return self

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


# --------------------------------------------------- threaded reply (#101) ---

QUOTE = '<div id="quote">On a day, they wrote:<blockquote>original text</blockquote></div>'
ORIGINAL_ID = "orig-1@example.invalid"


def _msg_file(tmp_path: Path) -> Path:
    path = tmp_path / "saved.msg"
    path.write_bytes(b"synthetic")
    return path


class _FakeReplyAccessor(_FakePropertyAccessor):
    def __init__(self, calls: list, properties: dict) -> None:
        super().__init__(calls, refuse=False)
        self._properties = properties

    def GetProperty(self, name: str):  # noqa: N802 - COM's spelling
        if name not in self._properties:
            raise RuntimeError("property not found")
        return self._properties[name]


class _FakeRecipient:
    def __init__(self, address: str, kind: int) -> None:
        self.Address, self.Type = address, kind


class _FakeRecipients:
    def __init__(self, items: list[_FakeRecipient]) -> None:
        self._items = items
        self.Count = len(items)

    def Item(self, i: int) -> _FakeRecipient:  # noqa: N802 - COM's spelling
        return self._items[i - 1]


class _FakeReplyItem(_FakeComMail):
    """What ``Reply()`` returns: already quoting the original; ``Display`` adds
    the signature on top of the quote, the way Outlook does."""

    def __init__(self, in_reply_to: str = f"<{ORIGINAL_ID}>", quote_html: str | None = None) -> None:
        super().__init__()
        self.Subject = "RE: Original subject"
        self.Recipients = _FakeRecipients([
            _FakeRecipient("sender@example.invalid", mapi.OL_TO),
            _FakeRecipient("other@example.invalid", mapi.OL_CC),
        ])
        self._quote = f"<html><body>{QUOTE}</body></html>" if quote_html is None else quote_html
        self._html = self._quote
        properties = {mapi.DASL_IN_REPLY_TO_ID: in_reply_to} if in_reply_to else {}
        self.PropertyAccessor = _FakeReplyAccessor(self.calls, properties)

    def Display(self, modal: bool) -> None:  # noqa: N802 - COM's spelling
        self.calls.append(("Display", modal))
        self._html = self._quote.replace("<body>", f"<body>{SIG}", 1)

    def Close(self, mode: int) -> None:  # noqa: N802 - COM's spelling
        self.calls.append(("Close", mode))


class _FakeOriginal:
    """The mail being answered."""

    def __init__(self, reply: _FakeReplyItem, refuse_reply: bool = False) -> None:
        self.reply = reply
        self.refuse_reply = refuse_reply
        self.Class = mapi.OL_CLASS_MAIL_ITEM
        self.calls: list = []
        self.PropertyAccessor = _FakeReplyAccessor(
            self.calls, {mapi.DASL_INTERNET_MESSAGE_ID: f"<{ORIGINAL_ID}>"},
        )

    def _answer(self, name: str) -> _FakeReplyItem:
        self.calls.append((name,))
        if self.refuse_reply:
            raise RuntimeError("there is no account to reply from")
        return self.reply

    def Reply(self) -> _FakeReplyItem:  # noqa: N802 - COM's spelling
        return self._answer("Reply")

    def ReplyAll(self) -> _FakeReplyItem:  # noqa: N802 - COM's spelling
        return self._answer("ReplyAll")

    def Close(self, mode: int) -> None:  # noqa: N802 - COM's spelling
        self.calls.append(("Close", mode))


class _FakeSharedNamespace:
    def __init__(self, original: _FakeOriginal | None) -> None:
        self.original = original
        self.opened: list[str] = []

    def OpenSharedItem(self, path: str) -> _FakeOriginal:  # noqa: N802 - COM's spelling
        self.opened.append(path)
        if self.original is None:
            raise RuntimeError("Cannot open the file")
        return self.original


def _reply_client(monkeypatch, original: _FakeOriginal | None):
    """A client whose Outlook is entirely fake; the only real Outlook here is the user's."""
    app = _FakeApplication(_FakeComMail())
    monkeypatch.setattr(process, "get_active_application", lambda: app)
    client = OutlookClient()
    namespace = _FakeSharedNamespace(original)
    monkeypatch.setattr(client, "_namespace", lambda: namespace)
    looked_up: list[str] = []

    def _find(message_id: str, folder_name=None):
        looked_up.append(message_id)
        return original

    monkeypatch.setattr(client, "find_by_message_id", _find)
    return client, app, namespace, looked_up


def _reply(client: OutlookClient, **overrides) -> CreatedDraft:
    kwargs = dict(
        to=None, cc=None, bcc=[SELF], subject=None, body_html="<p>My answer</p>",
        attachments=[], ref="tok", display=True,
        reply_to=ReplyTarget(REPLY_BY_MESSAGE_ID, ORIGINAL_ID),
    )
    kwargs.update(overrides)
    return client.create_draft(**kwargs)


# -- spec --

def test_a_reply_spec_needs_no_recipient_or_subject():
    spec = draft.parse_spec({
        "body_text": "Hello", "reply_to": {"message_id": f" <{ORIGINAL_ID}> "}, "reply_all": True,
    })

    assert (spec.to, spec.cc, spec.subject) == (None, None, None)
    assert spec.reply_to == ReplyTarget(REPLY_BY_MESSAGE_ID, ORIGINAL_ID, reply_all=True)


def test_reply_overrides_are_kept_when_given(tmp_path):
    msg = _msg_file(tmp_path)
    spec = draft.parse_spec({
        "body_text": "Hello", "reply_to": {"msg_path": str(msg)},
        "to": ["x@example.invalid"], "cc": [], "subject": "Own subject",
    })

    assert (spec.to, spec.cc, spec.subject) == (["x@example.invalid"], [], "Own subject")
    assert spec.reply_to == ReplyTarget(REPLY_BY_MSG_PATH, str(msg))


@pytest.mark.parametrize(
    ("overrides", "message"),
    [
        ({"reply_to": "<id>"}, "`reply_to` must be an object"),
        ({"reply_to": {}}, "exactly one of `message_id` or `msg_path`"),
        ({"reply_to": {"message_id": "a", "msg_path": "b"}}, "exactly one of `message_id` or `msg_path`"),
        ({"reply_to": {"message_id": " "}}, "`reply_to.message_id` must be a non-empty string"),
        ({"reply_to": {"message_id": "<>"}}, "`reply_to.message_id` must be a non-empty string"),
        ({"reply_to": {"messageid": "a"}}, "unknown `reply_to` keys: messageid"),
        ({"reply_to": {"msg_path": "relative/saved.msg"}}, "must be an absolute path"),
        ({"reply_to": {"msg_path": "Z:/definitely/not/here.msg"}}, "not an existing .msg file"),
        ({"reply_all": True}, "`reply_all` needs `reply_to`"),
        ({"reply_to": {"message_id": "a"}, "reply_all": "yes"}, "`reply_all` must be true or false"),
    ],
)
def test_an_unusable_reply_spec_is_refused_with_the_field_named(overrides, message):
    with pytest.raises(draft.SpecError, match=re.escape(message)):
        draft.parse_spec({"body_text": "Hello", **overrides})


def test_a_msg_path_must_be_a_msg_file(tmp_path):
    other = tmp_path / "saved.txt"
    other.write_text("x", encoding="utf-8")
    with pytest.raises(draft.SpecError, match="not an existing .msg file"):
        draft.parse_spec({"body_text": "Hello", "reply_to": {"msg_path": str(other)}})


@pytest.mark.parametrize("missing", ["to", "subject"])
def test_a_new_mail_still_needs_its_recipient_and_subject(missing):
    data = _spec()
    del data[missing]
    with pytest.raises(draft.SpecError, match=f"`{missing}`"):
        draft.parse_spec(data)


def test_a_new_mail_spec_is_unchanged_by_the_reply_keys():
    spec = draft.parse_spec(_spec(cc=["c@example.invalid"]))

    assert spec.reply_to is None
    assert (spec.to, spec.cc, spec.subject) == (["someone@example.invalid"], ["c@example.invalid"], "A subject")


# -- client --

def test_a_reply_by_message_id_is_outlooks_own_reply_with_the_body_above_the_quote(monkeypatch):
    reply = _FakeReplyItem()
    original = _FakeOriginal(reply)
    client, app, _, looked_up = _reply_client(monkeypatch, original)

    created = _reply(client)

    assert looked_up == [ORIGINAL_ID] and app.created == []
    assert original.calls == [("Reply",)]
    assert reply.calls == [
        ("SetProperty", DASL_X_ARCHIVE_REF, "tok"), ("Display", False), ("HTMLBody=",), ("Save",),
    ]
    assert reply.HTMLBody == (
        f"<html><body>{mark_body_html('<p>My answer</p>')}{SIG}{QUOTE}</body></html>"
    )
    assert (reply.To, reply.CC, reply.Subject, reply.BCC) == ("", "", "RE: Original subject", SELF)
    assert ("Send",) not in reply.calls
    assert created.subject == "RE: Original subject"
    assert (created.to, created.cc) == (["sender@example.invalid"], ["other@example.invalid"])
    assert (created.replied_to_message_id, created.thread_header) == (ORIGINAL_ID, "set")


def test_reply_all_asks_outlook_for_reply_all(monkeypatch):
    original = _FakeOriginal(_FakeReplyItem())
    client, *_ = _reply_client(monkeypatch, original)

    _reply(client, reply_to=ReplyTarget(REPLY_BY_MESSAGE_ID, ORIGINAL_ID, reply_all=True))

    assert original.calls == [("ReplyAll",)]


def test_given_recipients_and_subject_override_outlooks(monkeypatch):
    reply = _FakeReplyItem()
    client, *_ = _reply_client(monkeypatch, _FakeOriginal(reply))

    _reply(client, to=["x@example.invalid"], cc=[], subject="Own subject")

    assert (reply.To, reply.CC, reply.Subject) == ("x@example.invalid", "", "Own subject")


def _sent_by_me(*, to=("them@example.invalid",), cc=(), unreadable=0, reply_all=False):
    """An original the user sent: Outlook's reply to it is addressed to the user
    (it answers the sender), while the original's own recipients are the people
    the user wrote to (issue #105)."""
    reply = _FakeReplyItem()
    reply.Recipients = _FakeRecipients([_FakeRecipient(SELF.upper(), mapi.OL_TO)])
    original = _FakeOriginal(reply)
    original.Recipients = _FakeRecipients([
        *(_FakeRecipient(a, mapi.OL_TO) for a in to),
        *(_FakeRecipient(a, mapi.OL_CC) for a in cc),
        *(_FakeRecipient("", mapi.OL_TO) for _ in range(unreadable)),
        _FakeRecipient(SELF, mapi.OL_BCC),
    ])
    return reply, original


def test_a_reply_to_a_mail_the_user_sent_is_addressed_to_its_recipients_not_to_the_user(monkeypatch):
    reply, original = _sent_by_me(to=("them@example.invalid", "too@example.invalid"), cc=("cc@example.invalid",))
    client, *_ = _reply_client(monkeypatch, original)

    created = _reply(client, self_address=SELF)

    assert (reply.To, reply.CC) == ("them@example.invalid; too@example.invalid", "")
    assert created.recipients_from == "original_recipients"
    assert ("Send",) not in reply.calls


def test_reply_all_to_a_mail_the_user_sent_keeps_the_original_cc(monkeypatch):
    reply, original = _sent_by_me(cc=("cc@example.invalid",))
    client, *_ = _reply_client(monkeypatch, original)

    _reply(client, self_address=SELF, reply_to=ReplyTarget(REPLY_BY_MESSAGE_ID, ORIGINAL_ID, reply_all=True))

    assert (reply.To, reply.CC) == ("them@example.invalid", "cc@example.invalid")


def test_a_reply_to_a_mail_the_user_sent_whose_recipients_cannot_be_read_is_refused(monkeypatch):
    reply, original = _sent_by_me(unreadable=1)
    client, *_ = _reply_client(monkeypatch, original)

    with pytest.raises(ReplySourceError) as caught:
        _reply(client, self_address=SELF)

    assert caught.value.code == "reply_unavailable" and "give `to` explicitly" in str(caught.value)
    assert ("Save",) not in reply.calls and reply.calls[-1] == ("Close", mapi.OL_DISCARD)


def test_a_given_to_overrides_the_original_recipients_of_a_mail_the_user_sent(monkeypatch):
    reply, original = _sent_by_me(unreadable=1)  # unreadable, but the caller named the recipient
    client, *_ = _reply_client(monkeypatch, original)

    created = _reply(client, self_address=SELF, to=["x@example.invalid"], cc=[])

    assert (reply.To, reply.CC) == ("x@example.invalid", "")
    assert created.recipients_from == "caller"


def test_a_reply_to_incoming_mail_is_left_with_outlooks_recipients(monkeypatch):
    reply = _FakeReplyItem()
    client, *_ = _reply_client(monkeypatch, _FakeOriginal(reply))

    created = _reply(client, self_address=SELF)

    assert (reply.To, reply.CC) == ("", "")
    assert created.recipients_from == "sender"


def test_the_reply_document_says_who_the_reply_was_addressed_from():
    class _Client(FakeDraftClient):
        def create_draft(self, **kwargs) -> CreatedDraft:
            self.drafts.append(kwargs)
            return CreatedDraft(entry_id="e", to=["them@example.invalid"], recipients_from="original_recipients",
                                thread_header="set")

    client = _Client()
    spec = draft.parse_spec({"body_text": "Hi", "reply_to": {"message_id": ORIGINAL_ID}})

    doc = draft.create(client, spec, SELF)

    assert doc["recipients_from"] == "original_recipients" and doc["to"] == ["them@example.invalid"]
    assert client.drafts[0]["self_address"] == SELF


def test_a_reply_by_msg_path_opens_the_file_and_closes_it_again_discarding(monkeypatch, tmp_path):
    reply = _FakeReplyItem()
    original = _FakeOriginal(reply)
    client, app, namespace, looked_up = _reply_client(monkeypatch, original)
    path = str(_msg_file(tmp_path))

    _reply(client, reply_to=ReplyTarget(REPLY_BY_MSG_PATH, path))

    assert namespace.opened == [path] and looked_up == []
    assert original.calls == [("Reply",), ("Close", mapi.OL_DISCARD)]
    assert ("Save",) in reply.calls


def test_an_inbox_original_is_never_closed(monkeypatch):
    original = _FakeOriginal(_FakeReplyItem())
    client, *_ = _reply_client(monkeypatch, original)

    _reply(client)

    assert ("Close", mapi.OL_DISCARD) not in original.calls


def test_a_reply_that_is_not_threaded_says_so(monkeypatch):
    client, *_ = _reply_client(monkeypatch, _FakeOriginal(_FakeReplyItem(in_reply_to="")))

    created = _reply(client)

    assert created.thread_header == "not_set"
    assert "no In-Reply-To" in created.thread_header_reason


def test_an_unknown_message_id_creates_nothing(monkeypatch):
    client, app, _, _ = _reply_client(monkeypatch, None)

    with pytest.raises(ReplySourceError) as caught:
        _reply(client)

    assert caught.value.code == "reply_source_not_found"
    assert app.created == [] and app.mail.calls == []


def test_an_unopenable_msg_file_creates_nothing(monkeypatch, tmp_path):
    client, app, _, _ = _reply_client(monkeypatch, None)

    with pytest.raises(ReplySourceError) as caught:
        _reply(client, reply_to=ReplyTarget(REPLY_BY_MSG_PATH, str(_msg_file(tmp_path))))

    assert caught.value.code == "reply_source_not_found"
    assert "Cannot open the file" in str(caught.value)
    assert app.created == [] and app.mail.calls == []


def test_an_original_that_is_not_a_mail_is_not_replied_to(monkeypatch, tmp_path):
    original = _FakeOriginal(_FakeReplyItem())
    original.Class = 26  # a calendar item
    client, *_ = _reply_client(monkeypatch, original)

    with pytest.raises(ReplySourceError) as caught:
        _reply(client, reply_to=ReplyTarget(REPLY_BY_MSG_PATH, str(_msg_file(tmp_path))))

    assert caught.value.code == "reply_source_not_found"
    assert original.calls == [("Close", mapi.OL_DISCARD)]


def test_outlook_refusing_to_reply_is_its_own_error_and_saves_nothing(monkeypatch, tmp_path):
    reply = _FakeReplyItem()
    original = _FakeOriginal(reply, refuse_reply=True)
    client, *_ = _reply_client(monkeypatch, original)

    with pytest.raises(ReplySourceError) as caught:
        _reply(client, reply_to=ReplyTarget(REPLY_BY_MSG_PATH, str(_msg_file(tmp_path))))

    assert caught.value.code == "reply_unavailable"
    assert "no account to reply from" in str(caught.value)
    assert original.calls[-1] == ("Close", mapi.OL_DISCARD)
    assert reply.calls == []


def test_a_reply_with_no_readable_body_is_refused_rather_than_losing_the_quote(monkeypatch):
    reply = _FakeReplyItem(quote_html="")
    client, *_ = _reply_client(monkeypatch, _FakeOriginal(reply))

    with pytest.raises(ReplySourceError) as caught:
        _reply(client, display=False)

    assert caught.value.code == "reply_unavailable"
    assert ("Save",) not in reply.calls and reply.calls[-1] == ("Close", mapi.OL_DISCARD)


def test_an_undisplayed_reply_still_keeps_the_quote(monkeypatch):
    reply = _FakeReplyItem()
    client, *_ = _reply_client(monkeypatch, _FakeOriginal(reply))

    created = _reply(client, display=False)

    assert created.displayed is False
    assert reply.HTMLBody == f"<html><body>{mark_body_html('<p>My answer</p>')}{QUOTE}</body></html>"


# -- document, update, fingerprint, process --

def test_the_reply_document_reports_what_was_replied_to():
    class _Client(FakeDraftClient):
        def create_draft(self, **kwargs) -> CreatedDraft:
            self.drafts.append(kwargs)
            return CreatedDraft(
                entry_id="draft-entry-1", displayed=True, subject="RE: Original subject",
                to=["sender@example.invalid"], cc=["other@example.invalid"],
                replied_to_message_id=ORIGINAL_ID, thread_header="set",
            )

    spec = draft.parse_spec({"body_text": "Hi", "reply_to": {"message_id": ORIGINAL_ID}})

    doc = draft.create(_Client(), spec, SELF)

    assert doc["subject"] == "RE: Original subject"
    assert (doc["to"], doc["cc"], doc["bcc"]) == (["sender@example.invalid"], ["other@example.invalid"], [SELF])
    assert doc["reply_to"] == {"message_id": ORIGINAL_ID, "reply_all": False}
    assert (doc["replied_to_message_id"], doc["thread_header"], doc["thread_header_reason"]) == (
        ORIGINAL_ID, "set", "",
    )


def test_a_new_mail_document_says_it_is_not_a_reply():
    doc = draft.create(FakeDraftClient(), draft.parse_spec(_spec()), SELF)

    assert (doc["reply_to"], doc["thread_header"]) == (None, "not_a_reply")
    assert "replied_to_message_id" not in doc


def test_updating_a_reply_keeps_its_recipients_subject_and_quote(monkeypatch):
    mail = _FakeExistingMail(f"<html><body>{mark_body_html('<p>Old</p>')}{SIG}{QUOTE}</body></html>")
    mail.To, mail.Subject = "sender@example.invalid", "RE: Original subject"

    _update(monkeypatch, mail, to=None, cc=None, subject=None, attachments=[])

    assert (mail.To, mail.Subject) == ("sender@example.invalid", "RE: Original subject")
    assert mail.HTMLBody == f"<html><body>{mark_body_html('<p>New</p>')}{SIG}{QUOTE}</body></html>"


def test_the_update_document_does_not_claim_recipients_it_left_alone():
    client = FakeDraftClient()
    spec = draft.parse_spec({"body_text": "Hi", "reply_to": {"message_id": ORIGINAL_ID}})

    doc = draft.update(client, "draft-entry-1", spec, SELF)

    assert (doc["to"], doc["cc"], doc["subject"]) == (None, None, None)
    assert doc["thread_header"] == "unchanged" and doc["updated"] is True
    assert client.updates[0]["to"] is None


def test_updating_a_reply_whose_comment_outlook_dropped_replaces_the_text_and_keeps_the_quote(monkeypatch):
    mail = _FakeExistingMail(WORD_REPLY)

    _update(monkeypatch, mail, to=None, cc=None, subject=None, attachments=[])

    assert mapi.marked_body_region(mail.HTMLBody) == "<p>New</p>"
    assert mail.HTMLBody.count("original text") == 1 and '<div id="quote">' in mail.HTMLBody


def test_the_fingerprint_binds_the_quoted_original_but_the_preview_region_does_not():
    from email_archiver import send
    from email_archiver.outlook.drafts import DraftSnapshot

    def snapshot(quote: str) -> DraftSnapshot:
        html = f"<html><body>{mark_body_html('<p>Answer</p>')}{quote}</body></html>"
        return DraftSnapshot(
            entry_id="e", subject="RE: S", to=["a@example.invalid"], cc=[], bcc=[SELF],
            html_body=html, body_region_html=mapi.marked_body_region(html), attachments=[],
        )

    first, changed = snapshot("<blockquote>one</blockquote>"), snapshot("<blockquote>two</blockquote>")

    assert send.fingerprint(first)["hash"] != send.fingerprint(changed)["hash"]
    assert send.fingerprint(first)["parts"]["body"] != send.fingerprint(changed)["parts"]["body"]
    assert send.read_document(first)["body"]["region_text"] == "Answer"
    assert "one" in send.read_document(first)["body"]["html"]


def test_a_reply_spec_on_the_command_line_creates_a_reply(batch_process):
    code, doc, clients = batch_process({"body_text": "Hi", "reply_to": {"message_id": ORIGINAL_ID}})

    assert code == main_batch.EXIT_OK
    assert clients[0].drafts[0]["reply_to"] == ReplyTarget(REPLY_BY_MESSAGE_ID, ORIGINAL_ID)
    assert clients[0].drafts[0]["to"] is None and clients[0].drafts[0]["subject"] is None
    assert doc["reply_to"] == {"message_id": ORIGINAL_ID, "reply_all": False}


@pytest.mark.parametrize("code", ["reply_source_not_found", "reply_unavailable"])
def test_a_reply_that_cannot_start_exits_2_with_its_own_code(batch_process, code):
    refusal = ReplySourceError(code, "synthetic reason")

    exit_code, doc, _ = batch_process(
        {"body_text": "Hi", "reply_to": {"message_id": ORIGINAL_ID}}, refuse_create=refusal,
    )

    assert exit_code == main_batch.EXIT_CANNOT_START
    assert doc["error"] == {"code": code, "message": "synthetic reason"}


def test_a_bad_reply_spec_exits_2_without_starting_outlook(batch_process):
    code, doc, clients = batch_process({"body_text": "Hi", "reply_all": True})

    assert code == main_batch.EXIT_CANNOT_START
    assert doc["error"]["code"] == "bad_input" and "reply_all" in doc["error"]["message"]
    assert clients == []

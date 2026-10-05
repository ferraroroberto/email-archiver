"""
Headless draft: turn a caller's JSON spec into an unsent Outlook draft.

This is what ``main_batch.py draft`` runs. Like ``batch.py`` it is pure
orchestration over :class:`~email_archiver.outlook.client.OutlookClient` — no
COM here — so a fake client drives it end to end in a unit test.

Design decisions:

- **Drafts only.** The draft is saved to Drafts and shown; the user reads it
  and presses Send. Nothing in this module or the client's draft surface sends.
- **The spec is validated before Outlook is touched.** A missing attachment, no
  recipient or an ambiguous body fails in milliseconds as ``bad_input``, not
  after a 60-second Outlook start — and never as a half-filled draft.
- **The sender's own address is always blind-copied.** The BCC copy landing in
  the Inbox is what the caller files afterwards, so a draft that silently lacks
  it would never be filed. An address that cannot be resolved stops the run
  before a draft exists (``resolve_self_address``).
- **Unknown spec keys are refused**, because a misspelt ``atachments`` would
  otherwise drop the attachment without a word.
- **The ref token is reported honestly.** ``ref_header`` says whether the
  ``X-Archive-Ref`` header was actually stamped, with the reason when not.
- **A reply is Outlook's own.** ``reply_to`` names the mail to answer (an Inbox
  item by Message-ID, or a saved ``.msg``) and the client builds the draft from
  ``Reply()`` / ``ReplyAll()``, so the thread link, the quote, the ``Re:`` subject
  and the recipients are Outlook's. ``to`` / ``cc`` / ``subject`` become optional
  overrides; a mail that cannot be found is ``reply_source_not_found`` and no
  draft exists. The document reports what was replied to and whether the draft
  holds the In-Reply-To, as honestly as ``ref_header``.
- **An update edits the same item** (``update``): a caller iterating on one
  mail gets one draft, not a trail of near-duplicates. Only an unsent item in
  Drafts whose body the tool marked is touched; anything else is refused before
  the first write.
"""
from __future__ import annotations

import html
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from email_archiver.batch import SCHEMA_VERSION, now_iso
from email_archiver.config import Mailbox, get_outlook_self_address
from email_archiver.outlook.drafts import (
    REPLY_BY_MESSAGE_ID,
    REPLY_BY_MSG_PATH,
    ReplyTarget,
)
from email_archiver.text import normalize_message_id

VERB = "draft"

REF_STAMPED = "stamped"
REF_NOT_STAMPED = "not_stamped"
REF_REASON_NONE_GIVEN = "no ref given"

THREAD_NOT_APPLICABLE = "not_a_reply"
THREAD_UNCHANGED = "unchanged"

_SPEC_KEYS = frozenset({
    "to", "cc", "bcc", "subject", "body_text", "body_html", "attachments", "ref", "display",
    "reply_to", "reply_all",
})
_BLANK_LINES = re.compile(r"\n\s*\n")


class SpecError(ValueError):
    """The caller's draft spec is unusable; reported as ``bad_input``."""


@dataclass
class DraftSpec:
    """A validated draft spec. ``attachments`` are absolute paths to files.

    ``to`` / ``cc`` / ``subject`` are ``None`` only on a reply, where they were
    not given and Outlook's own stay; a new mail always has all three.
    """

    to: list[str] | None
    subject: str | None
    body_html: str
    cc: list[str] | None = field(default_factory=list)
    bcc: list[str] = field(default_factory=list)
    attachments: list[str] = field(default_factory=list)
    ref: str | None = None
    display: bool = True
    reply_to: ReplyTarget | None = None


# ------------------------------------------------------------------- spec ---

def _address_list(data: dict[str, Any], key: str, *, required: bool = False) -> list[str]:
    value = data.get(key, [])
    if not isinstance(value, list) or not all(isinstance(v, str) for v in value):
        raise SpecError(f"`{key}` must be a list of address strings")
    addresses = [v.strip() for v in value if v.strip()]
    if len(addresses) != len(value):
        raise SpecError(f"`{key}` contains a blank address")
    if required and not addresses:
        raise SpecError(f"`{key}` needs at least one address")
    return addresses


def text_to_html(text: str) -> str:
    """Minimal HTML for a plain-text body: escaped, paragraphs and line breaks kept.

    Quotes stay literal (``quote=False``): element content doesn't need them
    escaped, and Outlook stores ``&#x27;`` back as ``'``, which would leave
    ``html_to_text`` unable to read the saved draft.
    """
    normalised = text.replace("\r\n", "\n").replace("\r", "\n").strip("\n")
    paragraphs = [p for p in _BLANK_LINES.split(normalised) if p.strip()]
    return "".join(
        "<p>" + "<br>".join(html.escape(line, quote=False) for line in p.split("\n")) + "</p>"
        for p in paragraphs
    )


_PARAGRAPH = re.compile(r"<p>(.*?)</p>", re.DOTALL)


def html_to_text(body_html: str) -> str | None:
    """The plain text ``text_to_html`` turned into ``body_html``, or ``None``.

    An exact inverse, not a renderer: it answers only for HTML that
    ``text_to_html`` could have written — proven by encoding the answer again
    and getting the same HTML back — so a body anyone reshaped reads as
    ``None`` rather than as a lossy approximation.
    """
    paragraphs = _PARAGRAPH.findall(body_html)
    if "".join(f"<p>{p}</p>" for p in paragraphs) != body_html:
        return None
    text = "\n\n".join(
        "\n".join(html.unescape(line) for line in p.split("<br>")) for p in paragraphs
    )
    return text if text_to_html(text) == body_html else None


def _reply_target(data: dict[str, Any]) -> ReplyTarget | None:
    """The validated ``reply_to`` / ``reply_all`` of a spec, or ``None``."""
    reply_all = data.get("reply_all", False)
    if not isinstance(reply_all, bool):
        raise SpecError("`reply_all` must be true or false")
    raw = data.get("reply_to")
    if raw is None:
        if reply_all:
            raise SpecError("`reply_all` needs `reply_to`")
        return None
    if not isinstance(raw, dict):
        raise SpecError("`reply_to` must be an object with `message_id` or `msg_path`")
    unknown = sorted(set(raw) - {REPLY_BY_MESSAGE_ID, REPLY_BY_MSG_PATH})
    if unknown:
        raise SpecError(f"unknown `reply_to` keys: {', '.join(unknown)}")
    kinds = [k for k in (REPLY_BY_MESSAGE_ID, REPLY_BY_MSG_PATH) if k in raw]
    if len(kinds) != 1:
        raise SpecError("`reply_to` needs exactly one of `message_id` or `msg_path`")
    kind, value = kinds[0], raw[kinds[0]]
    if not isinstance(value, str) or not value.strip():
        raise SpecError(f"`reply_to.{kind}` must be a non-empty string")
    if kind == REPLY_BY_MESSAGE_ID:
        value = normalize_message_id(value)
        if not value:
            raise SpecError("`reply_to.message_id` must be a non-empty string")
    else:
        path = Path(value.strip())
        if not path.is_absolute():
            raise SpecError(f"`reply_to.msg_path` must be an absolute path: {value}")
        if path.suffix.lower() != ".msg" or not path.is_file():
            raise SpecError(f"`reply_to.msg_path` is not an existing .msg file: {value}")
        value = str(path)
    return ReplyTarget(kind=kind, value=value, reply_all=reply_all)


def parse_spec(data: Any) -> DraftSpec:
    """Validate a decoded spec and return it as a :class:`DraftSpec`.

    Raises:
        SpecError: with a message naming the offending field.
    """
    if not isinstance(data, dict):
        raise SpecError("the spec must be a JSON object")
    unknown = sorted(set(data) - _SPEC_KEYS)
    if unknown:
        raise SpecError(f"unknown spec keys: {', '.join(unknown)}")

    reply_to = _reply_target(data)
    # A reply takes its recipients and subject from Outlook; a key that is
    # present overrides, one that is absent leaves Outlook's own.
    to = _address_list(data, "to", required=True) if reply_to is None or "to" in data else None
    cc = _address_list(data, "cc") if reply_to is None or "cc" in data else None
    bcc = _address_list(data, "bcc")

    subject = data.get("subject")
    if not (reply_to is not None and subject is None) and (
        not isinstance(subject, str) or not subject.strip()
    ):
        raise SpecError("`subject` must be a non-empty string")

    bodies = [k for k in ("body_text", "body_html") if k in data]
    if len(bodies) != 1:
        raise SpecError("exactly one of `body_text` or `body_html` is required")
    body = data[bodies[0]]
    if not isinstance(body, str):
        raise SpecError(f"`{bodies[0]}` must be a string")
    body_html = text_to_html(body) if bodies[0] == "body_text" else body

    raw_attachments = data.get("attachments", [])
    if not isinstance(raw_attachments, list) or not all(
        isinstance(a, str) and a.strip() for a in raw_attachments
    ):
        raise SpecError("`attachments` must be a list of file paths")
    attachments: list[str] = []
    for raw in raw_attachments:
        path = Path(raw)
        if not path.is_file():
            raise SpecError(f"attachment is not an existing file: {raw}")
        attachments.append(str(path.resolve()))

    ref = data.get("ref")
    if ref is not None and (not isinstance(ref, str) or not ref.strip()):
        raise SpecError("`ref` must be a non-empty string when given")

    display = data.get("display", True)
    if not isinstance(display, bool):
        raise SpecError("`display` must be true or false")

    return DraftSpec(
        to=to, cc=cc, bcc=bcc, subject=subject, body_html=body_html,
        attachments=attachments, ref=ref.strip() if ref else None, display=display,
        reply_to=reply_to,
    )


def load_spec(path: str) -> DraftSpec:
    """Read and validate a spec file. Every failure is a :class:`SpecError`."""
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, json.JSONDecodeError) as exc:
        raise SpecError(f"{path}: {exc}") from exc
    return parse_spec(data)


# ---------------------------------------------------------- self address ---

def resolve_self_address(
    cfg: dict[str, Any], client: Any, mailbox: Mailbox | None = None,
) -> str | None:
    """The address every draft blind-copies, or ``None`` when it cannot be known.

    A registry ``mailbox`` blind-copies its own address: that copy lands in
    its Inbox, which is where it is filed from (issue #109). With no registry,
    ``outlook.self_address`` wins when set; otherwise the default sending
    account's SMTP address as Outlook reports it
    (``OutlookClient.default_account_smtp``, which never goes through
    ``GetExchangeUser``). Anything without an ``@`` is not an address and
    resolves to ``None`` — the caller must then refuse to create the draft.
    """
    if mailbox is not None and not mailbox.synthesized:
        return mailbox.address
    address = get_outlook_self_address(cfg) or (client.default_account_smtp() or "").strip()
    return address if "@" in address else None


def with_self_bcc(bcc: list[str], self_address: str) -> list[str]:
    """The BCC list with ``self_address`` present exactly once, order kept.

    Deduplicated case-insensitively, so a caller that already blind-copied
    itself (in any spelling) does not get a second copy. The self address stays
    in BCC even when it is also a To/CC recipient: the BCC line is the
    convention the filing step relies on.
    """
    merged: list[str] = []
    seen: set[str] = set()
    for address in [*bcc, self_address]:
        key = address.casefold()
        if key not in seen:
            seen.add(key)
            merged.append(address)
    return merged


# ------------------------------------------------------------------ draft ---

def create(client: Any, spec: DraftSpec, self_address: str) -> dict[str, Any]:
    """Create the draft through ``client`` and return the ``draft`` document.

    The envelope (``verb`` / ``schema_version`` / ``generated_at``) matches the
    other batch documents and ``batch.error_document``, so a consumer parses
    one shape whichever verb it spawned.
    """
    bcc = with_self_bcc(spec.bcc, self_address)
    created = client.create_draft(
        to=spec.to, cc=spec.cc, bcc=bcc, subject=spec.subject,
        body_html=spec.body_html, attachments=spec.attachments,
        ref=spec.ref, display=spec.display, reply_to=spec.reply_to,
        self_address=self_address,
    )
    return _document(spec, bcc, created, updated=False)


def update(client: Any, entry_id: str, spec: DraftSpec, self_address: str) -> dict[str, Any]:
    """Re-fill the existing draft ``entry_id`` from ``spec``; the same document
    as :func:`create`, with ``updated: true`` and ``updated_at``.

    A ``reply_to`` in the spec is not re-applied: the draft keeps the thread
    link and quote it was created with, so a caller may reuse its create spec.

    Raises:
        DraftUpdateError: the item is gone, sent, outside Drafts, or has no
            marked body region — raised by the client before any write.
    """
    bcc = with_self_bcc(spec.bcc, self_address)
    updated = client.update_draft(
        entry_id=entry_id, to=spec.to, cc=spec.cc, bcc=bcc, subject=spec.subject,
        body_html=spec.body_html, attachments=spec.attachments,
        ref=spec.ref, display=spec.display,
    )
    return _document(spec, bcc, updated, updated=True)


def _reply_fields(spec: DraftSpec, result: Any, *, updated: bool) -> dict[str, Any]:
    """The reply part of the document: what was replied to, and the thread link.

    ``thread_header`` is ``set`` / ``not_set`` as read back from the saved draft,
    ``unchanged`` on an update (the link was made at creation) and
    ``not_a_reply`` for a new mail, with ``thread_header_reason`` saying why not.
    ``recipients_from`` says who a created reply was addressed from: ``sender``
    (Outlook's own reply), ``original_recipients`` (the original was sent by the
    user) or ``caller`` (a ``to`` was given).
    """
    if spec.reply_to is None:
        return {"reply_to": None, "thread_header": THREAD_NOT_APPLICABLE, "thread_header_reason": ""}
    fields: dict[str, Any] = {
        "reply_to": {
            spec.reply_to.kind: spec.reply_to.value,
            "reply_all": spec.reply_to.reply_all,
        },
    }
    if updated:
        return {**fields, "thread_header": THREAD_UNCHANGED, "thread_header_reason": ""}
    return {
        **fields,
        "replied_to_message_id": result.replied_to_message_id,
        "thread_header": result.thread_header,
        "thread_header_reason": result.thread_header_reason,
        "recipients_from": result.recipients_from,
    }


def _reported(given: list[str] | None, saved: list[str], replying: bool) -> list[str] | None:
    """A recipient line for the document: the caller's when given, the saved
    reply's when Outlook computed it, ``None`` when an update left it as it was."""
    if given is not None:
        return list(given)
    return list(saved) if replying else None


def _document(spec: DraftSpec, bcc: list[str], result: Any, *, updated: bool) -> dict[str, Any]:
    if spec.ref is None:
        ref_reason = REF_REASON_NONE_GIVEN
    else:
        ref_reason = "" if result.ref_stamped else result.ref_reason
    now = now_iso()
    replying = spec.reply_to is not None and not updated
    return {
        "verb": VERB,
        "schema_version": SCHEMA_VERSION,
        "generated_at": now,
        "entry_id": result.entry_id,
        # The account the draft sends from, for the caller's preview; "" when
        # Outlook would not say.
        "from_address": result.from_address,
        # A reply's own subject and recipients are what Outlook saved; ``None``
        # is an update that left them as the draft has them.
        "subject": result.subject if replying and spec.subject is None else spec.subject,
        "to": _reported(spec.to, result.to, replying),
        "cc": _reported(spec.cc, result.cc, replying),
        "bcc": bcc,
        "attachments": list(spec.attachments),
        "ref": spec.ref,
        "ref_header": REF_STAMPED if result.ref_stamped else REF_NOT_STAMPED,
        "ref_header_reason": ref_reason,
        **_reply_fields(spec, result, updated=updated),
        "displayed": result.displayed,
        "updated": updated,
        "updated_at" if updated else "created_at": now,
    }

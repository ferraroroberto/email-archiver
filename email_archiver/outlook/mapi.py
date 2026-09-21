"""
Outlook vocabulary shared by every COM surface, and the helpers that need no COM.

Constants (folder, class and recipient-type numbers, DASL property names, the
MAPI error code batch mode retries on), ``OutlookUnavailableError``, and the
pure string/HTML helpers the batch and draft surfaces build on: category
arithmetic, header parsing, the ``--since`` filter and the marked draft-body
region. Nothing here imports win32com, so all of it tests without Outlook.
"""
from __future__ import annotations

import re
from datetime import datetime, timezone
from typing import Any


class OutlookUnavailableError(RuntimeError):
    """Outlook could not be reached — batch mode cannot start.

    Distinct from "there is nothing to do": this means the COM object model
    never answered, so the caller must exit non-zero rather than report an
    empty Inbox it never actually read.
    """


def safe_com(read: Any, default: Any) -> Any:
    """Read one COM property, returning ``default`` if it raises.

    COM properties on a real mailbox fail for reasons that are not this app's
    business (an item still downloading, a store that denies a property, an
    add-in that broke one). A single missing field must degrade that field, not
    abort a batch run over the whole Inbox.
    """
    try:
        return read()
    except Exception:
        return default


# olFolderInbox. Batch mode reads exactly one folder and files out of it.
OL_FOLDER_INBOX = 6
# MailItem. Anything else in the Inbox (meeting requests, reports) is skipped.
OL_CLASS_MAIL_ITEM = 43
# DASL name for MAPI PR_INTERNET_MESSAGE_ID (0x1035001F), used both to read the
# id off a live item and to build the server-side Restrict filter.
DASL_INTERNET_MESSAGE_ID = (
    "http://schemas.microsoft.com/mapi/proptag/0x1035001F"
)
# DASL name for MAPI PidTagFlagStatus (0x1090), the same property the scanner
# reads back out of the archived .msg — so a plan and a later scan agree.
DASL_FLAG_STATUS = "http://schemas.microsoft.com/mapi/proptag/0x10900003"

# olMailItem, for Application.CreateItem. The draft verb creates nothing else.
OL_MAIL_ITEM = 0
# olFolderDrafts: the only folder a draft is updated, read or sent from.
OL_FOLDER_DRAFTS = 16
# OlMailRecipientType, for reading a draft's recipients back by line.
OL_TO = 1
OL_CC = 2
OL_BCC = 3
# olSave, for Inspector.Close: keep what the open window holds.
OL_SAVE = 0
# DASL name of an Internet header in the PS_INTERNET_HEADERS property set. A
# draft stamped with it carries ``X-Archive-Ref: <token>`` into the sent mail, so
# a caller can recognise its own copy when it lands back in the Inbox.
DASL_X_ARCHIVE_REF = (
    "http://schemas.microsoft.com/mapi/string/"
    "{00020386-0000-0000-C000-000000000046}/X-Archive-Ref"
)
ARCHIVE_REF_HEADER = "X-Archive-Ref"
# DASL name for MAPI PR_TRANSPORT_MESSAGE_HEADERS (0x007D001F): the raw header
# block of a mail that arrived through a server. It is where the Step 3 probe
# read X-Archive-Ref back on an IMAP mailbox (email-archiver#70).
DASL_TRANSPORT_HEADERS = "http://schemas.microsoft.com/mapi/proptag/0x007D001F"
# DASL name for the received time, used by the `plan --since` Restrict filter.
DASL_DATE_RECEIVED = "urn:schemas:httpmail:datereceived"


# MAPI_E_OBJECT_CHANGED. Outlook raises it from MailItem.Move (and Delete, and
# Save) when it considers the in-memory item to have been modified since it was
# obtained — which SaveAs can leave an item as, on the very reference the
# archiver just wrote to disk. Batch mode tells this one refusal apart from
# every other move failure because it is the only one worth retrying.
MAPI_E_OBJECT_CHANGED = 0x80040109


def is_message_changed_error(exc: Any) -> bool:
    """Whether a COM failure is MAPI_E_OBJECT_CHANGED (0x80040109).

    ``pywintypes.com_error`` carries the interesting code in its ``excepinfo``
    tuple (``args[2][5]``) rather than in the HRESULT, which for a scripted
    Outlook call is the generic ``DISP_E_EXCEPTION``. So every integer in
    ``args``, one level of nesting deep, is compared — masked to 32 bits, since
    the same code is spelled both signed (``-2147221239``) and unsigned
    depending on where it came from.

    Deliberately keyed on the code and never on the message text: the message
    is localised, and matching "has been changed" would silently stop matching
    on a non-English Outlook.
    """
    codes: list[int] = []
    for arg in getattr(exc, "args", ()) or ():
        if isinstance(arg, int):
            codes.append(arg)
        elif isinstance(arg, tuple):
            codes.extend(inner for inner in arg if isinstance(inner, int))
    return any(code & 0xFFFFFFFF == MAPI_E_OBJECT_CHANGED for code in codes)


def split_categories(raw: str | None) -> list[str]:
    """Split Outlook's semicolon-separated Categories string into a list."""
    if not raw:
        return []
    return [part.strip() for part in str(raw).split(";") if part.strip()]


def join_categories(names: list[str]) -> str:
    """Join category names back into the form Outlook stores."""
    return "; ".join(names)


def with_category(raw: str | None, name: str) -> str:
    """Return the Categories string with ``name`` present exactly once.

    Kept pure and separate from the COM call so the add/remove arithmetic — the
    part that can silently drop a user's own categories — is unit-testable
    without Outlook. Existing categories are preserved in order.
    """
    names = split_categories(raw)
    if not any(n.casefold() == name.casefold() for n in names):
        names.append(name)
    return join_categories(names)


def without_category(raw: str | None, name: str) -> str:
    """Return the Categories string with ``name`` removed, others untouched."""
    return join_categories(
        [n for n in split_categories(raw) if n.casefold() != name.casefold()]
    )


def header_value(headers: str | None, name: str) -> str:
    """The value of the first ``name:`` header in a raw header block, or ``""``.

    Matched case-insensitively, with folded continuation lines (RFC 5322: a
    line starting with whitespace continues the previous header) joined back.
    Pure so the parsing tests without Outlook.
    """
    prefix = f"{name.casefold()}:"
    value: str | None = None
    for line in (headers or "").splitlines():
        if value is not None:
            if line[:1] in (" ", "\t"):
                value += " " + line.strip()
                continue
            break
        if line.casefold().startswith(prefix):
            value = line[len(prefix):].strip()
    return value or ""


def received_since_filter(since: datetime) -> str:
    """The ``Items.Restrict`` DASL filter for mail received at/after ``since``.

    ``since`` is naive local time (or aware). DASL compares date literals in
    **UTC**: a probe against a real Inbox with boundaries placed a minute
    either side of real mails matched the client-side count exactly with the
    UTC literal and was off by the whole UTC offset with the local one.
    """
    utc = since.astimezone(timezone.utc)
    return f"@SQL=\"{DASL_DATE_RECEIVED}\" >= '{utc:%Y-%m-%d %H:%M}'"


_BODY_OPEN_TAG = re.compile(r"<body\b[^>]*>", re.IGNORECASE)


# The caller's body sits between these two, so `update_draft` can replace
# exactly what `create_draft` wrote and leave everything after it — the
# signature Outlook inserted — alone. The closing comment, not the bare
# `</div>`, is what ends the region: a caller's own HTML may nest divs.
DRAFT_BODY_OPEN = '<div id="archive-draft-body">'
DRAFT_BODY_CLOSE = "</div><!--/archive-draft-body-->"
_DRAFT_BODY_OPEN_TAG = re.compile(r"""<div\s+id=["']?archive-draft-body["']?\s*>""", re.IGNORECASE)
_DRAFT_BODY_CLOSE_TAG = re.compile(r"</div>\s*<!--\s*/archive-draft-body\s*-->", re.IGNORECASE)


def mark_body_html(body_html: str) -> str:
    """``body_html`` wrapped in the markers ``replace_marked_body_html`` finds."""
    return f"{DRAFT_BODY_OPEN}{body_html}{DRAFT_BODY_CLOSE}"


def _marked_region_span(existing: str) -> tuple[re.Match, re.Match] | None:
    """The opening and (last) closing marker of the body region, or ``None``."""
    opened = _DRAFT_BODY_OPEN_TAG.search(existing)
    if opened is None:
        return None
    closes = list(_DRAFT_BODY_CLOSE_TAG.finditer(existing, opened.end()))
    return (opened, closes[-1]) if closes else None


def marked_body_region(existing_html: str | None) -> str | None:
    """The HTML inside the marked body region, or ``None`` without one —
    bounded exactly as ``replace_marked_body_html`` bounds what it replaces."""
    existing = existing_html or ""
    span = _marked_region_span(existing)
    return None if span is None else existing[span[0].end():span[1].start()]


def recipient_address(recipient: Any) -> str:
    """The SMTP address of one draft recipient, or ``""`` when it has none.

    A recipient Outlook has not resolved yet carries the typed address in
    ``Name`` with an empty ``Address`` (observed live on a fresh draft), so
    ``Address``, then ``AddressEntry.Address``, then ``Name`` are tried, and
    the first that holds an ``@`` wins. ``GetExchangeUser`` is never called:
    it can raise the address-book security modal.
    """
    for read in (
        lambda: recipient.Address,
        lambda: recipient.AddressEntry.Address,
        lambda: recipient.Name,
    ):
        value = str(safe_com(read, "") or "").strip()
        if "@" in value:
            return value
    return ""


def replace_marked_body_html(existing_html: str | None, body_html: str) -> str | None:
    """``existing_html`` with its marked body region replaced by ``body_html``.

    ``None`` when the region cannot be found — a draft created before the
    markers existed, or one whose HTML Outlook rewrote. The caller refuses
    then: replacing a guessed region could eat the signature or text the user
    typed. The last closing marker is used, so a body that itself quotes the
    marker text still ends where the tool's own region ends. Pure, like
    ``insert_body_html``.
    """
    existing = existing_html or ""
    span = _marked_region_span(existing)
    if span is None:
        return None
    return f"{existing[:span[0].start()]}{mark_body_html(body_html)}{existing[span[1].end():]}"


def insert_body_html(existing_html: str | None, body_html: str) -> str:
    """Return ``existing_html`` with ``body_html`` inserted at the top of its body.

    ``existing_html`` is what Outlook put in a new draft's ``HTMLBody`` when it
    displayed it, which is where the account's default signature lives. Putting
    the caller's text straight after the ``<body>`` tag keeps that signature
    below the message, where the user expects it. With no ``<body>`` tag to
    anchor on (the draft was never displayed, or the store returned nothing) the
    body is wrapped in a document of its own. Pure so it tests without Outlook.
    """
    match = _BODY_OPEN_TAG.search(existing_html or "")
    if match is None:
        return f"<html><body>{body_html}</body></html>"
    return f"{existing_html[:match.end()]}{body_html}{existing_html[match.end():]}"

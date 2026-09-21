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
from email_archiver.config import get_outlook_self_address

VERB = "draft"

REF_STAMPED = "stamped"
REF_NOT_STAMPED = "not_stamped"
REF_REASON_NONE_GIVEN = "no ref given"

_SPEC_KEYS = frozenset(
    {"to", "cc", "bcc", "subject", "body_text", "body_html", "attachments", "ref", "display"}
)
_BLANK_LINES = re.compile(r"\n\s*\n")


class SpecError(ValueError):
    """The caller's draft spec is unusable; reported as ``bad_input``."""


@dataclass
class DraftSpec:
    """A validated draft spec. ``attachments`` are absolute paths to files."""

    to: list[str]
    subject: str
    body_html: str
    cc: list[str] = field(default_factory=list)
    bcc: list[str] = field(default_factory=list)
    attachments: list[str] = field(default_factory=list)
    ref: str | None = None
    display: bool = True


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
    """Minimal HTML for a plain-text body: escaped, paragraphs and line breaks kept."""
    normalised = text.replace("\r\n", "\n").replace("\r", "\n").strip("\n")
    paragraphs = [p for p in _BLANK_LINES.split(normalised) if p.strip()]
    return "".join(
        "<p>" + "<br>".join(html.escape(line) for line in p.split("\n")) + "</p>"
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

    to = _address_list(data, "to", required=True)
    cc = _address_list(data, "cc")
    bcc = _address_list(data, "bcc")

    subject = data.get("subject")
    if not isinstance(subject, str) or not subject.strip():
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

def resolve_self_address(cfg: dict[str, Any], client: Any) -> str | None:
    """The address every draft blind-copies, or ``None`` when it cannot be known.

    ``outlook.self_address`` wins when set; otherwise the default sending
    account's SMTP address as Outlook reports it
    (``OutlookClient.default_account_smtp``, which never goes through
    ``GetExchangeUser``). Anything without an ``@`` is not an address and
    resolves to ``None`` — the caller must then refuse to create the draft.
    """
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
        ref=spec.ref, display=spec.display,
    )
    return _document(spec, bcc, created, updated=False)


def update(client: Any, entry_id: str, spec: DraftSpec, self_address: str) -> dict[str, Any]:
    """Re-fill the existing draft ``entry_id`` from ``spec``; the same document
    as :func:`create`, with ``updated: true`` and ``updated_at``.

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


def _document(spec: DraftSpec, bcc: list[str], result: Any, *, updated: bool) -> dict[str, Any]:
    if spec.ref is None:
        ref_reason = REF_REASON_NONE_GIVEN
    else:
        ref_reason = "" if result.ref_stamped else result.ref_reason
    now = now_iso()
    return {
        "verb": VERB,
        "schema_version": SCHEMA_VERSION,
        "generated_at": now,
        "entry_id": result.entry_id,
        "subject": spec.subject,
        "to": list(spec.to),
        "cc": list(spec.cc),
        "bcc": bcc,
        "attachments": list(spec.attachments),
        "ref": spec.ref,
        "ref_header": REF_STAMPED if result.ref_stamped else REF_NOT_STAMPED,
        "ref_header_reason": ref_reason,
        "displayed": result.displayed,
        "updated": updated,
        "updated_at" if updated else "created_at": now,
    }

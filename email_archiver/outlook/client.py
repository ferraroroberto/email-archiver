"""
Outlook COM integration layer.

Design decisions:
- All COM calls are isolated here; no other module imports win32com.
- OutlookClient.get_selected_email() returns a plain EmailData dataclass,
  so the rest of the app never touches COM objects after this layer.
- SMTP address resolution handles Exchange (on-prem / O365) where
  SenderEmailAddress may return a cryptic X.500/EX address instead of SMTP.
- is_running() checks the process list without starting Outlook, which is
  important for the fast-launch requirement. ensure_running() is its deliberate
  opposite, used only by batch mode, which has no user to open Outlook for it.
  It also owns the lifetime of what it starts: an outlook.exe that never
  publishes its COM object is terminated by the handle ensure_running() holds,
  never by image name, so an Outlook it merely attached to is out of reach.
- The batch surface (ensure_running / iter_inbox / iter_inbox_received_since /
  find_by_message_id / archive_ref / refetch / save_item / move_to /
  set_category / clear_category / is_open_in_inspector) lives here
  too, so batch.py stays pure orchestration and can be driven by a fake client
  in tests.
- The draft surface (default_account_smtp / create_draft / update_draft /
  open_draft / read_draft / snapshot_draft / close_inspectors_of) follows the
  same split for email_archiver/draft.py and email_archiver/send.py. No code
  path in this module sends mail: the one ``Send`` call lives in
  email_archiver/outlook/sending.py, reached only by the ``send`` verb.
"""
from __future__ import annotations

import hashlib
import logging
import os
import re
import tempfile
import time
from collections.abc import Iterator
from dataclasses import dataclass
from datetime import datetime, timezone
from typing import Any

from email_archiver.text import clean_subject as _clean_subject
from email_archiver.text import normalize_message_id

logger = logging.getLogger(__name__)


# ----------------------------------------------------------------- types ----

@dataclass
class EmailData:
    subject: str = ""
    sender: str = ""
    recipients: str = ""           # semicolon-separated SMTP addresses
    date_sent: datetime | None = None


@dataclass
class InboxMail:
    """One live Outlook mail, as batch mode sees it.

    Everything the batch verbs need is read off the COM item once, here, so the
    orchestration layer never touches COM. ``item`` is the only COM reference
    that escapes, and only because the archiver needs the real MailItem to call
    ``SaveAs``.
    """

    message_id: str = ""           # normalised; "" when the mail carries none
    entry_id: str = ""
    subject: str = ""              # cleaned (Re:/Fwd: stripped)
    sender: str = ""
    recipients: str = ""
    date_sent: datetime | None = None
    # ReceivedTime, naive local like date_sent. Not reported in a plan; it is
    # what `plan --since` compares against.
    date_received: datetime | None = None
    body_preview: str = ""
    attachment_count: int = 0
    flag_status: int = 0
    item: Any = None


class OutlookUnavailableError(RuntimeError):
    """Outlook could not be reached — batch mode cannot start.

    Distinct from "there is nothing to do": this means the COM object model
    never answered, so the caller must exit non-zero rather than report an
    empty Inbox it never actually read.
    """


# -------------------------------------------------------------- helpers -----


def _safe_com(read: Any, default: Any) -> Any:
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


def sent_datetime(mail_item: Any) -> datetime | None:
    """The mail's sent time as a stdlib datetime, or ``None``.

    Tries ``SentOn`` (when it was actually sent) before ``ReceivedTime``.
    ``archiver._get_sent_date_prefix`` formats this same value for its
    ``YYYY-MM-DD -`` filename prefix, so the date a plan reports and the date
    a prefix carries cannot disagree. pywintypes datetimes are rebuilt field
    by field rather than passed through, which is what the caller sees as a
    plain ``datetime``.
    """
    for attr in ("SentOn", "ReceivedTime"):
        value = _com_datetime(mail_item, attr)
        if value is not None:
            return value
    return None


def _com_datetime(mail_item: Any, attr: str) -> datetime | None:
    """One COM date property as a naive stdlib datetime, or ``None``."""
    value = _safe_com(lambda: getattr(mail_item, attr), None)
    if value is None:
        return None
    try:
        return datetime(
            value.year, value.month, value.day,
            value.hour, value.minute, value.second,
        )
    except (AttributeError, TypeError, ValueError):
        return None


def _resolve_smtp(address_entry: Any) -> str:
    """
    Attempt to resolve an AddressEntry to its primary SMTP address.
    Falls back to the raw Address string if Exchange resolution fails.
    """
    try:
        ex_user = address_entry.GetExchangeUser()
        if ex_user is not None:
            smtp = ex_user.PrimarySmtpAddress
            if smtp:
                return smtp
    except Exception:
        pass

    try:
        return address_entry.Address or ""
    except Exception:
        return ""


def _get_sender_smtp(mail_item: Any) -> str:
    """Return the sender's SMTP address, resolving Exchange entries."""
    try:
        addr_type = mail_item.SenderEmailType
        if addr_type == "EX":
            smtp = _resolve_smtp(mail_item.Sender)
            if smtp:
                return smtp
    except Exception:
        pass
    try:
        return mail_item.SenderEmailAddress or ""
    except Exception:
        return ""


def _get_recipients_smtp(mail_item: Any) -> str:
    """Return semicolon-separated SMTP addresses for all recipients."""
    addresses: list[str] = []
    try:
        for recipient in mail_item.Recipients:
            try:
                addr_type = recipient.AddressEntry.AddressEntryUserType
                # 0 = Exchange user, resolve via GetExchangeUser
                if addr_type == 0:
                    smtp = _resolve_smtp(recipient.AddressEntry)
                    addresses.append(smtp or recipient.Address)
                else:
                    addresses.append(recipient.Address or "")
            except Exception:
                try:
                    addresses.append(recipient.Address or "")
                except Exception:
                    pass
    except Exception as exc:
        logger.warning("Error reading recipients: %s", exc)
    return "; ".join(a for a in addresses if a)


def _tasklist_has_image(image_name: str) -> bool:
    """
    Return True if `image_name` (e.g. "OUTLOOK.EXE") appears in the Windows
    process list, using the `tasklist` console tool.

    A failed query is *not* the same fact as "the process is not running", so
    every failure is logged before falling back to False — otherwise a broken
    query is indistinguishable from a quiet machine.
    """
    import subprocess  # noqa: PLC0415
    import sys  # noqa: PLC0415

    is_windows = sys.platform == "win32"
    try:
        result = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {image_name}", "/NH"],
            capture_output=True, timeout=5,
            # tasklist writes the OEM code page, not the parent's locale.
            # text=True decodes with the ambient locale, which under
            # PYTHONUTF8=1 is UTF-8: the OEM bytes then fail to decode, stdout
            # comes back None, and the result silently reads as "not running".
            encoding="oem" if is_windows else "utf-8",
            errors="replace",
            creationflags=subprocess.CREATE_NO_WINDOW if is_windows else 0,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        logger.warning(
            "Could not run tasklist to check for %s (%s: %s) - "
            "treating as not running.", image_name, type(exc).__name__, exc,
        )
        return False

    if result.returncode != 0:
        logger.warning(
            "tasklist exited %s while checking for %s (%s) - "
            "treating as not running.",
            result.returncode, image_name, (result.stderr or "").strip(),
        )
        return False

    if not (result.stdout or "").strip():
        logger.warning(
            "tasklist returned no output while checking for %s - cannot tell "
            "whether it is running; treating as not running.", image_name,
        )
        return False

    return image_name.upper() in result.stdout.upper()


# --------------------------------------------------- batch-mode helpers ----

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

# How long ensure_running() waits for a freshly launched Outlook to publish its
# COM object. Outlook's first start on a cold profile is genuinely slow.
DEFAULT_START_TIMEOUT_SECONDS = 60.0
_POLL_INTERVAL_SECONDS = 1.0
# How long the teardown waits for a terminated Outlook to actually go away
# before giving up and saying so in the error it raises.
_TERMINATE_WAIT_SECONDS = 10.0

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
        value = str(_safe_com(read, "") or "").strip()
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


@dataclass
class CreatedDraft:
    """What Outlook reports back about a draft ``create_draft`` saved."""

    entry_id: str = ""
    ref_stamped: bool = False
    ref_reason: str = ""           # why the ref header is absent; "" when stamped
    displayed: bool = False


@dataclass
class DraftAttachment:
    """One attachment of a draft as stored: its bytes, not Outlook's
    ``Attachment.Size``, which counts MAPI overhead and never equals the file."""

    name: str
    size_bytes: int
    sha256: str


@dataclass
class DraftSnapshot:
    """An unsent draft read back from the store, as plain data.

    ``to`` / ``cc`` / ``bcc`` are the addresses on each line, in the order
    Outlook lists them; ``unreadable_recipients`` counts recipients with no
    readable address or line, which a caller must treat as a fact of its own,
    never as a shorter list. ``body_region_html`` is the marked region
    ``create_draft`` wrote, or ``None`` when the markers are gone.
    """

    entry_id: str
    subject: str
    to: list[str]
    cc: list[str]
    bcc: list[str]
    html_body: str
    body_region_html: str | None
    attachments: list[DraftAttachment]
    unreadable_recipients: int = 0


# Why `open_draft` refused. Each is its own `error.code` in `main_batch.py`,
# and each is raised before the item is changed.
DRAFT_NOT_FOUND = "draft_not_found"          # the EntryID no longer resolves
DRAFT_NOT_EDITABLE = "draft_not_editable"    # sent, or not in Drafts
DRAFT_BODY_UNMARKED = "draft_body_unmarked"  # no marked body region to replace
_UNMARKED_MESSAGE = (
    "the draft has no marked body region (created before draft --update existed, "
    "or its HTML was rewritten); refusing to guess where the body ends"
)


class DraftUpdateError(Exception):
    """An existing draft that cannot be updated; the item is left untouched."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


def _outlook_executable() -> str | None:
    """Path to the registered outlook.exe, or None when it cannot be found.

    Reads the ``App Paths`` registry entry Windows itself uses to resolve
    ``outlook.exe``, falling back to a PATH lookup. Deliberately not
    ``Dispatch("Outlook.Application")``: Dispatch starts a *hidden* instance
    that behaves differently (no explorer, add-ins in a different state) and
    that the user cannot see or interact with.
    """
    import shutil  # noqa: PLC0415
    import sys  # noqa: PLC0415

    if sys.platform == "win32":
        try:
            import winreg  # noqa: PLC0415

            for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                try:
                    with winreg.OpenKey(
                        hive,
                        r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"
                        r"\OUTLOOK.EXE",
                    ) as key:
                        path, _ = winreg.QueryValueEx(key, "")
                        if path:
                            return str(path).strip('"')
                except OSError:
                    continue
        except ImportError:  # pragma: no cover - winreg ships with CPython
            pass

    return shutil.which("outlook.exe")


def _get_active_application() -> Any | None:
    """Return the running Outlook Application, or None if none is published.

    ``GetActiveObject`` only ever attaches to an Outlook the user (or we)
    started; unlike ``Dispatch`` it never starts one, which is what makes it
    safe to call in a poll loop.
    """
    try:
        import win32com.client  # noqa: PLC0415

        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        return None


def _spawn_outlook(exe: str) -> Any:
    """Start ``exe`` and return the handle that owns the started process.

    A seam on purpose: the teardown in ``ensure_running()`` is unit-tested
    against a fake spawner, because the only Outlook on this machine is the
    user's own and no test may be able to reach it.
    """
    import subprocess  # noqa: PLC0415
    import sys  # noqa: PLC0415

    # Deliberately WITHOUT CREATE_NO_WINDOW: this is the one spawn in the
    # project whose window is meant to be visible — a hidden Outlook is
    # exactly what ensure_running() exists to avoid.
    return subprocess.Popen(  # noqa: S603
        [exe],
        creationflags=(
            subprocess.CREATE_NEW_PROCESS_GROUP
            if sys.platform == "win32"
            else 0
        ),
    )


def _terminate_spawned_outlook(proc: Any) -> str:
    """End the Outlook *this run started*, and report what happened to it.

    Takes the handle returned by ``_spawn_outlook`` and nothing else. Ending it
    through that handle rather than by image name or a PID lookup is the whole
    point: an Outlook that was already running — the user's own, on their
    own desktop — is not reachable from here, and a live handle keeps its
    PID reserved, so the call cannot land on a reused PID either.

    Returns one sentence for the ``OutlookUnavailableError`` message, so the
    ``outlook_unavailable`` error document records whether the process was
    cleaned up or is still out there.
    """
    pid = getattr(proc, "pid", None)
    try:
        if proc.poll() is not None:
            logger.info("The outlook.exe started (PID %s) had already exited.", pid)
            return f"The outlook.exe it started (PID {pid}) had already exited."
        proc.terminate()
        proc.wait(timeout=_TERMINATE_WAIT_SECONDS)
    except Exception as exc:
        logger.warning(
            "Could not terminate the outlook.exe started (PID %s): %s", pid, exc
        )
        return (
            f"The outlook.exe it started (PID {pid}) could NOT be terminated "
            f"({type(exc).__name__}: {exc}) and may still be running."
        )
    logger.warning("Terminated the outlook.exe this run started (PID %s).", pid)
    return f"The outlook.exe it started (PID {pid}) was terminated."


# ------------------------------------------------------------- client ------


def get_selected_mail_item() -> Any | None:
    """
    Return the currently selected Outlook MailItem via COM, or None if
    nothing is selected. Raises on COM errors so the caller can surface
    them appropriately.
    """
    import win32com.client  # noqa: PLC0415
    app = win32com.client.GetActiveObject("Outlook.Application")
    explorer = app.ActiveExplorer()
    if explorer is None or explorer.Selection.Count == 0:
        return None
    # Dispatch forces full COM interface resolution → MailItem with SaveAs
    return win32com.client.Dispatch(explorer.Selection.Item(1))


class OutlookClient:
    """
    Wraps Outlook COM automation.

    Lazy-initialised: COM objects are created only when needed, so importing
    this module is free (important for the fast-startup requirement).
    """

    def is_running(self) -> bool:
        """
        Check if Outlook.exe is in the process list without starting it.
        Uses psutil for reliability; falls back to a tasklist call.
        """
        try:
            import psutil
            return any(
                p.name().lower() == "outlook.exe"
                for p in psutil.process_iter(["name"])
            )
        except ImportError:
            pass

        # Fallback: subprocess tasklist (slower but no extra dep)
        return _tasklist_has_image("OUTLOOK.EXE")

    def get_selected_email(self) -> EmailData | None:
        """
        Return metadata for the currently selected email in Outlook.
        Returns None if Outlook is not running, no email is selected, or
        if the selected item is not a MailItem.
        """
        try:
            import win32com.client  # noqa: PLC0415
        except ImportError:
            logger.error("pywin32 not installed. Run: pip install pywin32")
            return None

        if not self.is_running():
            logger.warning("Outlook is not running.")
            return None

        try:
            item = get_selected_mail_item()
        except Exception as exc:
            logger.error("Cannot access Outlook or selection: %s", exc)
            return None

        if item is None:
            logger.warning("No email selected in Outlook.")
            return None

        try:
            if item.Class != OL_CLASS_MAIL_ITEM:
                logger.warning(
                    "Selected item is not a MailItem (class=%s).", item.Class
                )
                return None

        except Exception as exc:
            logger.error("Error accessing Outlook selection: %s", exc)
            return None

        try:
            return EmailData(
                subject=_clean_subject(item.Subject),
                sender=_get_sender_smtp(item),
                recipients=_get_recipients_smtp(item),
                date_sent=sent_datetime(item),
            )

        except Exception as exc:
            logger.error("Error extracting email metadata: %s", exc)
            return None

    # ------------------------------------------------------ batch surface ---
    #
    # Used only by email_archiver/batch.py. Everything below keeps COM inside
    # this class: the batch layer receives InboxMail dataclasses and opaque
    # item handles, so it can be driven end to end by a fake client.

    def ensure_running(
        self, timeout: float = DEFAULT_START_TIMEOUT_SECONDS
    ) -> Any:
        """Return the Outlook Application object, starting Outlook if needed.

        The opposite of ``is_running()``, and used only by batch mode, which
        runs unattended and has nobody to open Outlook for it. Starts the
        *registered* ``outlook.exe`` (visible, with its explorer, exactly as
        the user's own shortcut would) and then polls ``GetActiveObject`` until
        the COM object appears or ``timeout`` elapses.

        When the wait times out, the process this call started is terminated
        before the error is raised — see ``_terminate_spawned_outlook``. An
        Outlook that was already running is only ever attached to, never
        started and never terminated.

        Raises:
            OutlookUnavailableError: Outlook could not be started or never
                published its COM object inside the timeout. A loud failure on
                purpose — every batch verb needs Outlook, so continuing would
                report an empty Inbox nobody ever read. The message says what
                became of the process this call started.
        """
        app = _get_active_application()
        if app is not None:
            return app

        exe = _outlook_executable()
        if exe is None:
            raise OutlookUnavailableError(
                "Outlook is not running and outlook.exe could not be located "
                "(no App Paths registry entry and not on PATH)."
            )

        logger.info("Outlook is not running; starting %s", exe)
        try:
            proc = _spawn_outlook(exe)
        except OSError as exc:
            raise OutlookUnavailableError(
                f"Could not start Outlook ({exe}): {exc}"
            ) from exc

        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            time.sleep(_POLL_INTERVAL_SECONDS)
            app = _get_active_application()
            if app is not None:
                logger.info("Outlook is up.")
                return app

        # Nothing else owns this process's lifetime. Left running it outlives
        # the run — on a scheduled unattended run, an invisible orphan
        # holding the profile and OST against the user's own Outlook, one more
        # every time the wait times out (#78).
        teardown = _terminate_spawned_outlook(proc)
        raise OutlookUnavailableError(
            f"Outlook was started but did not publish its COM object within "
            f"{timeout:.0f}s. It may be showing a profile or password prompt, "
            f"or have no desktop to show one on. {teardown}"
        )

    def _namespace(self) -> Any:
        """Return the MAPI namespace of the running Outlook."""
        app = _get_active_application()
        if app is None:
            raise OutlookUnavailableError(
                "Outlook is no longer reachable over COM."
            )
        return app.GetNamespace("MAPI")

    def _inbox(self) -> Any:
        return self._namespace().GetDefaultFolder(OL_FOLDER_INBOX)

    def _named_folder(self, name: str, *, create: bool = True) -> Any:
        """Return the folder ``name`` directly under the mailbox root.

        Matched case-insensitively so a mailbox whose Archive folder is spelled
        differently is still found rather than duplicated. Created when absent
        and ``create`` is set, because ``apply`` must have somewhere to move a
        mail it has already written to disk.
        """
        root = self._inbox().Parent
        folders = root.Folders
        for i in range(1, folders.Count + 1):
            folder = folders.Item(i)
            if str(folder.Name).casefold() == name.casefold():
                return folder
        if not create:
            return None
        logger.info("Creating Outlook folder %r under %s", name, root.Name)
        return folders.Add(name)

    def _resolve_folder(self, folder_name: str | None) -> Any:
        """Inbox for ``None``, else the named folder under the mailbox root."""
        return self._inbox() if folder_name is None else self._named_folder(folder_name)

    def iter_inbox(self, preview_len: int = 500) -> Iterator[InboxMail]:
        """Yield every MailItem in the Inbox as an :class:`InboxMail`.

        ``preview_len`` is the scanner's ``scanning.body_preview_length`` so the
        preview a plan shows and the preview the index stores are the same
        string. Non-mail items (meeting requests, delivery reports) are skipped:
        they have no ``SaveAs``-able shape this app archives.
        """
        yield from self._iter_items(self._inbox().Items, preview_len)

    def iter_inbox_received_since(
        self, since: datetime, preview_len: int = 500
    ) -> Iterator[InboxMail]:
        """Yield the Inbox mails received at or after ``since``.

        Narrowed server-side with ``Items.Restrict`` so a large Inbox is not
        walked. The literal is minute-precise, so the result can be a little
        wider than ``since`` — never narrower — and ``batch.plan`` applies the
        exact check to every mail either way. A store that rejects the filter
        falls back to the whole Inbox for that same check; which path ran is
        logged.
        """
        items = self._inbox().Items
        flt = received_since_filter(since)
        try:
            restricted = items.Restrict(flt)
        except Exception as exc:
            logger.warning(
                "The store rejected the received-date filter (%s: %s); walking "
                "the whole Inbox and checking each mail's date instead.",
                type(exc).__name__, exc,
            )
            yield from self._iter_items(items, preview_len)
            return
        logger.info("Received-date filter applied server-side: %s", flt)
        yield from self._iter_items(restricted, preview_len)

    def _iter_items(self, items: Any, preview_len: int) -> Iterator[InboxMail]:
        import win32com.client  # noqa: PLC0415

        # Index rather than `for item in items`: the collection is live, and a
        # positional walk keeps the enumeration stable while the mails sit
        # still (batch mode never moves anything during a plan).
        for i in range(1, items.Count + 1):
            try:
                raw = items.Item(i)
                if raw.Class != OL_CLASS_MAIL_ITEM:
                    continue
                # Dispatch forces full interface resolution → a MailItem with
                # SaveAs, the same reason get_selected_mail_item does it.
                yield self._read_mail(win32com.client.Dispatch(raw), preview_len)
            except Exception as exc:
                logger.warning("Skipping unreadable item %d in folder: %s", i, exc)

    def read_mail(self, item: Any, preview_len: int = 0) -> InboxMail:
        """Read one already-obtained MailItem into an :class:`InboxMail`.

        ``apply`` needs the mail's cleaned subject for the filename after it has
        found the item by Message-ID; ``preview_len`` defaults to 0 because that
        path has no use for the body.
        """
        return self._read_mail(item, preview_len)

    def _read_mail(self, item: Any, preview_len: int) -> InboxMail:
        """Read every field batch mode needs off one COM MailItem."""
        return InboxMail(
            message_id=self._message_id(item),
            entry_id=_safe_com(lambda: str(item.EntryID), ""),
            subject=_clean_subject(_safe_com(lambda: item.Subject, "")),
            sender=_get_sender_smtp(item),
            recipients=_get_recipients_smtp(item),
            date_sent=sent_datetime(item),
            date_received=_com_datetime(item, "ReceivedTime"),
            body_preview=_safe_com(lambda: (item.Body or "")[:preview_len], ""),
            attachment_count=_safe_com(lambda: int(item.Attachments.Count), 0),
            flag_status=self._flag_status(item),
            item=item,
        )

    @staticmethod
    def _message_id(item: Any) -> str:
        """The mail's normalised Internet Message-ID, or ``""`` when it has none."""
        return normalize_message_id(
            _safe_com(
                lambda: item.PropertyAccessor.GetProperty(DASL_INTERNET_MESSAGE_ID),
                None,
            )
        )

    @staticmethod
    def _flag_status(item: Any) -> int:
        """Outlook's follow-up flag, read from the same MAPI property the
        scanner reads back out of the archived .msg (0x1090). Absent means
        unflagged, so 0 is a real answer, not a swallowed error."""
        value = _safe_com(
            lambda: item.PropertyAccessor.GetProperty(DASL_FLAG_STATUS), None
        )
        try:
            return int(value or 0)
        except (TypeError, ValueError):
            return 0

    def find_by_message_id(
        self, message_id: str, folder_name: str | None = None
    ) -> Any | None:
        """Return the MailItem with this Message-ID in a folder, or ``None``.

        ``folder_name`` is ``None`` for the Inbox, else a folder under the
        mailbox root. An empty ``message_id`` never matches — a mail with no
        Message-ID has no identity to search by, and matching every such mail
        against each other would file the wrong one.

        Tries a server-side ``Restrict`` first (an Archive folder holds
        thousands of mails), falling back to a linear walk when the filter
        cannot be built or the store rejects it.
        """
        if not message_id:
            return None
        folder = self._resolve_folder(folder_name)
        if folder is None:
            return None

        import win32com.client  # noqa: PLC0415

        # A quote inside the id would break out of the DASL string literal;
        # such ids are vanishingly rare and the linear walk handles them.
        if "'" not in message_id and '"' not in message_id:
            for candidate in (f"<{message_id}>", message_id):
                try:
                    found = folder.Items.Find(
                        f"@SQL=\"{DASL_INTERNET_MESSAGE_ID}\" = '{candidate}'"
                    )
                except Exception as exc:
                    logger.debug("Restrict lookup unavailable (%s); walking.", exc)
                    break
                if found is not None:
                    return win32com.client.Dispatch(found)

        # Read *only* the Message-ID while walking. Going through _read_mail
        # here would resolve every sender and recipient through
        # GetExchangeUser, which is the one call that can raise Outlook's
        # address-book security modal — turning a lookup that failed to use the
        # fast path into a hung process.
        items = folder.Items
        for i in range(1, items.Count + 1):
            try:
                raw = items.Item(i)
                if raw.Class != OL_CLASS_MAIL_ITEM:
                    continue
                if self._message_id(raw) == message_id:
                    return win32com.client.Dispatch(raw)
            except Exception as exc:
                logger.warning("Skipping unreadable item %d during lookup: %s", i, exc)
        return None

    def archive_ref(self, item: Any) -> str:
        """The mail's ``X-Archive-Ref`` header value, or ``""`` when absent.

        Read from the transport headers first — where the header survives on a
        mail that came back through a server — then from the named Internet
        header property ``create_draft`` stamps, which some stores promote the
        header into. Only ``PropertyAccessor`` reads: no address is resolved,
        so no ``GetExchangeUser`` security modal. Whether the header survives
        sending depends on the account; a caller keeps a fallback match.
        """
        accessor = _safe_com(lambda: item.PropertyAccessor, None)
        if accessor is None:
            return ""
        headers = _safe_com(lambda: accessor.GetProperty(DASL_TRANSPORT_HEADERS), "")
        value = header_value(headers, ARCHIVE_REF_HEADER)
        if value:
            return value
        return str(_safe_com(lambda: accessor.GetProperty(DASL_X_ARCHIVE_REF), "") or "").strip()

    def refetch(self, item: Any) -> Any | None:
        """Re-acquire a mail from the store by its ``EntryID``, or ``None``.

        ``MailItem.SaveAs`` can leave the in-memory item flagged as modified,
        and a later ``Move`` on *that* reference is then refused with
        MAPI_E_OBJECT_CHANGED — the mail stays in the Inbox with its files
        already on disk (issue #59). A reference read back through
        ``Session.GetItemFromID`` is a fresh object carrying none of that
        state, so it is what batch mode moves.

        ``None`` (never an exception) when the item has no readable EntryID or
        the store cannot hand it back: failing to re-acquire is a reason to
        move the original reference, not a reason to strand the mail.
        """
        entry_id = _safe_com(lambda: str(item.EntryID), "")
        if not entry_id:
            logger.warning("Cannot re-acquire a mail with no readable EntryID.")
            return None
        try:
            import win32com.client  # noqa: PLC0415

            fresh = self._namespace().GetItemFromID(entry_id)
        except Exception as exc:
            logger.warning("GetItemFromID failed for %s: %s", entry_id, exc)
            return None
        if fresh is None:
            return None
        return win32com.client.Dispatch(fresh)

    def save_item(self, item: Any) -> None:
        """Commit a mail's pending changes.

        Its own method rather than an inline ``item.Save()`` in batch mode:
        ``batch.py`` touches no COM object, and this is the call that clears
        the modified flag behind a MAPI_E_OBJECT_CHANGED before the one retry.
        """
        item.Save()

    def move_to(self, item: Any, folder_name: str | None) -> Any:
        """Move a mail to a folder and return the moved item.

        ``folder_name`` is ``None`` for the Inbox. The returned object is the
        item *in its new home* — Outlook rewrites ``EntryID`` on a move, so the
        caller must read the new id off this object and never off the original
        reference, which now points at nothing.
        """
        target = self._resolve_folder(folder_name)
        if target is None:
            raise OutlookUnavailableError(
                f"Outlook folder {folder_name!r} could not be resolved."
            )
        import win32com.client  # noqa: PLC0415

        return win32com.client.Dispatch(item.Move(target))

    def set_category(self, item: Any, name: str) -> None:
        """Add a category to a mail, preserving any it already carries."""
        item.Categories = with_category(_safe_com(lambda: item.Categories, ""), name)
        item.Save()

    def clear_category(self, item: Any, name: str) -> None:
        """Remove one category from a mail, leaving the others alone."""
        item.Categories = without_category(_safe_com(lambda: item.Categories, ""), name)
        item.Save()

    def entry_id(self, item: Any) -> str:
        """The mail's current EntryID (valid only for its current folder)."""
        return _safe_com(lambda: str(item.EntryID), "")

    def is_open_in_inspector(self, item: Any) -> bool | None:
        """Whether this mail is open in an Outlook window. ``None`` = unknown.

        An inspector holds its item open for the life of the window, and every
        write to that item is refused with MAPI_E_OBJECT_CHANGED meanwhile —
        including on a reference re-acquired by EntryID, which is the same
        underlying object, and including after a ``Save()``. So both of batch
        mode's move defences fail by construction on an open mail, and the
        only useful thing left to do is say so (issue #76).

        Three genuinely different answers, never two:

        ``True``   this mail was found in an open inspector — close it;
        ``False``  the inspectors were all read and none holds this mail;
        ``None``   the question could not be answered — no readable EntryID,
                   Outlook unreachable, or an inspector that would not be read.

        ``None`` is deliberately not folded into ``False``: a check that could
        not run is not evidence the window is closed, and reporting it as one
        would point the next operator at the wrong remedy.
        """
        entry_id = _safe_com(lambda: str(item.EntryID), "")
        if not entry_id:
            return None
        app = _get_active_application()
        if app is None:
            return None
        try:
            inspectors = app.Inspectors
            count = int(inspectors.Count)
        except Exception as exc:
            logger.warning("Could not read Outlook's open windows: %s", exc)
            return None

        unread = False
        for index in range(1, count + 1):
            current = _safe_com(
                lambda i=index: inspectors.Item(i).CurrentItem, None
            )
            if current is None:
                unread = True
                continue
            open_id = _safe_com(lambda c=current: str(c.EntryID), "")
            if not open_id:
                # An unsaved compose window has no EntryID. It cannot be the
                # mail being filed, so it is not a gap in the answer.
                continue
            if open_id == entry_id:
                return True
        # One window that would not be read is enough to make "none of them
        # holds it" a guess rather than a finding.
        return None if unread else False

    # ------------------------------------------------------ draft surface ---
    #
    # Used only by email_archiver/draft.py. Nothing here sends: a draft is
    # saved and shown, and the user presses Send.

    def default_account_smtp(self) -> str:
        """The default sending account's SMTP address, or ``""`` when unknown.

        The account whose delivery store is the profile's default store is the
        one a new mail sends from; a profile with a single account needs no
        match. Read from ``Account.SmtpAddress`` and never through
        ``GetExchangeUser``, which can raise the address-book security modal.
        ``""`` (never a guess) when several accounts exist and none owns the
        default store, or the address read back is not an SMTP address.
        """
        namespace = self._namespace()
        accounts = namespace.Accounts
        count = _safe_com(lambda: int(accounts.Count), 0)
        default_store_id = _safe_com(lambda: str(namespace.DefaultStore.StoreID), "")
        candidates: list[str] = []
        for i in range(1, count + 1):
            account = _safe_com(lambda i=i: accounts.Item(i), None)
            if account is None:
                continue
            smtp = _safe_com(lambda a=account: str(a.SmtpAddress or "").strip(), "")
            candidates.append(smtp)
            store_id = _safe_com(lambda a=account: str(a.DeliveryStore.StoreID), "")
            if default_store_id and store_id == default_store_id:
                return smtp if "@" in smtp else ""
        if len(candidates) == 1 and "@" in candidates[0]:
            return candidates[0]
        logger.warning(
            "Could not tell the default sending account apart among %d account(s).",
            count,
        )
        return ""

    def create_draft(
        self,
        *,
        to: list[str],
        cc: list[str],
        bcc: list[str],
        subject: str,
        body_html: str,
        attachments: list[str],
        ref: str | None,
        display: bool,
    ) -> CreatedDraft:
        """Create, save and (optionally) show a new mail. Never sends it.

        Every step that can fail on caller data (attachments, the ref header)
        runs before ``Display``, so a failure leaves no half-filled window on
        the user's screen. ``Display`` comes before the body is written because
        opening the compose window is what makes Outlook insert the account's
        default signature; the body is then put above it. ``Save`` runs last,
        so the finished draft sits in Drafts even if the window is closed.
        """
        app = _get_active_application()
        if app is None:
            raise OutlookUnavailableError("Outlook is no longer reachable over COM.")

        mail = app.CreateItem(OL_MAIL_ITEM)
        mail.To = "; ".join(to)
        mail.CC = "; ".join(cc)
        mail.BCC = "; ".join(bcc)
        mail.Subject = subject
        for path in attachments:
            mail.Attachments.Add(path)

        created = CreatedDraft()
        if ref:
            try:
                mail.PropertyAccessor.SetProperty(DASL_X_ARCHIVE_REF, ref)
                created.ref_stamped = True
            except Exception as exc:
                created.ref_reason = f"SetProperty refused: {type(exc).__name__}: {exc}"
                logger.warning("Could not stamp X-Archive-Ref: %s", exc)

        existing_html = ""
        if display:
            mail.Display(False)  # non-modal: this process does not wait on it
            created.displayed = True
            existing_html = _safe_com(lambda: mail.HTMLBody or "", "")
            logger.info(
                "Draft displayed; Outlook's own body is %d chars, signature "
                "marker %s.", len(existing_html),
                "present" if "_MailAutoSig" in existing_html else "absent",
            )
        mail.HTMLBody = insert_body_html(existing_html, mark_body_html(body_html))
        mail.Save()
        created.entry_id = _safe_com(lambda: str(mail.EntryID), "")
        return created

    def update_draft(
        self,
        *,
        entry_id: str,
        to: list[str],
        cc: list[str],
        bcc: list[str],
        subject: str,
        body_html: str,
        attachments: list[str],
        ref: str | None,
        display: bool,
    ) -> CreatedDraft:
        """Re-fill the unsent draft ``entry_id`` in place. Never sends it.

        Every check runs before the first write, so a refusal
        (:class:`DraftUpdateError`) leaves the item exactly as it was. Only the
        marked body region is replaced; the signature below it stays. The
        attachments are all removed and the spec's re-added, so the draft ends
        up carrying exactly what the caller asked for.
        """
        mail = self.open_draft(entry_id)
        if replace_marked_body_html(_safe_com(lambda: mail.HTMLBody or "", ""), body_html) is None:
            # Checked on the stored item first too, so an unmarked draft is
            # refused without closing the user's open window on it.
            raise DraftUpdateError(DRAFT_BODY_UNMARKED, _UNMARKED_MESSAGE)
        if self.close_inspectors_of(entry_id):
            # The window held its own copy of the item: re-read what it saved,
            # or the region check below would run on stale HTML.
            mail = self.open_draft(entry_id)
        new_html = replace_marked_body_html(_safe_com(lambda: mail.HTMLBody or "", ""), body_html)
        if new_html is None:
            raise DraftUpdateError(DRAFT_BODY_UNMARKED, _UNMARKED_MESSAGE)

        mail.To = "; ".join(to)
        mail.CC = "; ".join(cc)
        mail.BCC = "; ".join(bcc)
        mail.Subject = subject
        while int(mail.Attachments.Count):
            mail.Attachments.Remove(1)
        for path in attachments:
            mail.Attachments.Add(path)

        updated = CreatedDraft()
        if ref:
            try:
                mail.PropertyAccessor.SetProperty(DASL_X_ARCHIVE_REF, ref)
                updated.ref_stamped = True
            except Exception as exc:
                updated.ref_reason = f"SetProperty refused: {type(exc).__name__}: {exc}"
                logger.warning("Could not stamp X-Archive-Ref: %s", exc)

        mail.HTMLBody = new_html
        mail.Save()
        if display:
            mail.Display(False)  # non-modal: this process does not wait on it
            updated.displayed = True
        updated.entry_id = _safe_com(lambda: str(mail.EntryID), "") or entry_id
        logger.info("Draft updated in place.")
        return updated

    def open_draft(self, entry_id: str) -> Any:
        """The unsent mail ``entry_id`` in Drafts, or :class:`DraftUpdateError`.

        The one lookup behind ``update_draft``, ``read_draft`` and the ``send``
        verb: by EntryID only, never by subject or search.
        """
        namespace = self._namespace()
        try:
            mail = namespace.GetItemFromID(entry_id)
        except Exception as exc:
            raise DraftUpdateError(
                DRAFT_NOT_FOUND, f"no Outlook item for EntryID {entry_id}: {type(exc).__name__}: {exc}"
            ) from exc
        if mail is None:
            raise DraftUpdateError(DRAFT_NOT_FOUND, f"no Outlook item for EntryID {entry_id}")
        if _safe_com(lambda: bool(mail.Sent), True):
            raise DraftUpdateError(DRAFT_NOT_EDITABLE, "the item has been sent; only an unsent draft is used")
        drafts_id = _safe_com(lambda: str(namespace.GetDefaultFolder(OL_FOLDER_DRAFTS).EntryID), "")
        parent_id = _safe_com(lambda: str(mail.Parent.EntryID), "")
        if not drafts_id or parent_id != drafts_id:
            raise DraftUpdateError(DRAFT_NOT_EDITABLE, "the item is not in the Drafts folder")
        return mail

    def read_draft(self, entry_id: str) -> DraftSnapshot:
        """The unsent draft ``entry_id`` as stored. Writes nothing, shows nothing.

        A window open on the draft keeps its own unsaved copy, which this does
        not see; the ``send`` verb closes such windows (saving) before it reads.
        """
        return self.snapshot_draft(self.open_draft(entry_id), entry_id)

    def snapshot_draft(self, mail: Any, entry_id: str) -> DraftSnapshot:
        """Read one draft item into a :class:`DraftSnapshot`.

        Attachment bytes are saved to a temporary directory, hashed and
        removed, so the snapshot binds their content and not only their name.
        An attachment that cannot be saved raises: a snapshot that silently
        skipped one would bind less than it claims.
        """
        lines: dict[int, list[str]] = {OL_TO: [], OL_CC: [], OL_BCC: []}
        unreadable = 0
        recipients = mail.Recipients
        for index in range(1, int(recipients.Count) + 1):
            recipient = recipients.Item(index)
            address = recipient_address(recipient)
            kind = _safe_com(lambda r=recipient: int(r.Type), 0)
            if not address or kind not in lines:
                unreadable += 1
                continue
            lines[kind].append(address)

        attachments: list[DraftAttachment] = []
        with tempfile.TemporaryDirectory(prefix="email-archiver-read-") as scratch:
            items = mail.Attachments
            for index in range(1, int(items.Count) + 1):
                item = items.Item(index)
                path = os.path.join(scratch, str(index))
                item.SaveAsFile(path)
                with open(path, "rb") as fh:
                    data = fh.read()
                attachments.append(DraftAttachment(
                    name=str(_safe_com(lambda i=item: i.FileName, "") or ""),
                    size_bytes=len(data), sha256=hashlib.sha256(data).hexdigest(),
                ))

        html_body = str(mail.HTMLBody or "")
        return DraftSnapshot(
            entry_id=entry_id, subject=str(mail.Subject or ""),
            to=lines[OL_TO], cc=lines[OL_CC], bcc=lines[OL_BCC],
            html_body=html_body, body_region_html=marked_body_region(html_body),
            attachments=attachments, unreadable_recipients=unreadable,
        )

    def close_inspectors_of(self, entry_id: str) -> int:
        """Close every open compose window showing ``entry_id``, saving it.

        An open window keeps its own copy of the draft: left open over an
        update it still shows the old text, and Send pressed there would send
        that. Closed with olSave so nothing the user typed is discarded before
        the update replaces it. Returns how many were closed.
        """
        app = _get_active_application()
        inspectors = _safe_com(lambda: app.Inspectors, None) if app is not None else None
        count = _safe_com(lambda: int(inspectors.Count), 0) if inspectors is not None else 0
        closed = 0
        for i in range(count, 0, -1):
            inspector = _safe_com(lambda i=i: inspectors.Item(i), None)
            if inspector is None:
                continue
            if _safe_com(lambda ins=inspector: str(ins.CurrentItem.EntryID), "") == entry_id:
                inspector.Close(OL_SAVE)
                closed += 1
        if closed:
            logger.info("Closed %d open window(s) on the draft before updating it.", closed)
        return closed

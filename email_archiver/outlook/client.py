"""
Outlook COM integration layer.

Design decisions:
- All COM calls are isolated here; no other module imports win32com.
- OutlookClient.get_selected_email() returns a plain EmailData dataclass,
  so the rest of the app never touches COM objects after this layer.
- SMTP address resolution handles Exchange (on-prem / O365) where
  SenderEmailAddress may return a cryptic X.500/EX address instead of SMTP.
- Split by concern (issue #88): this module is the read path and the batch
  surface. The Outlook process lifecycle is in outlook/process.py, the draft
  surface in outlook/drafts.py (a mixin OutlookClient inherits), and the
  vocabulary and COM-free helpers they share in outlook/mapi.py.
- The batch surface (ensure_running / iter_inbox / iter_inbox_received_since /
  find_by_message_id / archive_ref / refetch / save_item / move_to /
  set_category / clear_category / is_open_in_inspector) lives here
  too, so batch.py stays pure orchestration and can be driven by a fake client
  in tests.
- The draft surface (default_account_smtp / create_draft / update_draft /
  open_draft / read_draft / snapshot_draft / close_inspectors_of) is inherited
  from outlook/drafts.py, for email_archiver/draft.py and email_archiver/send.py.
"""
from __future__ import annotations

import logging
from collections.abc import Iterator
from dataclasses import dataclass
from datetime import datetime
from typing import Any

from email_archiver.outlook import process
from email_archiver.outlook.drafts import DraftSurface
from email_archiver.outlook.mapi import (
    ARCHIVE_REF_HEADER,
    DASL_FLAG_STATUS,
    DASL_INTERNET_MESSAGE_ID,
    DASL_TRANSPORT_HEADERS,
    DASL_X_ARCHIVE_REF,
    OL_CLASS_MAIL_ITEM,
    OL_FOLDER_INBOX,
    OutlookUnavailableError,
    header_value,
    received_since_filter,
    safe_com,
    with_category,
    without_category,
)
from email_archiver.outlook.process import DEFAULT_START_TIMEOUT_SECONDS
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


# -------------------------------------------------------------- helpers -----


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
    value = safe_com(lambda: getattr(mail_item, attr), None)
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


class OutlookClient(DraftSurface):
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
        return process.tasklist_has_image("OUTLOOK.EXE")

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

        See :func:`email_archiver.outlook.process.ensure_running`, which owns
        the start, the poll and the teardown of what it started.
        """
        return process.ensure_running(timeout)

    def _namespace(self) -> Any:
        """Return the MAPI namespace of the running Outlook."""
        app = process.get_active_application()
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
            entry_id=safe_com(lambda: str(item.EntryID), ""),
            subject=_clean_subject(safe_com(lambda: item.Subject, "")),
            sender=_get_sender_smtp(item),
            recipients=_get_recipients_smtp(item),
            date_sent=sent_datetime(item),
            date_received=_com_datetime(item, "ReceivedTime"),
            body_preview=safe_com(lambda: (item.Body or "")[:preview_len], ""),
            attachment_count=safe_com(lambda: int(item.Attachments.Count), 0),
            flag_status=self._flag_status(item),
            item=item,
        )

    @staticmethod
    def _message_id(item: Any) -> str:
        """The mail's normalised Internet Message-ID, or ``""`` when it has none."""
        return normalize_message_id(
            safe_com(
                lambda: item.PropertyAccessor.GetProperty(DASL_INTERNET_MESSAGE_ID),
                None,
            )
        )

    @staticmethod
    def _flag_status(item: Any) -> int:
        """Outlook's follow-up flag, read from the same MAPI property the
        scanner reads back out of the archived .msg (0x1090). Absent means
        unflagged, so 0 is a real answer, not a swallowed error."""
        value = safe_com(
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
        accessor = safe_com(lambda: item.PropertyAccessor, None)
        if accessor is None:
            return ""
        headers = safe_com(lambda: accessor.GetProperty(DASL_TRANSPORT_HEADERS), "")
        value = header_value(headers, ARCHIVE_REF_HEADER)
        if value:
            return value
        return str(safe_com(lambda: accessor.GetProperty(DASL_X_ARCHIVE_REF), "") or "").strip()

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
        entry_id = safe_com(lambda: str(item.EntryID), "")
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
        item.Categories = with_category(safe_com(lambda: item.Categories, ""), name)
        item.Save()

    def clear_category(self, item: Any, name: str) -> None:
        """Remove one category from a mail, leaving the others alone."""
        item.Categories = without_category(safe_com(lambda: item.Categories, ""), name)
        item.Save()

    def entry_id(self, item: Any) -> str:
        """The mail's current EntryID (valid only for its current folder)."""
        return safe_com(lambda: str(item.EntryID), "")

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
        entry_id = safe_com(lambda: str(item.EntryID), "")
        if not entry_id:
            return None
        app = process.get_active_application()
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
            current = safe_com(
                lambda i=index: inspectors.Item(i).CurrentItem, None
            )
            if current is None:
                unread = True
                continue
            open_id = safe_com(lambda c=current: str(c.EntryID), "")
            if not open_id:
                # An unsaved compose window has no EntryID. It cannot be the
                # mail being filed, so it is not a gap in the answer.
                continue
            if open_id == entry_id:
                return True
        # One window that would not be read is enough to make "none of them
        # holds it" a guess rather than a finding.
        return None if unread else False

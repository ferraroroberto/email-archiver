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
- The batch surface (ensure_running / iter_inbox / find_by_message_id /
  move_to / set_category / clear_category) lives here too, so batch.py stays
  pure orchestration and can be driven by a fake client in tests.
"""
from __future__ import annotations

import logging
import time
from collections.abc import Iterator
from dataclasses import dataclass
from datetime import datetime
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
    raw_item: Any = None           # the COM MailItem – only used by archiver


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


def _sent_datetime(mail_item: Any) -> datetime | None:
    """The mail's sent time as a stdlib datetime, or ``None``.

    Tries ``SentOn`` (when it was actually sent) before ``ReceivedTime``, the
    same order ``archiver._get_sent_date_prefix`` uses, so the date a plan
    reports and the date a ``YYYY-MM-DD -`` prefix carries cannot disagree.
    pywintypes datetimes are rebuilt field by field rather than passed through,
    which is what the caller sees as a plain ``datetime``.
    """
    for attr in ("SentOn", "ReceivedTime"):
        value = _safe_com(lambda a=attr: getattr(mail_item, a), None)
        if value is None:
            continue
        try:
            return datetime(
                value.year, value.month, value.day,
                value.hour, value.minute, value.second,
            )
        except (AttributeError, TypeError, ValueError):
            continue
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

# How long ensure_running() waits for a freshly launched Outlook to publish its
# COM object. Outlook's first start on a cold profile is genuinely slow.
DEFAULT_START_TIMEOUT_SECONDS = 60.0
_POLL_INTERVAL_SECONDS = 1.0


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
            # Ensure it is a MailItem (Class == 43)
            if item.Class != 43:
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
                date_sent=_sent_datetime(item),
                raw_item=item,
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

        Raises:
            OutlookUnavailableError: Outlook could not be started or never
                published its COM object inside the timeout. A loud failure on
                purpose — every batch verb needs Outlook, so continuing would
                report an empty Inbox nobody ever read.
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
        import subprocess  # noqa: PLC0415
        import sys  # noqa: PLC0415

        try:
            # Deliberately WITHOUT CREATE_NO_WINDOW: this is the one spawn in
            # the project whose window is meant to be visible — a hidden
            # Outlook is exactly what this method exists to avoid.
            subprocess.Popen(  # noqa: S603
                [exe],
                creationflags=(
                    subprocess.CREATE_NEW_PROCESS_GROUP
                    if sys.platform == "win32"
                    else 0
                ),
            )
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

        raise OutlookUnavailableError(
            f"Outlook was started but did not publish its COM object within "
            f"{timeout:.0f}s. It may be showing a profile or password prompt."
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
        yield from self._iter_folder(self._inbox(), preview_len)

    def _iter_folder(self, folder: Any, preview_len: int) -> Iterator[InboxMail]:
        import win32com.client  # noqa: PLC0415

        items = folder.Items
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
            date_sent=_sent_datetime(item),
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

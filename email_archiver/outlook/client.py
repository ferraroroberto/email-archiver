"""
Outlook COM integration layer.

Design decisions:
- All COM calls are isolated here; no other module imports win32com.
- OutlookClient.get_selected_email() returns a plain EmailData dataclass,
  so the rest of the app never touches COM objects after this layer.
- SMTP address resolution handles Exchange (on-prem / O365) where
  SenderEmailAddress may return a cryptic X.500/EX address instead of SMTP.
- is_running() checks the process list without starting Outlook, which is
  important for the fast-launch requirement.
"""
from __future__ import annotations

import logging
import re
from dataclasses import dataclass, field
from datetime import datetime
from typing import Any

logger = logging.getLogger(__name__)


# ----------------------------------------------------------------- types ----

@dataclass
class EmailData:
    subject: str = ""
    sender: str = ""
    recipients: str = ""           # semicolon-separated SMTP addresses
    date_sent: datetime | None = None
    body: str = ""
    raw_item: Any = None           # the COM MailItem – only used by archiver


# -------------------------------------------------------------- helpers -----

_RE_REPLY_PREFIX = re.compile(r"^\s*(re|rv|fwd?)\s*:?\s*", re.IGNORECASE)


def _clean_subject(raw: str | None) -> str:
    if not raw:
        return ""
    return _RE_REPLY_PREFIX.sub("", raw.strip()).strip()


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
        import subprocess
        try:
            result = subprocess.run(
                ["tasklist", "/FI", "IMAGENAME eq OUTLOOK.EXE", "/NH"],
                capture_output=True, text=True, timeout=5
            )
            return "OUTLOOK.EXE" in result.stdout.upper()
        except Exception:
            return False

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
            app = win32com.client.GetActiveObject("Outlook.Application")
        except Exception as exc:
            logger.error("Cannot connect to running Outlook instance: %s", exc)
            return None

        try:
            explorer = app.ActiveExplorer()
            if explorer is None:
                logger.warning("No active Outlook explorer window.")
                return None

            selection = explorer.Selection
            if selection.Count == 0:
                logger.warning("No email selected in Outlook.")
                return None

            # Dispatch forces win32com to resolve the full COM interface for
            # this object. Without it, late-bound dynamic dispatch returns a
            # generic Item wrapper that lacks MailItem-specific methods like
            # SaveAs, which causes AttributeError: Item.SaveAs at archive time.
            item = win32com.client.Dispatch(selection.Item(1))

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
            date_sent: datetime | None = None
            try:
                # ReceivedTime is a pywintypes datetime; convert to stdlib
                rt = item.ReceivedTime
                date_sent = datetime(
                    rt.year, rt.month, rt.day,
                    rt.hour, rt.minute, rt.second
                )
            except Exception:
                pass

            return EmailData(
                subject=_clean_subject(item.Subject),
                sender=_get_sender_smtp(item),
                recipients=_get_recipients_smtp(item),
                date_sent=date_sent,
                body=item.Body[:2000] if item.Body else "",
                raw_item=item,
            )

        except Exception as exc:
            logger.error("Error extracting email metadata: %s", exc)
            return None

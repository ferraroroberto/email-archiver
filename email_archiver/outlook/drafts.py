"""
The draft surface of :class:`~email_archiver.outlook.client.OutlookClient`.

``DraftSurface`` is a mixin: ``OutlookClient`` inherits these methods, so
email_archiver/draft.py and email_archiver/send.py keep calling them on the
client, and a fake client in the tests still stands in for the whole thing. It
relies on the host class for ``_namespace()``. No code path in this module
sends mail: the one ``Send`` call lives in email_archiver/outlook/sending.py,
reached only by the ``send`` verb.
"""
from __future__ import annotations

import hashlib
import logging
import os
import tempfile
from dataclasses import dataclass
from typing import Any

from email_archiver.outlook import process
from email_archiver.outlook.mapi import (
    DASL_X_ARCHIVE_REF,
    OL_BCC,
    OL_CC,
    OL_FOLDER_DRAFTS,
    OL_MAIL_ITEM,
    OL_SAVE,
    OL_TO,
    OutlookUnavailableError,
    insert_body_html,
    mark_body_html,
    marked_body_region,
    recipient_address,
    replace_marked_body_html,
    safe_com,
)

logger = logging.getLogger(__name__)


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


class DraftSurface:
    """Draft methods ``OutlookClient`` inherits; the host supplies ``_namespace()``.

    Used by email_archiver/draft.py and, for reading a draft back,
    email_archiver/send.py. Nothing here sends: a draft is saved and shown, and
    the user presses Send.
    """

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
        count = safe_com(lambda: int(accounts.Count), 0)
        default_store_id = safe_com(lambda: str(namespace.DefaultStore.StoreID), "")
        candidates: list[str] = []
        for i in range(1, count + 1):
            account = safe_com(lambda i=i: accounts.Item(i), None)
            if account is None:
                continue
            smtp = safe_com(lambda a=account: str(a.SmtpAddress or "").strip(), "")
            candidates.append(smtp)
            store_id = safe_com(lambda a=account: str(a.DeliveryStore.StoreID), "")
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
        app = process.get_active_application()
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
            existing_html = safe_com(lambda: mail.HTMLBody or "", "")
            logger.info(
                "Draft displayed; Outlook's own body is %d chars, signature "
                "marker %s.", len(existing_html),
                "present" if "_MailAutoSig" in existing_html else "absent",
            )
        mail.HTMLBody = insert_body_html(existing_html, mark_body_html(body_html))
        mail.Save()
        created.entry_id = safe_com(lambda: str(mail.EntryID), "")
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
        if replace_marked_body_html(safe_com(lambda: mail.HTMLBody or "", ""), body_html) is None:
            # Checked on the stored item first too, so an unmarked draft is
            # refused without closing the user's open window on it.
            raise DraftUpdateError(DRAFT_BODY_UNMARKED, _UNMARKED_MESSAGE)
        if self.close_inspectors_of(entry_id):
            # The window held its own copy of the item: re-read what it saved,
            # or the region check below would run on stale HTML.
            mail = self.open_draft(entry_id)
        new_html = replace_marked_body_html(safe_com(lambda: mail.HTMLBody or "", ""), body_html)
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
        updated.entry_id = safe_com(lambda: str(mail.EntryID), "") or entry_id
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
        if safe_com(lambda: bool(mail.Sent), True):
            raise DraftUpdateError(DRAFT_NOT_EDITABLE, "the item has been sent; only an unsent draft is used")
        drafts_id = safe_com(lambda: str(namespace.GetDefaultFolder(OL_FOLDER_DRAFTS).EntryID), "")
        parent_id = safe_com(lambda: str(mail.Parent.EntryID), "")
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
            kind = safe_com(lambda r=recipient: int(r.Type), 0)
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
                    name=str(safe_com(lambda i=item: i.FileName, "") or ""),
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
        app = process.get_active_application()
        inspectors = safe_com(lambda: app.Inspectors, None) if app is not None else None
        count = safe_com(lambda: int(inspectors.Count), 0) if inspectors is not None else 0
        closed = 0
        for i in range(count, 0, -1):
            inspector = safe_com(lambda i=i: inspectors.Item(i), None)
            if inspector is None:
                continue
            if safe_com(lambda ins=inspector: str(ins.CurrentItem.EntryID), "") == entry_id:
                inspector.Close(OL_SAVE)
                closed += 1
        if closed:
            logger.info("Closed %d open window(s) on the draft before updating it.", closed)
        return closed

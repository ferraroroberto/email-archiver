"""
The draft surface of :class:`~email_archiver.outlook.client.OutlookClient`.

``DraftSurface`` is a mixin: ``OutlookClient`` inherits these methods, so
email_archiver/draft.py and email_archiver/send.py keep calling them on the
client, and a fake client in the tests still stands in for the whole thing. It
relies on the host class for ``_namespace()`` and ``store()`` (the resolved
mailbox's store, which every folder here is taken from). No code path in this module
sends mail: the one ``Send`` call lives in email_archiver/outlook/sending.py,
reached only by the ``send`` verb.
"""
from __future__ import annotations

import hashlib
import logging
import os
import tempfile
from dataclasses import dataclass, field
from typing import Any

from email_archiver.outlook import process
from email_archiver.outlook.mapi import (
    AccountNotInOutlookError,
    DASL_IN_REPLY_TO_ID,
    DASL_INTERNET_MESSAGE_ID,
    DASL_X_ARCHIVE_REF,
    OL_BCC,
    OL_CC,
    OL_CLASS_MAIL_ITEM,
    OL_DISCARD,
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
from email_archiver.outlook.stores import find_account
from email_archiver.text import normalize_message_id

logger = logging.getLogger(__name__)

# IDispatch::Invoke's DISPATCH_PROPERTYPUTREF: how an object-valued property
# such as SendUsingAccount is assigned when pywin32's plain put is refused.
_DISPATCH_PROPERTYPUTREF = 8


REPLY_BY_MESSAGE_ID = "message_id"
REPLY_BY_MSG_PATH = "msg_path"

THREAD_SET = "set"
THREAD_NOT_SET = "not_set"

# Who a reply was addressed from. Outlook's own ``Reply()`` answers the
# original's sender, which on a mail the user sent is the user (issue #105).
RECIPIENTS_SENDER = "sender"                      # Outlook's reply: the original's sender
RECIPIENTS_ORIGINAL = "original_recipients"       # the original was sent by the user: its To / CC
RECIPIENTS_CALLER = "caller"                      # the caller gave `to`


@dataclass
class ReplyTarget:
    """The mail a draft answers: an Inbox item by Message-ID, or a saved ``.msg``.

    ``kind`` is :data:`REPLY_BY_MESSAGE_ID` or :data:`REPLY_BY_MSG_PATH`;
    ``value`` is the Message-ID or the absolute file path.
    """

    kind: str
    value: str
    reply_all: bool = False


@dataclass
class CreatedDraft:
    """What Outlook reports back about a draft ``create_draft`` saved.

    The reply fields are filled only for a reply: ``to`` / ``cc`` / ``subject``
    are what the saved draft actually carries (Outlook computed them, unless the
    caller overrode them), ``replied_to_message_id`` is the original's own id and
    ``thread_header`` says whether the draft holds it as its In-Reply-To.
    ``recipients_from`` says who the reply was addressed from (``RECIPIENTS_*``).
    """

    entry_id: str = ""
    ref_stamped: bool = False
    ref_reason: str = ""           # why the ref header is absent; "" when stamped
    displayed: bool = False
    subject: str = ""
    to: list[str] = field(default_factory=list)
    cc: list[str] = field(default_factory=list)
    replied_to_message_id: str = ""
    thread_header: str = ""        # THREAD_SET / THREAD_NOT_SET; "" for a new mail
    thread_header_reason: str = ""
    recipients_from: str = ""      # RECIPIENTS_*; "" for a new mail
    from_address: str = ""         # the account it sends from; "" when unreadable


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
    # SendUsingAccount's SMTP address; "" when unset or unreadable. Not part
    # of the fingerprint: `send` checks it against the mailbox on its own.
    sending_account: str = ""


# Why `open_draft` refused. Each is its own `error.code` in `main_batch.py`,
# and each is raised before the item is changed.
DRAFT_NOT_FOUND = "draft_not_found"          # the EntryID no longer resolves
DRAFT_NOT_EDITABLE = "draft_not_editable"    # sent, or not in Drafts
DRAFT_BODY_UNMARKED = "draft_body_unmarked"  # no marked body region to replace
_UNMARKED_MESSAGE = (
    "the draft has no marked body region (created before draft --update existed, "
    "or its HTML was rewritten); refusing to guess where the body ends"
)


# Why a reply could not be started. Each is its own `error.code` in
# `main_batch.py`, and each is raised before anything is saved.
REPLY_SOURCE_NOT_FOUND = "reply_source_not_found"  # the original cannot be found or opened
REPLY_UNAVAILABLE = "reply_unavailable"            # the original opened but Outlook would not reply to it


# The draft does not send as the mailbox it was made for: Outlook refused the
# account, or (send) the stored draft's account is not the mailbox's. Its own
# `error.code`; nothing is kept or sent (issue #109).
SENDING_ACCOUNT_MISMATCH = "sending_account_mismatch"


class DraftAccountError(Exception):
    """A draft that would not send as its mailbox; it was discarded unsaved."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


class ReplySourceError(Exception):
    """The mail to reply to is unusable; no draft was saved."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


class DraftUpdateError(Exception):
    """An existing draft that cannot be updated; the item is left untouched."""

    def __init__(self, code: str, message: str) -> None:
        super().__init__(message)
        self.code = code


class DraftSurface:
    """Draft methods ``OutlookClient`` inherits; the host supplies
    ``_namespace()``, ``store()`` and ``mailbox``.

    Used by email_archiver/draft.py and, for reading a draft back,
    email_archiver/send.py. Nothing here sends: a draft is saved and shown, and
    the user presses Send.
    """

    def default_account_smtp(self) -> str:
        """The SMTP address of the account delivering to this client's store,
        or ``""`` when unknown.

        With no registry the store is the profile's default one, whose account
        is the one a new mail sends from; a profile with a single account needs
        no match. Read from ``Account.SmtpAddress`` and never through
        ``GetExchangeUser``, which can raise the address-book security modal.
        ``""`` (never a guess) when several accounts exist and none owns the
        store, or the address read back is not an SMTP address.
        """
        accounts = self._namespace().Accounts
        count = safe_com(lambda: int(accounts.Count), 0)
        mailbox_store_id = safe_com(lambda: str(self.store().StoreID), "")
        candidates: list[str] = []
        for i in range(1, count + 1):
            account = safe_com(lambda i=i: accounts.Item(i), None)
            if account is None:
                continue
            smtp = safe_com(lambda a=account: str(a.SmtpAddress or "").strip(), "")
            candidates.append(smtp)
            store_id = safe_com(lambda a=account: str(a.DeliveryStore.StoreID), "")
            if mailbox_store_id and store_id == mailbox_store_id:
                return smtp if "@" in smtp else ""
        if len(candidates) == 1 and "@" in candidates[0]:
            return candidates[0]
        logger.warning(
            "Could not tell the default sending account apart among %d account(s).",
            count,
        )
        return ""

    def sending_account(self) -> Any | None:
        """The Outlook ``Account`` a draft for this client's mailbox sends from.

        ``None`` for the mailbox synthesized when there is no registry: a new
        mail keeps Outlook's own default account, exactly as before. A registry
        mailbox always names its account explicitly, matched on
        ``Account.SmtpAddress`` — never trusting Outlook to infer it.

        Raises:
            AccountNotInOutlookError: no account sends from the mailbox's address.
        """
        mailbox = getattr(self, "mailbox", None)
        if mailbox is None or mailbox.synthesized:
            return None
        account = find_account(self._namespace(), mailbox.address)
        if account is None:
            raise AccountNotInOutlookError(
                f"mailbox {mailbox.alias!r} is in Outlook but no account sends from its "
                "address, so a draft cannot send as it; nothing was created"
            )
        return account

    def create_draft(
        self,
        *,
        to: list[str] | None,
        cc: list[str] | None,
        bcc: list[str],
        subject: str | None,
        body_html: str,
        attachments: list[str],
        ref: str | None,
        display: bool,
        reply_to: ReplyTarget | None = None,
        self_address: str = "",
    ) -> CreatedDraft:
        """Create, save and (optionally) show a new mail. Never sends it.

        With ``reply_to`` the item is Outlook's own ``Reply()`` / ``ReplyAll()``
        of that mail instead of a blank one, so it carries the thread link, the
        quoted original, the ``Re:`` subject and the original's recipients.
        ``to`` / ``cc`` / ``subject`` of ``None`` keep what Outlook computed;
        anything given overrides it. The caller's body goes above the quote.
        Outlook answers the sender, so a reply to a mail ``self_address`` sent
        would come back addressed to ``self_address``: with no ``to`` given it
        is addressed to that mail's own recipients instead (see
        :meth:`_start_reply`).

        Every step that can fail on caller data (attachments, the ref header)
        runs before ``Display``, so a failure leaves no half-filled window on
        the user's screen. ``Display`` comes before the body is written because
        opening the compose window is what makes Outlook insert the account's
        default signature; the body is then put above it. ``Save`` runs last,
        so the finished draft sits in Drafts even if the window is closed.

        For a registry mailbox (issue #109) the item's ``SendUsingAccount`` is
        set to that mailbox's account right after it is created — before
        ``Display``, so the signature is that account's — and read back; a
        draft that would send as anyone else is discarded unsaved
        (:class:`DraftAccountError`). After ``Save`` it is checked to be in
        that mailbox's own Drafts, and moved there if Outlook put it elsewhere.
        """
        app = process.get_active_application()
        if app is None:
            raise OutlookUnavailableError("Outlook is no longer reachable over COM.")
        # Before anything exists: a mailbox with no account leaves nothing behind.
        account = self.sending_account()

        replied_to_id = ""
        recipients_from = ""
        if reply_to is None:
            mail = app.CreateItem(OL_MAIL_ITEM)
        else:
            mail, replied_to_id, recipients_from = self._start_reply(
                reply_to, self_address=self_address, to_given=to is not None,
            )
        if account is not None:
            _send_as(mail, account, self.mailbox.address)
        if to is not None:
            mail.To = "; ".join(to)
        if cc is not None:
            mail.CC = "; ".join(cc)
        mail.BCC = "; ".join(bcc)
        if subject is not None:
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
        elif reply_to is not None:
            existing_html = safe_com(lambda: mail.HTMLBody or "", "")
        if reply_to is not None and not existing_html:
            # Writing the body without Outlook's own HTML would replace the
            # quote with nothing: refuse before anything is saved.
            safe_com(lambda: mail.Close(OL_DISCARD), None)
            raise ReplySourceError(
                REPLY_UNAVAILABLE,
                "Outlook's reply carries no readable body, so the quoted original "
                "cannot be kept; nothing was saved",
            )
        mail.HTMLBody = insert_body_html(existing_html, mark_body_html(body_html))
        mail.Save()
        if account is not None:
            mail = self._into_mailbox_drafts(mail, display)
        created.entry_id = safe_com(lambda: str(mail.EntryID), "")
        # A synthesized mailbox's new mail keeps Outlook's default account,
        # which may not read back on a fresh item: report the account owning
        # the store instead. A registry mailbox's was proven by _send_as.
        created.from_address = sending_address(mail) or (
            safe_com(self.default_account_smtp, "") if account is None else ""
        )
        if reply_to is not None:
            self._describe_reply(mail, created, replied_to_id)
            created.recipients_from = recipients_from
        return created

    def _into_mailbox_drafts(self, mail: Any, display: bool) -> Any:
        """``mail`` (just saved) in this mailbox's own Drafts.

        Outlook decides where a saved draft lands; for an account other than
        the default one that may be the default store's Drafts, where
        ``read`` / ``send`` for this mailbox would refuse it. Moved over then,
        closing (saving) its window first and showing it again after.
        """
        drafts = self.store().GetDefaultFolder(OL_FOLDER_DRAFTS)
        drafts_id = safe_com(lambda: str(drafts.EntryID), "")
        if drafts_id and safe_com(lambda: str(mail.Parent.EntryID), "") == drafts_id:
            return mail
        logger.info("Outlook saved the draft outside this mailbox's Drafts; moving it there.")
        entry_id = safe_com(lambda: str(mail.EntryID), "")
        if display and entry_id:
            self.close_inspectors_of(entry_id)
        moved = mail.Move(drafts)
        if display:
            moved.Display(False)
        return moved

    def _start_reply(
        self, target: ReplyTarget, *, self_address: str = "", to_given: bool = False,
    ) -> tuple[Any, str, str]:
        """Outlook's reply item to ``target``, the original's Message-ID, and
        who the reply is addressed from (``RECIPIENTS_*``).

        Raises :class:`ReplySourceError` when the original cannot be found or
        opened, or Outlook will not reply to it (a ``.msg`` opened outside any
        store may have nowhere to reply from). ``Reply()`` only builds an
        unsaved item, so a refusal here has left nothing behind. An original
        opened from a file is closed again, discarding, once the reply exists.

        ``Reply()`` answers the original's sender. When the original was sent
        by ``self_address`` (a saved sent mail, or one in Sent Items) that
        addresses the reply to the user, so the To / CC are taken from the
        original's own recipients instead — Reply All keeps its CC, a plain
        reply does not. Unreadable original recipients refuse the reply
        (``reply_unavailable``) rather than leave it addressed to the user; a
        caller that gave ``to`` has overridden the recipients and is not asked.
        """
        original = self._open_reply_source(target)
        try:
            replied_to_id = normalize_message_id(safe_com(
                lambda: original.PropertyAccessor.GetProperty(DASL_INTERNET_MESSAGE_ID), None,
            ))
            try:
                mail = original.ReplyAll() if target.reply_all else original.Reply()
            except Exception as exc:
                raise ReplySourceError(
                    REPLY_UNAVAILABLE,
                    f"Outlook could not reply to the original: {type(exc).__name__}: {exc}",
                ) from exc
            if mail is None:
                raise ReplySourceError(REPLY_UNAVAILABLE, "Outlook returned no reply item for the original")
            recipients_from = RECIPIENTS_CALLER if to_given else RECIPIENTS_SENDER
            if not to_given and self_address and _addressed_to(mail, self_address):
                self._readdress_to_original_recipients(original, mail, target.reply_all)
                recipients_from = RECIPIENTS_ORIGINAL
        finally:
            if target.kind == REPLY_BY_MSG_PATH:
                safe_com(lambda: original.Close(OL_DISCARD), None)
        return mail, replied_to_id, recipients_from

    @staticmethod
    def _readdress_to_original_recipients(original: Any, mail: Any, reply_all: bool) -> None:
        """Point ``mail`` (a reply addressed to the user) at the recipients of
        ``original``, or refuse: a draft addressed to the user is never kept."""
        lines, unreadable = _recipient_lines(original)
        to, cc = lines[OL_TO], lines[OL_CC] if reply_all else []
        if not to or unreadable:
            safe_com(lambda: mail.Close(OL_DISCARD), None)
            raise ReplySourceError(
                REPLY_UNAVAILABLE,
                "the original was sent by you, so Outlook would address the reply to you, and its "
                f"own recipients cannot be read ({len(to)} readable To, {unreadable} unreadable); "
                "give `to` explicitly. Nothing was saved",
            )
        mail.To = "; ".join(to)
        mail.CC = "; ".join(cc)
        logger.info(
            "Reply to a mail you sent: addressed to its %d recipient(s) instead of to you.", len(to),
        )

    def _open_reply_source(self, target: ReplyTarget) -> Any:
        """The original as an Outlook MailItem, or :class:`ReplySourceError`."""
        if target.kind == REPLY_BY_MSG_PATH:
            try:
                original = self._namespace().OpenSharedItem(target.value)
            except Exception as exc:
                raise ReplySourceError(
                    REPLY_SOURCE_NOT_FOUND,
                    f"Outlook could not open the .msg file: {type(exc).__name__}: {exc}",
                ) from exc
        else:
            original = self.find_by_message_id(normalize_message_id(target.value), None)
        if original is None:
            raise ReplySourceError(
                REPLY_SOURCE_NOT_FOUND, f"no Inbox mail with Message-ID {target.value}",
            )
        if safe_com(lambda: int(original.Class), OL_CLASS_MAIL_ITEM) != OL_CLASS_MAIL_ITEM:
            if target.kind == REPLY_BY_MSG_PATH:
                safe_com(lambda: original.Close(OL_DISCARD), None)
            raise ReplySourceError(REPLY_SOURCE_NOT_FOUND, "the original is not a mail item")
        return original

    @staticmethod
    def _describe_reply(mail: Any, created: CreatedDraft, replied_to_id: str) -> None:
        """Fill the reply fields of ``created`` from the saved draft.

        The In-Reply-To is read back from the draft itself (PR_IN_REPLY_TO_ID),
        so ``thread_header`` reports what is stored, not what was hoped for.
        """
        lines, _ = _recipient_lines(mail)
        created.to, created.cc = lines[OL_TO], lines[OL_CC]
        created.subject = str(safe_com(lambda: mail.Subject or "", ""))
        created.replied_to_message_id = replied_to_id
        in_reply_to = normalize_message_id(safe_com(
            lambda: mail.PropertyAccessor.GetProperty(DASL_IN_REPLY_TO_ID), None,
        ))
        if in_reply_to:
            created.thread_header = THREAD_SET
        else:
            created.thread_header = THREAD_NOT_SET
            created.thread_header_reason = "the saved draft has no In-Reply-To property"

    def update_draft(
        self,
        *,
        entry_id: str,
        to: list[str] | None,
        cc: list[str] | None,
        bcc: list[str],
        subject: str | None,
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
        up carrying exactly what the caller asked for. ``to`` / ``cc`` /
        ``subject`` of ``None`` stay as the draft has them (a reply's own), and
        a reply's quoted original, which sits after the marked region, is never
        touched.
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

        if to is not None:
            mail.To = "; ".join(to)
        if cc is not None:
            mail.CC = "; ".join(cc)
        mail.BCC = "; ".join(bcc)
        if subject is not None:
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
        updated.from_address = sending_address(mail)
        logger.info("Draft updated in place.")
        return updated

    def open_draft(self, entry_id: str) -> Any:
        """The unsent mail ``entry_id`` in this mailbox's Drafts, or
        :class:`DraftUpdateError`.

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
        drafts_id = safe_com(lambda: str(self.store().GetDefaultFolder(OL_FOLDER_DRAFTS).EntryID), "")
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
        lines, unreadable = _recipient_lines(mail)

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
            sending_account=sending_address(mail),
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


def sending_address(mail: Any) -> str:
    """The SMTP address of ``mail``'s ``SendUsingAccount``; ``""`` when unset
    or unreadable."""
    return safe_com(lambda: str(mail.SendUsingAccount.SmtpAddress or "").strip(), "")


def _send_as(mail: Any, account: Any, address: str) -> None:
    """Make ``mail`` send as ``account`` and prove it by reading it back.

    pywin32's plain property put can be refused for an object-valued property,
    or appear to work and change nothing; the PROPERTYPUTREF form is tried
    then. Whatever the route, only a read-back equal to ``address`` is
    accepted: otherwise the unsaved item is discarded and nothing is kept.

    Raises:
        DraftAccountError: the draft would not send as ``address``.
    """
    try:
        mail.SendUsingAccount = account
    except Exception as exc:
        logger.info("Plain SendUsingAccount put refused (%s); trying PROPERTYPUTREF.", exc)
    if sending_address(mail).casefold() != address.casefold():
        try:
            oleobj = mail._oleobj_
            oleobj.Invoke(oleobj.GetIDsOfNames("SendUsingAccount"), 0, _DISPATCH_PROPERTYPUTREF, 0, account)
        except Exception as exc:
            logger.info("SendUsingAccount PROPERTYPUTREF refused too: %s", exc)
    if sending_address(mail).casefold() != address.casefold():
        safe_com(lambda: mail.Close(OL_DISCARD), None)
        raise DraftAccountError(
            SENDING_ACCOUNT_MISMATCH,
            "Outlook would not set the draft to send as this mailbox's account; "
            "nothing was saved",
        )
    logger.info("Draft set to send as this mailbox's account.")


def _addressed_to(mail: Any, address: str) -> bool:
    """Whether ``address`` is on ``mail``'s To line (case-insensitive)."""
    lines, _ = _recipient_lines(mail)
    return address.strip().casefold() in {a.casefold() for a in lines[OL_TO]}


def _recipient_lines(mail: Any) -> tuple[dict[int, list[str]], int]:
    """A mail's recipient addresses by line (To / CC / BCC), and how many
    recipients had no readable address or line."""
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
    return lines, unreadable

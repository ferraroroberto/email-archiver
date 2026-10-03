"""
Which Outlook store a mailbox lives in (issue #108).

Every folder the client touches (Inbox, Sent, Drafts, the archive folder) is
taken from one resolved ``Store``, never from the namespace's default folders,
so a profile holding several mailboxes cannot silently act on the wrong one.

A registry mailbox is matched by address: the account whose ``SmtpAddress``
equals it, through that account's delivery store, else the one store whose
``DisplayName`` equals it (an IMAP store is named after its address). Only the
mailbox synthesized when there is no registry file resolves to the profile's
default store — that is today's behaviour, and the one place ``DefaultStore``
is read. Reads ``Account.SmtpAddress`` only, never ``GetExchangeUser``, which
can raise the address-book security modal.
"""
from __future__ import annotations

import logging
from typing import Any

from email_archiver.config import Mailbox
from email_archiver.outlook.mapi import MailboxNotInOutlookError, safe_com

logger = logging.getLogger(__name__)


def resolve_store(namespace: Any, mailbox: Mailbox | None) -> Any:
    """The ``Store`` holding ``mailbox``.

    ``None`` or the synthesized mailbox is the profile's default store.

    Raises:
        MailboxNotInOutlookError: no store matches the mailbox's address, or
            several do and none is tied to it by an account.
    """
    if mailbox is None or mailbox.synthesized:
        store = namespace.DefaultStore
        alias = mailbox.alias if mailbox is not None else "default"
    else:
        store = find_store(namespace, mailbox.address)
        if store is None:
            raise MailboxNotInOutlookError(
                f"mailbox {mailbox.alias!r} has no store in the running Outlook profile; "
                "add its account to Outlook or fix its address in config/mailboxes.json"
            )
        alias = mailbox.alias
    logger.info(
        "ℹ️ Resolved mailbox %s → store %s", alias,
        safe_com(lambda: str(store.DisplayName), "<unnamed>"),
    )
    return store


def find_account(namespace: Any, address: str) -> Any | None:
    """The Outlook ``Account`` whose SMTP address is ``address``, or ``None``."""
    target = address.strip().casefold()
    accounts = safe_com(lambda: namespace.Accounts, None)
    count = safe_com(lambda: int(accounts.Count), 0) if accounts is not None else 0
    for i in range(1, count + 1):
        account = safe_com(lambda i=i: accounts.Item(i), None)
        if account is None:
            continue
        if safe_com(lambda a=account: str(a.SmtpAddress or "").strip(), "").casefold() == target:
            return account
    return None


def find_store(namespace: Any, address: str) -> Any | None:
    """The store for ``address``, or ``None`` when the profile has none.

    Raises:
        MailboxNotInOutlookError: no account owns the address and more than one
            store is named after it, so picking one would be a guess.
    """
    target = address.strip().casefold()
    account = find_account(namespace, address)
    if account is not None:
        store = safe_com(lambda: account.DeliveryStore, None)
        if store is not None:
            return store
        logger.warning("The account for %s has no readable delivery store.", address)

    stores = safe_com(lambda: namespace.Stores, None)
    count = safe_com(lambda: int(stores.Count), 0) if stores is not None else 0
    named = []
    for i in range(1, count + 1):
        store = safe_com(lambda i=i: stores.Item(i), None)
        if store is None:
            continue
        if safe_com(lambda s=store: str(s.DisplayName).strip(), "").casefold() == target:
            named.append(store)
    if len(named) > 1:
        raise MailboxNotInOutlookError(
            f"{len(named)} Outlook stores are named after this mailbox's address and "
            "no account owns it, so which one is meant cannot be told"
        )
    return named[0] if named else None

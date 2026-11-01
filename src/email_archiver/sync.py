"""
Ask Outlook to run Send/Receive and report how it went (issue #132).

This is what ``main_batch.py sync`` runs. A mail that is already on the server
stays invisible to ``find``, ``plan`` and ``apply`` until Outlook syncs on its
own; this verb makes it sync and waits, bounded, for the end.

Like ``find.py`` it is pure orchestration over
:class:`~email_archiver.outlook.client.OutlookClient` — no COM here — so a fake
client drives it in a unit test, with an injected clock and sleep.

Design decisions:

- **Reads, moves, saves, drafts and sends no mail.** The one COM call that
  starts a sync lives in ``outlook/syncing.py`` and nowhere else.
- **One trigger for the whole profile.** Send/Receive is global to the
  profile, so there is no per-mailbox trigger; the document says
  ``trigger.scope: "profile"`` and every mailbox that is present shares the
  one observed outcome rather than implying a per-account observation.
- **Outcomes are honest.** ``completed`` only when the end was observed;
  ``timed_out`` at the deadline; ``unknown`` when the end cannot be observed or
  a check could not run. ``unknown`` is never folded into ``completed``.
  ``mailbox_not_in_outlook`` is that mailbox's alone.
- **Never kills or restarts Outlook** to force a sync.
"""
from __future__ import annotations

import logging
import time
from collections.abc import Callable, Sequence
from typing import Any

from email_archiver.batch_common import error_document, header
from email_archiver.config import Mailbox
from email_archiver.outlook.mapi import (
    SYNC_ENDED,
    SYNC_UNOBSERVABLE,
    MailboxNotInOutlookError,
)

logger = logging.getLogger(__name__)

VERB = "sync"
DEFAULT_TIMEOUT_SECONDS = 120.0
POLL_INTERVAL_SECONDS = 0.25

OUTCOME_COMPLETED = "completed"
OUTCOME_TIMED_OUT = "timed_out"
OUTCOME_UNKNOWN = "unknown"
OUTCOME_MAILBOX_NOT_IN_OUTLOOK = "mailbox_not_in_outlook"
OUTCOME_DEFERRED = "deferred_no_desktop"

ERROR_DEFERRED_NO_DESKTOP = "deferred_no_desktop"
SCOPE_PROFILE = "profile"


def deferred_document(desktop: dict[str, Any]) -> dict[str, Any]:
    """The document for a session with no desktop: nothing was started."""
    document = error_document(
        VERB, ERROR_DEFERRED_NO_DESKTOP,
        "This is a Windows session with no desktop (session 0), so Outlook was "
        "not started and no Send/Receive was triggered. Run it again from a "
        "logged-on desktop session.",
    )
    document["desktop"] = desktop
    document["outcome"] = OUTCOME_DEFERRED
    return document


def _entry(
    mailbox: Mailbox, address: str, outcome: str, waited: float, reason: str,
) -> dict[str, Any]:
    return {
        "alias": mailbox.alias,
        "address": address or mailbox.address,
        "outcome": outcome,
        "waited_seconds": round(waited, 2),
        "reason": reason,
    }


def _wait(
    handle: Any, timeout: float, clock: Callable[[], float], sleep: Callable[[float], None],
) -> tuple[str, float, str]:
    """Poll ``handle`` until it ends, cannot be observed, or ``timeout`` passes.

    Returns ``(outcome, seconds waited, reason)``. The sleep is cut to the time
    left, so the call returns at the deadline, not a poll interval after it.
    """
    start = clock()
    deadline = start + timeout
    while True:
        try:
            state, reason = handle.poll()
        except Exception as exc:  # a handle that cannot be read is not "still running"
            return OUTCOME_UNKNOWN, clock() - start, f"{type(exc).__name__}: {exc}"
        if state == SYNC_ENDED:
            return OUTCOME_COMPLETED, clock() - start, ""
        if state == SYNC_UNOBSERVABLE:
            return OUTCOME_UNKNOWN, clock() - start, reason
        now = clock()
        if now >= deadline:
            return (
                OUTCOME_TIMED_OUT, now - start,
                f"Outlook's Send/Receive did not report an end within {timeout:g}s; "
                "it may still be running.",
            )
        sleep(min(POLL_INTERVAL_SECONDS, deadline - now))


def run(
    client_for: Callable[[Mailbox], Any],
    mailboxes: Sequence[Mailbox],
    *,
    desktop: dict[str, Any],
    timeout: float = DEFAULT_TIMEOUT_SECONDS,
    clock: Callable[[], float] = time.monotonic,
    sleep: Callable[[float], None] = time.sleep,
) -> dict[str, Any]:
    """Trigger one profile-wide Send/Receive and report each mailbox's outcome.

    ``client_for`` builds a client bound to one mailbox. A mailbox whose store
    is not in Outlook reports ``mailbox_not_in_outlook`` and the others carry
    on; any other failure of that check is ``unknown`` with its reason. When no
    mailbox is present nothing is triggered.
    """
    present: list[tuple[Mailbox, Any, str]] = []
    entries: dict[str, dict[str, Any]] = {}
    for mailbox in mailboxes:
        client = client_for(mailbox)
        try:
            client.store()
        except MailboxNotInOutlookError as exc:
            entries[mailbox.alias] = _entry(
                mailbox, "", OUTCOME_MAILBOX_NOT_IN_OUTLOOK, 0.0, str(exc),
            )
            continue
        except Exception as exc:  # a check that could not run is not "missing"
            entries[mailbox.alias] = _entry(
                mailbox, "", OUTCOME_UNKNOWN, 0.0, f"{type(exc).__name__}: {exc}",
            )
            continue
        address = ""
        if mailbox.synthesized:
            try:
                address = client.default_account_smtp() or ""
            except Exception as exc:  # the address is a label; its absence changes no outcome
                logger.warning("Sync: could not read the default account address: %s", exc)
        present.append((mailbox, client, address))

    trigger: dict[str, Any] = {
        "method": None, "scope": SCOPE_PROFILE, "signal": "none", "started": False,
    }
    outcome, waited, reason = OUTCOME_UNKNOWN, 0.0, ""
    if present:
        try:
            handle = present[0][1].start_sync()
        except Exception as exc:
            # A sync that may have half-started is not "not started" either, but
            # nothing was observed, so every present mailbox is unknown.
            outcome, reason = OUTCOME_UNKNOWN, (
                f"Send/Receive could not be triggered ({type(exc).__name__}: {exc}); "
                "it may have started."
            )
            logger.error("Sync: trigger failed: %s: %s", type(exc).__name__, exc)
        else:
            trigger.update(method=handle.method, signal=handle.signal, started=True)
            logger.info(
                "Sync: started via %s (signal %s), waiting up to %gs.",
                handle.method, handle.signal, timeout,
            )
            outcome, waited, reason = _wait(handle, timeout, clock, sleep)
    for mailbox, _client, address in present:
        entries[mailbox.alias] = _entry(mailbox, address, outcome, waited, reason)

    results = [entries[mailbox.alias] for mailbox in mailboxes]
    logger.info(
        "Sync: %s.", ", ".join(f"{e['alias']} {e['outcome']}" for e in results) or "no mailbox",
    )
    return {
        **header(VERB),
        "desktop": desktop,
        "timeout_seconds": timeout,
        "trigger": trigger,
        "mailboxes": results,
    }

"""
The one place this app starts an Outlook Send/Receive over COM (issue #132).

Kept out of ``client.py`` on purpose, like ``sending.py`` for ``Send``: only the
``sync`` verb may start a sync, and a test pins that nothing else references
``SendAndReceive`` or ``SyncObjects``. The only caller is
:func:`email_archiver.sync.run`, through ``OutlookClient.start_sync``.

Nothing here reads, moves, saves or sends a mail, and nothing ends or restarts
Outlook.

How the end is observed (a hypothesis until the live probe of issue #132
confirms it): each ``Namespace.SyncObjects`` group raises ``SyncEnd`` when its
sync finishes; the handle watches those events while pumping COM messages on
the thread that owns the apartment. Folder item counts are deliberately not a
signal: a count that stops changing is not proof the sync ended. With no usable
``SyncObject``, or events that cannot be attached, ``SendAndReceive(False)``
is the trigger and the handle reports ``unobservable``.
"""
from __future__ import annotations

import logging
from collections.abc import Callable
from typing import Any

from email_archiver.outlook.mapi import SYNC_ENDED, SYNC_RUNNING, SYNC_UNOBSERVABLE

logger = logging.getLogger(__name__)

METHOD_SYNC_OBJECTS = "SyncObject.Start"
METHOD_SEND_AND_RECEIVE = "Namespace.SendAndReceive"
SIGNAL_SYNC_OBJECT = "sync_object"
SIGNAL_NONE = "none"


class _SyncEvents:
    """Event sink ``DispatchWithEvents`` mixes into a ``SyncObject``.

    Flags are class defaults overwritten on the instance, so the sink needs no
    ``__init__`` (which the mixed-in class would not run).
    """

    ended = False
    error = ""

    def OnSyncEnd(self) -> None:  # noqa: N802 - COM event name
        self.ended = True

    def OnError(self, code: Any = None, description: Any = None) -> None:  # noqa: N802
        self.error = f"{code}: {description}"


def _attach_events(sync_object: Any) -> Any:
    """``sync_object`` wrapped so ``SyncEnd`` / ``Error`` set flags on it."""
    import win32com.client  # noqa: PLC0415

    return win32com.client.DispatchWithEvents(sync_object, _SyncEvents)


def _pump_messages() -> None:
    """Deliver pending COM events; they only arrive while this thread pumps."""
    try:
        import pythoncom  # noqa: PLC0415
    except ImportError:  # pragma: no cover - win32 only
        return
    pythoncom.PumpWaitingMessages()


class SyncHandle:
    """What :func:`start` started, and whether its end can be watched."""

    def __init__(
        self,
        method: str,
        signal: str,
        listeners: list[Any],
        pump: Callable[[], None],
        unobservable_reason: str = "",
    ) -> None:
        self.method = method
        self.signal = signal
        self._listeners = listeners
        self._pump = pump
        self._unobservable_reason = unobservable_reason

    def poll(self) -> tuple[str, str]:
        """``(state, reason)``: ``running``, ``ended`` or ``unobservable``."""
        if not self._listeners:
            return SYNC_UNOBSERVABLE, self._unobservable_reason
        self._pump()
        errors = [lst.error for lst in self._listeners if lst.error]
        if errors:
            return SYNC_UNOBSERVABLE, f"Outlook reported a sync error ({errors[0]})."
        if all(lst.ended for lst in self._listeners):
            return SYNC_ENDED, ""
        return SYNC_RUNNING, ""


def _sync_objects(namespace: Any) -> list[Any]:
    """The profile's sync groups, or ``[]`` when Outlook exposes none."""
    try:
        collection = namespace.SyncObjects
        return [collection.Item(i) for i in range(1, int(collection.Count) + 1)]
    except Exception as exc:
        logger.warning("Sync: could not read Namespace.SyncObjects: %s: %s", type(exc).__name__, exc)
        return []


def _send_and_receive(namespace: Any, reason: str, pump: Callable[[], None]) -> SyncHandle:
    logger.warning("Sync: %s; triggering Namespace.SendAndReceive, whose end cannot be observed.", reason)
    namespace.SendAndReceive(False)  # False: no progress dialog asked for
    return SyncHandle(METHOD_SEND_AND_RECEIVE, SIGNAL_NONE, [], pump, f"{reason}.")


def start(
    namespace: Any,
    *,
    attach: Callable[[Any], Any] = _attach_events,
    pump: Callable[[], None] = _pump_messages,
) -> SyncHandle:
    """Start Send/Receive for the whole profile and return a handle to watch it.

    ``attach`` and ``pump`` are the seams a test replaces; the defaults are the
    real COM event sink and message pump.
    """
    sync_objects = _sync_objects(namespace)
    if not sync_objects:
        return _send_and_receive(namespace, "the profile exposes no SyncObject", pump)
    try:
        watched = [attach(sync_object) for sync_object in sync_objects]
    except Exception as exc:
        # Nothing has been started yet, so falling back starts exactly one sync.
        return _send_and_receive(
            namespace, f"SyncEnd events could not be attached ({type(exc).__name__}: {exc})", pump,
        )
    for sync_object in watched:
        sync_object.Start()
    logger.info("Sync: started %d SyncObject group(s).", len(watched))
    return SyncHandle(METHOD_SYNC_OBJECTS, SIGNAL_SYNC_OBJECT, watched, pump)

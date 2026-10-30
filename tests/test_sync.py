"""Tests for the ``sync`` verb (issue #132).

Nothing here touches Outlook: the orchestration runs on a fake client with an
injected clock and sleep, the COM adapter on a fake namespace with injected
event-attach and message-pump seams, and the verb's process contract on a stub
``pythoncom``. The fake client exposes only the surface ``sync`` is allowed to
use, so a stray move, save, draft or send fails loudly. Everything is
synthetic.
"""
from __future__ import annotations

import sys
import types
from pathlib import Path

import pytest

from email_archiver import sync
from email_archiver.cli import batch as main_batch
from email_archiver.config import load_mailbox_registry, parse_mailbox_registry
from email_archiver.outlook import process, syncing
from email_archiver.outlook.mapi import (
    SYNC_ENDED,
    SYNC_RUNNING,
    SYNC_UNOBSERVABLE,
    MailboxNotInOutlookError,
    OutlookUnavailableError,
)
from tests.batch.batch_support import archive_root, cfg  # noqa: F401 - fixtures are used by name
from tests.test_mailboxes import OWNER, _registry_doc, batch_process  # noqa: F401

REPO_ROOT = Path(__file__).resolve().parent.parent
INTERACTIVE = {"session_id": 1, "status": "interactive"}


# ------------------------------------------------------------------ fakes ---

class FakeClock:
    """Time that only moves when the loop sleeps."""

    def __init__(self) -> None:
        self.now = 0.0
        self.sleeps: list[float] = []

    def __call__(self) -> float:
        return self.now

    def sleep(self, seconds: float) -> None:
        self.sleeps.append(seconds)
        self.now += seconds


class FakeHandle:
    """A sync handle that reports ``states`` in order, then repeats the last."""

    method = "Fake.Start"
    signal = "fake_signal"

    def __init__(self, *states: tuple[str, str]) -> None:
        self._states = list(states)
        self.polls = 0

    def poll(self) -> tuple[str, str]:
        state = self._states[min(self.polls, len(self._states) - 1)]
        self.polls += 1
        return state


def _running() -> tuple[str, str]:
    return SYNC_RUNNING, ""


def _ended() -> tuple[str, str]:
    return SYNC_ENDED, ""


class FakeSyncClient:
    """The only surface ``sync`` may use on a client: no move, save, draft or send."""

    def __init__(self, mailbox, *, store_error=None, handle=None, trigger_error=None,
                 start_error=None):
        self.mailbox = mailbox
        self._store_error = store_error
        self._handle = handle
        self._trigger_error = trigger_error
        self._start_error = start_error
        self.triggers = 0
        self.started: list[float] = []

    def ensure_running(self, timeout):
        if self._start_error is not None:
            raise self._start_error

    def store(self):
        if self._store_error is not None:
            raise self._store_error
        return object()

    def default_account_smtp(self) -> str:
        return OWNER

    def start_sync(self):
        self.triggers += 1
        if self._trigger_error is not None:
            raise self._trigger_error
        return self._handle

    def __getattr__(self, name):
        raise AssertionError(f"sync must not use client.{name}")


def _registry():
    return parse_mailbox_registry(_registry_doc())


def _run(clients: dict, *, timeout=120.0, clock=None, mailboxes=None):
    clock = clock or FakeClock()
    registry = _registry()
    doc = sync.run(
        lambda mailbox: clients[mailbox.alias],
        mailboxes or list(registry.mailboxes.values()),
        desktop=INTERACTIVE, timeout=timeout, clock=clock, sleep=clock.sleep,
    )
    return doc, clock


def _outcomes(doc) -> dict[str, str]:
    return {e["alias"]: e["outcome"] for e in doc["mailboxes"]}


def _clients(handle, **overrides) -> dict:
    registry = _registry()
    clients = {a: FakeSyncClient(m, handle=handle) for a, m in registry.mailboxes.items()}
    for alias, kwargs in overrides.items():
        clients[alias] = FakeSyncClient(registry.mailboxes[alias], **kwargs)
    return clients


# ----------------------------------------------------------- orchestration ---

def test_an_ended_sync_completes_every_present_mailbox_with_the_waited_time():
    handle = FakeHandle(_running(), _running(), _ended())

    doc, clock = _run(_clients(handle))

    assert _outcomes(doc) == {"owner": "completed", "second": "completed"}
    assert [e["waited_seconds"] for e in doc["mailboxes"]] == [0.5, 0.5]
    assert clock.sleeps == [sync.POLL_INTERVAL_SECONDS] * 2
    assert doc["verb"] == "sync" and doc["schema_version"] == 1
    assert doc["desktop"] == INTERACTIVE and doc["timeout_seconds"] == 120.0
    assert "mailbox" not in doc


def test_the_trigger_is_started_once_for_the_whole_profile_and_says_so():
    handle = FakeHandle(_ended())
    clients = _clients(handle)

    doc, _ = _run(clients)

    assert sum(c.triggers for c in clients.values()) == 1
    assert doc["trigger"] == {
        "method": "Fake.Start", "scope": "profile", "signal": "fake_signal", "started": True,
    }


def test_a_sync_that_never_ends_times_out_at_the_deadline_and_returns():
    handle = FakeHandle(_running())

    doc, clock = _run(_clients(handle), timeout=1.0)

    assert _outcomes(doc) == {"owner": "timed_out", "second": "timed_out"}
    assert doc["mailboxes"][0]["waited_seconds"] == pytest.approx(1.0)
    assert clock.now <= 1.0 + sync.POLL_INTERVAL_SECONDS
    assert "did not report an end" in doc["mailboxes"][0]["reason"]


def test_the_last_sleep_is_cut_to_the_time_left():
    clock = FakeClock()

    doc, _ = _run(_clients(FakeHandle(_running())), timeout=0.3, clock=clock)

    assert clock.sleeps[0] == sync.POLL_INTERVAL_SECONDS
    assert clock.sleeps[1] == pytest.approx(0.05)
    assert doc["mailboxes"][0]["waited_seconds"] == pytest.approx(0.3)


def test_an_unobservable_end_is_unknown_and_never_completed():
    handle = FakeHandle((SYNC_UNOBSERVABLE, "the profile exposes no SyncObject."))

    doc, clock = _run(_clients(handle))

    assert _outcomes(doc) == {"owner": "unknown", "second": "unknown"}
    assert doc["mailboxes"][0]["reason"] == "the profile exposes no SyncObject."
    assert clock.sleeps == []  # no point waiting for an end that cannot be seen


def test_a_handle_that_cannot_be_read_is_unknown_not_still_running():
    class Broken(FakeHandle):
        def poll(self):
            raise RuntimeError("COM went away")

    doc, _ = _run(_clients(Broken()))

    assert _outcomes(doc) == {"owner": "unknown", "second": "unknown"}
    assert "COM went away" in doc["mailboxes"][0]["reason"]


def test_a_trigger_that_fails_is_unknown_for_the_present_mailboxes():
    clients = _clients(None, owner={"trigger_error": RuntimeError("refused")})

    doc, _ = _run(clients)

    assert _outcomes(doc) == {"owner": "unknown", "second": "unknown"}
    assert "refused" in doc["mailboxes"][0]["reason"]
    assert doc["trigger"]["started"] is False


def test_a_store_check_that_could_not_run_is_unknown_for_that_mailbox_only():
    clients = _clients(FakeHandle(_ended()), second={"store_error": RuntimeError("COM hiccup")})

    doc, _ = _run(clients)

    assert _outcomes(doc) == {"owner": "completed", "second": "unknown"}
    second = doc["mailboxes"][1]
    assert "COM hiccup" in second["reason"] and second["waited_seconds"] == 0


def test_a_mailbox_missing_from_outlook_is_reported_alone_and_the_other_still_syncs():
    clients = _clients(
        FakeHandle(_running(), _ended()),
        second={"store_error": MailboxNotInOutlookError("no store for second")},
    )

    doc, _ = _run(clients)

    assert _outcomes(doc) == {"owner": "completed", "second": "mailbox_not_in_outlook"}
    assert doc["mailboxes"][1]["waited_seconds"] == 0
    assert doc["mailboxes"][1]["reason"] == "no store for second"
    assert clients["owner"].triggers == 1


def test_with_every_mailbox_missing_nothing_is_triggered():
    missing = {"store_error": MailboxNotInOutlookError("no store")}
    clients = _clients(None, owner=missing, second=missing)

    doc, _ = _run(clients)

    assert _outcomes(doc) == {
        "owner": "mailbox_not_in_outlook", "second": "mailbox_not_in_outlook",
    }
    assert all(c.triggers == 0 for c in clients.values())
    assert doc["trigger"] == {
        "method": None, "scope": "profile", "signal": "none", "started": False,
    }


def test_the_synthesized_mailbox_reports_the_default_account_address():
    registry = load_mailbox_registry(Path("absent.json"))
    mailbox = registry.select(None)
    clients = {mailbox.alias: FakeSyncClient(mailbox, handle=FakeHandle(_ended()))}

    doc, _ = _run(clients, mailboxes=[mailbox])

    assert doc["mailboxes"][0]["alias"] == "default"
    assert doc["mailboxes"][0]["address"] == OWNER


def test_the_deferred_document_names_the_state_and_starts_nothing():
    desktop = {"session_id": 0, "status": "session_0"}

    doc = sync.deferred_document(desktop)

    assert doc["verb"] == "sync" and doc["schema_version"] == 1
    assert doc["desktop"] == desktop
    assert doc["outcome"] == "deferred_no_desktop"
    assert doc["error"]["code"] == "deferred_no_desktop"
    assert "mailboxes" not in doc


# ----------------------------------------------------------------- adapter ---

class FakeSyncObject:
    """A SyncObject whose ``ended`` / ``error`` an event would set."""

    def __init__(self) -> None:
        self.ended = False
        self.error = ""
        self.starts = 0

    def Start(self) -> None:  # noqa: N802 - COM-shaped API
        self.starts += 1


class FakeSyncObjects:
    def __init__(self, items) -> None:
        self._items = list(items)

    @property
    def Count(self) -> int:  # noqa: N802 - COM-shaped API
        return len(self._items)

    def Item(self, index: int):  # noqa: N802 - COM-shaped API
        return self._items[index - 1]


class FakeNamespace:
    def __init__(self, items=(), *, sync_objects_error=None) -> None:
        self._items = list(items)
        self._error = sync_objects_error
        self.send_and_receive: list[tuple] = []

    @property
    def SyncObjects(self):  # noqa: N802 - COM-shaped API
        if self._error is not None:
            raise self._error
        return FakeSyncObjects(self._items)

    def SendAndReceive(self, *args) -> None:  # noqa: N802 - COM-shaped API
        self.send_and_receive.append(args)


def test_a_sync_end_event_ends_the_handle():
    group = FakeSyncObject()
    pumps = []

    def pump():
        pumps.append(1)
        if len(pumps) == 2:
            group.ended = True  # the event arrives while the thread pumps

    handle = syncing.start(FakeNamespace([group]), attach=lambda o: o, pump=pump)

    assert (handle.method, handle.signal) == ("SyncObject.Start", "sync_object")
    assert group.starts == 1
    assert handle.poll() == (SYNC_RUNNING, "")
    assert handle.poll() == (SYNC_ENDED, "")


def test_every_group_must_end_before_the_handle_does():
    first, second = FakeSyncObject(), FakeSyncObject()
    handle = syncing.start(FakeNamespace([first, second]), attach=lambda o: o, pump=lambda: None)

    first.ended = True
    assert handle.poll()[0] == SYNC_RUNNING
    second.ended = True
    assert handle.poll()[0] == SYNC_ENDED
    assert (first.starts, second.starts) == (1, 1)


def test_an_outlook_sync_error_is_unobservable_not_ended():
    group = FakeSyncObject()
    handle = syncing.start(FakeNamespace([group]), attach=lambda o: o, pump=lambda: None)

    group.error = "-2147467259: The server is not available"
    group.ended = True

    state, reason = handle.poll()
    assert state == SYNC_UNOBSERVABLE
    assert "server is not available" in reason


@pytest.mark.parametrize("namespace", [
    FakeNamespace([]),
    FakeNamespace(sync_objects_error=RuntimeError("no SyncObjects")),
])
def test_no_usable_sync_object_triggers_send_and_receive_and_is_unobservable(namespace):
    handle = syncing.start(namespace, attach=pytest.fail, pump=lambda: None)

    assert namespace.send_and_receive == [(False,)]
    assert (handle.method, handle.signal) == ("Namespace.SendAndReceive", "none")
    state, reason = handle.poll()
    assert state == SYNC_UNOBSERVABLE and "no SyncObject" in reason


def test_events_that_cannot_be_attached_fall_back_without_starting_any_group():
    group = FakeSyncObject()
    namespace = FakeNamespace([group])

    def refuse(_sync_object):
        raise RuntimeError("no event support")

    handle = syncing.start(namespace, attach=refuse, pump=lambda: None)

    assert group.starts == 0
    assert namespace.send_and_receive == [(False,)]
    state, reason = handle.poll()
    assert state == SYNC_UNOBSERVABLE and "no event support" in reason


def test_the_event_sink_sets_its_flags():
    sink = syncing._SyncEvents()
    assert (sink.ended, sink.error) == (False, "")

    sink.OnSyncEnd()
    sink.OnError(5, "boom")

    assert sink.ended is True and sink.error == "5: boom"


def test_only_the_syncing_module_can_start_a_sync():
    # Spelled in halves so a grep of the repo for them finds only the adapter.
    needles = ("Send" + "AndReceive", "Sync" + "Objects")
    holders = sorted(
        str(path.relative_to(REPO_ROOT)).replace("\\", "/")
        for path in [*REPO_ROOT.glob("*.py"), *(REPO_ROOT / "email_archiver").rglob("*.py")]
        if any(needle in path.read_text(encoding="utf-8") for needle in needles)
    )
    assert holders == ["email_archiver/outlook/syncing.py"]


def test_start_sync_is_called_only_from_the_sync_module():
    needle = "." + "start_sync("
    holders = sorted(
        str(path.relative_to(REPO_ROOT)).replace("\\", "/")
        for path in [*REPO_ROOT.glob("*.py"), *(REPO_ROOT / "email_archiver").rglob("*.py")]
        if needle in path.read_text(encoding="utf-8")
    )
    assert holders == ["email_archiver/sync.py"]


# ------------------------------------------------------------ desktop state ---

@pytest.mark.parametrize("session_id, status", [
    (0, "session_0"), (1, "interactive"), (3, "interactive"), (None, "unknown"),
])
def test_desktop_state_names_each_session(monkeypatch, session_id, status):
    monkeypatch.setattr(process, "windows_session_id", lambda: session_id)

    assert process.desktop_state() == {"session_id": session_id, "status": status}


def test_the_real_session_id_is_a_number_or_unknown():
    assert process.windows_session_id() is None or process.windows_session_id() >= 0


# --------------------------------------------------------- process contract ---

def _factory(built, *, handle=None, **kwargs):
    def make(mailbox):
        client = FakeSyncClient(mailbox, handle=handle or FakeHandle(_ended()), **kwargs)
        built.append(client)
        return client
    return make


@pytest.fixture
def interactive(monkeypatch):
    monkeypatch.setattr(process, "windows_session_id", lambda: 1)


def test_the_verb_syncs_every_registry_mailbox_and_exits_0(batch_process, interactive):
    built: list = []

    code, doc, _ = batch_process(["sync"], registry=_registry_doc(), client_factory=_factory(built))

    assert code == main_batch.EXIT_OK
    assert doc["verb"] == "sync" and "mailbox" not in doc
    assert _outcomes(doc) == {"owner": "completed", "second": "completed"}
    assert doc["desktop"]["status"] == "interactive"


def test_a_timed_out_sync_is_a_result_not_a_failure(batch_process, interactive):
    built: list = []

    code, doc, _ = batch_process(
        ["sync", "--timeout", "0.3"], registry=_registry_doc(),
        client_factory=_factory(built, handle=FakeHandle(_running())),
    )

    assert code == main_batch.EXIT_OK
    assert _outcomes(doc) == {"owner": "timed_out", "second": "timed_out"}
    assert doc["timeout_seconds"] == 0.3


def test_mailbox_selects_one_by_any_alias_and_any_case(batch_process, interactive):
    built: list = []

    code, doc, _ = batch_process(
        ["sync", "--mailbox", "ME"], registry=_registry_doc(), client_factory=_factory(built),
    )

    assert code == main_batch.EXIT_OK
    assert _outcomes(doc) == {"owner": "completed"}
    assert {c.mailbox.alias for c in built} == {"owner"}


def test_an_unknown_mailbox_exits_2_before_any_client_is_built(batch_process, interactive):
    built: list = []

    code, doc, _ = batch_process(
        ["sync", "--mailbox", "nobody"], registry=_registry_doc(), client_factory=_factory(built),
    )

    assert code == main_batch.EXIT_CANNOT_START
    assert doc["error"]["code"] == "mailbox_unknown"
    assert built == []


def test_without_a_registry_file_the_one_synthesized_mailbox_is_synced(batch_process, interactive):
    built: list = []

    code, doc, _ = batch_process(["sync"], registry=None, client_factory=_factory(built))

    assert code == main_batch.EXIT_OK
    assert [(e["alias"], e["address"]) for e in doc["mailboxes"]] == [("default", OWNER)]


@pytest.mark.parametrize("raw", ["0", "-5", "abc", "nan", "inf", ""])
def test_a_bad_timeout_is_bad_input_before_outlook(batch_process, interactive, raw):
    code, doc, built = batch_process(
        ["sync", "--timeout", raw], registry=_registry_doc(),
        client_factory=lambda mb: pytest.fail("Outlook touched"),
    )

    assert code == main_batch.EXIT_CANNOT_START
    assert doc["error"]["code"] == "bad_input"
    assert built == []


def test_session_0_defers_with_exit_3_and_starts_nothing(batch_process, monkeypatch):
    monkeypatch.setattr(process, "windows_session_id", lambda: 0)
    monkeypatch.setitem(sys.modules, "pythoncom", types.SimpleNamespace(
        CoInitialize=lambda: pytest.fail("COM initialised in session 0"),
        CoUninitialize=lambda: None,
    ))

    code, doc, built = batch_process(
        ["sync"], registry=_registry_doc(),
        client_factory=lambda mb: pytest.fail("Outlook client built in session 0"),
    )

    assert code == main_batch.EXIT_DEFERRED == 3
    assert doc["desktop"] == {"session_id": 0, "status": "session_0"}
    assert doc["outcome"] == "deferred_no_desktop"
    assert doc["error"]["code"] == "deferred_no_desktop"
    assert built == []


def test_an_unreadable_session_state_proceeds_and_says_unknown(batch_process, monkeypatch):
    monkeypatch.setattr(process, "windows_session_id", lambda: None)
    built: list = []

    code, doc, _ = batch_process(["sync"], registry=_registry_doc(), client_factory=_factory(built))

    assert code == main_batch.EXIT_OK
    assert doc["desktop"] == {"session_id": None, "status": "unknown"}
    assert _outcomes(doc) == {"owner": "completed", "second": "completed"}


def test_outlook_that_cannot_be_started_exits_2(batch_process, interactive):
    built: list = []

    code, doc, _ = batch_process(
        ["sync"], registry=_registry_doc(),
        client_factory=_factory(built, start_error=OutlookUnavailableError("never came up")),
    )

    assert code == main_batch.EXIT_CANNOT_START
    assert doc["error"]["code"] == "outlook_unavailable"
    assert all(c.triggers == 0 for c in built)


def test_a_malformed_registry_exits_2_before_outlook(batch_process, interactive):
    code, doc, built = batch_process(
        ["sync"], registry="{ not json", client_factory=lambda mb: pytest.fail("Outlook touched"),
    )

    assert code == main_batch.EXIT_CANNOT_START
    assert doc["error"]["code"] == "mailbox_registry_invalid"
    assert built == []

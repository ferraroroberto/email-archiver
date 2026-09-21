"""
Tests for OutlookClient.ensure_running()'s ownership of what it starts.

ensure_running() used to spawn outlook.exe and forget it. When the COM object
never arrived it raised, and the process it had started stayed up — on a
scheduled unattended run, an invisible orphan holding the mail profile and OST
against the user's own Outlook, one more every time the wait timed out. See
issue #78.

Everything here is driven through a fake spawner. No test in this module may be
able to reach a real Outlook: the only one on a developer's machine is the
user's own, and the bug being fixed is precisely about ending the wrong process.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from email_archiver.outlook import client as client_mod
from email_archiver.outlook import process as process_mod
from email_archiver.outlook.client import OutlookClient
from email_archiver.outlook.mapi import OutlookUnavailableError
from email_archiver.outlook.process import _terminate_spawned_outlook

FAKE_EXE = r"C:\fake\outlook.exe"


class _FakeProc:
    """Stands in for the Popen handle ``_spawn_outlook`` hands back.

    Records every lifetime call so a test can assert not just that the right
    process was ended, but that no other one was touched at all.
    """

    def __init__(
        self,
        pid: int = 4321,
        exit_code: int | None = None,
        terminate_error: Exception | None = None,
        wait_error: Exception | None = None,
    ):
        self.pid = pid
        self._exit_code = exit_code
        self._terminate_error = terminate_error
        self._wait_error = wait_error
        self.terminate_calls = 0
        self.wait_calls = 0

    def poll(self) -> int | None:
        return self._exit_code

    def terminate(self) -> None:
        self.terminate_calls += 1
        if self._terminate_error is not None:
            raise self._terminate_error
        self._exit_code = 1

    def wait(self, timeout: float | None = None) -> int | None:
        self.wait_calls += 1
        if self._wait_error is not None:
            raise self._wait_error
        return self._exit_code


def _arrange(monkeypatch, *, active, proc: _FakeProc | None):
    """Point ensure_running() at fakes and return the spawn call log.

    ``active`` is either a value returned by every get_active_application()
    call, or a list consumed one call at a time (so a test can make Outlook
    appear partway through the poll loop).
    """
    calls: list[str] = []

    if isinstance(active, list):
        queue = list(active)

        def fake_active():
            return queue.pop(0) if queue else None
    else:

        def fake_active():
            return active

    def fake_spawn(exe: str):
        calls.append(exe)
        assert proc is not None, "ensure_running spawned when it should not have"
        return proc

    monkeypatch.setattr(process_mod, "get_active_application", fake_active)
    monkeypatch.setattr(process_mod, "_outlook_executable", lambda: FAKE_EXE)
    monkeypatch.setattr(process_mod, "_spawn_outlook", fake_spawn)
    monkeypatch.setattr(process_mod, "_POLL_INTERVAL_SECONDS", 0.0)
    return calls


# ------------------------------------------------- the attach path ---------


def test_attaching_to_a_running_outlook_starts_and_ends_nothing(monkeypatch):
    """The user's own Outlook is attached to, never started, never touched."""
    sentinel = object()
    calls = _arrange(monkeypatch, active=sentinel, proc=None)

    assert OutlookClient().ensure_running(timeout=5.0) is sentinel
    assert calls == [], "attaching must never spawn an outlook.exe"


def test_outlook_appearing_during_the_poll_is_never_terminated(monkeypatch):
    """A started Outlook that does publish its COM object is left running."""
    sentinel = object()
    proc = _FakeProc()
    _arrange(monkeypatch, active=[None, None, sentinel], proc=proc)

    assert OutlookClient().ensure_running(timeout=5.0) is sentinel
    assert proc.terminate_calls == 0


# ----------------------------------------------- the timeout teardown ------


def test_timeout_terminates_the_process_it_started(monkeypatch):
    proc = _FakeProc(pid=39980)
    _arrange(monkeypatch, active=None, proc=proc)

    with pytest.raises(OutlookUnavailableError) as excinfo:
        OutlookClient().ensure_running(timeout=0.01)

    assert proc.terminate_calls == 1
    message = str(excinfo.value)
    assert "39980" in message
    assert "was terminated" in message


def test_timeout_reports_a_process_that_had_already_exited(monkeypatch):
    """Nothing to end: Outlook handed off to another instance and quit."""
    proc = _FakeProc(pid=111, exit_code=0)
    _arrange(monkeypatch, active=None, proc=proc)

    with pytest.raises(OutlookUnavailableError) as excinfo:
        OutlookClient().ensure_running(timeout=0.01)

    assert proc.terminate_calls == 0
    assert "had already exited" in str(excinfo.value)


def test_a_teardown_that_fails_is_reported_not_swallowed(monkeypatch):
    """The orphan is still out there, and the error document has to say so."""
    proc = _FakeProc(pid=222, terminate_error=PermissionError("denied"))
    _arrange(monkeypatch, active=None, proc=proc)

    with pytest.raises(OutlookUnavailableError) as excinfo:
        OutlookClient().ensure_running(timeout=0.01)

    message = str(excinfo.value)
    assert "could NOT be terminated" in message
    assert "PermissionError" in message
    assert "may still be running" in message


def test_a_process_that_will_not_die_is_reported_as_still_running(monkeypatch):
    """terminate() returned but the process outlived the wait."""
    proc = _FakeProc(pid=333, wait_error=TimeoutError("still alive"))
    _arrange(monkeypatch, active=None, proc=proc)

    with pytest.raises(OutlookUnavailableError) as excinfo:
        OutlookClient().ensure_running(timeout=0.01)

    assert proc.terminate_calls == 1
    assert "could NOT be terminated" in str(excinfo.value)


def test_a_spawn_that_fails_outright_raises_without_a_teardown(monkeypatch):
    monkeypatch.setattr(process_mod, "get_active_application", lambda: None)
    monkeypatch.setattr(process_mod, "_outlook_executable", lambda: FAKE_EXE)

    def boom(exe: str):
        raise OSError("not executable")

    monkeypatch.setattr(process_mod, "_spawn_outlook", boom)

    with pytest.raises(OutlookUnavailableError, match="Could not start Outlook"):
        OutlookClient().ensure_running(timeout=0.01)


# ------------------------------------- only ever the handed-over handle ----


def test_the_teardown_touches_only_the_process_it_was_given():
    """The constraint the whole fix rests on, stated as a test.

    ``someone_elses`` stands for the interactive Outlook already running on the
    user's desktop: the teardown is handed one process and must not reach past
    it to any other.
    """
    ours = _FakeProc(pid=1001)
    someone_elses = _FakeProc(pid=1002)

    _terminate_spawned_outlook(ours)

    assert ours.terminate_calls == 1
    assert someone_elses.terminate_calls == 0
    assert someone_elses.poll() is None


def test_the_client_never_kills_outlook_by_image_name():
    """A regression guard with teeth, against the tempting wrong fix.

    Ending outlook.exe by image name — taskkill /IM, or any process-list sweep
    — would take the user's own Outlook down with the orphan. The module reads
    the process list (is_running's tasklist fallback) and must never do more
    than read it.
    """
    source = "".join(
        Path(mod.__file__).read_text(encoding="utf-8")
        for mod in (client_mod, process_mod)
    )
    assert "taskkill" not in source.lower()
    assert "/IM" not in source

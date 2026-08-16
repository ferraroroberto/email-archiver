"""
Tests for the `tasklist` fallback behind OutlookClient.is_running().

The fallback shells out to a native Windows console tool, which writes the OEM
code page rather than the parent's locale. Decoding it with `text=True` breaks
the moment the parent runs in UTF-8 mode (PYTHONUTF8=1): stdout comes back
None and a running Outlook is silently reported as not running. See issue #38.
"""
from __future__ import annotations

import logging
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

import pytest

from email_archiver.outlook import client as client_mod
from email_archiver.outlook.client import OutlookClient, _tasklist_has_image

REPO_ROOT = Path(__file__).resolve().parents[1]
IS_WINDOWS = sys.platform == "win32"


class _FakeCompleted:
    def __init__(self, stdout: str | None = "", stderr: str = "", returncode: int = 0):
        self.stdout = stdout
        self.stderr = stderr
        self.returncode = returncode


def _capture_run(monkeypatch, result: _FakeCompleted | Exception) -> dict:
    """Replace subprocess.run and return the dict its kwargs land in."""
    seen: dict = {}

    def fake_run(*args, **kwargs):
        seen["args"] = args
        seen["kwargs"] = kwargs
        if isinstance(result, Exception):
            raise result
        return result

    monkeypatch.setattr(subprocess, "run", fake_run)
    return seen


# ------------------------------------------------------------ decoding -----


def test_tasklist_pins_its_own_decoding(monkeypatch):
    """The call must never inherit the ambient locale via text=True."""
    seen = _capture_run(monkeypatch, _FakeCompleted(stdout="OUTLOOK.EXE  1234 Console\n"))

    assert _tasklist_has_image("OUTLOOK.EXE") is True

    kwargs = seen["kwargs"]
    assert "text" not in kwargs
    assert "universal_newlines" not in kwargs
    assert kwargs["errors"] == "replace"
    assert kwargs["encoding"] == ("oem" if IS_WINDOWS else "utf-8")
    assert kwargs["capture_output"] is True
    assert kwargs["timeout"] == 5


@pytest.mark.skipif(not IS_WINDOWS, reason="CREATE_NO_WINDOW is Windows-only")
def test_tasklist_keeps_console_window_suppressed(monkeypatch):
    """A console-less parent must not get a window flashed at it per spawn."""
    seen = _capture_run(monkeypatch, _FakeCompleted(stdout="INFO: No tasks\n"))

    _tasklist_has_image("OUTLOOK.EXE")

    assert seen["kwargs"]["creationflags"] == subprocess.CREATE_NO_WINDOW


# ------------------------------------------------- failed query is loud -----


def test_no_match_is_quiet(monkeypatch, caplog):
    """A successful query that found nothing is a fact, not a failure."""
    _capture_run(monkeypatch, _FakeCompleted(
        stdout="INFO: No tasks are running which match the specified criteria.\n"))

    with caplog.at_level(logging.WARNING, logger=client_mod.__name__):
        assert _tasklist_has_image("OUTLOOK.EXE") is False

    assert caplog.records == []


def test_undecodable_output_is_logged(monkeypatch, caplog):
    """stdout=None is what a decoding failure actually looks like — not a fact."""
    _capture_run(monkeypatch, _FakeCompleted(stdout=None))

    with caplog.at_level(logging.WARNING, logger=client_mod.__name__):
        assert _tasklist_has_image("OUTLOOK.EXE") is False

    assert any("no output" in r.getMessage() for r in caplog.records)


def test_nonzero_exit_is_logged(monkeypatch, caplog):
    _capture_run(monkeypatch, _FakeCompleted(stdout="", stderr="boom", returncode=1))

    with caplog.at_level(logging.WARNING, logger=client_mod.__name__):
        assert _tasklist_has_image("OUTLOOK.EXE") is False

    assert any("exited 1" in r.getMessage() for r in caplog.records)


@pytest.mark.parametrize("exc", [
    OSError("tasklist not found"),
    subprocess.TimeoutExpired(cmd="tasklist", timeout=5),
])
def test_query_errors_are_logged(monkeypatch, caplog, exc):
    _capture_run(monkeypatch, exc)

    with caplog.at_level(logging.WARNING, logger=client_mod.__name__):
        assert _tasklist_has_image("OUTLOOK.EXE") is False

    assert any("Could not run tasklist" in r.getMessage() for r in caplog.records)


def test_programming_errors_are_not_swallowed(monkeypatch):
    """A genuine bug must surface, not be reported as 'Outlook not running'."""
    _capture_run(monkeypatch, TypeError("bad call"))

    with pytest.raises(TypeError):
        _tasklist_has_image("OUTLOOK.EXE")


# ------------------------------------------------------------- wiring ------


def test_is_running_uses_tasklist_when_psutil_missing(monkeypatch):
    monkeypatch.setitem(sys.modules, "psutil", None)  # `import psutil` -> ImportError
    asked: list[str] = []
    monkeypatch.setattr(client_mod, "_tasklist_has_image",
                        lambda name: asked.append(name) or True)

    assert OutlookClient().is_running() is True
    assert asked == ["OUTLOOK.EXE"]


# ------------------------------------------- the real thing, end to end -----


@pytest.mark.skipif(not IS_WINDOWS, reason="tasklist/OEM decoding is Windows-only")
def test_detects_running_process_under_pythonutf8(tmp_path):
    """
    Regression proof for #38, with no mocks in the path under test.

    A process whose image name carries a non-ASCII character makes tasklist
    emit a byte that is valid cp850 and invalid UTF-8 — the exact condition
    that used to hand `text=True` a None stdout. The query runs in a child
    interpreter with PYTHONUTF8=1, because that is the environment where the
    old code reported a running process as not running.
    """
    ping = Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "PING.EXE"
    if not ping.exists():
        pytest.skip("PING.EXE not available on this machine")

    probe = tmp_path / "prueba_ñ.exe"
    shutil.copy2(ping, probe)

    proc = subprocess.Popen(
        [str(probe), "-n", "30", "127.0.0.1"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        creationflags=subprocess.CREATE_NO_WINDOW,
    )
    try:
        deadline = time.monotonic() + 10
        while not _tasklist_has_image(probe.name) and time.monotonic() < deadline:
            time.sleep(0.2)

        env = {**os.environ, "PYTHONUTF8": "1", "PYTHONIOENCODING": "utf-8"}
        child = subprocess.run(
            [sys.executable, "-c",
             "import sys;"
             "from email_archiver.outlook.client import _tasklist_has_image as q;"
             "print(q(sys.argv[1]))",
             probe.name],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            cwd=str(REPO_ROOT), env=env, timeout=60,
        )
        assert child.stdout.strip() == "True", (
            f"child stdout={child.stdout!r} stderr={child.stderr!r}"
        )
    finally:
        proc.kill()
        proc.wait(timeout=10)


@pytest.mark.skipif(not IS_WINDOWS, reason="tasklist is Windows-only")
def test_absent_process_reports_false():
    assert _tasklist_has_image("definitely_not_running_38.exe") is False

"""The headless scan hands over to folder-explorer only when its index is the configured one."""
from __future__ import annotations

import json
import subprocess
from pathlib import Path

import pytest

from email_archiver import folder_explorer
from email_archiver.cli import scan as main_scan


@pytest.fixture()
def checkout(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> Path:
    """An invented folder-explorer checkout: just the venv interpreter path, never run."""
    home = tmp_path / "fe"
    py = folder_explorer.python(home)
    py.parent.mkdir(parents=True)
    py.write_bytes(b"")
    monkeypatch.setenv(folder_explorer.HOME_ENV, str(home))
    return home


def _fake_run(conf: dict, index_code: int, calls: list):
    def run(cmd, **kwargs):
        calls.append(cmd[3:])
        if cmd[3:5] == ["email", "config"]:
            return subprocess.CompletedProcess(cmd, 0, json.dumps(conf) + "\n", "")
        return subprocess.CompletedProcess(cmd, index_code, "done\n", "")
    return run


def test_owner_needs_the_same_index_file(checkout: Path, tmp_path: Path) -> None:
    db = tmp_path / "emails.sqlite3"
    calls: list = []
    own = _fake_run({"roots": [], "db": str(db), "own": True}, 0, calls)
    assert folder_explorer.owner({"database": {"path": str(db)}}, own) == checkout
    assert folder_explorer.owner({"database": {"path": str(tmp_path / "other.db")}}, own) is None
    read_only = _fake_run({"roots": [], "db": str(db), "own": False}, 0, calls)
    assert folder_explorer.owner({"database": {"path": str(db)}}, read_only) is None


def test_no_checkout_means_no_owner(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv(folder_explorer.HOME_ENV, str(tmp_path / "nowhere"))
    assert folder_explorer.home() is None
    assert folder_explorer.owner({"database": {"path": str(tmp_path / "x.db")}}) is None


@pytest.mark.parametrize("fe_code,expected", [(0, main_scan.EXIT_OK), (2, main_scan.EXIT_BUSY), (1, main_scan.EXIT_FAILED)])
def test_headless_scan_delegates_and_maps_the_exit_code(checkout: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch,
                                                        fe_code: int, expected: int) -> None:
    db = tmp_path / "emails.sqlite3"
    calls: list = []
    monkeypatch.setattr(folder_explorer.subprocess, "run", _fake_run({"roots": [], "db": str(db), "own": True}, fe_code, calls))
    monkeypatch.setattr(main_scan, "load_config", lambda: {"database": {"path": str(db)}})
    monkeypatch.setattr(main_scan, "setup_logging", lambda _cfg: None)
    monkeypatch.setattr(main_scan, "_scan_here", lambda cfg: pytest.fail("scanned here while folder-explorer owns the index"))
    assert main_scan.main(["--no-ui"]) == expected
    assert calls[-1][:3] == ["email", "index", "--fresh"]


def test_headless_scan_runs_here_when_the_index_is_this_apps(checkout: Path, tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    calls: list = []
    monkeypatch.setattr(folder_explorer.subprocess, "run", _fake_run({"roots": [], "db": str(tmp_path / "fe.sqlite3"), "own": True}, 0, calls))
    monkeypatch.setattr(main_scan, "load_config", lambda: {"database": {"path": str(tmp_path / "emails.db")}})
    monkeypatch.setattr(main_scan, "setup_logging", lambda _cfg: None)
    ran: list = []
    monkeypatch.setattr(main_scan, "_scan_here", lambda cfg: ran.append(cfg) or 0)
    assert main_scan.main(["--no-ui"]) == 0 and len(ran) == 1
    assert all(c[:2] != ["email", "index"] for c in calls)

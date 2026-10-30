"""``python -m email_archiver <subcommand>`` and the root ``main_*.py`` wrappers.

A scheduled job and a spawning app call the root scripts by path, so each old
path must still reach the same function as its subcommand. Nothing here opens
a window, Outlook or a mailbox: the dispatcher is tested over stub ``main``s and
the wrappers over ``--help``.
"""
from __future__ import annotations

import importlib
import importlib.util
import os
import subprocess
import sys
from pathlib import Path
from types import ModuleType

import pytest

from email_archiver import cli

REPO_ROOT = Path(__file__).resolve().parents[1]
NO_WINDOW = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0

# subcommand -> root wrapper that must reach the same function
WRAPPERS = {
    "archive": "main_archive.py",
    "scan": "main_scan.py",
    "ui": "main_ui.py",
    "batch": "main_batch.py",
}


def _stub(monkeypatch: pytest.MonkeyPatch, name: str, result: int | None) -> list[object]:
    """Replace one subcommand's ``main`` with a recorder; returns the calls it saw."""
    module = importlib.import_module(f"email_archiver.cli.{cli.SUBCOMMANDS[name][0]}")
    calls: list[object] = []

    def fake(*args: object) -> int | None:
        calls.append(args)
        return result

    monkeypatch.setattr(module, "main", fake)
    return calls


@pytest.mark.parametrize("name", ["scan", "batch"])
def test_scan_and_batch_receive_the_rest_of_the_command_line(
    monkeypatch: pytest.MonkeyPatch, name: str
) -> None:
    calls = _stub(monkeypatch, name, 3)
    assert cli.main([name, "--flag", "x"]) == 3
    assert calls == [(["--flag", "x"],)]


@pytest.mark.parametrize("name", ["archive", "ui"])
def test_the_window_subcommands_take_no_arguments_and_exit_zero(
    monkeypatch: pytest.MonkeyPatch, name: str
) -> None:
    calls = _stub(monkeypatch, name, None)
    assert cli.main([name]) == 0
    assert calls == [()]


def test_a_missing_or_unknown_subcommand_is_a_usage_error(
    capsys: pytest.CaptureFixture[str],
) -> None:
    assert cli.main([]) == 2
    assert cli.main(["nope"]) == 2
    assert "unknown subcommand 'nope'" in capsys.readouterr().err


def test_help_lists_every_subcommand_and_exits_zero(
    capsys: pytest.CaptureFixture[str],
) -> None:
    assert cli.main(["--help"]) == 0
    out = capsys.readouterr().out
    assert all(name in out for name in WRAPPERS)


@pytest.mark.parametrize("subcommand,wrapper", sorted(WRAPPERS.items()))
def test_each_root_wrapper_reaches_its_subcommands_main(subcommand: str, wrapper: str) -> None:
    spec = importlib.util.spec_from_file_location(f"wrapper_{subcommand}", REPO_ROOT / wrapper)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    target: ModuleType = importlib.import_module(
        f"email_archiver.cli.{cli.SUBCOMMANDS[subcommand][0]}"
    )
    assert module.main is target.main


@pytest.mark.parametrize("argv", [["main_batch.py"], ["main_scan.py"], ["-m", "email_archiver", "batch"]])
def test_the_old_paths_and_the_module_entry_run_as_scripts(argv: list[str]) -> None:
    proc = subprocess.run(
        [sys.executable, *argv, "--help"], cwd=REPO_ROOT, capture_output=True,
        text=True, encoding="utf-8", errors="replace", timeout=60,
        stdin=subprocess.DEVNULL, creationflags=NO_WINDOW, check=False,
    )
    assert proc.returncode == 0, proc.stderr
    assert "usage:" in proc.stdout

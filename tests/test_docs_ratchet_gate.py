"""The gate's docs-ratchet stage maps the vendored script's exit codes.

The runner (timeout, unstartable script, venv interpreter) is shared with the
static-ratchet stage and pinned in test_static_ratchet_gate.py.

`scripts/docs_ratchet.py` exits 3 when it could not confirm (git unusable, a
non-UTF-8 file, a fence or comment that never closes): part of the tree was not
read. These tests pin that 3 fails as NOT CONFIRMED -- never a pass, never a
skip -- and that the pre-ship gate calls the wrapper that guarantees it rather
than the script. They run fake scripts in a temp directory: no mailbox, no Outlook.
"""

from __future__ import annotations

import tomllib
from pathlib import Path

import pytest

from scripts import docs_ratchet_gate as gate
from tests.test_static_ratchet_gate import _fake_script

REPO_ROOT = Path(__file__).resolve().parents[1]


def test_only_exit_zero_passes() -> None:
    assert gate.classify_exit(0)[0] is True
    for code in (1, 2, 3, 4, -1):
        assert gate.classify_exit(code)[0] is False, code


def test_exit_three_is_not_confirmed_and_never_a_skip() -> None:
    ok, label = gate.classify_exit(3)
    assert not ok
    assert "NOT CONFIRMED" in label
    assert "skip" not in label.lower()


@pytest.mark.parametrize(("code", "expected"), [(0, 0), (1, 1), (2, 1), (3, 1)])
def test_run_exits_nonzero_for_every_script_failure(
    tmp_path: Path, capsys: pytest.CaptureFixture[str], code: int, expected: int
) -> None:
    got = gate.run(_fake_script(tmp_path, code, "from the script"))
    out = capsys.readouterr().out
    assert got == expected
    assert "from the script" in out  # the script's own finding is echoed
    assert out.rstrip().splitlines()[-1].startswith("OK" if code == 0 else "FAIL")


def test_run_reports_a_tree_it_could_not_read(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    assert gate.run(_fake_script(tmp_path, 3, "docs/x.md: fence never closes")) == 1
    out = capsys.readouterr().out
    assert "docs/x.md: fence never closes" in out
    assert "NOT CONFIRMED" in out


def test_the_pre_ship_gate_calls_the_wrapper_not_the_script() -> None:
    gate_ps1 = (REPO_ROOT / "scripts" / "verify-before-ship.ps1").read_text(
        encoding="utf-8"
    )
    stage_lines = [
        line for line in gate_ps1.splitlines() if line.lstrip().startswith("Invoke-Stage")
    ]
    ratchet = [line for line in stage_lines if "docs_ratchet" in line]
    assert len(ratchet) == 1
    assert "-m scripts.docs_ratchet_gate" in ratchet[0]
    assert "docs_ratchet.py" not in ratchet[0]


def test_the_ratchet_is_declared_and_its_script_is_vendored() -> None:
    fleet = tomllib.loads((REPO_ROOT / ".fleet.toml").read_text(encoding="utf-8"))
    assert fleet["docs_ratchet"]["enabled"] is True
    baseline = REPO_ROOT / fleet["docs_ratchet"]["baseline"]
    assert baseline.is_file()
    entry = fleet["vendored"]["docs_ratchet"]
    assert entry["dest"] == "scripts/docs_ratchet.py"
    assert (REPO_ROOT / entry["dest"]).is_file()

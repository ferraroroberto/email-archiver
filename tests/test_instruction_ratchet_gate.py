"""The gate's instruction-ratchet stage maps the vendored script's exit codes.

The runner (timeout, unstartable script, venv interpreter) is shared with the
static-ratchet stage and pinned in test_static_ratchet_gate.py.

`scripts/instruction_ratchet.py` exits 3 when it could not confirm (git unusable,
a non-UTF-8 instruction file, a fence that never closes): part of the instruction
files was not read. These tests pin that 3 fails as NOT CONFIRMED -- never a pass,
never a skip -- and that the pre-ship gate calls the wrapper that guarantees it
rather than the script. They run fake scripts in a temp directory: no mailbox, no
Outlook.
"""

from __future__ import annotations

import tomllib
from pathlib import Path

import pytest

from scripts import instruction_ratchet_gate as gate
from tests.test_static_ratchet_gate import (
    _fake_script,
    assert_only_zero_passes_and_three_is_not_confirmed,
)

REPO_ROOT = Path(__file__).resolve().parents[1]


def test_only_exit_zero_passes_and_exit_three_is_never_a_skip() -> None:
    assert_only_zero_passes_and_three_is_not_confirmed(gate)


def test_run_fails_every_nonzero_exit_and_names_the_finding(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    assert gate.run(_fake_script(tmp_path, 0, "all clear")) == 0
    assert capsys.readouterr().out.rstrip().splitlines()[-1].startswith("OK")
    for code in (1, 2, 3):
        assert gate.run(_fake_script(tmp_path, code, "CLAUDE.md: R-03 rose")) == 1
        out = capsys.readouterr().out
        assert "CLAUDE.md: R-03 rose" in out  # the script's finding is echoed
        assert out.rstrip().splitlines()[-1].startswith("FAIL")
    assert "NOT CONFIRMED" in out  # the last code run was 3


def test_the_pre_ship_gate_calls_the_wrapper_after_the_structure_ratchet() -> None:
    gate_ps1 = (REPO_ROOT / "scripts" / "verify-before-ship.ps1").read_text(
        encoding="utf-8"
    )
    stage_lines = [
        line for line in gate_ps1.splitlines() if line.lstrip().startswith("Invoke-Stage")
    ]
    ratchet = [line for line in stage_lines if "instruction_ratchet" in line]
    assert len(ratchet) == 1
    assert "-m scripts.instruction_ratchet_gate" in ratchet[0]
    assert "instruction_ratchet.py" not in ratchet[0]
    assert stage_lines.index(ratchet[0]) == 1 + next(
        i for i, line in enumerate(stage_lines) if "structure_ratchet" in line
    )


def test_the_ratchet_is_declared_and_its_script_is_vendored() -> None:
    fleet = tomllib.loads((REPO_ROOT / ".fleet.toml").read_text(encoding="utf-8"))
    assert fleet["instruction_ratchet"]["enabled"] is True
    baseline = REPO_ROOT / fleet["instruction_ratchet"]["baseline"]
    assert baseline.is_file()
    entry = fleet["vendored"]["instruction_ratchet"]
    assert entry["dest"] == "scripts/instruction_ratchet.py"
    assert (REPO_ROOT / entry["dest"]).is_file()

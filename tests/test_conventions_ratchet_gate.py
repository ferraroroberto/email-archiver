"""The gate's conventions-ratchet stage maps the vendored script's exit codes.

`scripts/conventions_ratchet.py` exits 3 when it could not confirm (git unusable, a
tracked Python file that does not parse): a file that was not read may hold a
violation. The runner is shared with the static-ratchet stage and pinned in
test_static_ratchet_gate.py; these tests pin the verdicts and that the pre-ship
gate calls the wrapper, not the script.
"""

from __future__ import annotations

import tomllib
from pathlib import Path

import pytest

from scripts import conventions_ratchet_gate as gate
from tests.test_static_ratchet_gate import (
    _fake_script,
    assert_only_zero_passes_and_three_is_not_confirmed,
)

REPO_ROOT = Path(__file__).resolve().parents[1]


def test_only_exit_zero_passes_and_exit_three_is_never_a_skip() -> None:
    assert_only_zero_passes_and_three_is_not_confirmed(gate)


def test_run_fails_every_nonzero_exit_and_names_the_rule(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    assert gate.run(_fake_script(tmp_path, 0, "all clear")) == 0
    assert capsys.readouterr().out.rstrip().splitlines()[-1].startswith("OK")
    for code in (1, 2, 3):
        assert gate.run(_fake_script(tmp_path, code, "email_archiver/x.py:3: [no-drive-path] literal")) == 1
        out = capsys.readouterr().out
        assert "[no-drive-path]" in out  # the script's finding is echoed
        assert out.rstrip().splitlines()[-1].startswith("FAIL")


def test_the_gate_runs_the_wrapper_straight_after_the_dead_code_ratchet() -> None:
    gate_ps1 = (REPO_ROOT / "scripts" / "verify-before-ship.ps1").read_text(encoding="utf-8")
    stages = [line for line in gate_ps1.splitlines() if line.lstrip().startswith("Invoke-Stage")]
    ours = [line for line in stages if "conventions_ratchet" in line]
    assert len(ours) == 1
    assert "-m scripts.conventions_ratchet_gate" in ours[0]
    assert "conventions_ratchet.py" not in ours[0]
    assert stages.index(ours[0]) == 1 + next(
        i for i, line in enumerate(stages) if "dead_code_ratchet" in line
    )


def test_the_ratchet_is_declared_and_its_script_is_vendored() -> None:
    fleet = tomllib.loads((REPO_ROOT / ".fleet.toml").read_text(encoding="utf-8"))
    assert fleet["conventions_ratchet"]["enabled"] is True
    assert (REPO_ROOT / fleet["conventions_ratchet"]["baseline"]).is_file()
    entry = fleet["vendored"]["conventions_ratchet"]
    assert entry["dest"] == "scripts/conventions_ratchet.py"
    assert (REPO_ROOT / entry["dest"]).is_file()

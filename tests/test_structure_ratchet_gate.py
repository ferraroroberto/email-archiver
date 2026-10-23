"""The gate's structure-ratchet stage maps the vendored script's exit codes.

Exit 3 means part of the tree could not be measured. It must fail as NOT
CONFIRMED -- never a pass, never a skip -- and the pre-ship gate must call the
wrapper that guarantees it, not the script. The vendored script's own
behaviour is project-scaffolding's to test; here the script is a fake in a temp
directory, plus one run of the real one over this tree.
"""

from __future__ import annotations

import subprocess
import sys
import tomllib
from pathlib import Path

import pytest

from scripts import structure_ratchet_gate as gate
from tests.test_static_ratchet_gate import _fake_script

REPO_ROOT = Path(__file__).resolve().parents[1]
FLEET = tomllib.loads((REPO_ROOT / ".fleet.toml").read_text(encoding="utf-8"))


def test_only_exit_zero_passes_and_exit_three_is_never_a_skip() -> None:
    verdicts = {code: gate.classify_exit(code) for code in range(-1, 5)}
    assert [code for code, (ok, _) in verdicts.items() if ok] == [0]
    assert "NOT CONFIRMED" in verdicts[3][1]
    assert "skip" not in verdicts[3][1].lower()


def test_run_fails_every_nonzero_exit_and_echoes_the_finding(
    tmp_path: Path, capsys: pytest.CaptureFixture[str]
) -> None:
    assert gate.run(_fake_script(tmp_path, 0, "all clear")) == 0
    assert capsys.readouterr().out.rstrip().splitlines()[-1].startswith("OK")
    for code in (1, 2, 3):
        assert gate.run(_fake_script(tmp_path, code, "size:x.py grew")) == 1
        out = capsys.readouterr().out
        assert "size:x.py grew" in out
        assert out.rstrip().splitlines()[-1].startswith("FAIL")


def test_the_pre_ship_gate_calls_the_wrapper_not_the_script() -> None:
    gate_ps1 = (REPO_ROOT / "scripts" / "verify-before-ship.ps1").read_text(encoding="utf-8")
    stages = [line for line in gate_ps1.splitlines() if "structure_ratchet" in line and "Invoke-Stage" in line]
    assert len(stages) == 1
    assert "-m scripts.structure_ratchet_gate" in stages[0]
    assert "structure_ratchet.py" not in stages[0]


def test_the_declared_shape_and_baseline_match_the_vendored_script() -> None:
    assert FLEET["shape"] == "pipeline-cli"
    assert FLEET["structure_ratchet"]["enabled"] is True
    assert (REPO_ROOT / FLEET["structure_ratchet"]["baseline"]).is_file()
    entry = FLEET["vendored"]["structure_ratchet"]
    assert (REPO_ROOT / entry["dest"]).is_file()
    assert len(entry["sha"]) == 40


def test_the_vendored_script_passes_on_this_tree() -> None:
    proc = subprocess.run(
        [sys.executable, str(gate.SCRIPT)], cwd=REPO_ROOT, capture_output=True, text=True,
        encoding="utf-8", errors="replace", stdin=subprocess.DEVNULL,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0), timeout=120, check=False,
    )
    assert proc.returncode == 0, proc.stdout + proc.stderr


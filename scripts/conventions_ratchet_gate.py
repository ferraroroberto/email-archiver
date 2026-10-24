r"""The gate's conventions-ratchet stage: run the vendored script, map its exit code.

    & .\.venv\Scripts\python.exe -m scripts.conventions_ratchet_gate

`scripts/conventions_ratchet.py` is vendored and never edited here. It exits 0 on a
pass (or when the stage is not enabled), 1 on a rule count above the baseline or a
missing baseline, 2 on a bad table, baseline or arguments, and 3 when a tracked
Python file could not be read or parsed, or `git` is unusable. Only 0 passes; 3 is
reported as its own state, NOT CONFIRMED, and fails the gate. The mapping lives here
so a test can pin it, and `verify-before-ship.ps1` calls this wrapper, never the script.
"""

from __future__ import annotations

import sys
from pathlib import Path

from scripts import static_ratchet_gate as shared

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPT = REPO_ROOT / "scripts" / "conventions_ratchet.py"

_FAIL_LABELS = {
    shared.EXIT_NEW: "NEW convention hit, or missing baseline",
    shared.EXIT_CONFIG: "bad [conventions_ratchet] table, baseline or arguments",
    shared.EXIT_NOT_CONFIRMED: "NOT CONFIRMED -- a Python file could not be parsed, so it was not checked",
}


def classify_exit(code: int) -> tuple[bool, str]:
    """`(ok, label)` for one script exit code. Only 0 passes."""
    if code == shared.EXIT_PASS:
        return True, "no convention hit beyond the baseline"
    return False, _FAIL_LABELS.get(code, f"unexpected exit {code}")


def run(script: Path = SCRIPT) -> int:
    """The conventions-ratchet stage: the shared runner with this stage's verdicts."""
    return shared.run_stage(script, "conventions ratchet", classify_exit)


if __name__ == "__main__":
    sys.exit(run())

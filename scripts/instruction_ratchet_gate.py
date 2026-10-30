r"""The gate's instruction-ratchet stage: run the vendored script, map its exit code.

    & .\.venv\Scripts\python.exe -m scripts.instruction_ratchet_gate

`scripts/instruction_ratchet.py` is vendored and never edited here. It exits 0 on a
pass, 1 on a rule count that rose or a file past its size cap, 2 on a bad table,
baseline or arguments, and 3 when an instruction file could not be measured (`git`
unusable, an unreadable or non-UTF-8 file, an unclosed fence, an unreadable skill
description). Only 0 passes; 3 is reported as its own state, NOT CONFIRMED, and fails
the gate. The mapping lives here so a test can pin it, and `verify-before-ship.ps1`
calls this wrapper, never the script.
"""

from __future__ import annotations

import sys
from pathlib import Path

from scripts import static_ratchet_gate as shared

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPT = REPO_ROOT / "scripts" / "instruction_ratchet.py"

_FAIL_LABELS = {
    shared.EXIT_NEW: "NEW or grown rule hit, file past its size cap, or missing baseline",
    shared.EXIT_CONFIG: "bad [instruction_ratchet] table, baseline or arguments",
    shared.EXIT_NOT_CONFIRMED: "NOT CONFIRMED -- an instruction file could not be measured, so it was not checked",
}


def classify_exit(code: int) -> tuple[bool, str]:
    """`(ok, label)` for one script exit code. Only 0 passes."""
    if code == shared.EXIT_PASS:
        return True, "no instruction-file rule count above its baseline"
    return False, _FAIL_LABELS.get(code, f"unexpected exit {code}")


def run(script: Path = SCRIPT) -> int:
    """The instruction-ratchet stage: the shared runner with this stage's verdicts."""
    return shared.run_stage(script, "instruction ratchet", classify_exit)


if __name__ == "__main__":
    sys.exit(run())

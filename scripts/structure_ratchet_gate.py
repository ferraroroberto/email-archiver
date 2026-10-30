r"""The gate's structure-ratchet stage: run the vendored script, map its exit code.

    & .\.venv\Scripts\python.exe -m scripts.structure_ratchet_gate

`scripts/structure_ratchet.py` is vendored and never edited here. It exits 0 on a
pass, 1 on a new or grown offender or a root folder outside the declared shape,
2 on a bad table or baseline, and 3 when part of the tree could not be measured
(`git` unusable, an unreadable or unparseable file, an unknown `shape`). Only 0
passes; 3 is reported as its own state, NOT CONFIRMED, and fails the gate. The
mapping lives here so a test can pin it, and `verify-before-ship.ps1` calls this
wrapper, never the script.
"""

from __future__ import annotations

import sys
from pathlib import Path

from scripts import static_ratchet_gate as shared

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPT = REPO_ROOT / "scripts" / "structure_ratchet.py"

_FAIL_LABELS = {
    shared.EXIT_NEW: "NEW or grown offender, root folder outside the shape, or missing baseline",
    shared.EXIT_CONFIG: "bad [structure_ratchet] table, baseline or arguments",
    shared.EXIT_NOT_CONFIRMED: "NOT CONFIRMED -- part of the tree could not be measured, so it was not checked",
}


def classify_exit(code: int) -> tuple[bool, str]:
    """`(ok, label)` for one script exit code. Only 0 passes."""
    if code == shared.EXIT_PASS:
        return True, "no new or grown offender and the root layout matches its shape"
    return False, _FAIL_LABELS.get(code, f"unexpected exit {code}")


def run(script: Path = SCRIPT) -> int:
    """The structure-ratchet stage: the shared runner with this stage's verdicts."""
    return shared.run_stage(script, "structure ratchet", classify_exit)


if __name__ == "__main__":
    sys.exit(run())

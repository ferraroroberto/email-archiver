r"""The gate's docs-ratchet stage: run the vendored script, map its exit code.

    & .\.venv\Scripts\python.exe -m scripts.docs_ratchet_gate

`scripts/docs_ratchet.py` (vendored, never edited here) exits 0 on a pass, 1 on a
new marker or a README past its budget, 2 on a bad table or baseline and 3 when
it could not confirm (git unusable, a non-UTF-8 file, a fence or comment that
never closes). Exit 3 means part of the tree was not read, so it fails the gate
as its own state and never reads as a pass or a skip. This wrapper is the one
place that mapping lives, so a test can pin it; `verify-before-ship.ps1` calls
the wrapper, never the script. The runner itself is shared with the static-ratchet
stage.

Exit 0 only when the script exited 0; every other outcome exits 1 with its
reason on the last line.
"""

from __future__ import annotations

import sys
from pathlib import Path

from scripts.static_ratchet_gate import (
    EXIT_CONFIG, EXIT_NEW, EXIT_NOT_CONFIRMED, EXIT_PASS, run_stage,
)

REPO_ROOT = Path(__file__).resolve().parent.parent
SCRIPT = REPO_ROOT / "scripts" / "docs_ratchet.py"

_FAIL_LABELS = {
    EXIT_NEW: "NEW history marker, README past its budget, or missing baseline",
    EXIT_CONFIG: "bad [docs_ratchet] table, baseline or arguments",
    EXIT_NOT_CONFIRMED: "NOT CONFIRMED -- a file could not be read, so it was not checked",
}


def classify_exit(code: int) -> tuple[bool, str]:
    """`(ok, label)` for one script exit code. Only 0 passes."""
    if code == EXIT_PASS:
        return True, "no new history marker and the README within its budget"
    return False, _FAIL_LABELS.get(code, f"unexpected exit {code}")


def run(script: Path = SCRIPT) -> int:
    """The docs-ratchet stage: the shared runner with this stage's verdicts."""
    return run_stage(script, "docs ratchet", classify_exit)


if __name__ == "__main__":
    sys.exit(run())

r"""The gate's dead-code-ratchet stage: run the vendored script, map its exit code.

    & .\.venv\Scripts\python.exe -m scripts.dead_code_ratchet_gate

`scripts/dead_code_ratchet.py` is vendored and never edited here. Exit 1 is a dead
definition outside the baseline or a missing baseline; exit 3 is a tracked Python
file that could not be read or parsed, or `git` unusable. Only 0 passes, 3 is NOT
CONFIRMED (see `static_ratchet_gate`), and `verify-before-ship.ps1` calls this
wrapper, never the script.
"""

from __future__ import annotations

import sys
from pathlib import Path

from scripts import static_ratchet_gate as shared

SCRIPT = shared.REPO_ROOT / "scripts" / "dead_code_ratchet.py"

classify_exit = shared.verdicts(
    "no dead definition beyond the baseline",
    "NEW dead definition, or missing baseline",
    "bad [dead_code_ratchet] table, baseline or arguments",
)


def run(script: Path = SCRIPT) -> int:
    """The dead-code-ratchet stage: the shared runner with this stage's verdicts."""
    return shared.run_stage(script, "dead code ratchet", classify_exit)


if __name__ == "__main__":
    sys.exit(run())

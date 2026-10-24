r"""The gate's conventions-ratchet stage: run the vendored script, map its exit code.

    & .\.venv\Scripts\python.exe -m scripts.conventions_ratchet_gate

`scripts/conventions_ratchet.py` is vendored and never edited here. It exits 0 on a
pass or when the stage is not enabled; exit 1 is a rule count above the baseline or a
missing baseline; exit 3 is a tracked Python file that could not be read or parsed,
or `git` unusable. Only 0 passes, 3 is NOT CONFIRMED (see `static_ratchet_gate`), and
`verify-before-ship.ps1` calls this wrapper, never the script.
"""

from __future__ import annotations

import sys
from pathlib import Path

from scripts import static_ratchet_gate as shared

SCRIPT = shared.REPO_ROOT / "scripts" / "conventions_ratchet.py"

classify_exit = shared.verdicts(
    "no convention hit beyond the baseline",
    "NEW convention hit, or missing baseline",
    "bad [conventions_ratchet] table, baseline or arguments",
)


def run(script: Path = SCRIPT) -> int:
    """The conventions-ratchet stage: the shared runner with this stage's verdicts."""
    return shared.run_stage(script, "conventions ratchet", classify_exit)


if __name__ == "__main__":
    sys.exit(run())

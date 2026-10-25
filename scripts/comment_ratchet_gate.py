r"""The gate's comment-ratchet stage: run the vendored script, map its exit code.

    & .\.venv\Scripts\python.exe -m scripts.comment_ratchet_gate

`scripts/comment_ratchet.py` is vendored and never edited here. It exits 0 on a pass or
when the stage is not enabled; exit 1 is a file whose `#NNN` comment-reference count rose
above the baseline or a missing baseline; exit 3 is a file that could not be scanned, or
`git` unusable. Only 0 passes, 3 is NOT CONFIRMED (see `static_ratchet_gate`), and
`verify-before-ship.ps1` calls this wrapper, never the script.
"""

from __future__ import annotations

import sys
from pathlib import Path

from scripts import static_ratchet_gate as shared

SCRIPT = shared.REPO_ROOT / "scripts" / "comment_ratchet.py"

classify_exit = shared.verdicts(
    "no issue reference in a comment beyond the baseline",
    "NEW issue reference in a comment, or missing baseline",
    "bad [comment_ratchet] table, baseline or arguments",
)


def run(script: Path = SCRIPT) -> int:
    """The comment-ratchet stage: the shared runner with this stage's verdicts."""
    return shared.run_stage(script, "comment ratchet", classify_exit)


if __name__ == "__main__":
    sys.exit(run())

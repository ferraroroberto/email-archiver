"""Log records reach a captured stderr as UTF-8 (issue #87).

``main_batch.py`` is spawned by task-os with its streams captured and no
``PYTHONUTF8``, and task-os decodes stderr as UTF-8. Without the fix, Python
writes stderr in the ANSI code page under a pipe, so an accented archive path
comes back as U+FFFD and an arrow as a literal ``\\u2192`` escape. The test runs
a real child process with stderr piped, because the encoding a pipe gets is the
whole point and cannot be faked in-process.
"""
from __future__ import annotations

import os
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[1]

_CHILD = """
import logging, sys
from email_archiver.config import setup_logging
setup_logging({"logging": {"level": "INFO", "file": sys.argv[1]}})
logging.getLogger("t").info("Archived email 003 \\u2192 Documentos/Comunicaci\\u00f3n.msg")
"""


def test_a_captured_stderr_gets_log_records_as_utf8(tmp_path):
    env = {k: v for k, v in os.environ.items() if k not in ("PYTHONUTF8", "PYTHONIOENCODING")}
    env["PYTHONPATH"] = str(REPO_ROOT)
    flags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
    proc = subprocess.run(
        [sys.executable, "-c", _CHILD, str(tmp_path / "a.log")],
        cwd=REPO_ROOT, env=env, capture_output=True, timeout=60,
        creationflags=flags,
    )

    assert proc.returncode == 0, proc.stderr
    stderr = proc.stderr.decode("utf-8", errors="replace")
    assert "Archived email 003 → Documentos/Comunicación.msg" in stderr
    assert "�" not in stderr

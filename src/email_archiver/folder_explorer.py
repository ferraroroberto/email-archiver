"""Hand the headless scan to folder-explorer once it owns the email index.

Folder-explorer indexes the email folders itself and is the one writer of its
index. When this app's ``database.path`` points at that index, scanning here
would make two writers of one file, so ``main_scan.py --no-ui`` runs
folder-explorer's ``email index --fresh`` instead (a live listing, so a mail
filed a moment ago is seen). In every other case the scan here runs as before.

Folder-explorer is the checkout named by ``FOLDER_EXPLORER_HOME``, else a
sibling ``folder-explorer`` folder beside this repo; it is asked where its
index is with ``python -m app.cli email config --json``.
"""
from __future__ import annotations

import json
import logging
import os
import subprocess
import sys
from pathlib import Path
from typing import Any, Callable, Optional

from email_archiver.config import PROJECT_ROOT

logger = logging.getLogger(__name__)

HOME_ENV = "FOLDER_EXPLORER_HOME"
NO_WINDOW = subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0

# folder-explorer's `email index` exit codes.
FE_OK, FE_BUSY = 0, 2
# How long the config probe and the index run may take, and how long the run
# waits for another email run to let go of folder-explorer's lock. The index
# run stays inside the 900 s a consumer allows for the whole re-index.
PROBE_TIMEOUT_S = 60
INDEX_TIMEOUT_S = 840
LOCK_WAIT_S = 600

Runner = Callable[..., "subprocess.CompletedProcess[str]"]


def home() -> Optional[Path]:
    """The folder-explorer checkout with its venv, or None when there is none."""
    raw = os.environ.get(HOME_ENV, "").strip()
    path = Path(raw) if raw else PROJECT_ROOT.parent / "folder-explorer"
    return path if python(path).is_file() else None


def python(checkout: Path) -> Path:
    return checkout / ".venv" / ("Scripts/python.exe" if sys.platform == "win32" else "bin/python")


def _run(checkout: Path, args: list[str], timeout: float, run: Optional[Runner]) -> "subprocess.CompletedProcess[str]":
    env = dict(os.environ, PYTHONUTF8="1", PYTHONUNBUFFERED="1")
    cmd = [str(python(checkout)), "-m", "app.cli", *args]
    return (run or subprocess.run)(cmd, cwd=str(checkout), env=env, capture_output=True, text=True, encoding="utf-8",
                                   timeout=timeout, creationflags=NO_WINDOW, stdin=subprocess.DEVNULL)


def index_config(checkout: Path, run: Optional[Runner] = None) -> Optional[dict[str, Any]]:
    """folder-explorer's ``{roots, db, own}``, or None when it cannot say."""
    try:
        proc = _run(checkout, ["email", "config", "--json"], PROBE_TIMEOUT_S, run)
        return json.loads((proc.stdout or "").strip().splitlines()[-1]) if proc.returncode == 0 else None
    except (OSError, subprocess.TimeoutExpired, ValueError, IndexError) as exc:
        logger.info("ℹ️ folder-explorer did not report its email index: %s", exc)
        return None


def _same_file(a: str, b: str) -> bool:
    return os.path.normcase(os.path.realpath(a)) == os.path.normcase(os.path.realpath(b))


def owner(cfg: dict[str, Any], run: Optional[Runner] = None) -> Optional[Path]:
    """The folder-explorer checkout that owns ``database.path``, or None when the index is this app's."""
    checkout = home()
    if checkout is None:
        return None
    conf = index_config(checkout, run)
    db = str(cfg.get("database", {}).get("path") or "")
    if conf and conf.get("own") and conf.get("db") and db and _same_file(db, str(conf["db"])):
        return checkout
    return None


def index(checkout: Path, run: Optional[Runner] = None) -> int:
    """Run folder-explorer's fresh email index; returns its exit code (2: busy)."""
    try:
        proc = _run(checkout, ["email", "index", "--fresh", "--wait", str(LOCK_WAIT_S)], INDEX_TIMEOUT_S, run)
    except subprocess.TimeoutExpired:
        logger.error("folder-explorer's email index ran past %d s and was stopped", INDEX_TIMEOUT_S)
        return 1
    except OSError as exc:
        logger.error("folder-explorer's email index could not start: %s", exc)
        return 1
    for line in (proc.stdout or "").splitlines()[-5:]:
        logger.info("folder-explorer: %s", line)
    return proc.returncode

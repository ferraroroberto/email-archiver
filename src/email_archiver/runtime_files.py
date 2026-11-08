"""Writers for runtime files whose path the runtime-write ratchet cannot resolve.

The SQLite index, the log file and the archive tree sit where ``config.json``
or the caller says, so the ratchet cannot follow them statically and each write
would need its own exempt marker. These writers take the path, refuse it at
runtime when it resolves under one of the code folders the ratchet guards, and
then write, so the guarantee a marker only claimed is enforced.

Stdlib-only and import-free of the rest of the package, so ``config`` and
``database`` can both use it.
"""
from __future__ import annotations

import logging
import os
import sqlite3
from pathlib import Path
from typing import Any, Union

PathLike = Union[str, "os.PathLike[str]"]

REPO_ROOT = Path(__file__).resolve().parents[2]

# The folders `[runtime_write_ratchet] code_folders` guards (its default set).
CODE_FOLDERS = ("src", "app", "scripts", "hooks", "skills")


IN_MEMORY = ":memory:"


def checked(path: PathLike) -> Path:
    """``path`` resolved, or ``ValueError`` when it sits under a code folder."""
    resolved = Path(path).expanduser().resolve()
    for name in CODE_FOLDERS:
        folder = REPO_ROOT / name
        if resolved == folder or folder in resolved.parents:
            raise ValueError(
                f"Refusing to write under the code folder {folder}: {resolved}. "
                "Runtime output belongs in data/, config/ or the archive tree."
            )
    return resolved


def make_dirs(path: PathLike, *, parents: bool = True, exist_ok: bool = True) -> Path:
    """Create the folder ``path``; refuses a path under a code folder."""
    target = checked(path)
    # runtime-write-exempt: caller- or config-named folder; checked() above raises when it is under a code folder
    target.mkdir(parents=parents, exist_ok=exist_ok)
    return target


def file_handler(path: PathLike, encoding: str = "utf-8") -> logging.FileHandler:
    """A ``logging.FileHandler`` on ``path``; refuses a path under a code folder."""
    target = checked(path)
    # runtime-write-exempt: config-named log file; checked() above raises when it is under a code folder
    return logging.FileHandler(target, encoding=encoding)


def connect_sqlite(path: PathLike, **kwargs: Any) -> sqlite3.Connection:
    """Open the SQLite file ``path``; refuses a path under a code folder.

    ``:memory:`` is no file, so it passes through untouched."""
    if str(path) == IN_MEMORY:
        return sqlite3.connect(IN_MEMORY, **kwargs)
    target = checked(path)
    # runtime-write-exempt: config-named database; checked() above raises when it is under a code folder
    return sqlite3.connect(str(target), **kwargs)

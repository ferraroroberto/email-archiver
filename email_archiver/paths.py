"""
Archive-root path safety, in one place.

Two different operations write to (or rename inside) the archive tree —
``batch.revert``'s deletes and ``renumber``'s renames — and both must refuse
anything that resolves outside the configured ``archive.root_paths``. That
guard is a safety primitive, so it is defined exactly once here rather than
copied into each module: a second spelling of "is this inside the archive?" is
a second chance to get it wrong.

The refusal is deliberately not an exception. ``revert`` reports it per file
and carries on with the rest of the bundle; ``renumber`` reports it per folder
and leaves that folder untouched.
"""
from __future__ import annotations

import logging
import os
from pathlib import Path
from typing import Any

from email_archiver.config import get_archive_roots

logger = logging.getLogger(__name__)

# Why a path was refused.
REFUSED_OUTSIDE_ROOTS = "outside_archive_roots"
REFUSED_UNRESOLVABLE = "unresolvable_path"


def is_within(child: Path, root: Path) -> bool:
    """Whether ``child`` is ``root`` or lives underneath it.

    Compared through ``os.path.normcase`` rather than ``Path.is_relative_to``
    alone: on Windows the two paths routinely differ in case (the config spells
    a root one way, ``resolve()`` another) and a case-sensitive comparison
    would refuse a file that is genuinely inside the archive. The trailing
    separator is what stops ``C:\\Archive`` from matching ``C:\\ArchiveOther``.
    """
    child_s = os.path.normcase(str(child))
    root_s = os.path.normcase(str(root))
    if child_s == root_s:
        return True
    return child_s.startswith(root_s.rstrip("\\/") + os.sep)


def resolve_under_roots(
    raw_path: str, roots: list[Path]
) -> tuple[Path | None, str | None]:
    """Resolve a path and require it to sit under one of the archive roots.

    Returns ``(path, None)`` when it does, ``(None, reason)`` when it does not.
    """
    try:
        resolved = Path(raw_path).resolve()
    except (OSError, ValueError):
        return None, REFUSED_UNRESOLVABLE
    for root in roots:
        if is_within(resolved, root):
            return resolved, None
    return None, REFUSED_OUTSIDE_ROOTS


def archive_roots(cfg: dict[str, Any]) -> list[Path]:
    """The configured archive roots, resolved; an unresolvable one is dropped
    with a warning rather than silently widening or narrowing the guard."""
    resolved: list[Path] = []
    for raw in get_archive_roots(cfg):
        try:
            resolved.append(Path(raw).resolve())
        except (OSError, ValueError):
            logger.warning("Ignoring unresolvable archive root: %r", raw)
    return resolved

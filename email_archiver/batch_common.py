"""
What every batch-family document shares: its version, its envelope, its error
shape and the renumbering step ``apply`` and ``revert`` both end with.

Kept apart from ``email_archiver.batch`` so the verbs that build documents
(``batch``, ``draft``, ``find``, ``send``) and the modules that support them
import one small base, never each other.
"""
from __future__ import annotations

import logging
import os
from datetime import datetime
from typing import Any

from email_archiver.config import (
    Mailbox,
    get_mailbox_archive_folder,
    get_outlook_category,
)
from email_archiver.database.repository import EmailRepository
from email_archiver.paths import archive_roots as _archive_roots
from email_archiver.renumber import renumber_folder

logger = logging.getLogger(__name__)

# Bumped when the shape of a document below changes incompatibly, so a consumer
# spawning this process can refuse a version it does not understand instead of
# reading a field that silently moved.
SCHEMA_VERSION = 1


# `error.code` values more than one verb reports. The move landed and only the
# tag did not: its own code because the two are genuinely different states to
# recover from — after a move_failed the mail is still in the Inbox, after a
# category_failed it is already filed and only looks untouched in Outlook.
ERROR_NOT_IN_INBOX = "not_in_inbox"            # already moved, or never there
ERROR_CATEGORY_FAILED = "category_failed"


def now_iso() -> str:
    """The local-time timestamp every batch-family document carries."""
    return datetime.now().astimezone().isoformat(timespec="seconds")


def header(verb: str) -> dict[str, Any]:
    """The three keys every verb's document opens with, so a schema change to
    them lands here and nowhere else."""
    return {
        "verb": verb,
        "schema_version": SCHEMA_VERSION,
        "generated_at": now_iso(),
    }


def envelope(
    verb: str, cfg: dict[str, Any], category: str | None = None,
    mailbox: Mailbox | None = None,
) -> dict[str, Any]:
    """The fields every verb's document opens with. ``category`` is the one
    this run actually used, when the caller overrode the config's;
    ``archive_folder`` is the one of ``mailbox`` (``main_batch`` adds the
    ``mailbox`` key itself, to every Outlook verb's document)."""
    return {
        **header(verb),
        "archive_folder": get_mailbox_archive_folder(cfg, mailbox),
        "category": category or get_outlook_category(cfg),
    }


def unique_folders(paths: list[str]) -> list[str]:
    """The distinct folders in ``paths``, first spelling wins.

    Deduplicated case-insensitively because Windows hands the same folder back
    spelled several ways, and renumbering one folder twice in a run would
    reorder what the first pass just fixed.
    """
    seen: set[str] = set()
    folders: list[str] = []
    for raw in paths:
        folder = os.path.normpath(raw)
        key = os.path.normcase(folder)
        if key in seen:
            continue
        seen.add(key)
        folders.append(folder)
    return folders


def renumber_folders(
    cfg: dict[str, Any], repo: EmailRepository, folders: list[str]
) -> tuple[dict[str, list[dict[str, Any]]], list[dict[str, Any]]]:
    """Renumber each touched folder once; return ``(renumbered, refused)``.

    ``renumbered`` is the map a consumer heals its stored paths from, keyed by
    folder. A folder that could not be renumbered at all is **not** an empty
    map in there — an empty map means "already in order", which is a different
    fact — it goes into ``refused`` with its reason.
    """
    roots = _archive_roots(cfg)
    renumbered: dict[str, list[dict[str, Any]]] = {}
    refused: list[dict[str, Any]] = []

    for folder in folders:
        try:
            result = renumber_folder(folder, repo, roots)
        except Exception as exc:  # one folder's failure is not the run's
            logger.exception("Renumbering %r failed", folder)
            refused.append({
                "folder_path": folder,
                "reason": f"{type(exc).__name__}: {exc}",
            })
            continue
        if result.refused:
            refused.append({"folder_path": folder, "reason": result.refused})
            continue
        renumbered[result.folder_path] = result.renamed

    return renumbered, refused


def error_document(verb: str, code: str, message: str, **details: Any) -> dict[str, Any]:
    """The document printed when the run could not start at all.

    Deliberately the same envelope shape minus the results, so a consumer
    parses one JSON document either way and branches on ``"error" in doc``
    rather than on the exit code alone. ``details`` add keys to ``error``
    (``send``'s refusal names the parts that differ under ``differs``).
    """
    return {
        **header(verb),
        "error": {"code": code, "message": message, **details},
    }

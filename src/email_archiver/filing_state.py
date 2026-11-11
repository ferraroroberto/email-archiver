"""
The per-mail filing state of ``apply`` and ``dismiss``, declared once.

Both verbs walk each mail through look-up, write, move and tag. What a caller
reads back (``ok``, ``moved``, ``categorized``) is a fact about how far that
walk got, so it is derived here from what the run observed instead of being
set by whichever branch happened to run last. ``docs/filing-state.md`` has the
state table.

The module is pure: it reads no Outlook, disk or index, so a table test drives
every state.
"""
from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Any


class FilingState(str, Enum):
    """Where one mail's filing run ended. The set is closed."""

    # The decision could not be acted on; nothing was looked up or written.
    REFUSED = "refused"
    # The look-up ran and the mail is not in the folder it was asked for.
    NOT_FOUND = "not_found"
    # The look-up itself raised, so whether the mail is there is not known.
    LOOKUP_FAILED = "lookup_failed"
    # The files could not be written (``apply`` only); the mail is untouched.
    ARCHIVE_FAILED = "archive_failed"
    # The move did not land: the mail is still in the Inbox (``apply`` has
    # written its files by then, ``dismiss`` writes none).
    MOVE_FAILED = "move_failed"
    # The move landed and the tag did not stick.
    MOVED_UNTAGGED = "moved_untagged"
    # Moved and tagged: the mail is filed.
    FILED = "filed"
    # A Sent Items mail archived but never moved, and the tag did not stick.
    KEPT_IN_SENT_UNTAGGED = "kept_in_sent_untagged"
    # A Sent Items mail archived and tagged where it is.
    FILED_KEPT_IN_SENT = "filed_kept_in_sent"


class Lookup(str, Enum):
    """What the Outlook look-up of the mail answered."""

    NOT_ASKED = "not_asked"
    FOUND = "found"
    MISSING = "missing"
    RAISED = "raised"


@dataclass(frozen=True)
class FilingFacts:
    """Everything one run observed about one mail; the only input of the derive."""

    refused: bool = False
    lookup: Lookup = Lookup.NOT_ASKED
    archive_failed: bool = False
    moved: bool = False
    kept_in_sent: bool = False
    categorized: bool = False


# ok, moved, categorized: the three result keys each state stands for.
_RESULT_FLAGS: dict[FilingState, tuple[bool, bool, bool]] = {
    FilingState.REFUSED: (False, False, False),
    FilingState.NOT_FOUND: (False, False, False),
    FilingState.LOOKUP_FAILED: (False, False, False),
    FilingState.ARCHIVE_FAILED: (False, False, False),
    FilingState.MOVE_FAILED: (False, False, False),
    FilingState.MOVED_UNTAGGED: (False, True, False),
    FilingState.FILED: (True, True, True),
    FilingState.KEPT_IN_SENT_UNTAGGED: (False, False, False),
    FilingState.FILED_KEPT_IN_SENT: (True, False, True),
}


def _check_possible(facts: FilingFacts) -> None:
    """Refuse a combination the walk cannot reach, so a bug is loud and local."""
    reached_move = facts.moved or facts.kept_in_sent or facts.categorized
    if facts.refused and (facts.lookup is not Lookup.NOT_ASKED or facts.archive_failed or reached_move):
        raise ValueError(f"a refused decision cannot have progressed: {facts}")
    if not facts.refused and facts.lookup is Lookup.NOT_ASKED:
        raise ValueError(f"no look-up was observed and the decision was not refused: {facts}")
    if facts.lookup is not Lookup.FOUND and (facts.archive_failed or reached_move):
        raise ValueError(f"a mail that was not found cannot have been archived or moved: {facts}")
    if facts.archive_failed and reached_move:
        raise ValueError(f"a mail whose files were not written cannot have been moved: {facts}")
    if facts.moved and facts.kept_in_sent:
        raise ValueError(f"a mail is either moved or kept in Sent Items: {facts}")
    if facts.categorized and not (facts.moved or facts.kept_in_sent):
        raise ValueError(f"a mail cannot be tagged before it is filed: {facts}")


def derive_filing_state(facts: FilingFacts) -> FilingState:
    """The one place that maps what a run observed to a filing state.

    Raises:
        ValueError: ``facts`` describe a walk that cannot happen.
    """
    _check_possible(facts)
    if facts.refused:
        return FilingState.REFUSED
    if facts.lookup is Lookup.RAISED:
        return FilingState.LOOKUP_FAILED
    if facts.lookup is Lookup.MISSING:
        return FilingState.NOT_FOUND
    if facts.archive_failed:
        return FilingState.ARCHIVE_FAILED
    if facts.moved:
        return FilingState.FILED if facts.categorized else FilingState.MOVED_UNTAGGED
    if facts.kept_in_sent:
        return (
            FilingState.FILED_KEPT_IN_SENT if facts.categorized
            else FilingState.KEPT_IN_SENT_UNTAGGED
        )
    return FilingState.MOVE_FAILED


def settle(result: dict[str, Any], facts: FilingFacts) -> FilingState:
    """Set ``ok``, ``moved`` and ``categorized`` in ``result`` from ``facts``.

    The only writer of those three keys; the rest of the result (error, files,
    ``move_via``, ``entry_id``) stays with the step that learned it.
    """
    state = derive_filing_state(facts)
    result["ok"], result["moved"], result["categorized"] = _RESULT_FLAGS[state]
    return state

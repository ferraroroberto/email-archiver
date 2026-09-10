"""
Renumber an archive folder so its sequence numbers read as the mails' order.

The ``NNN`` in ``NNN - name.ext`` is meant to be browsable: opening a project
folder in Explorer and reading down the list should be reading the thread in
the order it happened. Two things break that, and neither is a bug in how a
single mail is archived:

- ``batch.revert`` deletes a bundle and leaves a hole in the sequence;
- ``batch.apply`` always allocates ``max + 1``, so a mail filed into a folder
  after the fact takes the highest number even when it is older than what is
  already there.

``renumber_folder`` is the repair: order the folder's numbered bundles by sent
date, hand out contiguous numbers from the folder's lowest existing one, and
return the old → new map so a consumer that stored ``.msg`` paths can heal its
own references.

Design decisions:

- **The unit is a bundle, not a file.** A number and every file carrying it
  move together, so an email never parts company with its attachments. A
  number that carries no ``.msg`` at all (an attachment whose mail was deleted
  by hand, a document filed into the sequence) is still a bundle: it holds its
  slot and its number moves with everything else, because leaving it behind
  would park it on a number that now belongs to a different mail.
- **An unknown date keeps its position.** A bundle with no readable date — no
  ``.msg``, or one whose date will not parse — inherits the date of the bundle
  before it and is ordered stably after it. It stays where the owner put it
  instead of being swept to one end of the folder.
- **The base is the folder's lowest existing number, not 001.** A "part 2"
  folder that starts at 079 keeps starting at 079: its numbers continue another
  folder's sequence, and renumbering it from 001 would destroy that.
- **Two mails on one number are split by date, and their attachments are
  matched to the right one.** Which files on a shared number belong to which
  mail cannot be read off the filenames, so each candidate ``.msg`` is asked
  what attachments it carries and the on-disk names are matched against that.
  An attachment that still cannot be placed follows the earlier mail, and the
  result says so — a guess that is reported is recoverable, a silent one is not.
- **Renames go through a temporary name first.** Closing a gap shifts a run of
  bundles down by one, so a file's target name is very often the name of the
  file next to it. Every changing file is moved to a same-length ``~XX``
  placeholder first and to its final name second, so no rename ever lands on a
  name that is still occupied. Same length on purpose: these paths sit near
  Windows' MAX_PATH and a longer temporary name could fail to be created at all.
- **Only files that actually change are touched.** A OneDrive-backed folder is
  listed once and a bundle whose number does not move is not renamed, not
  re-read, and not re-indexed.
"""
from __future__ import annotations

import logging
import os
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from typing import Any

from email_archiver.archiver.archiver import (
    SequencePrefix,
    sanitize_filename,
    split_sequence_prefix,
)
from email_archiver.database.repository import EmailRepository
from email_archiver.paths import resolve_under_roots
from email_archiver.scanner.scanner import read_msg_facts

logger = logging.getLogger(__name__)

_MSG_SUFFIX = ".msg"

# Why a whole folder was refused. The path reasons come from `paths`; this one
# is renumber's own.
REFUSED_UNREADABLE = "unreadable_folder"

# How an attachment on a shared number was placed, reported per split bundle.
PLACED_BY_MSG = "matched_to_msg"        # the .msg lists it as its attachment
PLACED_BY_FALLBACK = "assigned_to_first"  # nothing matched; earliest mail took it

# Placeholder numbers for the first rename phase. Exactly three characters, so
# a path that fits today still fits mid-rename, and starting with `~` so a
# placeholder can never be mistaken for a real sequence number.
_TEMP_ALPHABET = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
_MAX_TEMP = len(_TEMP_ALPHABET) ** 2

# The sort key an unknown date gets when nothing before it has one either:
# it stays at the front of the folder, which is where it already is.
_EPOCH = datetime.min.replace(tzinfo=timezone.utc)


class RenumberError(RuntimeError):
    """A folder that could not be renumbered safely, mid-flight."""


# ----------------------------------------------------------------- types ----

@dataclass
class RenumberResult:
    """What one folder's renumber did — or would do, under ``dry_run``."""

    folder_path: str
    dry_run: bool = False
    # Set when nothing was attempted at all: the folder is outside the archive
    # roots, unresolvable, or unreadable. Never folded into "renamed nothing".
    refused: str | None = None
    # The old → new map, one entry per bundle whose number changed.
    renamed: list[dict[str, Any]] = field(default_factory=list)
    # Filenames with no sequence prefix — left exactly as they are.
    skipped: list[str] = field(default_factory=list)
    bundles: int = 0
    unchanged: int = 0
    files_renamed: int = 0
    index_rows_updated: int = 0
    # Rows that described a file no longer at that name and whose name a rename
    # has now given to a different mail. Counted separately from the updates:
    # "the index followed the renames" and "the index was also wrong" are two
    # facts, and folding the second into the first hides it.
    index_rows_dropped: int = 0
    first_number: str = ""
    last_number: str = ""

    def counts(self) -> dict[str, int]:
        return {
            "bundles": self.bundles,
            "renamed": len(self.renamed),
            "unchanged": self.unchanged,
            "files_renamed": self.files_renamed,
            "skipped": len(self.skipped),
            "index_rows_updated": self.index_rows_updated,
            "index_rows_dropped": self.index_rows_dropped,
        }


@dataclass
class _Bundle:
    """One sequence number's worth of files, with the mail that anchors it."""

    number: str
    position: int
    msg: str | None = None                       # filename, not a path
    attachments: list[str] = field(default_factory=list)
    prefixes: dict[str, SequencePrefix] = field(default_factory=dict)
    message_id: str = ""
    sent: datetime | None = None
    sort_key: datetime = _EPOCH
    placement: str = ""                          # only set on a split bundle

    def files(self) -> list[str]:
        return ([self.msg] if self.msg else []) + self.attachments


@dataclass
class _Move:
    """One file's two-phase rename."""

    old: str
    temp: str
    new: str
    is_msg: bool


# --------------------------------------------------------------- reading ----

def _parse_folder(
    folder: str,
) -> tuple[dict[str, dict[str, list[str]]], dict[str, SequencePrefix], list[str]]:
    """One listing pass: numbered files grouped by number, plus the rest.

    Returns ``(groups, prefixes, skipped)`` where ``groups`` maps a sequence
    number to its ``msgs``/``others`` filenames, ``prefixes`` maps every
    numbered filename to its parsed prefix, and ``skipped`` holds the filenames
    that carry no sequence number.
    """
    groups: dict[str, dict[str, list[str]]] = {}
    prefixes: dict[str, SequencePrefix] = {}
    skipped: list[str] = []

    for name in sorted(os.listdir(folder)):
        if not os.path.isfile(os.path.join(folder, name)):
            continue
        parsed = split_sequence_prefix(name)
        if parsed is None:
            skipped.append(name)
            continue
        prefixes[name] = parsed
        bucket = groups.setdefault(parsed.number, {"msgs": [], "others": []})
        key = "msgs" if name.lower().endswith(_MSG_SUFFIX) else "others"
        bucket[key].append(name)

    return groups, prefixes, skipped


def _parse_sent(raw: str) -> datetime | None:
    """An indexed ``date_sent`` as a comparable, timezone-aware datetime.

    Naive values are read as local time — that is what the scanner wrote, and
    comparing a naive value against an aware one raises rather than mis-sorts,
    so this is the one place the two can meet.
    """
    if not raw:
        return None
    try:
        value = datetime.fromisoformat(raw)
    except ValueError:
        logger.debug("Unparseable date_sent %r; treating it as unknown", raw)
        return None
    return value if value.tzinfo is not None else value.astimezone()


@dataclass
class _MsgFacts:
    sent: datetime | None
    message_id: str
    attachment_names: tuple[str, ...]


def _facts_for(
    folder: str,
    name: str,
    indexed: dict[str, Any],
    *,
    need_attachments: bool = False,
) -> _MsgFacts:
    """What is known about one ``.msg``: its sent date, id and attachments.

    The index answers first — it is already read and costs nothing. The file
    itself is opened only when the index cannot answer: no row (archived since
    the last scan), no usable date, or an attachment list is needed to place a
    shared number's files.
    """
    record = indexed.get(name.lower())
    sent = _parse_sent(record.date_sent) if record is not None else None
    message_id = (record.message_id or "") if record is not None else ""
    attachment_names: tuple[str, ...] = ()

    if sent is None or need_attachments:
        facts = read_msg_facts(os.path.join(folder, name))
        if facts is not None:
            sent = sent or _parse_sent(facts.date_sent)
            message_id = message_id or facts.message_id
            attachment_names = facts.attachment_names

    if sent is None:
        logger.info("No sent date for %r; it keeps its position in the folder", name)
    return _MsgFacts(sent, message_id, attachment_names)


# -------------------------------------------------------------- bundling ----

def _match_key(stem: str) -> str:
    """A filename stem reduced to what two spellings of it have in common.

    The archiver sanitises an attachment's name, lowercases its suffix, cuts
    the stem to fit MAX_PATH (leaving a trailing ``...``) and appends ``_2`` on
    a collision, so the name on disk is rarely the name the mail carries. Both
    sides go through this before they are compared.
    """
    return sanitize_filename(stem, max_len=60).casefold().rstrip("._ ")


# Below this, a shared prefix says nothing: "re" and "report.pdf" are not the
# same attachment, and one of the two names is usually a MAX_PATH truncation.
_MIN_PREFIX_MATCH = 4


def _claims(stem: str, keys: set[str]) -> bool:
    """Whether a mail carrying ``keys`` plausibly owns the file stem ``stem``."""
    if not stem:
        return False
    if stem in keys:
        return True
    return any(
        len(key) >= _MIN_PREFIX_MATCH
        and len(stem) >= _MIN_PREFIX_MATCH
        and (stem.startswith(key) or key.startswith(stem))
        for key in keys
    )


def _place_attachments(
    others: list[str],
    prefixes: dict[str, SequencePrefix],
    candidates: list[tuple[str, _MsgFacts]],
) -> tuple[dict[str, list[str]], str]:
    """Split one number's non-``.msg`` files between the mails sharing it.

    Each candidate mail is asked which attachments it carries; a file on disk
    goes to the mail that names it. Anything still unplaced goes to the first
    (earliest) mail, and the caller is told that is what happened — the
    alternative, leaving it on a number that is about to belong to someone
    else, is worse and silent.
    """
    owned: dict[str, list[str]] = {name: [] for name, _ in candidates}
    known = {
        name: {_match_key(Path(att).stem) for att in facts.attachment_names}
        for name, facts in candidates
    }
    placement = PLACED_BY_MSG

    for other in others:
        stem = _match_key(Path(prefixes[other].rest).stem)
        owner = next(
            (name for name, _ in candidates if _claims(stem, known[name])), None
        )
        if owner is None:
            owner = candidates[0][0]
            placement = PLACED_BY_FALLBACK
            logger.info(
                "No mail on this number claims %r; it follows the earlier mail",
                other,
            )
        owned[owner].append(other)

    return owned, placement


def _build_bundles(
    folder: str,
    groups: dict[str, dict[str, list[str]]],
    prefixes: dict[str, SequencePrefix],
    indexed: dict[str, Any],
) -> list[_Bundle]:
    """Every numbered bundle in the folder, in its current on-disk order."""
    bundles: list[_Bundle] = []

    for number in sorted(groups, key=int):
        msgs = groups[number]["msgs"]
        others = groups[number]["others"]

        if not msgs:
            # No mail anchors this number. It still holds a slot: its files
            # would otherwise end up sharing a number with a different mail.
            bundles.append(_Bundle(
                number=number,
                position=len(bundles),
                attachments=list(others),
                prefixes={n: prefixes[n] for n in others},
            ))
            continue

        facts = {
            name: _facts_for(folder, name, indexed, need_attachments=len(msgs) > 1)
            for name in msgs
        }
        ordered = sorted(msgs, key=lambda n: (facts[n].sent or _EPOCH, n))

        if len(ordered) == 1:
            owned = {ordered[0]: list(others)}
            placement = ""
        else:
            logger.info(
                "Sequence number %s carries %d mails; splitting them by date",
                number, len(ordered),
            )
            owned, placement = _place_attachments(
                others, prefixes, [(n, facts[n]) for n in ordered]
            )

        for name in ordered:
            files = [name, *owned[name]]
            bundles.append(_Bundle(
                number=number,
                position=len(bundles),
                msg=name,
                attachments=list(owned[name]),
                prefixes={n: prefixes[n] for n in files},
                message_id=facts[name].message_id,
                sent=facts[name].sent,
                placement=placement if len(ordered) > 1 else "",
            ))

    return bundles


def _order(bundles: list[_Bundle]) -> list[_Bundle]:
    """Chronological order, with an undated bundle keeping its place.

    An undated bundle inherits the date of the last dated bundle before it, so
    a stable sort leaves it sitting exactly where it sits now — right after
    that neighbour — rather than at one end of the folder.
    """
    carried = _EPOCH
    for bundle in bundles:
        if bundle.sent is not None:
            carried = bundle.sent
        bundle.sort_key = carried
    return sorted(bundles, key=lambda b: (b.sort_key, b.position))


# --------------------------------------------------------------- renaming ---

def _temp_number(index: int) -> str:
    """A three-character placeholder number that no real file can carry."""
    if index >= _MAX_TEMP:
        raise RenumberError(
            f"more than {_MAX_TEMP} files to rename in one folder; "
            "renumber it in smaller pieces"
        )
    high, low = divmod(index, len(_TEMP_ALPHABET))
    return f"~{_TEMP_ALPHABET[high]}{_TEMP_ALPHABET[low]}"


def _plan(
    folder: str, ordered: list[_Bundle], base: int
) -> tuple[list[_Move], list[dict[str, Any]], int]:
    """The renames to run and the map to report, for one ordered folder."""
    moves: list[_Move] = []
    mapped: list[dict[str, Any]] = []
    unchanged = 0

    for offset, bundle in enumerate(ordered):
        number = f"{base + offset:03d}"
        if number == bundle.number:
            unchanged += 1
            continue

        entry: dict[str, Any] = {
            "from": None,
            "to": None,
            "message_id": bundle.message_id,
            "attachments": [],
        }
        for name in bundle.files():
            new_name = bundle.prefixes[name].with_number(number)
            temp_name = bundle.prefixes[name].with_number(_temp_number(len(moves)))
            is_msg = name == bundle.msg
            moves.append(_Move(old=name, temp=temp_name, new=new_name, is_msg=is_msg))
            old_path = os.path.join(folder, name)
            new_path = os.path.join(folder, new_name)
            if is_msg:
                entry["from"], entry["to"] = old_path, new_path
            else:
                entry["attachments"].append([old_path, new_path])
        if bundle.placement:
            entry["attachments_placed"] = bundle.placement
        mapped.append(entry)

    return moves, mapped, unchanged


def _rename(folder: str, moves: list[_Move]) -> None:
    """Run the two phases, refusing to start if a placeholder is taken."""
    for move in moves:
        temp_path = os.path.join(folder, move.temp)
        if os.path.exists(temp_path):
            raise RenumberError(
                f"the placeholder name {move.temp!r} is already taken in this "
                "folder; nothing was renamed"
            )

    for move in moves:
        os.rename(os.path.join(folder, move.old), os.path.join(folder, move.temp))
    try:
        for move in moves:
            os.rename(os.path.join(folder, move.temp), os.path.join(folder, move.new))
    except OSError:
        left = [m.temp for m in moves if os.path.exists(os.path.join(folder, m.temp))]
        logger.error(
            "Renumber stopped part-way; %d file(s) are still under a placeholder "
            "name in this folder: %s", len(left), left,
        )
        raise


def _sibling(stored_path: str, name: str) -> str:
    """``name`` in the folder ``stored_path`` is in — the index's spelling of it,
    not this caller's spelling of the same folder."""
    return os.path.join(os.path.dirname(stored_path), name)


def _reindex(
    moves: list[_Move], indexed: dict[str, Any], repo: EmailRepository
) -> tuple[int, int]:
    """Follow the renames in the index. Returns ``(rows updated, rows dropped)``.

    Two phases, for the same reason the files themselves go through a
    placeholder: a row's target path is very often another row's current path,
    and ``emails.file_path`` is UNIQUE — updating them one by one in bundle
    order hits that constraint the moment a run shifts down by one. (Observed
    the first time this ran against a real folder: the files were renamed and
    the whole index update rolled back.)

    A row still sitting on a name a rename is about to take, which is not
    itself one of the rows being moved, cannot be describing a file that is
    still there — nothing on disk was at that name (a target never collides
    with a file that stays put), so the row is stale and is dropped rather than
    left to collide or to point at somebody else's mail.
    """
    msg_moves = [move for move in moves if move.is_msg]
    sources = {move.old.lower() for move in msg_moves}
    pairs = [
        (indexed[move.old.lower()], move)
        for move in msg_moves
        if move.old.lower() in indexed
    ]

    dropped = 0
    for move in msg_moves:
        stale = indexed.get(move.new.lower())
        if stale is not None and move.new.lower() not in sources:
            logger.info(
                "Dropping the index row for %r: that name now belongs to a "
                "different mail and the file it described is gone", move.new,
            )
            dropped += repo.delete_by_path(stale.file_path)

    for record, move in pairs:
        repo.update_path(record.file_path, _sibling(record.file_path, move.temp))
    updated = 0
    for record, move in pairs:
        updated += repo.update_path(
            _sibling(record.file_path, move.temp),
            _sibling(record.file_path, move.new),
        )

    if updated or dropped:
        repo.commit()
    return updated, dropped


# ------------------------------------------------------------------ entry ---

def renumber_folder(
    folder_path: str,
    repo: EmailRepository,
    roots: list[Path],
    *,
    dry_run: bool = False,
) -> RenumberResult:
    """Give ``folder_path``'s numbered bundles contiguous numbers in date order.

    ``roots`` is the configured ``archive.root_paths``, resolved — a folder that
    does not sit inside one is refused untouched, so a caller cannot rename its
    way across the rest of the disk. It is a positional argument rather than an
    option because a renumber with no root check is never the right call.

    ``dry_run`` computes the whole plan and reports the same map without
    renaming anything or touching the index.
    """
    resolved, refusal = resolve_under_roots(folder_path, roots)
    normalised = os.path.normpath(folder_path)
    if resolved is None:
        logger.warning("Refusing to renumber %r: %s", normalised, refusal)
        return RenumberResult(
            folder_path=normalised, dry_run=dry_run, refused=refusal
        )

    try:
        groups, prefixes, skipped = _parse_folder(normalised)
    except OSError as exc:
        logger.error("Cannot list %r: %s", normalised, exc)
        return RenumberResult(
            folder_path=normalised, dry_run=dry_run, refused=REFUSED_UNREADABLE
        )

    result = RenumberResult(
        folder_path=normalised, dry_run=dry_run, skipped=skipped
    )
    if not groups:
        logger.info("Nothing numbered in %r; leaving it alone", normalised)
        return result

    indexed = {r.filename.lower(): r for r in repo.find_in_folder(normalised)}
    ordered = _order(_build_bundles(normalised, groups, prefixes, indexed))

    base = min(int(number) for number in groups)
    last = base + len(ordered) - 1
    if last > 999:
        raise RenumberError(
            f"{len(ordered)} bundles from {base:03d} would need a four-digit "
            "sequence number; this folder cannot be renumbered in place"
        )

    moves, mapped, unchanged = _plan(normalised, ordered, base)
    result.renamed = mapped
    result.bundles = len(ordered)
    result.unchanged = unchanged
    result.files_renamed = len(moves)
    result.first_number = f"{base:03d}"
    result.last_number = f"{last:03d}"

    if dry_run:
        logger.info(
            "Renumber (dry run) of %r: %d bundle(s), %d would change, "
            "%d file(s) would be renamed.",
            normalised, result.bundles, len(mapped), len(moves),
        )
        return result

    if moves:
        _rename(normalised, moves)
        result.index_rows_updated, result.index_rows_dropped = _reindex(
            moves, indexed, repo
        )

    logger.info(
        "Renumbered %r: %d bundle(s) now %s–%s, %d changed, %d file(s) renamed, "
        "%d index row(s) updated, %d stale row(s) dropped.",
        normalised, result.bundles, result.first_number, result.last_number,
        len(mapped), len(moves), result.index_rows_updated,
        result.index_rows_dropped,
    )
    return result

"""Tests for renumbering a folder into date order (issue #61).

The sequence number in an archive folder is meant to read as the chronological
order of the mails in it, and two normal batch operations break that: a
``revert`` leaves a hole, and an ``apply`` always takes ``max + 1`` even for a
mail older than everything already filed. ``renumber_folder`` is the repair, and
what it returns is an interface — another local app stores ``.msg`` paths and
heals them from that map — so these tests pin the map's shape as much as the
renames on disk.

Everything on screen here is synthetic: no real folder name, address or subject.
"""
from __future__ import annotations

import os
from pathlib import Path

import pytest

from email_archiver import renumber as renumber_module
from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRecord, EmailRepository
from email_archiver.paths import REFUSED_OUTSIDE_ROOTS
from email_archiver.renumber import renumber_folder
from email_archiver.scanner.scanner import MsgFacts


# -------------------------------------------------------------- fixtures ----

@pytest.fixture
def archive_root(tmp_path) -> Path:
    root = tmp_path / "archive"
    root.mkdir()
    return root


@pytest.fixture
def roots(archive_root) -> list[Path]:
    return [archive_root.resolve()]


@pytest.fixture
def folder(archive_root) -> Path:
    target = archive_root / "Project Alpha"
    target.mkdir()
    return target


@pytest.fixture
def repo(tmp_path):
    conn = init_db(tmp_path / "emails.db")
    try:
        yield EmailRepository(conn)
    finally:
        conn.close()


def _write(folder: Path, name: str) -> Path:
    path = folder / name
    path.write_text(f"synthetic {name}", encoding="utf-8")
    return path


def _index(
    repo: EmailRepository,
    folder: Path,
    name: str,
    date_sent: str,
    message_id: str = "",
) -> None:
    """Put one row in the index, the way a scan would have."""
    repo.upsert_email(EmailRecord(
        file_path=str(folder / name),
        folder_path=str(folder),
        filename=name,
        subject=Path(name).stem,
        date_sent=date_sent,
        file_mtime=1.0,
        message_id=(
            message_id
            or f"{Path(name).stem}@example.invalid".replace(" ", "-")
        ),
    ))
    repo.commit()


def _seed(repo: EmailRepository, folder: Path, mails: dict[str, str]) -> None:
    """Write and index a whole folder: ``{filename: date_sent}``."""
    for name, date_sent in mails.items():
        _write(folder, name)
        if date_sent:
            _index(repo, folder, name, date_sent)


def _names(folder: Path) -> list[str]:
    return sorted(p.name for p in folder.iterdir())


# ------------------------------------------------------------ the basics ----

def test_a_gap_left_by_a_revert_closes(repo, folder, roots):
    _seed(repo, folder, {
        "001 - one.msg": "2026-01-01T09:00:00+01:00",
        "003 - three.msg": "2026-01-03T09:00:00+01:00",
        "004 - four.msg": "2026-01-04T09:00:00+01:00",
    })

    result = renumber_folder(str(folder), repo, roots)

    assert result.refused is None
    assert _names(folder) == ["001 - one.msg", "002 - three.msg", "003 - four.msg"]
    assert result.first_number == "001"
    assert result.last_number == "003"
    assert result.bundles == 3
    assert result.unchanged == 1


def test_an_older_mail_filed_last_moves_to_its_date_position(repo, folder, roots):
    """The apply half of the problem: a mail archived after a correction took
    ``max + 1`` even though it is older than what was already there."""
    _seed(repo, folder, {
        "001 - first.msg": "2026-01-01T09:00:00+01:00",
        "002 - third.msg": "2026-01-03T09:00:00+01:00",
        "003 - latecomer.msg": "2026-01-02T09:00:00+01:00",
    })

    renumber_folder(str(folder), repo, roots)

    assert _names(folder) == [
        "001 - first.msg", "002 - latecomer.msg", "003 - third.msg",
    ]


def test_attachments_follow_their_mail(repo, folder, roots):
    _seed(repo, folder, {
        "001 - one.msg": "2026-01-01T09:00:00+01:00",
        "003 - three.msg": "2026-01-03T09:00:00+01:00",
    })
    _write(folder, "003 - report.pdf")
    _write(folder, "003 - notes.docx")

    result = renumber_folder(str(folder), repo, roots)

    assert _names(folder) == [
        "001 - one.msg", "002 - notes.docx", "002 - report.pdf", "002 - three.msg",
    ]
    entry = result.renamed[0]
    assert Path(entry["from"]).name == "003 - three.msg"
    assert Path(entry["to"]).name == "002 - three.msg"
    assert sorted(Path(old).name for old, _ in entry["attachments"]) == [
        "003 - notes.docx", "003 - report.pdf",
    ]
    assert sorted(Path(new).name for _, new in entry["attachments"]) == [
        "002 - notes.docx", "002 - report.pdf",
    ]


def test_the_map_reports_only_what_changed(repo, folder, roots):
    _seed(repo, folder, {
        "001 - one.msg": "2026-01-01T09:00:00+01:00",
        "003 - three.msg": "2026-01-03T09:00:00+01:00",
    })

    result = renumber_folder(str(folder), repo, roots)

    assert [Path(e["from"]).name for e in result.renamed] == ["003 - three.msg"]
    assert set(result.renamed[0]) == {"from", "to", "message_id", "attachments"}
    assert result.renamed[0]["message_id"] == "003---three@example.invalid"


def test_a_folder_already_in_order_is_not_touched(repo, folder, roots):
    _seed(repo, folder, {
        "001 - one.msg": "2026-01-01T09:00:00+01:00",
        "002 - two.msg": "2026-01-02T09:00:00+01:00",
    })
    before = {p.name: p.stat().st_mtime_ns for p in folder.iterdir()}

    result = renumber_folder(str(folder), repo, roots)

    assert result.renamed == []
    assert result.files_renamed == 0
    assert {p.name: p.stat().st_mtime_ns for p in folder.iterdir()} == before


# --------------------------------------------------------------- the base ---

def test_the_base_is_the_folders_lowest_existing_number(repo, folder, roots):
    """A "part 2" folder continues another folder's sequence: renumbering it
    from 001 would destroy exactly the fact its numbers carry."""
    _seed(repo, folder, {
        "079 - one.msg": "2026-01-01T09:00:00+01:00",
        "081 - two.msg": "2026-01-02T09:00:00+01:00",
    })

    result = renumber_folder(str(folder), repo, roots)

    assert _names(folder) == ["079 - one.msg", "080 - two.msg"]
    assert (result.first_number, result.last_number) == ("079", "080")


# ------------------------------------------------------- both name forms ----

def test_each_file_keeps_its_own_dated_or_undated_form(repo, folder, roots):
    _seed(repo, folder, {
        "2026-01-01 - 001 - dated.msg": "2026-01-01T09:00:00+01:00",
        "003 - undated.msg": "2026-01-03T09:00:00+01:00",
        "2026-01-05 - 005 - dated too.msg": "2026-01-05T09:00:00+01:00",
    })
    _write(folder, "2026-01-05 - 005 - attached.pdf")

    renumber_folder(str(folder), repo, roots)

    assert _names(folder) == [
        "002 - undated.msg",
        "2026-01-01 - 001 - dated.msg",
        "2026-01-05 - 003 - attached.pdf",
        "2026-01-05 - 003 - dated too.msg",
    ]


# ------------------------------------------------------ two on one number ---

def test_two_mails_sharing_one_number_end_up_distinct(repo, folder, roots):
    _seed(repo, folder, {
        "087 - earlier.msg": "2026-09-08T16:00:00+02:00",
        "087 - later.msg": "2026-09-09T05:00:00+02:00",
        "088 - after.msg": "2026-09-09T13:00:00+02:00",
    })

    result = renumber_folder(str(folder), repo, roots)

    assert _names(folder) == [
        "087 - earlier.msg", "088 - later.msg", "089 - after.msg",
    ]
    assert result.bundles == 3
    assert (result.first_number, result.last_number) == ("087", "089")


def test_a_shared_numbers_attachments_go_to_the_mail_that_carries_them(
    repo, folder, roots, monkeypatch
):
    """Filenames cannot say whose attachment is whose, so each candidate mail
    is asked what it carries. Getting this wrong files somebody's diagrams
    under an unrelated mail — silently, and permanently."""
    _seed(repo, folder, {
        "087 - earlier.msg": "2026-09-08T16:00:00+02:00",
        "087 - later.msg": "2026-09-09T05:00:00+02:00",
    })
    _write(folder, "087 - diagram one.png")

    facts = {
        "087 - earlier.msg": MsgFacts("", "", ("image001.png",)),
        "087 - later.msg": MsgFacts("", "", ("diagram one.png",)),
    }
    monkeypatch.setattr(
        renumber_module, "read_msg_facts",
        lambda path: facts.get(os.path.basename(path)),
    )

    result = renumber_folder(str(folder), repo, roots)

    assert _names(folder) == [
        "087 - earlier.msg", "088 - diagram one.png", "088 - later.msg",
    ]
    entry = result.renamed[0]
    assert Path(entry["to"]).name == "088 - later.msg"
    assert entry["attachments_placed"] == renumber_module.PLACED_BY_MSG


def test_an_unclaimed_attachment_follows_the_earlier_mail_and_says_so(
    repo, folder, roots, monkeypatch
):
    """A guess that is reported is recoverable; a silent one is not."""
    _seed(repo, folder, {
        "087 - earlier.msg": "2026-09-08T16:00:00+02:00",
        "087 - later.msg": "2026-09-09T05:00:00+02:00",
    })
    _write(folder, "087 - orphan.png")
    monkeypatch.setattr(
        renumber_module, "read_msg_facts", lambda path: MsgFacts("", "", ())
    )

    result = renumber_folder(str(folder), repo, roots)

    assert _names(folder) == [
        "087 - earlier.msg", "087 - orphan.png", "088 - later.msg",
    ]
    # The earlier mail did not move, so only the split-off mail is in the map;
    # the fallback is still reported, on the bundle that was split.
    assert result.renamed[0]["attachments_placed"] == (
        renumber_module.PLACED_BY_FALLBACK
    )


# ------------------------------------------------- unnumbered and orphans ---

def test_a_file_with_no_sequence_prefix_is_left_alone_and_reported(
    repo, folder, roots
):
    _seed(repo, folder, {
        "001 - one.msg": "2026-01-01T09:00:00+01:00",
        "003 - three.msg": "2026-01-03T09:00:00+01:00",
    })
    _write(folder, "052a - hand named.msg")
    _write(folder, "notes.txt")

    result = renumber_folder(str(folder), repo, roots)

    assert sorted(result.skipped) == ["052a - hand named.msg", "notes.txt"]
    assert (folder / "052a - hand named.msg").exists()
    assert (folder / "notes.txt").exists()


def test_a_number_with_no_mail_holds_its_slot(repo, folder, roots):
    """An attachment whose mail was deleted by hand still owns its number. Left
    behind, it would end up sharing a number with an unrelated mail."""
    _seed(repo, folder, {
        "001 - one.msg": "2026-01-01T09:00:00+01:00",
        "004 - four.msg": "2026-01-04T09:00:00+01:00",
    })
    _write(folder, "003 - stranded.docx")

    result = renumber_folder(str(folder), repo, roots)

    assert _names(folder) == [
        "001 - one.msg", "002 - stranded.docx", "003 - four.msg",
    ]
    assert result.bundles == 3
    # The stranded bundle's number does move here (003 -> 002), so its entry
    # in the map is reached — and being mail-less, it carries no .msg path, a
    # consumer healing stored .msg paths has nothing to heal for it.
    stranded_entries = [e for e in result.renamed if e["from"] is None]
    assert len(stranded_entries) == 1


def test_an_undated_bundle_keeps_its_position_rather_than_moving_to_one_end(
    repo, folder, roots
):
    _seed(repo, folder, {
        "001 - one.msg": "2026-01-01T09:00:00+01:00",
        "003 - three.msg": "2026-01-03T09:00:00+01:00",
    })
    # Indexed with no date and unreadable as a .msg: genuinely unknown.
    _write(folder, "002 - undatable.msg")

    renumber_folder(str(folder), repo, roots)

    assert _names(folder) == [
        "001 - one.msg", "002 - undatable.msg", "003 - three.msg",
    ]


# -------------------------------------------------------------- dry run -----

def test_dry_run_changes_nothing_and_returns_the_same_map(repo, folder, roots):
    _seed(repo, folder, {
        "001 - one.msg": "2026-01-01T09:00:00+01:00",
        "003 - three.msg": "2026-01-03T09:00:00+01:00",
    })
    _write(folder, "003 - report.pdf")
    before = _names(folder)

    planned = renumber_folder(str(folder), repo, roots, dry_run=True)

    assert planned.dry_run is True
    assert _names(folder) == before
    assert planned.index_rows_updated == 0
    assert repo.find_in_folder(str(folder))[0].filename == "001 - one.msg"

    done = renumber_folder(str(folder), repo, roots)

    assert done.renamed == planned.renamed
    assert (done.first_number, done.last_number) == (
        planned.first_number, planned.last_number
    )
    assert _names(folder) != before


# --------------------------------------------------------------- the index --

def test_the_index_rows_follow_the_renames(repo, folder, roots):
    _seed(repo, folder, {
        "001 - one.msg": "2026-01-01T09:00:00+01:00",
        "003 - three.msg": "2026-01-03T09:00:00+01:00",
        "004 - four.msg": "2026-01-04T09:00:00+01:00",
    })

    result = renumber_folder(str(folder), repo, roots)

    assert result.index_rows_updated == 2
    rows = {r.filename: r.file_path for r in repo.find_in_folder(str(folder))}
    assert set(rows) == {"001 - one.msg", "002 - three.msg", "003 - four.msg"}
    assert all(Path(path).exists() for path in rows.values())
    # The row is updated, not replaced: its Message-ID is still the mail's.
    assert repo.find_path_by_message_id("003---three@example.invalid") == str(
        folder / "002 - three.msg"
    )


def test_two_mails_swapping_names_do_not_break_the_unique_index(repo, folder, roots):
    """Regression: ``emails.file_path`` is UNIQUE, and two mails on one thread
    carry the same subject, so a reorder hands one of them the exact path the
    other still holds. Updating the rows one at a time raised ``IntegrityError``
    and rolled the whole index update back — *after* the files on disk had been
    renamed, leaving the index describing a folder that no longer existed.
    Found the first time this ran against a real folder.
    """
    _write(folder, "001 - thread.msg")
    _write(folder, "002 - thread.msg")
    _index(repo, folder, "001 - thread.msg", "2026-01-02T09:00:00+01:00",
           message_id="later@example.invalid")
    _index(repo, folder, "002 - thread.msg", "2026-01-01T09:00:00+01:00",
           message_id="earlier@example.invalid")

    result = renumber_folder(str(folder), repo, roots)

    # The names are the same two; which mail is behind each one swapped.
    assert _names(folder) == ["001 - thread.msg", "002 - thread.msg"]
    assert result.index_rows_updated == 2
    assert (folder / "001 - thread.msg").read_text(encoding="utf-8") == (
        "synthetic 002 - thread.msg"
    )
    assert repo.find_path_by_message_id("earlier@example.invalid") == str(
        folder / "001 - thread.msg"
    )
    assert repo.find_path_by_message_id("later@example.invalid") == str(
        folder / "002 - thread.msg"
    )


def test_a_stale_row_on_a_name_a_rename_takes_is_dropped_not_left_lying(
    repo, folder, roots
):
    """A row whose file is gone keeps pointing at a name another mail is about
    to be given. Left alone it breaks the UNIQUE constraint, and were it to
    survive it would describe somebody else's mail."""
    _seed(repo, folder, {
        "001 - one.msg": "2026-01-01T09:00:00+01:00",
        "003 - thread.msg": "2026-01-03T09:00:00+01:00",
    })
    # Indexed, but its file is long gone — and 003 is headed for exactly its name.
    _index(repo, folder, "002 - thread.msg", "2026-01-02T09:00:00+01:00",
           message_id="gone@example.invalid")

    result = renumber_folder(str(folder), repo, roots)

    assert result.index_rows_dropped == 1
    assert result.index_rows_updated == 1
    rows = {r.filename for r in repo.find_in_folder(str(folder))}
    assert rows == {"001 - one.msg", "002 - thread.msg"} == set(_names(folder))
    assert repo.find_path_by_message_id("gone@example.invalid") is None


def test_a_renamed_mail_that_was_never_indexed_is_not_an_error(repo, folder, roots):
    _seed(repo, folder, {"001 - one.msg": "2026-01-01T09:00:00+01:00"})
    _write(folder, "003 - never scanned.msg")

    result = renumber_folder(str(folder), repo, roots)

    assert _names(folder) == ["001 - one.msg", "002 - never scanned.msg"]
    assert result.index_rows_updated == 0


# --------------------------------------------------------------- refusals ---

def test_a_folder_outside_the_archive_roots_is_refused_untouched(
    repo, tmp_path, roots
):
    outsider = tmp_path / "not-the-archive"
    outsider.mkdir()
    (outsider / "003 - precious.msg").write_text("do not touch me", encoding="utf-8")

    result = renumber_folder(str(outsider), repo, roots)

    assert result.refused == REFUSED_OUTSIDE_ROOTS
    assert result.renamed == []
    assert _names(outsider) == ["003 - precious.msg"]


def test_a_traversal_out_of_an_archive_root_is_refused(repo, archive_root, roots,
                                                       tmp_path):
    outsider = tmp_path / "not-the-archive"
    outsider.mkdir()
    sneaky = str(archive_root / ".." / "not-the-archive")

    assert renumber_folder(sneaky, repo, roots).refused == REFUSED_OUTSIDE_ROOTS


def test_an_unreadable_folder_is_refused_rather_than_reported_empty(repo, roots,
                                                                    archive_root):
    missing = archive_root / "never created"

    result = renumber_folder(str(missing), repo, roots)

    assert result.refused == renumber_module.REFUSED_UNREADABLE
    assert result.renamed == []


def test_a_folder_with_nothing_numbered_is_a_no_op_not_a_refusal(repo, folder,
                                                                 roots):
    _write(folder, "notes.txt")

    result = renumber_folder(str(folder), repo, roots)

    assert result.refused is None
    assert result.bundles == 0
    assert result.skipped == ["notes.txt"]


# ------------------------------------------------------ the rename itself ---

def test_a_whole_run_shifting_down_by_one_never_collides(repo, folder, roots):
    """Every target name is the name of the file next to it — which is why the
    renames go through a placeholder first."""
    _seed(repo, folder, {
        f"{n:03d} - mail.msg": f"2026-01-{n:02d}T09:00:00+01:00"
        for n in range(2, 12)
    })

    result = renumber_folder(str(folder), repo, roots)

    assert _names(folder) == [f"{n:03d} - mail.msg" for n in range(2, 12)]
    assert result.files_renamed == 0, "already contiguous from its own base"

    (folder / "002 - mail.msg").unlink()
    result = renumber_folder(str(folder), repo, roots)

    assert _names(folder) == [f"{n:03d} - mail.msg" for n in range(3, 12)]
    assert result.files_renamed == 0, "the base moved with the folder's lowest"


def test_closing_an_interior_gap_shifts_the_whole_tail_down(repo, folder, roots):
    _seed(repo, folder, {
        f"{n:03d} - mail {n}.msg": f"2026-01-{n:02d}T09:00:00+01:00"
        for n in [1, 2, 4, 5, 6, 7]
    })

    result = renumber_folder(str(folder), repo, roots)

    assert _names(folder) == [
        "001 - mail 1.msg", "002 - mail 2.msg", "003 - mail 4.msg",
        "004 - mail 5.msg", "005 - mail 6.msg", "006 - mail 7.msg",
    ]
    assert result.files_renamed == 4
    assert not [p for p in folder.iterdir() if p.name.startswith("~")]

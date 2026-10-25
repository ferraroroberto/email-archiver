"""``renumber``: the verb itself and the ``--renumber`` flag of ``apply`` and ``revert``.
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path


from email_archiver import batch, batch_common
from email_archiver import renumber as renumber_module
from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRecord, EmailRepository
from email_archiver.scanner.scanner import MsgFacts
from tests.batch.batch_support import (
    FakeOutlookClient,
    _index_archived_file,
    _mail,
)
from tests.batch.batch_support import (  # noqa: F401 - fixtures are used by name
    _no_retry_pauses,
    archive_root,
    cfg,
)


# --------------------------------------------------------------- renumber ---

def test_apply_without_the_flag_reports_no_renumber_keys_at_all(cfg, archive_root):
    """The flag is opt-in: without it the document is byte-for-byte what a
    consumer built against before this shipped."""
    dest = archive_root / "Project Alpha"
    client = FakeOutlookClient([_mail("a@example.invalid", "Untouched")])

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ])

    assert "renumbered" not in doc
    assert "renumber_refused" not in doc


def test_apply_with_renumber_puts_an_older_mail_in_its_date_position(
    cfg, archive_root, monkeypatch
):
    """The reported case: a mail filed into a folder after a correction takes
    `max + 1` even though it is older than what is already there.

    The mail `apply` just wrote has no index row yet — nothing has scanned it —
    so its date comes from the file. Outlook writes a real .msg there and
    extract-msg reads it; the fake client writes a placeholder, so the read is
    faked here too rather than silently degrading to "unknown date".
    """
    monkeypatch.setattr(
        renumber_module, "read_msg_facts",
        lambda path: MsgFacts("2026-03-14T09:00:00+01:00", "a@example.invalid", ()),
    )
    dest = archive_root / "Project Alpha"
    dest.mkdir(parents=True)
    newer = dest / "001 - Already filed.msg"
    newer.write_text("synthetic .msg", encoding="utf-8")

    conn = init_db(cfg["database"]["path"])
    EmailRepository(conn).upsert_email(EmailRecord(
        file_path=str(newer),
        folder_path=str(dest),
        filename=newer.name,
        subject="Already filed",
        date_sent="2026-03-20T09:00:00+01:00",
        file_mtime=1.0,
        message_id="already@example.invalid",
    ))
    conn.commit()
    conn.close()

    client = FakeOutlookClient([
        _mail("a@example.invalid", "Older one", sent=datetime(2026, 3, 14, 9, 0)),
    ])
    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
    ], renumber=True)

    assert doc["results"][0]["ok"] is True
    # It went in as 002 and came out as 001; the mail already there moved up.
    assert sorted(p.name for p in dest.iterdir()) == [
        "001 - Older one.msg", "002 - Already filed.msg",
    ]
    assert doc["renumber_refused"] == []
    moved = {
        Path(e["from"]).name: (Path(e["to"]).name, e["message_id"])
        for e in doc["renumbered"][str(dest)]
    }
    assert moved == {
        "001 - Already filed.msg": (
            "002 - Already filed.msg", "already@example.invalid",
        ),
        "002 - Older one.msg": ("001 - Older one.msg", "a@example.invalid"),
    }


def test_apply_with_renumber_lists_each_destination_folder_once(
    cfg, archive_root, monkeypatch
):
    """Two mails filed into the same folder must renumber it once for the
    whole run, not once per mail — a second pass would reorder what the
    first pass just fixed. ``doc["renumbered"]`` is keyed by folder path, so
    it can never show a folder twice however broken the dedupe in
    ``_unique_folders`` is; only counting the actual ``renumber_folder``
    calls pins the dedupe itself."""
    dest = archive_root / "Project Alpha"
    other = archive_root / "Project Beta"
    client = FakeOutlookClient([
        _mail("a@example.invalid", "One", sent=datetime(2026, 3, 14, 9, 0)),
        _mail("b@example.invalid", "Two", sent=datetime(2026, 3, 15, 9, 0)),
        _mail("c@example.invalid", "Three", sent=datetime(2026, 3, 16, 9, 0)),
    ])

    real_renumber_folder = batch_common.renumber_folder
    calls: list[str] = []

    def _counting_renumber_folder(folder_path, *args, **kwargs):
        calls.append(str(folder_path))
        return real_renumber_folder(folder_path, *args, **kwargs)

    monkeypatch.setattr(batch_common, "renumber_folder", _counting_renumber_folder)

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest)},
        {"message_id": "b@example.invalid", "folder_path": str(dest)},
        {"message_id": "c@example.invalid", "folder_path": str(other)},
    ], renumber=True)

    assert calls == [str(dest), str(other)]
    assert sorted(doc["renumbered"]) == sorted([str(dest), str(other)])


def test_apply_with_renumber_never_reaches_a_folder_outside_the_archive_roots(
    cfg, tmp_path
):
    """`apply` refuses a folder outside the roots before writing (issue #71),
    so the renumber that follows has nothing there to touch either."""
    outside = tmp_path / "not-the-archive"
    client = FakeOutlookClient([_mail("a@example.invalid", "Elsewhere")])

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(outside)},
    ], renumber=True)

    assert doc["results"][0]["error"]["code"] == batch.ERROR_BAD_DECISION
    assert not outside.exists()
    assert doc["renumbered"] == {}
    assert doc["renumber_refused"] == []


def test_revert_with_renumber_closes_the_gap_it_just_made(cfg, archive_root):
    dest = archive_root / "Project Alpha"
    client = FakeOutlookClient([
        _mail("a@example.invalid", "First", sent=datetime(2026, 3, 14, 9, 0)),
        _mail("b@example.invalid", "Second", sent=datetime(2026, 3, 15, 9, 0)),
        _mail("c@example.invalid", "Third", sent=datetime(2026, 3, 16, 9, 0)),
    ])
    applied = batch.apply(client, cfg, [
        {"message_id": mid, "folder_path": str(dest)}
        for mid in ("a@example.invalid", "b@example.invalid", "c@example.invalid")
    ])
    for result in applied["results"]:
        _index_archived_file(
            cfg, result, message_id=result["message_id"],
            subject=Path(result["files"][0]).stem,
        )
    assert sorted(p.name for p in dest.iterdir()) == [
        "001 - First.msg", "002 - Second.msg", "003 - Third.msg",
    ]

    doc = batch.revert(client, cfg, [
        {"message_id": "b@example.invalid", "files": applied["results"][1]["files"]},
    ], renumber=True)

    assert doc["results"][0]["ok"] is True
    assert sorted(p.name for p in dest.iterdir()) == [
        "001 - First.msg", "002 - Third.msg",
    ]
    entry = doc["renumbered"][str(dest)][0]
    assert Path(entry["from"]).name == "003 - Third.msg"
    assert Path(entry["to"]).name == "002 - Third.msg"
    assert doc["renumber_refused"] == []


def test_revert_with_renumber_ignores_a_folder_it_deleted_nothing_from(
    cfg, archive_root
):
    """No delete, no gap: a refused or already-missing file must not drag a
    whole folder through a renumber it did not need."""
    dest = archive_root / "Project Alpha"
    dest.mkdir(parents=True)
    (dest / "003 - Untouched.msg").write_text("synthetic .msg", encoding="utf-8")

    doc = batch.revert(FakeOutlookClient([]), cfg, [
        {"message_id": "a@example.invalid",
         "files": [str(dest / "009 - never existed.msg")]},
    ], renumber=True)

    assert doc["results"][0]["missing"]
    assert doc["renumbered"] == {}
    assert sorted(p.name for p in dest.iterdir()) == ["003 - Untouched.msg"]


def test_the_renumber_verb_reports_the_map_and_a_dry_run_changes_nothing(
    cfg, archive_root
):
    dest = archive_root / "Project Alpha"
    dest.mkdir(parents=True)
    for name in ("001 - one.msg", "003 - three.msg"):
        (dest / name).write_text("synthetic .msg", encoding="utf-8")

    conn = init_db(cfg["database"]["path"])
    repo = EmailRepository(conn)
    for name, date_sent in (
        ("001 - one.msg", "2026-01-01T09:00:00+01:00"),
        ("003 - three.msg", "2026-01-03T09:00:00+01:00"),
    ):
        repo.upsert_email(EmailRecord(
            file_path=str(dest / name), folder_path=str(dest), filename=name,
            subject=name, date_sent=date_sent, file_mtime=1.0,
            message_id=f"{name}@example.invalid",
        ))
    conn.commit()
    conn.close()

    planned = batch.renumber(cfg, str(dest), dry_run=True)
    assert planned["verb"] == "renumber"
    assert planned["schema_version"] == batch.SCHEMA_VERSION
    assert planned["dry_run"] is True
    assert planned["counts"]["bundles"] == 2
    assert planned["counts"]["renamed"] == 1
    assert sorted(p.name for p in dest.iterdir()) == [
        "001 - one.msg", "003 - three.msg",
    ]

    done = batch.renumber(cfg, str(dest))
    assert done["renumbered"] == planned["renumbered"]
    assert done["counts"]["index_rows_updated"] == 1
    assert sorted(p.name for p in dest.iterdir()) == [
        "001 - one.msg", "002 - three.msg",
    ]


def test_the_renumber_verb_refuses_a_folder_outside_the_archive_roots(cfg, tmp_path):
    outside = tmp_path / "not-the-archive"
    outside.mkdir()

    doc = batch.renumber(cfg, str(outside))

    assert doc["renumbered"] == {}
    assert doc["renumber_refused"] == [
        {"folder_path": str(outside), "reason": batch.REFUSED_OUTSIDE_ROOTS}
    ]

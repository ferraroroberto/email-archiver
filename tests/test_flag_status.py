"""Tests for the Outlook follow-up flag the scanner records (issue #51).

A sister tool (task-os) turns a flagged, archived mail into a task by reading
``emails.flag_status`` from this index, read-only. Three things have to hold for
that to work, and each has its own failure mode:

- the column exists on an **existing** database, not only a fresh one --
  ``init_db`` runs ``CREATE TABLE IF NOT EXISTS``, so a DDL edit alone is a
  silent no-op against the ~18k-row index this repo actually has;
- the property read distinguishes *unflagged* from *unreadable* -- MAPI omits
  0x1090 entirely on an unflagged message, so "absent" is a real answer;
- a pre-existing row reads NULL, not 0. NULL means "indexed before the scanner
  read the property"; 0 means "read, and not flagged". Folding the first into
  the second would assert a fact nobody established.
"""
from __future__ import annotations

import sqlite3

from email_archiver.database.models import _COLUMNS, init_db
from email_archiver.database.repository import EmailRecord, EmailRepository
from email_archiver.scanner.scanner import FLAG_FOLLOWUP, _flag_status

# The `emails` table as it shipped before this change, used to prove the
# migration on a database that predates the column.
_OLD_DDL = """
CREATE TABLE emails (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    file_path    TEXT    UNIQUE NOT NULL,
    folder_path  TEXT    NOT NULL,
    filename     TEXT    NOT NULL,
    subject      TEXT,
    sender       TEXT,
    recipients   TEXT,
    date_sent    TEXT,
    body_preview TEXT,
    file_mtime   REAL    NOT NULL,
    indexed_at   TEXT    NOT NULL DEFAULT (datetime('now'))
);
"""


class _FakeMsg:
    """Stands in for ``extract_msg.Message``: only ``getPropertyVal`` is used."""

    def __init__(self, props: dict[str, object]) -> None:
        self._props = props

    def getPropertyVal(self, key: str):  # noqa: N802 - extract_msg's spelling
        return self._props.get(key)


def _columns(conn: sqlite3.Connection) -> set[str]:
    return {r["name"] for r in conn.execute("PRAGMA table_info(emails)")}


# ------------------------------------------------------- the property read ---

def test_a_flagged_message_reads_as_flagged():
    assert _flag_status(_FakeMsg({"10900003": 2})) == FLAG_FOLLOWUP


def test_an_absent_property_is_unflagged_not_an_error():
    # MAPI writes 0x1090 only when a flag is set, so a plain message has no
    # such property at all. That is an answer, not a failure.
    assert _flag_status(_FakeMsg({})) == 0


def test_a_property_of_the_wrong_shape_does_not_raise():
    # A .msg that stores something unexpected must not abort the whole scan.
    for value in ("", "not-a-number", None, object()):
        assert _flag_status(_FakeMsg({"10900003": value})) == 0


def test_the_read_does_not_swallow_a_completed_flag():
    # 1 = flagged-then-completed. It is not 2, so it is not follow-up work, but
    # it must be carried through as itself rather than flattened to 0.
    assert _flag_status(_FakeMsg({"10900003": 1})) == 1


# ------------------------------------------------------------ the migration --

def test_a_fresh_database_has_the_column(tmp_path):
    conn = init_db(str(tmp_path / "emails.db"))
    assert "flag_status" in _columns(conn)


def test_an_existing_database_gains_the_column_without_losing_rows(tmp_path):
    """The case a DDL edit alone would silently miss."""
    path = str(tmp_path / "emails.db")
    old = sqlite3.connect(path)
    old.executescript(_OLD_DDL)
    old.execute(
        "INSERT INTO emails (file_path, folder_path, filename, file_mtime)"
        " VALUES ('a.msg', 'C:/archive', 'a.msg', 1.0)"
    )
    old.commit()
    old.close()

    conn = init_db(path)

    assert "flag_status" in _columns(conn)
    row = conn.execute("SELECT * FROM emails WHERE file_path = 'a.msg'").fetchone()
    assert row is not None, "the migration must not lose the existing row"
    assert row["flag_status"] is None, (
        "a row indexed before the property was read must stay NULL -- 0 would "
        "claim it was read and found unflagged"
    )


def test_the_migration_is_idempotent(tmp_path):
    path = str(tmp_path / "emails.db")
    init_db(path).close()
    # A duplicate-column error here would break every later scan.
    conn = init_db(path)
    assert "flag_status" in _columns(conn)


def test_every_migrated_column_is_also_in_the_create_table(tmp_path):
    """A fresh database and a migrated one must end up identical."""
    fresh = _columns(init_db(str(tmp_path / "fresh.db")))
    assert set(_COLUMNS) <= fresh


# ------------------------------------------------------------ the round trip --

def test_the_flag_survives_a_write_and_an_update(tmp_path):
    conn = init_db(str(tmp_path / "emails.db"))
    repo = EmailRepository(conn)
    rec = EmailRecord(
        file_path="a.msg", folder_path="C:/archive", filename="a.msg",
        file_mtime=1.0, flag_status=FLAG_FOLLOWUP,
    )
    repo.upsert_email(rec)
    conn.commit()

    stored = conn.execute("SELECT flag_status FROM emails WHERE file_path='a.msg'")
    assert stored.fetchone()["flag_status"] == FLAG_FOLLOWUP

    # Re-archiving the same path with the flag cleared must update, not keep a
    # stale 2 -- otherwise a task would be raised for mail no longer flagged.
    repo.upsert_email(EmailRecord(
        file_path="a.msg", folder_path="C:/archive", filename="a.msg",
        file_mtime=2.0, flag_status=0,
    ))
    conn.commit()
    stored = conn.execute("SELECT flag_status FROM emails WHERE file_path='a.msg'")
    assert stored.fetchone()["flag_status"] == 0


def test_an_unflagged_record_defaults_to_read_and_unflagged(tmp_path):
    conn = init_db(str(tmp_path / "emails.db"))
    EmailRepository(conn).upsert_email(EmailRecord(
        file_path="a.msg", folder_path="C:/archive", filename="a.msg",
        file_mtime=1.0,
    ))
    conn.commit()
    row = conn.execute("SELECT flag_status FROM emails").fetchone()
    assert row["flag_status"] == 0, (
        "the scanner always reads the property, so a row it wrote is 0, not NULL"
    )

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

from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRecord, EmailRepository
from email_archiver.scanner.scanner import FLAG_FOLLOWUP, _flag_status

from .conftest import _FakeMsg

# The column migration itself (fresh db has it / an existing db gains it
# without losing rows / the migration is idempotent / every migrated column
# is also in the create table) is covered once, parametrized over every
# entry in `models._COLUMNS`, in test_column_migrations.py.


# ------------------------------------------------------- the property read ---

def test_a_flagged_message_reads_as_flagged():
    assert _flag_status(_FakeMsg(props={"10900003": 2})) == FLAG_FOLLOWUP


def test_an_absent_property_is_unflagged_not_an_error():
    # MAPI writes 0x1090 only when a flag is set, so a plain message has no
    # such property at all. That is an answer, not a failure.
    assert _flag_status(_FakeMsg(props={})) == 0


def test_a_property_of_the_wrong_shape_does_not_raise():
    # A .msg that stores something unexpected must not abort the whole scan.
    for value in ("", "not-a-number", None, object()):
        assert _flag_status(_FakeMsg(props={"10900003": value})) == 0


def test_the_read_does_not_swallow_a_completed_flag():
    # 1 = flagged-then-completed. It is not 2, so it is not follow-up work, but
    # it must be carried through as itself rather than flattened to 0.
    assert _flag_status(_FakeMsg(props={"10900003": 1})) == 1


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

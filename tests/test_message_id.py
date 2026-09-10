"""Tests for the Internet Message-ID the scanner records (issue #53).

Batch mode pairs a live Outlook mail with the file archived from it by
Message-ID, because Outlook rewrites ``EntryID`` the moment a mail is moved
between folders — which is exactly what ``apply`` does to every mail it files.
Three things have to hold, each with its own failure mode:

- the column exists on an **existing** database, not only a fresh one — the same
  ``CREATE TABLE IF NOT EXISTS`` no-op that the follow-up flag hit;
- both sides normalise the id the same way, or a plan reports every already
  archived mail as new and files it a second time;
- an empty id never matches. A ``.msg`` with no Message-ID header stores ``""``
  and a row indexed before the column existed stores NULL; treating either as a
  match would pair two unrelated mails.
"""
from __future__ import annotations

import sqlite3

from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRecord, EmailRepository
from email_archiver.scanner.scanner import _message_id
from email_archiver.text import normalize_message_id

from .conftest import _FakeMsg
from .test_column_migrations import DDL_BEFORE_MESSAGE_ID as _OLD_DDL

# The column migration itself (fresh db has it / an existing db gains it
# without losing rows / the migration is idempotent / every migrated column
# is also in the create table) is covered once, parametrized over every
# entry in `models._COLUMNS`, in test_column_migrations.py.


# ------------------------------------------------------------ normalisation --

def test_the_angle_brackets_are_stripped():
    assert normalize_message_id("<abc123@mail.example>") == "abc123@mail.example"


def test_an_id_without_brackets_is_left_alone():
    assert normalize_message_id("abc123@mail.example") == "abc123@mail.example"


def test_surrounding_whitespace_is_stripped():
    assert normalize_message_id("  <abc123@mail.example>\n") == "abc123@mail.example"


def test_case_is_preserved():
    # RFC 5322 makes the left-hand side case-sensitive; folding it would merge
    # ids that are genuinely different mails.
    assert normalize_message_id("<AbC@Mail.Example>") == "AbC@Mail.Example"


def test_a_missing_id_is_the_empty_string_not_none():
    for raw in (None, "", "   ", "<>"):
        assert normalize_message_id(raw) == ""


# --------------------------------------------------------------- the read ----

def test_the_scanner_reads_the_header_extract_msg_exposes():
    assert _message_id(_FakeMsg("<abc@mail.example>")) == "abc@mail.example"


def test_the_scanner_falls_back_to_the_raw_proptag():
    # Some message classes have no `messageId` property at all.
    msg = _FakeMsg(props={"1035001F": "<abc@mail.example>"})
    assert _message_id(msg) == "abc@mail.example"


def test_an_empty_property_falls_back_rather_than_returning_blank():
    msg = _FakeMsg("", props={"1035001F": "<abc@mail.example>"})
    assert _message_id(msg) == "abc@mail.example"


def test_a_message_with_no_id_anywhere_reads_as_empty_not_an_error():
    assert _message_id(_FakeMsg(None)) == ""
    assert _message_id(_FakeMsg()) == ""


# ------------------------------------------------------------ the migration --
# The generic column-migration behaviour is covered once, parametrized over
# `models._COLUMNS`, in test_column_migrations.py. What's left here is
# message_id-specific: the index-creation ordering.

def test_the_index_is_created_after_the_column_exists(tmp_path):
    """The order that a naive DDL edit gets wrong.

    `init_db` runs `_DDL` before the ALTERs, so an index on `message_id`
    declared there would be created against a column an existing table does not
    have yet -- and the whole script would fail, breaking every later scan.
    """
    path = str(tmp_path / "emails.db")
    old = sqlite3.connect(path)
    old.executescript(_OLD_DDL)
    old.commit()
    old.close()

    conn = init_db(path)
    indexes = {r[1] for r in conn.execute("PRAGMA index_list(emails)")}
    assert "idx_emails_message_id" in indexes


# ------------------------------------------------------------- the lookup ----

def _repo_with(conn, **overrides) -> EmailRepository:
    repo = EmailRepository(conn)
    rec = EmailRecord(
        file_path="C:/archive/Project Alpha/001 - kickoff.msg",
        folder_path="C:/archive/Project Alpha",
        filename="001 - kickoff.msg",
        file_mtime=1.0,
    )
    for key, value in overrides.items():
        setattr(rec, key, value)
    repo.upsert_email(rec)
    conn.commit()
    return repo


def test_a_known_message_id_resolves_to_its_archived_file():
    conn = init_db(":memory:")
    repo = _repo_with(conn, message_id="abc@mail.example")
    assert repo.find_path_by_message_id("abc@mail.example") == (
        "C:/archive/Project Alpha/001 - kickoff.msg"
    )


def test_an_unknown_message_id_resolves_to_nothing():
    conn = init_db(":memory:")
    repo = _repo_with(conn, message_id="abc@mail.example")
    assert repo.find_path_by_message_id("other@mail.example") is None


def test_the_empty_message_id_never_matches():
    """Two mails with no Message-ID are not the same mail."""
    conn = init_db(":memory:")
    repo = _repo_with(conn, message_id="")
    assert repo.find_path_by_message_id("") is None


def test_a_null_row_never_matches_the_empty_id():
    conn = init_db(":memory:")
    conn.execute(
        "INSERT INTO emails (file_path, folder_path, filename, file_mtime)"
        " VALUES ('a.msg', 'C:/archive', 'a.msg', 1.0)"
    )
    conn.commit()
    assert EmailRepository(conn).find_path_by_message_id("") is None


def test_the_id_survives_a_write_and_an_update():
    conn = init_db(":memory:")
    repo = _repo_with(conn, message_id="abc@mail.example")
    # Re-indexing the same path with a corrected id must update, not keep the
    # stale one -- otherwise a plan pairs the mail with the wrong file forever.
    _repo_with(conn, message_id="def@mail.example")
    assert repo.find_path_by_message_id("abc@mail.example") is None
    assert repo.find_path_by_message_id("def@mail.example") is not None

"""Migration tests for columns ALTERed onto an existing `emails` table.

Every entry in ``models._COLUMNS`` (issue #51's ``flag_status``, issue #53's
``message_id``) got its own copy of "fresh db has it" / "existing db gains it
without losing rows" / "the migration is idempotent" / "every migrated column
is also in the create table" -- one column-name swap apart. Parametrizing
over ``_COLUMNS`` replaces all four copies with one (issue #63); the next
column added there is covered for free.
"""
from __future__ import annotations

import sqlite3

import pytest

from email_archiver.database.models import _COLUMNS, init_db

from .conftest import _columns

# The `emails` table before `flag_status` was added (issue #51).
_DDL_BEFORE_FLAG_STATUS = """
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

# The `emails` table after `flag_status` landed but before `message_id` did
# (issue #53). Also used by test_message_id.py's index-ordering regression,
# which needs a schema that has flag_status but not yet message_id.
DDL_BEFORE_MESSAGE_ID = """
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
    indexed_at   TEXT    NOT NULL DEFAULT (datetime('now')),
    flag_status  INTEGER
);
"""

# The "old" DDL each column was migrated onto -- i.e. the schema the instant
# before that column existed.
_OLD_DDL_BY_COLUMN = {
    "flag_status": _DDL_BEFORE_FLAG_STATUS,
    "message_id": DDL_BEFORE_MESSAGE_ID,
}


@pytest.mark.parametrize("column", list(_COLUMNS))
def test_a_fresh_database_has_the_column(column, tmp_path):
    conn = init_db(str(tmp_path / "emails.db"))
    assert column in _columns(conn)


@pytest.mark.parametrize("column", list(_COLUMNS))
def test_an_existing_database_gains_the_column_without_losing_rows(column, tmp_path):
    """The case a DDL edit alone would silently miss."""
    path = str(tmp_path / "emails.db")
    old = sqlite3.connect(path)
    old.executescript(_OLD_DDL_BY_COLUMN[column])
    old.execute(
        "INSERT INTO emails (file_path, folder_path, filename, file_mtime)"
        " VALUES ('a.msg', 'C:/archive', 'a.msg', 1.0)"
    )
    old.commit()
    old.close()

    conn = init_db(path)

    assert column in _columns(conn)
    row = conn.execute("SELECT * FROM emails WHERE file_path = 'a.msg'").fetchone()
    assert row is not None, "the migration must not lose the existing row"
    assert row[column] is None, (
        "a row indexed before this column was read must stay NULL -- a "
        "concrete default would claim it was read"
    )


@pytest.mark.parametrize("column", list(_COLUMNS))
def test_the_migration_is_idempotent(column, tmp_path):
    path = str(tmp_path / "emails.db")
    init_db(path).close()
    # A duplicate-column error here would break every later scan.
    conn = init_db(path)
    assert column in _columns(conn)


def test_every_migrated_column_is_also_in_the_create_table(tmp_path):
    """A fresh database and a migrated one must end up identical."""
    fresh = _columns(init_db(str(tmp_path / "fresh.db")))
    assert set(_COLUMNS) <= fresh

"""A writer meeting "database is locked" queues and retries instead of failing.

Two connections share one temp database. The first holds the write lock; the
second has a zero busy timeout so the lock surfaces at once, and an injected
sleep stands in for the clock: it releases the first writer on the Nth wait.
"""
from __future__ import annotations

import logging
import sqlite3

import pytest

from email_archiver.database.models import get_connection, init_db
from email_archiver.database.repository import (
    LOCK_RETRIES,
    EmailRecord,
    EmailRepository,
)


def _rec(name: str) -> EmailRecord:
    return EmailRecord(
        file_path=name, folder_path="archive", filename=name, file_mtime=1.0
    )


def _two_writers(tmp_path, release_after: int | None):
    db = tmp_path / "emails.db"
    init_db(str(db)).close()
    holder = get_connection(db)
    holder.execute("BEGIN IMMEDIATE")  # first writer owns the lock
    waits: list[float] = []

    def fake_sleep(seconds: float) -> None:
        waits.append(seconds)
        if release_after is not None and len(waits) == release_after:
            holder.rollback()

    second = EmailRepository(get_connection(db, timeout=0), sleep=fake_sleep)
    return db, second, waits


def test_second_writer_retries_until_first_releases(tmp_path, caplog):
    db, second, waits = _two_writers(tmp_path, release_after=2)
    with caplog.at_level(logging.INFO):
        second.upsert_email(_rec("a.msg"))
        second.commit()

    assert waits == [1.0, 2.0]
    assert "locked" in caplog.text and "retry 1/" in caplog.text
    row = get_connection(db).execute("SELECT COUNT(*) FROM emails").fetchone()
    assert row[0] == 1


def test_lock_that_never_clears_still_raises(tmp_path):
    _, second, waits = _two_writers(tmp_path, release_after=None)
    with pytest.raises(sqlite3.OperationalError, match="locked"):
        second.upsert_email(_rec("a.msg"))
    assert len(waits) == LOCK_RETRIES


def test_other_operational_errors_are_not_retried(tmp_path):
    _, second, waits = _two_writers(tmp_path, release_after=None)
    with pytest.raises(sqlite3.OperationalError, match="no such table"):
        second._retry_locked(
            lambda: second._conn.execute("SELECT * FROM missing"), "probe"
        )
    assert waits == []

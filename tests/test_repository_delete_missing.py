"""Regression tests for ``EmailRepository.delete_missing_emails`` (issue #43).

The old implementation bound one SQL parameter per known path, which raises
``sqlite3.OperationalError: too many SQL variables`` once the archive holds
more files than SQLite's ``SQLITE_LIMIT_VARIABLE_NUMBER`` (32,766) -- after
the entire indexing pass has already run, aborting the purge. The fix routes
through a temp table instead, independent of file count.
"""
from __future__ import annotations

from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRecord, EmailRepository


def _repo(tmp_path):
    conn = init_db(str(tmp_path / "emails.db"))
    return EmailRepository(conn), conn


def _seed(repo, file_path):
    repo.upsert_email(
        EmailRecord(
            file_path=file_path,
            folder_path="C:/archive",
            filename=file_path,
            file_mtime=1.0,
        )
    )


def test_deletes_only_rows_missing_from_known_paths(tmp_path):
    repo, conn = _repo(tmp_path)
    _seed(repo, "a.msg")
    _seed(repo, "b.msg")
    conn.commit()

    deleted = repo.delete_missing_emails(["a.msg"])
    conn.commit()

    assert deleted == 1
    assert repo.get_mtime("a.msg") == 1.0
    assert repo.get_mtime("b.msg") is None


def test_empty_known_paths_is_a_no_op(tmp_path):
    repo, conn = _repo(tmp_path)
    _seed(repo, "a.msg")
    conn.commit()

    assert repo.delete_missing_emails([]) == 0
    assert repo.get_mtime("a.msg") == 1.0


def test_survives_more_known_paths_than_sqlites_variable_limit(tmp_path):
    # SQLITE_LIMIT_VARIABLE_NUMBER is 32766 in the shipped 3.50.4 build --
    # the pre-fix query raised OperationalError past that many bound
    # parameters. Only a couple of real rows are needed; the rest of the
    # list just has to push the count past the limit.
    repo, conn = _repo(tmp_path)
    _seed(repo, "keep.msg")
    _seed(repo, "gone.msg")
    conn.commit()

    known_paths = ["keep.msg"] + [f"synthetic_{i}.msg" for i in range(40_000)]
    deleted = repo.delete_missing_emails(known_paths)
    conn.commit()

    assert deleted == 1
    assert repo.get_mtime("keep.msg") == 1.0
    assert repo.get_mtime("gone.msg") is None

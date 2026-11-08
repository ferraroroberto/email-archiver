"""The runtime writers refuse a path under a code folder and write elsewhere."""
from __future__ import annotations

import logging

import pytest

from email_archiver import runtime_files


def test_make_dirs_creates_a_folder_outside_the_code_folders(tmp_path):
    target = tmp_path / "archive" / "mail"
    runtime_files.make_dirs(target)
    assert target.is_dir()


@pytest.mark.parametrize("folder", runtime_files.CODE_FOLDERS)
def test_make_dirs_refuses_a_code_folder(folder):
    target = runtime_files.REPO_ROOT / folder / "runtime_files_probe"
    with pytest.raises(ValueError, match="code folder"):
        runtime_files.make_dirs(target)
    assert not target.exists()


def test_connect_sqlite_refuses_a_code_folder():
    with pytest.raises(ValueError, match="code folder"):
        runtime_files.connect_sqlite(runtime_files.REPO_ROOT / "src" / "probe.db")


def test_connect_sqlite_opens_a_database_elsewhere(tmp_path):
    conn = runtime_files.connect_sqlite(tmp_path / "index.db")
    conn.close()
    assert (tmp_path / "index.db").exists()


def test_file_handler_refuses_a_code_folder():
    with pytest.raises(ValueError, match="code folder"):
        runtime_files.file_handler(runtime_files.REPO_ROOT / "scripts" / "probe.log")


def test_file_handler_writes_elsewhere(tmp_path):
    handler = runtime_files.file_handler(tmp_path / "run.log")
    try:
        handler.emit(logging.LogRecord("t", logging.INFO, __file__, 1, "hello", None, None))
    finally:
        handler.close()
    assert "hello" in (tmp_path / "run.log").read_text(encoding="utf-8")

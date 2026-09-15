"""A configured archive root that is not there must not cost the index its rows.

An unmounted or unsynced OneDrive root reads as "nothing there". Before issue
#69 the scanner only warned about it and then ran the post-scan purge against
the paths it *could* see, deleting every index row under the missing root. The
headless entry point also exited 0 whatever happened, so an unattended nightly
run could quietly empty half the index and report success.
"""
from __future__ import annotations

from pathlib import Path

import pytest

import main_scan
from email_archiver.database.models import init_db
from email_archiver.database.repository import EmailRecord, EmailRepository
from email_archiver.scanner.scanner import FolderScanner


def _cfg(tmp_path, roots):
    return {
        "archive": {"root_paths": [str(r) for r in roots]},
        "database": {"path": str(tmp_path / "emails.db")},
        "scanning": {"batch_size": 100, "body_preview_length": 100},
    }


def _seed(db_path, file_path):
    conn = init_db(db_path)
    EmailRepository(conn).upsert_email(
        EmailRecord(
            file_path=file_path,
            folder_path=str(Path(file_path).parent),
            filename=Path(file_path).name,
            file_mtime=1.0,
        )
    )
    conn.commit()
    conn.close()


def _mtime(db_path, file_path):
    conn = init_db(db_path)
    try:
        return EmailRepository(conn).get_mtime(file_path)
    finally:
        conn.close()


@pytest.fixture
def two_roots(tmp_path):
    """A present root holding one (unreadable) .msg, and a root that is gone.

    The present root must yield at least one known path: the purge is a no-op
    on an empty list, which would let the test pass on pre-fix code.
    """
    present = tmp_path / "present"
    present.mkdir()
    (present / "001 - on disk.msg").write_bytes(b"not a real msg")
    missing = tmp_path / "missing"
    cfg = _cfg(tmp_path, [present, missing])
    under_missing = str(missing / "Project" / "001 - archived.msg")
    _seed(cfg["database"]["path"], under_missing)
    return cfg, missing, under_missing


def test_scan_leaves_rows_under_a_missing_root_untouched(two_roots):
    cfg, missing, under_missing = two_roots

    stats = FolderScanner(cfg).scan()

    assert _mtime(cfg["database"]["path"], under_missing) == 1.0
    assert stats.deleted == 0
    assert stats.missing_roots == [str(missing)]


def test_headless_scan_exits_non_zero_and_names_the_missing_root(
    two_roots, monkeypatch, caplog
):
    cfg, missing, under_missing = two_roots
    monkeypatch.setattr(main_scan, "load_config", lambda: cfg)
    monkeypatch.setattr(main_scan, "setup_logging", lambda _cfg: None)

    with caplog.at_level("ERROR"):
        code = main_scan.main(["--no-ui"])

    assert code == main_scan.EXIT_ROOT_MISSING
    assert str(missing) in caplog.text
    assert _mtime(cfg["database"]["path"], under_missing) == 1.0


def test_headless_scan_exits_zero_when_every_root_is_present(
    tmp_path, monkeypatch, capsys
):
    root = tmp_path / "root"
    root.mkdir()
    (root / "001 - unreadable.msg").write_bytes(b"not a real msg")
    cfg = _cfg(tmp_path, [root])
    monkeypatch.setattr(main_scan, "load_config", lambda: cfg)
    monkeypatch.setattr(main_scan, "setup_logging", lambda _cfg: None)

    # A per-file extraction error is not a failed scan.
    assert main_scan.main(["--no-ui"]) == main_scan.EXIT_OK
    # capsys stdout is not a TTY, as under a captured job log.
    out = capsys.readouterr().out
    assert "\r" not in out
    assert "1 / 1" in out


def test_headless_scan_exits_non_zero_with_no_roots(tmp_path, monkeypatch):
    cfg = _cfg(tmp_path, [])
    monkeypatch.setattr(main_scan, "load_config", lambda: cfg)
    monkeypatch.setattr(main_scan, "setup_logging", lambda _cfg: None)

    assert main_scan.main(["--no-ui"]) == main_scan.EXIT_CONFIG_ERROR


def test_headless_scan_exits_non_zero_without_a_config(monkeypatch):
    def _missing():
        raise FileNotFoundError("Config file not found")

    monkeypatch.setattr(main_scan, "load_config", _missing)

    assert main_scan.main(["--no-ui"]) == main_scan.EXIT_CONFIG_ERROR

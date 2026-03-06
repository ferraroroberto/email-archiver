"""
Database schema and connection factory.

Design notes:
- SQLite FTS5 is used for full-text search over subject/sender/recipients/body.
  FTS5 is built into Python's sqlite3 on Windows; no extra dependency needed.
- Triggers keep the FTS index in sync with the emails table automatically.
- The folders table is a denormalized summary updated by the scanner for fast
  folder-level scoring in the suggestion engine.
- WAL journal mode allows concurrent reads during a long scan without blocking
  the archive command.
"""
from __future__ import annotations

import sqlite3
from pathlib import Path


_DDL = """
PRAGMA journal_mode = WAL;
PRAGMA synchronous  = NORMAL;
PRAGMA foreign_keys = ON;

-- ------------------------------------------------------------------ emails --
CREATE TABLE IF NOT EXISTS emails (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    file_path    TEXT    UNIQUE NOT NULL,   -- absolute path to .msg file
    folder_path  TEXT    NOT NULL,          -- parent directory (archive target)
    filename     TEXT    NOT NULL,
    subject      TEXT,
    sender       TEXT,
    recipients   TEXT,
    date_sent    TEXT,                      -- ISO-8601 or empty
    body_preview TEXT,                      -- first N chars of plain-text body
    file_mtime   REAL    NOT NULL,          -- os.stat().st_mtime for change detection
    indexed_at   TEXT    NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_emails_folder  ON emails(folder_path);
CREATE INDEX IF NOT EXISTS idx_emails_mtime   ON emails(file_mtime);

-- ------------------------------------------------------- FTS5 search index --
-- content= mode: FTS5 mirrors the emails table; triggers keep it in sync.
CREATE VIRTUAL TABLE IF NOT EXISTS emails_fts USING fts5(
    subject,
    sender,
    recipients,
    body_preview,
    content = emails,
    content_rowid = id,
    tokenize = 'unicode61 remove_diacritics 1'
);

CREATE TRIGGER IF NOT EXISTS emails_ai
AFTER INSERT ON emails BEGIN
    INSERT INTO emails_fts(rowid, subject, sender, recipients, body_preview)
    VALUES (new.id, new.subject, new.sender, new.recipients, new.body_preview);
END;

CREATE TRIGGER IF NOT EXISTS emails_ad
AFTER DELETE ON emails BEGIN
    INSERT INTO emails_fts(emails_fts, rowid, subject, sender, recipients, body_preview)
    VALUES ('delete', old.id, old.subject, old.sender, old.recipients, old.body_preview);
END;

CREATE TRIGGER IF NOT EXISTS emails_au
AFTER UPDATE ON emails BEGIN
    INSERT INTO emails_fts(emails_fts, rowid, subject, sender, recipients, body_preview)
    VALUES ('delete', old.id, old.subject, old.sender, old.recipients, old.body_preview);
    INSERT INTO emails_fts(rowid, subject, sender, recipients, body_preview)
    VALUES (new.id, new.subject, new.sender, new.recipients, new.body_preview);
END;

-- --------------------------------------------------------------- folders ---
-- Aggregated per-folder stats used by the suggestion engine for scoring.
CREATE TABLE IF NOT EXISTS folders (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_path  TEXT    UNIQUE NOT NULL,
    email_count  INTEGER NOT NULL DEFAULT 0,
    last_updated TEXT    NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_folders_path ON folders(folder_path);
"""


def get_connection(db_path: str | Path) -> sqlite3.Connection:
    """
    Open a SQLite connection with sensible defaults.
    Returns a connection with row_factory=sqlite3.Row so columns are
    accessible by name.
    """
    conn = sqlite3.connect(str(db_path), check_same_thread=False)
    conn.row_factory = sqlite3.Row
    # Increase cache to speed up FTS5 queries on large datasets
    conn.execute("PRAGMA cache_size = -32768")  # 32 MB page cache
    return conn


def init_db(db_path: str | Path) -> sqlite3.Connection:
    """Create all tables/indexes/triggers if they don't exist yet."""
    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    conn = get_connection(db_path)
    conn.executescript(_DDL)
    conn.commit()
    return conn

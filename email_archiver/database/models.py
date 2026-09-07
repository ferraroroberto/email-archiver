"""
Database schema and connection factory.

Design notes:
- SQLite FTS5 is used for full-text search over subject/sender/recipients/body.
  FTS5 is built into Python's sqlite3 on Windows; no extra dependency needed.
- Triggers keep the FTS index in sync with the emails table automatically.
- The folders table is a plain registry of folders the scanner has seen
  (path + last-scanned timestamp); the suggestion engine scores folders by
  aggregating per-email FTS matches in Python, not by reading this table.
- WAL journal mode allows concurrent reads during a long scan without blocking
  the archive command.
"""
from __future__ import annotations

import logging
import sqlite3
from pathlib import Path

logger = logging.getLogger(__name__)


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
    indexed_at   TEXT    NOT NULL DEFAULT (datetime('now')),
    flag_status  INTEGER                    -- MAPI PidTagFlagStatus; see _COLUMNS
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
-- Plain registry of folders the scanner has seen (path + last-scanned time).
CREATE TABLE IF NOT EXISTS folders (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_path  TEXT    UNIQUE NOT NULL,
    last_updated TEXT    NOT NULL DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_folders_path ON folders(folder_path);
"""


# Columns added to `emails` after the table shipped. The DDL above only runs
# under CREATE TABLE IF NOT EXISTS, so on an existing database it is a no-op and
# a new column would never appear -- every install here predates these, so they
# have to be ALTERed in explicitly. Keep the two in sync: a column listed here
# must also be in the CREATE TABLE, so a fresh database and a migrated one end
# up identical.
_COLUMNS: dict[str, str] = {
    # MAPI PidTagFlagStatus (0x1090): 2 = flagged for follow-up, 1 = completed,
    # 0 = not flagged. Deliberately nullable with no default -- NULL means "this
    # row was indexed before the scanner read the property", which is a
    # different fact from 0, "read, and not flagged". Rows indexed before this
    # build keep NULL until their file changes and is re-read.
    "flag_status": "INTEGER",
}


def _add_missing_columns(conn: sqlite3.Connection) -> list[str]:
    """ALTER IN the `_COLUMNS` an existing `emails` table is missing.

    Returns the columns actually added, so a caller can log a real migration.
    Idempotent: a second run finds nothing to do.
    """
    have = {row["name"] for row in conn.execute("PRAGMA table_info(emails)")}
    added = []
    for name, decl in _COLUMNS.items():
        if name not in have:
            conn.execute(f"ALTER TABLE emails ADD COLUMN {name} {decl}")
            added.append(name)
    return added


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
    """Create all tables/indexes/triggers if they don't exist yet, then bring an
    older `emails` table up to the current column set."""
    Path(db_path).parent.mkdir(parents=True, exist_ok=True)
    conn = get_connection(db_path)
    conn.executescript(_DDL)
    added = _add_missing_columns(conn)
    if added:
        logger.info("Added column(s) to emails: %s", ", ".join(added))
    conn.commit()
    return conn

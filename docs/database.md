# Database schema

```sql
-- One row per indexed .msg file
CREATE TABLE emails (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    file_path    TEXT UNIQUE NOT NULL,   -- absolute path
    folder_path  TEXT NOT NULL,          -- parent directory
    filename     TEXT NOT NULL,
    subject      TEXT,
    sender       TEXT,
    recipients   TEXT,
    date_sent    TEXT,                   -- ISO-8601
    body_preview TEXT,                   -- first 500 chars of plain text
    file_mtime   REAL NOT NULL,          -- for incremental scan (os.stat)
    indexed_at   TEXT NOT NULL DEFAULT (datetime('now')),
    flag_status  INTEGER,                 -- Outlook follow-up flag; see below
    message_id   TEXT                     -- Internet Message-ID; see below
);

-- FTS5 full-text index (kept in sync via triggers)
CREATE VIRTUAL TABLE emails_fts USING fts5(
    subject, sender, recipients, body_preview,
    content = emails, content_rowid = id,
    tokenize = 'unicode61 remove_diacritics 1'
);

-- Plain registry of folders the scanner has seen
CREATE TABLE folders (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_path  TEXT UNIQUE NOT NULL,
    last_updated TEXT NOT NULL DEFAULT (datetime('now'))
);
```

The FTS5 index is automatically kept in sync with the `emails` table via `AFTER INSERT`, `AFTER UPDATE`, and `AFTER DELETE` triggers — no manual maintenance needed.

WAL journal mode is enabled so the archive command can read the DB while a scan is running in another process without blocking.

## The follow-up flag (`flag_status`)

`flag_status` records Outlook's follow-up flag, read from the archived `.msg` as MAPI `PidTagFlagStatus` (`0x1090`): **2** = flagged for follow-up, **1** = flagged then completed, **0** = read, and not flagged. It exists so other tools can turn a flagged mail into a task — task-os polls it read-only and raises one Inbox task per flagged message.

Two things about it are easy to get wrong:

- **Flag first, then archive.** The flag is written into the `.msg` at archive time, by the same `SaveAs` call that creates the file. The archived file is a snapshot: flagging a message in Outlook *after* it has been archived changes nothing on disk, and no re-scan will see it.
- **`NULL` is not `0`.** A row indexed before this column existed reads `NULL`, meaning *the property was never read* — a different fact from `0`, *read, and not flagged*. Existing rows keep `NULL` until their file changes and is re-indexed. That is deliberate and costs nothing: Outlook does not preserve the flag through filing, so an already-archived backlog carries no flags to find (a random sample of 600 of ~18k archived files found none).

The column is added to an existing database automatically on the next run — `init_db` ALTERs in any column the `emails` table is missing, because its `CREATE TABLE IF NOT EXISTS` is a no-op once the table exists.

## The Internet Message-ID (`message_id`)

`message_id` records the mail's `Message-ID` header (MAPI `PR_INTERNET_MESSAGE_ID`, `0x1035001F`) read back out of the archived `.msg`, stored **without** its angle brackets so both sides of the app spell it the same way. It exists for [batch mode](batch-mode.md#batch-mode-headless): it is the only identity that survives a mail being moved between Outlook folders, so it is what lets `plan` recognise a mail that is already filed and what pairs a `revert` with the files written for it.

- **`NULL` is not `""`.** A row indexed before this column existed reads `NULL`, *the header was never read*; a `.msg` that genuinely carries no `Message-ID` reads `""`. Neither ever matches a lookup — two mails with no Message-ID are not the same mail.
- Added to an existing database the same way `flag_status` is, plus an index on the column created **after** the ALTER — declaring it alongside the `CREATE TABLE` would run it against a column an existing database does not have yet and take the whole scan down with it.
- **The existing backlog is not backfilled.** Scanning is incremental on `mtime`, so a `.msg` indexed before this column existed is skipped and keeps its `NULL` — its Message-ID is in the file, but nothing has read it. Everything archived from now on is indexed with its id on the next scan (a new file is never skipped), so `plan`'s already-archived detection is complete for the batch workflow's own loop and blind only to mail filed before it. To cover the backlog too, force a full re-read by deleting `data/emails.db` and re-scanning — the same cost as a first scan.

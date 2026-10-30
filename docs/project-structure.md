# Project structure

```
archiver/
│
├── config/
│   ├── config.example.yaml     ← Template: copy to config.yaml
│   ├── config.yaml             ← Your local config (git-ignored)
│   ├── mailboxes.sample.json   ← Template: copy to mailboxes.json
│   └── mailboxes.json          ← Your mailbox registry (git-ignored, optional)
│
├── data/
│   └── emails.db                ← SQLite database (auto-created on first scan)
│
├── logs/
│   └── archiver.log             ← Log file (plain FileHandler, not rotated)
│
├── email_archiver/              ← Main package
│   ├── config.py                ← YAML loader, path resolution, logging setup, mailbox registry
│   ├── text.py                  ← Shared subject + Message-ID normalisation
│   ├── paths.py                 ← Archive-root guard shared by revert and renumber
│   ├── batch.py                 ← Headless plan/apply/revert/renumber orchestration (no COM, no tkinter)
│   ├── batch_common.py          ← What every batch-family document shares: schema version, envelope, error shape, folder renumbering
│   ├── batch_move.py            ← Moving an archived mail out of the Inbox: the retries and the `message_changed` error
│   ├── dismiss.py               ← Headless move-only verb: Inbox mail to the archive folder, tagged, no .msg kept (no COM)
│   ├── draft.py                 ← Headless draft: spec validation, BCC-self, the draft document (no COM)
│   ├── find.py                  ← Headless read-only live search over a mailbox's Outlook folders (no COM)
│   ├── send.py                  ← Headless read-back + guarded send: fingerprint, approval check (no COM)
│   ├── sync.py                  ← Headless Send/Receive: one profile-wide trigger, bounded wait, per-mailbox outcome (no COM)
│   ├── renumber.py              ← Re-sequence a folder into sent-date order
│   ├── explorer.py              ← Foremost open Explorer window → folder path
│   │
│   ├── cli/
│   │   ├── __init__.py          ← `python -m email_archiver <archive|scan|ui|batch>` dispatcher
│   │   ├── archive.py           ← Archive Email dialog entry
│   │   ├── scan.py              ← Scan Archive entry, headless exit codes
│   │   ├── launcher.py          ← Full launcher window entry
│   │   └── batch.py             ← Headless JSON verbs entry
│   │
│   ├── database/
│   │   ├── models.py            ← SQLite schema, FTS5 setup, connection factory
│   │   └── repository.py       ← All SQL queries (EmailRepository class)
│   │
│   ├── scanner/
│   │   └── scanner.py          ← Incremental .msg file indexer (FolderScanner)
│   │
│   ├── outlook/
│   │   ├── client.py           ← Outlook COM isolation (OutlookClient): read path + batch surface
│   │   ├── drafts.py           ← Draft surface OutlookClient inherits (DraftSurface)
│   │   ├── process.py          ← Outlook process: ensure_running; the session-0 desktop check
│   │   ├── mapi.py             ← Shared constants and COM-free helpers
│   │   ├── stores.py           ← Which Outlook store a mailbox lives in (by address)
│   │   ├── sending.py          ← The one COM Send call (send verb only)
│   │   └── syncing.py          ← The one place a Send/Receive is started (sync verb only)
│   │
│   ├── archiver/
│   │   └── archiver.py         ← File saving logic (EmailArchiver)
│   │
│   ├── engine/
│   │   └── suggester.py        ← Three-stage folder ranking (SuggestionEngine)
│   │
│   └── ui/
│       ├── app.py              ← ArchiveDialog, ScanWindow, LauncherApp
│       └── dialogs.py          ← Native folder-picker wrapper
│
├── scripts/                     ← pre-ship gate (verify-before-ship.ps1) and the vendored ratchets
├── tests/                       ← pytest suite (filename fitting, sequencing, date-prefix toggle, Explorer picker, Message-ID, batch verbs)
├── docs/
│   ├── architecture.mmd         ← Hand-authored Mermaid diagram of internal structure
│   └── *.md                     ← Reference docs, indexed in the README
│
├── main_archive.py              ← Stream Deck entry: Archive Email (thin wrapper over cli/archive.py)
├── main_scan.py                 ← Stream Deck entry: Scan Archive (thin wrapper over cli/scan.py)
├── main_ui.py                   ← Full launcher, both buttons (thin wrapper over cli/launcher.py)
├── main_batch.py                ← Headless JSON entry (thin wrapper over cli/batch.py)
├── launch_archive.bat           ← Runs pythonw main_archive.py (no console)
├── launch_scan.bat              ← Runs pythonw main_scan.py (no console)
├── run-scan-nightly.bat         ← Scheduled job: headless scan in the foreground, propagates the exit code
└── requirements.txt
```

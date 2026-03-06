# Email Archiver

A modular Python application for automatically indexing and archiving Outlook emails into a structured OneDrive folder system.

---

## Background

Email archiving used to be a fully manual process: select an email in Outlook, navigate to the right project folder in Windows Explorer, drag, save attachments separately, rename everything consistently. With a deeply nested OneDrive archive of nearly 18,000 emails across hundreds of project folders, this was taking significant time every day.

A first version of this automation was built around 2021–2022 using a different approach: an Excel spreadsheet as the database (cached as a Pickle file for speed), `fuzzywuzzy` for fuzzy subject matching, and three separate scripts — one to classify/index (`email-automation-classify.py`), one to archive with AI suggestion (`email-automation-archive.py`), and one to save to the folder currently open in Windows Explorer (`email-automation-save.py`). That code lives in `old_refactored/` as a reference.

The original approach had several pain points over time:

- **Slow startup**: loading a large Excel file via Pandas took several seconds — too slow for a Stream Deck button workflow
- **No incremental scan**: every re-scan re-read every `.msg` file from scratch
- **Brittle architecture**: all logic was in flat scripts with duplicated helpers
- **Fuzzy matching only**: `fuzz.token_set_ratio` over subject strings worked but missed context from sender, recipients and body
- **Excel as a database**: not designed for tens of thousands of rows or concurrent access

This rewrite (March 2026) replaces all of that with a clean modular architecture, SQLite with full-text search, and an instant-startup UI designed for Stream Deck use.

---

## What it does

Two commands, each launchable from a Stream Deck button or the command line:

### 1. Scan Archive
Walks the entire OneDrive archive from a configured root folder, opens every `.msg` file it finds, extracts metadata (subject, sender, recipients, date, body preview), and stores it in a local SQLite database with a full-text search index. Subsequent scans are incremental — only new or modified files are processed.

### 2. Archive Email
Connects to the running Outlook instance, reads the currently selected email, queries the database to find the most relevant project folders, and presents ranked suggestions. You click a folder (or browse manually), confirm, and the email is saved as a `.msg` file plus all attachments — all consistently numbered — in under two seconds.

---

## File naming convention

When an email is archived in a folder, files are named with a zero-padded 3-digit sequence number derived from the highest existing prefix in that folder:

```
023_Project_Alpha_meeting_notes.msg
023_01_invoice.pdf
023_02_signed_contract.docx
```

- The email gets `NNN_sanitized_subject.msg`
- Each real attachment gets `NNN_NN_original_filename.ext`
- Embedded images (inline in HTML body) are skipped automatically

---

## Project structure

```
archiver/
│
├── config/
│   ├── config.example.yaml     ← Template: copy to config.yaml
│   └── config.yaml             ← Your local config (git-ignored)
│
├── data/
│   └── emails.db                ← SQLite database (auto-created on first scan)
│
├── logs/
│   └── archiver.log             ← Rotating log file
│
├── email_archiver/              ← Main package
│   ├── config.py                ← YAML loader, path resolution, logging setup
│   │
│   ├── database/
│   │   ├── models.py            ← SQLite schema, FTS5 setup, connection factory
│   │   └── repository.py       ← All SQL queries (EmailRepository class)
│   │
│   ├── scanner/
│   │   └── scanner.py          ← Incremental .msg file indexer (FolderScanner)
│   │
│   ├── outlook/
│   │   └── client.py           ← Outlook COM isolation (OutlookClient)
│   │
│   ├── archiver/
│   │   └── archiver.py         ← File saving logic (EmailArchiver)
│   │
│   ├── engine/
│   │   └── suggester.py        ← Two-stage folder ranking (SuggestionEngine)
│   │
│   └── ui/
│       ├── app.py              ← ArchiveDialog, ScanWindow, LauncherApp
│       └── dialogs.py          ← Confirm, browse, progress dialogs
│
├── old_refactored/              ← Original scripts kept for reference
│   ├── email-automation-classify.py   (scanner, Excel/Pickle based)
│   ├── email-automation-archive.py    (archiver with fuzzywuzzy suggestions)
│   ├── email-automation-save.py       (save to active Explorer window)
│   ├── utils.py                       (shared helpers)
│   ├── window.py                      (Explorer window detection)
│   ├── outlook_macros.vb              (VBA macros for Outlook)
│   └── refactor.md                    (original design prompt, in Spanish)
│
├── main_archive.py              ← Stream Deck entry: Archive Email
├── main_scan.py                 ← Stream Deck entry: Scan Archive
├── main_ui.py                   ← Full launcher (both buttons)
├── launch_archive.bat           ← Runs pythonw main_archive.py (no console)
├── launch_scan.bat              ← Runs pythonw main_scan.py (no console)
└── requirements.txt
```

---

## Setup

### Prerequisites

- Windows 10/11
- Python 3.10+ (with a virtual environment recommended)
- Microsoft Outlook Desktop installed and configured
- OneDrive synced locally

### Install dependencies

```powershell
cd email-archiver
pip install -r requirements.txt
```

Dependencies:

| Package | Purpose |
|---|---|
| `pyyaml` | Config file loading |
| `pywin32` | Outlook COM automation (`win32com`) |
| `extract-msg` | Read `.msg` files without Outlook (scanner) |
| `rapidfuzz` | Fuzzy folder-name matching (suggestion boost) |
| `psutil` | Detect if Outlook.exe is running |

### Configure

Copy the example config and set your archive root path:

```powershell
copy config\config.example.yaml config\config.yaml
```

Then edit `config/config.yaml` and set the archive root:

```yaml
archive:
  root_path: "C:/Users/YourName/OneDrive/Archive"
```

All other defaults are sensible out of the box.

### First scan

The first scan reads every `.msg` file in your archive. If your files are stored as OneDrive cloud placeholders ("Files On-Demand"), each file must be downloaded as it is opened — this makes the first scan slow (minutes to hours depending on archive size and connection speed). Subsequent scans skip unchanged files and complete in seconds.

To speed up the first scan, right-click your archive root in Windows Explorer → **Always keep on this device** to force OneDrive to sync everything locally first.

```powershell
python main_scan.py
# or headless:
python main_scan.py --no-ui
```

---

## Usage

### Stream Deck buttons

Point each button to the corresponding `.bat` file:

| Button | File | Action |
|---|---|---|
| Scan Archive | `launch_scan.bat` | Opens scan progress window |
| Archive Email | `launch_archive.bat` | Opens archive suggestion dialog |

The `.bat` files use `pythonw` so no console window flashes on screen.

### Command line

```powershell
# Open the archive dialog (reads selected Outlook email)
python main_archive.py

# Open the scan window
python main_scan.py

# Scan without any UI (prints progress to stdout)
python main_scan.py --no-ui

# Full launcher with both buttons + DB stats
python main_ui.py
```

---

## How the suggestion engine works

When you trigger "Archive Email", the app runs a two-stage ranking pipeline:

**Stage 1 — FTS5 full-text search (70% of final score)**

SQLite's built-in FTS5 engine searches the entire index of ~18,000 emails using BM25 relevance scoring. The query is built from the incoming email's subject + sender + recipients, tokenised and joined with OR so partial matches still contribute. Results are aggregated per folder (sum of per-email BM25 scores) and normalised to [0, 1].

BM25 weights: subject ×10, sender ×3, recipients ×3, body preview ×1.

**Stage 2 — Folder name boost (30% of final score)**

`rapidfuzz.token_set_ratio` compares the email subject against the folder's leaf directory name. A folder called `Project Alpha` that matches the subject "RE: Project Alpha – Budget Q4" gets a high boost. This handles the common case where a new email in a project thread has no prior emails in the DB yet.

**Final score = 0.70 × FTS_score + 0.30 × folder_name_score**

Up to 3 suggestions are shown, ranked by final score, each displaying the folder path, match percentage, number of similar past emails, and a sample subject from that folder.

---

## Database schema

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
    indexed_at   TEXT NOT NULL DEFAULT (datetime('now'))
);

-- FTS5 full-text index (kept in sync via triggers)
CREATE VIRTUAL TABLE emails_fts USING fts5(
    subject, sender, recipients, body_preview,
    content = emails, content_rowid = id,
    tokenize = 'unicode61 remove_diacritics 1'
);

-- Aggregated folder stats (used by suggestion engine)
CREATE TABLE folders (
    id           INTEGER PRIMARY KEY AUTOINCREMENT,
    folder_path  TEXT UNIQUE NOT NULL,
    email_count  INTEGER NOT NULL DEFAULT 0,
    last_updated TEXT NOT NULL DEFAULT (datetime('now'))
);
```

The FTS5 index is automatically kept in sync with the `emails` table via `AFTER INSERT`, `AFTER UPDATE`, and `AFTER DELETE` triggers — no manual maintenance needed.

WAL journal mode is enabled so the archive command can read the DB while a scan is running in another process without blocking.

---

## Incremental scanning behaviour

| Scenario | What happens |
|---|---|
| File unchanged | `mtime` matches DB → skipped instantly (no file open) |
| New file | Not in DB → parsed and inserted |
| File modified | `mtime` differs → re-parsed and updated |
| File deleted | After a complete scan, `DELETE WHERE file_path NOT IN (all found paths)` purges stale entries |
| Scan cancelled | Purge step is skipped — safe, no phantom deletions |

---

## What changed from the old version

| | Old (`old_refactored/`) | New |
|---|---|---|
| **Database** | Excel + Pickle (Pandas) | SQLite + FTS5 |
| **Startup time** | 3–8 s (Excel load) | < 0.5 s |
| **Matching** | `fuzzywuzzy` on subject only | BM25 full-text + folder name fuzzy boost |
| **Incremental scan** | No — re-read all files every time | Yes — `mtime` check, skips unchanged files |
| **Architecture** | 3 flat scripts + shared utils | Package with 6 separated modules |
| **UI** | Blocking `window.mainloop()` per dialog | Background thread, non-blocking |
| **Attachments** | `NNN - filename.ext` | `NNN_NN_filename.ext` |
| **Exchange resolution** | Partial | Full `GetExchangeUser()` SMTP fallback |
| **Logging** | `print()` statements | Structured `logging` to file + console |
| **Config** | Hardcoded `.txt` params files | `config/config.yaml` |

---

## Notes on the old code (`old_refactored/`)

The original scripts are preserved exactly as they were and are not imported or used by the new application. They serve as a reference for:

- The original matching logic (`fuzz.token_set_ratio` with subject + sender + recipient scoring and `Folder_Name_Score`)
- The `get_first_explorer_folder_path()` approach for detecting the active Windows Explorer window (used in `email-automation-save.py` to archive to the currently open folder — not replicated in the new version but available if needed)
- The VBA macros in `outlook_macros.vb` for any Outlook-side automation
- The original design brief in `refactor.md`

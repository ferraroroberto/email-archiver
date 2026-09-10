# Email Archiver

A modular Python application for automatically indexing and archiving Outlook emails into a structured OneDrive folder system.

---

## Background

Email archiving used to be a fully manual process: select an email in Outlook, navigate to the right project folder in Windows Explorer, drag, save attachments separately, rename everything consistently. With a deeply nested OneDrive archive of nearly 18,000 emails across hundreds of project folders, this was taking significant time every day.

A first version of this automation was built around 2021–2022 using a different approach: an Excel spreadsheet as the database (cached as a Pickle file for speed), `fuzzywuzzy` for fuzzy subject matching, and three separate scripts — one to classify/index (`email-automation-classify.py`), one to archive with AI suggestion (`email-automation-archive.py`), and one to save to the folder currently open in Windows Explorer (`email-automation-save.py`).

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
Connects to the running Outlook instance, reads the currently selected email, queries the database to find the most relevant project folders, and presents ranked suggestions. You click a folder, confirm, and the email is saved as a `.msg` file plus all attachments — all consistently numbered — in under two seconds. The window closes immediately after saving (no confirmation popup); the log file records what was saved.

When none of the suggestions is right there are two escapes, both in the dialog's bottom bar:

| Button | What it does |
|---|---|
| **Browse folder…** | Opens the native folder picker, starting at the first configured archive root |
| **Explorer folder** | Archives straight into the folder shown in the **foremost open File Explorer window** — the one you looked at last. No picker, no typing. |

`Explorer folder` is for the case where the right destination is already open on screen. It only considers real filesystem folders (virtual shell locations like *This PC* or *Quick Access* are skipped), it is **not** limited to the configured archive roots, and if no Explorer window is showing a real folder it says so rather than guessing.

---

## File naming convention

When an email is archived in a folder, files are named `NNN - <name>`, using ` - ` (space-dash-space) as the separator:

```
023 - Project_Alpha_meeting_notes.msg
023 - invoice.pdf
023 - signed_contract.docx
```

- `NNN` is a zero-padded 3-digit sequence number, derived from the highest existing prefix in that destination folder. It groups a bundle — an email and its attachments always share one `NNN` — and increments per folder.
- The email gets `NNN - sanitized_subject.msg`
- Each real attachment gets `NNN - original_filename.ext`
- Existing archived files are never renamed — this only affects what gets written going forward
- Embedded images (inline in HTML body) are skipped automatically; real attachments (e.g. PDFs) are saved even when the client sets a ContentId
- The filename is dynamically shortened with a trailing `...` if the destination folder is deep enough that the full path would otherwise exceed Windows' 260-char `MAX_PATH` limit (e.g. `042 - Long_subject_starts_here....msg`)

### Optional date prefix

The email's sent date can be prefixed to every name in the bundle, for folders shared with a document archive that files everything as `YYYY-MM-DD - <slug>`:

```
2026-03-14 - 023 - Project_Alpha_meeting_notes.msg
2026-03-14 - 023 - invoice.pdf
2026-03-14 - 023 - signed_contract.docx
```

`naming.date_prefix` in `config/config.yaml` controls the default form:

```yaml
naming:
  date_prefix: auto   # default; true → always dated, false → never dated
```

- **`auto` (default)** — infer the form per destination folder from what it already holds: majority of existing numbered files wins (dated vs. undated), falling back to the undated form when the folder is empty, has no numbered files, or ties. This is what makes an unattended batch run (see [Batch mode](#batch-mode-headless), which has no checkbox) file into a mixed archive correctly without any per-mail configuration — a folder that already files everything as `YYYY-MM-DD - NNN - <name>` (for instance a tree shared with a dated document archive) keeps getting the dated form, the rest keep the undated form.
- **`true` / `false`** — force that form for every folder, ignoring its contents.
- **Per archive** — the **Date prefix (YYYY-MM-DD)** checkbox in the archive dialog's bottom bar starts pre-ticked to whatever the highlighted suggestion's own folder would get (inferred when the config is `auto`, the fixed config value otherwise), and can still be flipped by hand before clicking `Archive`, `Browse folder…` or `Explorer folder`; nothing is written to the config either way.

Precedence when a decision is not left to inference: an explicit per-archive choice (the ticked/unticked checkbox, or a batch `apply` decision's `date_prefix`) always wins over the config, `auto` inference is next, and the config's fixed boolean is the last resort.

- `YYYY-MM-DD` is the email's **sent** date (`SentOn`, falling back to `ReceivedTime`), normalized to local time — not the date it was archived. Attachments inherit the parent email's date, so a bundle stays contiguous.
- `NNN` keeps exactly the same meaning and derivation with the prefix on
- If the email's sent date cannot be resolved, that archive falls back to the undated form and logs why, rather than inventing a placeholder date
- Sequence allocation recognizes **both** the `NNN - ` and the `YYYY-MM-DD - NNN - ` forms no matter which one is being written, so a folder holding a mix of both still picks the correct next number and never collides — the toggle is safe to flip at any time, and inference reads that same single folder listing
- Path shortening accounts for the 13 extra characters only when the prefix is actually applied

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
│   └── archiver.log             ← Log file (plain FileHandler, not rotated)
│
├── email_archiver/              ← Main package
│   ├── config.py                ← YAML loader, path resolution, logging setup
│   ├── text.py                  ← Shared subject + Message-ID normalisation
│   ├── paths.py                 ← Archive-root guard shared by revert and renumber
│   ├── batch.py                 ← Headless plan/apply/revert/renumber orchestration (no COM, no tkinter)
│   ├── renumber.py              ← Re-sequence a folder into sent-date order
│   ├── explorer.py              ← Foremost open Explorer window → folder path
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
│   │   └── suggester.py        ← Three-stage folder ranking (SuggestionEngine)
│   │
│   └── ui/
│       ├── app.py              ← ArchiveDialog, ScanWindow, LauncherApp
│       └── dialogs.py          ← Native folder-picker wrapper
│
├── tests/                       ← pytest suite (filename fitting, sequencing, date-prefix toggle, Explorer picker, is_running regression, Message-ID, batch verbs)
├── docs/
│   └── architecture.mmd         ← Hand-authored Mermaid diagram of internal structure
│
├── main_archive.py              ← Stream Deck entry: Archive Email
├── main_scan.py                 ← Stream Deck entry: Scan Archive
├── main_ui.py                   ← Full launcher (both buttons)
├── main_batch.py                ← Headless entry: plan / apply / revert / renumber (JSON)
├── launch_archive.bat           ← Runs pythonw main_archive.py (no console)
├── launch_scan.bat              ← Runs pythonw main_scan.py (no console)
└── requirements.txt
```

---

## Setup

### Prerequisites

- Windows 10/11
- Python 3.10+
- Microsoft Outlook Desktop installed and configured
- OneDrive synced locally

### Install dependencies

Create the virtual environment and install dependencies:

```powershell
cd email-archiver
python -m venv .venv
& .\.venv\Scripts\python.exe -m pip install -r requirements.txt
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

Then edit `config/config.yaml` and set one or more archive roots:

```yaml
archive:
  # List of absolute paths to your email archive roots (OneDrive local sync folders)
  root_paths:
    - "C:/Users/YourName/OneDrive/Documentos/"
    - "C:/Users/YourName/OneDrive/Archive/"
```

All other defaults are sensible out of the box. The one knob worth knowing about is `naming.date_prefix` (default `auto`) — see [File naming convention](#file-naming-convention).

### First scan

The first scan reads every `.msg` file in your archive. If your files are stored as OneDrive cloud placeholders ("Files On-Demand"), each file must be downloaded as it is opened — this makes the first scan slow (minutes to hours depending on archive size and connection speed). Subsequent scans skip unchanged files and complete in seconds.

To speed up the first scan, right-click your archive root in Windows Explorer → **Always keep on this device** to force OneDrive to sync everything locally first.

```powershell
& .\.venv\Scripts\python.exe main_scan.py
# or headless:
& .\.venv\Scripts\python.exe main_scan.py --no-ui
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
& .\.venv\Scripts\python.exe main_archive.py

# Open the scan window
& .\.venv\Scripts\python.exe main_scan.py

# Scan without any UI (prints progress to stdout)
& .\.venv\Scripts\python.exe main_scan.py --no-ui

# Full launcher with both buttons + DB stats
& .\.venv\Scripts\python.exe main_ui.py
```

---

## Verification

The project ships a pytest suite (`tests/`) covering the filename fitter, sequencing, the date-prefix toggle, the Explorer picker, the `is_running()` regression, the Message-ID column and its migration, renumbering (ordering, both name forms, split numbers, the dry run, the index), and every batch verb end to end against a fake Outlook client. Run it before declaring any change done:

```powershell
& .\.venv\Scripts\python.exe -m pytest tests/
```

`pytest` is a dev-only dependency — see `requirements.txt`.

---

## Batch mode (headless)

`main_batch.py` is the archiver's headless face: four verbs — three that file the **whole Inbox** in one run instead of one selected mail at a time, plus one that repairs a folder's numbering — each printing exactly one JSON document on stdout. It exists so another local app can drive the archiver as a subprocess — the archiver stays the sole owner of Outlook COM, the suggestion engine and the naming rules, and the caller only decides *which folder* each mail goes to.

```powershell
& .\.venv\Scripts\python.exe main_batch.py plan --candidates 5 > plan.json
& .\.venv\Scripts\python.exe main_batch.py apply --decisions decisions.json
& .\.venv\Scripts\python.exe main_batch.py revert --items revert.json
& .\.venv\Scripts\python.exe main_batch.py renumber --folder "<a folder>" --dry-run
```

| Verb | Reads | Does |
|---|---|---|
| `plan` | nothing | Starts Outlook if it is closed, enumerates the Inbox, and returns every mail with its metadata and the top `--candidates` folder suggestions (default 10), each carrying its own `date_prefix`. Read-only. |
| `apply` | `[{message_id, folder_path, date_prefix}]` | Archives each mail into `folder_path`, then moves it to the Outlook `Archive` folder and tags it with the category. A mail already in the index is finished rather than filed again — see [Retrying a mail that was written but never moved](#retrying-a-mail-that-was-written-but-never-moved). |
| `revert` | `[{message_id, files}]` | Deletes exactly the listed files, removes their index rows, removes the category and moves the mail back to the Inbox. |
| `renumber` | nothing | Re-sequences one folder into sent-date order and prints the old → new map. Touches no Outlook and no COM — see [Renumbering a folder](#renumbering-a-folder). |

An `apply` result can be handed straight back to `revert` — the object with its `results` list is accepted as-is, no reshaping needed.

### Identity: the Message-ID, not the EntryID

Outlook rewrites a mail's `EntryID` when it is moved between folders, which is exactly what `apply` does to every mail it files. So the identity that ties the three verbs together is the **Internet Message-ID** (MAPI `PR_INTERNET_MESSAGE_ID`), stored without its angle brackets. The scanner records it for every `.msg` it indexes, which is what lets `plan` mark a mail as `already_archived` instead of offering to file it a second time. A mail carrying no Message-ID (rare — drafts, some system mail) is reported under `skipped` with `reason: "no_message_id"` and left alone, never guessed at.

### Exit codes and error handling

| Exit | Meaning |
|---|---|
| `0` | The run completed and stdout carries its document. Individual mails may still have failed — each result has its own `error` with a `code`. One failing mail never aborts the run. |
| `2` | The run could not start, or `renumber` could not finish its one folder. stdout carries `{"error": {"code", "message"}}` instead of results; the code is one of `config_missing` (no config, **or** one that could not be loaded — a malformed `config.yaml` lands here too, with the parser error in `message`), `bad_input`, `outlook_unavailable`, `com_unavailable`. A folder outside `archive.root_paths` is a `bad_input`, and so is a renumber that stopped part-way: no map is printed, because a map the disk may not match is worse than none. |

Per-mail `error.code` values: `bad_decision` (the entry had no `message_id` or `folder_path`), `not_in_inbox`, `not_in_archive_folder`, `archive_failed`, `move_failed`, `category_failed`. The last two are deliberately distinct: after a `move_failed` the mail is still in the Inbox, after a `category_failed` it is already filed and only *looks* untouched in Outlook.

`apply` fills in a result's `files` **before** it moves the mail, so a mail that was written to disk but failed to move is still fully revertible — that is why a result can carry `ok: false` and a non-empty `files` at the same time.

Every document carries `schema_version`, so a consumer can refuse a shape it does not understand rather than read a field that silently moved.

### Retrying a mail that was written but never moved

A mail can end up with its files on disk and its place in the Inbox kept: `SaveAs` leaves the in-memory Outlook item flagged as modified often enough that the `Move` right after it is refused with *"the operation cannot be performed because the message has been changed"* (MAPI `0x80040109`). That mail is reported `move_failed`, and every later `plan` reports it `already_archived` — its `.msg` is in the index — so without a retry path it could never leave the Inbox.

Two things make it retryable:

- **`apply` moves a reference it re-acquires from the store by EntryID**, not the one it just wrote to disk, and on a `0x80040109` from that reference too it saves it once and retries the move. Which of the three paths finished the move is logged and reported per result as `move_via`: `refetched` (the re-acquired reference was enough), `saved_retry` (it took a `Save()` and one more attempt) or `original` (the re-acquire failed, so the original reference was moved). A move that fails for any *other* reason is still a plain `move_failed`, never retried.
- **`apply` accepts a decision for a mail `plan` reported `already_archived`** and finishes it: it writes nothing, moves the mail to the `Archive` folder, tags it, and reports the existing file as `files` with `reused: true` and an empty `sequence_number` (none was allocated). This is what a consumer offers as *retry* on such a mail. The one exception is an index row whose file is gone — no longer proof of anything, so that mail is archived for real instead.

One case is beyond both defences and is called out by name in the result: when the running Outlook process itself is holding the item, every write to it is refused for the life of that process — a re-acquired reference and a `Save()` are refused exactly like the original. The mail keeps its `move_failed`, and the message says what to do: **restart Outlook, then apply the same decision again**. It then completes on the first attempt, and because the files are already on disk it writes nothing.

`plan` also states `in_inbox` on every mail it reports. It is always `true` — `plan` enumerates the Inbox — but it is the fact that separates an `already_archived` mail *still sitting in the Inbox* from one that is properly filed and gone, so a consumer can key a retry offer off the document rather than off an assumption.

### Renumbering a folder

The `NNN` prefix is meant to be browsable: opening a project folder and reading down the list should be reading the thread in the order it happened. Two ordinary batch operations break that, and neither is a bug in how a single mail is archived — `revert` deletes a bundle and leaves a hole, and `apply` always allocates `max + 1`, so a mail filed into a folder after a correction takes the highest number even when it is older than everything already there.

`renumber` is the repair:

```powershell
& .\.venv\Scripts\python.exe main_batch.py renumber --folder "<a folder>" --dry-run
& .\.venv\Scripts\python.exe main_batch.py renumber --folder "<a folder>"
```

`--dry-run` computes and prints exactly the same map without renaming anything, so a run is always previewable. The same repair can be chained onto the two verbs that cause the problem — `apply --renumber` renumbers every destination folder the run wrote into, `revert --renumber` closes the gap in every folder the run deleted from — once per folder, after the verb's own work. Without the flag both verbs behave exactly as before and neither reports anything new.

**The ordering rule**, in one place:

- **Sent date decides.** Bundles are ordered by the mail's sent date, taken from the index and read out of the `.msg` itself when the index has no row for it (a file archived since the last scan). A bundle whose date cannot be established at all keeps its current position rather than being swept to one end.
- **The unit is a bundle** — a sequence number and every file carrying it. An email never parts company with its attachments. A number that carries no `.msg` at all (an attachment whose mail was deleted by hand, a document filed into the sequence) still holds a slot, because leaving it behind would park it on a number that now belongs to a different mail.
- **Numbers are contiguous from the folder's lowest existing one**, not from `001`. A folder that starts at `079` because it continues another folder's sequence keeps starting at `079`.
- **Each file keeps its own name form** — `NNN - ` stays undated, `YYYY-MM-DD - NNN - ` keeps its date. A folder holding a mix keeps the mix.
- **Two mails that ended up on one number are split by date**, and their attachments are matched to the right one by asking each `.msg` what it carries. An attachment that still cannot be placed follows the earlier mail and the map says so, as `attachments_placed`.
- **Files with no sequence prefix are left alone** and reported under `skipped`.
- **A folder outside `archive.root_paths` is refused**, untouched.

**The consumer must heal its stored paths from the map.** A renumber renames files another app may be holding paths to; the archiver updates its own index and nothing else. The map is reported under `renumbered`, keyed by folder, in all three verbs that produce one:

```json
"renumbered": {
  "<folder>": [
    {
      "from": "<folder>\\003 - subject.msg",
      "to":   "<folder>\\002 - subject.msg",
      "message_id": "abc123@mail.example",
      "attachments": [["<folder>\\003 - report.pdf", "<folder>\\002 - report.pdf"]]
    }
  ]
}
```

Only bundles that actually changed appear. `from`/`to` are `null` for a bundle that has no `.msg` (its files are all in `attachments`), so a consumer healing `.msg` paths simply has nothing to do for it.

**This includes the same document's own `files` lists.** An `apply --renumber` result reports each mail's `files` under the names it was *written* with, and the renumber that ran afterwards may have moved some of them — so hand-an-`apply`-result-straight-to-`revert` needs the map applied to those paths first when the flag was used. Without `--renumber` the `files` are final, as they always were. A folder that could **not** be renumbered is never an empty map in there — an empty map means "already in order" — it is listed under `renumber_refused` with its reason. The key is additive: `schema_version` is unchanged, and a consumer that does not know about it reads the rest of the document exactly as before.

### Safety

- `renumber` renames **only** inside the folder it is given, and only when that folder resolves inside `archive.root_paths`. It never reads or writes file contents, lists the folder once, and renames only the files whose number actually changes.
- Renames run in two phases through a same-length `~XX` placeholder, because closing a gap gives a file the name of the file next to it. The placeholder is the same length as the number it replaces, so a path that fits Windows' `MAX_PATH` today still fits mid-rename. A run that stops part-way logs exactly which files are still under a placeholder.
- The index follows the same two phases: `emails.file_path` is UNIQUE, and two mails on one thread share a subject, so a reorder routinely hands one of them the exact path the other still holds. A row left on a name a rename is about to take, whose own file is gone, is dropped and counted separately as `index_rows_dropped` — "the index followed the renames" and "the index was also wrong" are two different facts.
- `revert` deletes **only** the files it is given, and only those that resolve inside `archive.root_paths`. Anything else is refused per file with a reason (`outside_archive_roots`, `unresolvable_path`) and left on disk.
- A file already gone is reported as `missing`, not as an error — running a revert twice is not a failure to explain.
- Deleting a `.msg` also removes its row from the index in the same step (`index_rows_removed` in the result), so `plan` offers the mail again right away instead of waiting for the next full scan to notice the file is gone. A file whose delete failed keeps its row — the file is still there, so the row is still correct.
- Outlook closed at `plan` time is **started** (the registered `outlook.exe`, visible, exactly as your own shortcut would) and waited for, bounded by `--start-timeout` (60 s by default). A still-unreachable Outlook is a loud exit 2, never an empty Inbox nobody read.
- Batch mode is meant to be spawned with a timeout. A COM modal — the address-book security prompt, a profile chooser — then blocks *this* process, which the caller can kill, and never the caller.
- The `date_prefix` flag on an `apply` decision is per mail and comes from the caller. Batch mode deliberately does **not** re-read the global `naming.date_prefix` toggle itself for `apply` — the right form depends on the destination folder, so each `plan` candidate already carries the `date_prefix` that folder would get under the configured mode (inferred per folder when `naming.date_prefix: auto`, the fixed config value otherwise); the caller normally just hands that value straight back on the decision for the folder it picked.

### Configuration

```yaml
outlook:
  archive_folder: "Archive"           # created under the mailbox root if missing
  category: "Archived by task-os"     # stamped by apply, removed by revert
```

Both keys are optional and fall back to the values above. The folder is matched case-insensitively so a mailbox that already has one is used rather than duplicated.

---

## How the suggestion engine works

When you trigger "Archive Email", the app runs a three-stage ranking pipeline:

**Stage 1 — FTS5 full-text search (45% of final score)**

SQLite's built-in FTS5 engine searches the entire index of ~18,000 emails using BM25 relevance scoring. The query is built from the incoming email's subject + sender + recipients, tokenised and joined with OR so partial matches still contribute. Results are aggregated per folder (sum of per-email BM25 scores) and normalised to [0, 1].

BM25 weights: subject ×10, sender ×3, recipients ×3, body preview ×1.

**Stage 2 — Subject thread score (30% of final score)**

`rapidfuzz.token_set_ratio` compares the incoming email's subject against the subjects of recent emails already stored in each candidate folder (up to 5 samples). A high score means the same conversation thread is already in that folder — the strongest signal for routing emails in an ongoing thread. Re:/Fwd: prefixes are handled naturally by token_set_ratio.

**Stage 3 — Folder name boost (25% of final score)**

`rapidfuzz.token_set_ratio` compares the email subject against the folder's leaf directory name. A folder called `Project Alpha` that matches the subject "RE: Project Alpha – Budget Q4" gets a high boost. This handles the common case where a new project thread has no prior emails in the DB yet.

**Final score = 0.45 × FTS_score + 0.30 × subject_thread_score + 0.25 × folder_name_score**

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

### The follow-up flag (`flag_status`)

`flag_status` records Outlook's follow-up flag, read from the archived `.msg` as MAPI `PidTagFlagStatus` (`0x1090`): **2** = flagged for follow-up, **1** = flagged then completed, **0** = read, and not flagged. It exists so other tools can turn a flagged mail into a task — task-os polls it read-only and raises one Inbox task per flagged message.

Two things about it are easy to get wrong:

- **Flag first, then archive.** The flag is written into the `.msg` at archive time, by the same `SaveAs` call that creates the file. The archived file is a snapshot: flagging a message in Outlook *after* it has been archived changes nothing on disk, and no re-scan will see it.
- **`NULL` is not `0`.** A row indexed before this column existed reads `NULL`, meaning *the property was never read* — a different fact from `0`, *read, and not flagged*. Existing rows keep `NULL` until their file changes and is re-indexed. That is deliberate and costs nothing: Outlook does not preserve the flag through filing, so an already-archived backlog carries no flags to find (a random sample of 600 of ~18k archived files found none).

The column is added to an existing database automatically on the next run — `init_db` ALTERs in any column the `emails` table is missing, because its `CREATE TABLE IF NOT EXISTS` is a no-op once the table exists.

### The Internet Message-ID (`message_id`)

`message_id` records the mail's `Message-ID` header (MAPI `PR_INTERNET_MESSAGE_ID`, `0x1035001F`) read back out of the archived `.msg`, stored **without** its angle brackets so both sides of the app spell it the same way. It exists for [batch mode](#batch-mode-headless): it is the only identity that survives a mail being moved between Outlook folders, so it is what lets `plan` recognise a mail that is already filed and what pairs a `revert` with the files written for it.

- **`NULL` is not `""`.** A row indexed before this column existed reads `NULL`, *the header was never read*; a `.msg` that genuinely carries no `Message-ID` reads `""`. Neither ever matches a lookup — two mails with no Message-ID are not the same mail.
- Added to an existing database the same way `flag_status` was, plus an index on the column created **after** the ALTER — declaring it alongside the `CREATE TABLE` would run it against a column an existing database does not have yet and take the whole scan down with it.
- **The existing backlog is not backfilled.** Scanning is incremental on `mtime`, so a `.msg` indexed before this column existed is skipped and keeps its `NULL` — its Message-ID is in the file, but nothing has read it. Everything archived from now on is indexed with its id on the next scan (a new file is never skipped), so `plan`'s already-archived detection is complete for the batch workflow's own loop and blind only to mail filed before it. To cover the backlog too, force a full re-read by deleting `data/emails.db` and re-scanning — the same cost as a first scan.

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

| | Old version (pre-2026) | New |
|---|---|---|
| **Database** | Excel + Pickle (Pandas) | SQLite + FTS5 |
| **Startup time** | 3–8 s (Excel load) | < 0.5 s |
| **Matching** | `fuzzywuzzy` on subject only | BM25 full-text + folder name fuzzy boost |
| **Incremental scan** | No — re-read all files every time | Yes — `mtime` check, skips unchanged files |
| **Architecture** | 3 flat scripts + shared utils | Package with 6 separated modules |
| **UI** | Blocking `window.mainloop()` per dialog | Background thread, non-blocking |
| **Attachments** | (varies) | `NNN - filename.ext` (email: `NNN - subject.msg`), optional `YYYY-MM-DD - ` prefix |
| **Save to open Explorer folder** | Separate script (`email-automation-save.py`) | `Explorer folder` button in the archive dialog |
| **Exchange resolution** | Partial | Full `GetExchangeUser()` SMTP fallback |
| **Logging** | `print()` statements | Structured `logging` to file + console |
| **Config** | Hardcoded `.txt` params files | `config/config.yaml` |


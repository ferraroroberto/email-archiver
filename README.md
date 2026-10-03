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
- **Per archive** — the **Date prefix (YYYY-MM-DD)** checkbox in the archive dialog's bottom bar starts pre-ticked to whatever the highlighted suggestion's own folder would get (inferred when the config is `auto`, the fixed config value otherwise), and can still be flipped by hand before clicking `Archive`, `Browse folder…` or `Explorer folder`; nothing is written to the config either way. Until you flip it, that tick is only a preview: `Browse folder…` and `Explorer folder` resolve the form for the folder you actually pick, never for the card you last hovered.

Precedence when a decision is not left to inference: an explicit per-archive choice (the checkbox once flipped by hand, or a batch `apply` decision's `date_prefix`) always wins over the config, `auto` inference is next, and the config's fixed boolean is the last resort.

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
│   ├── draft.py                 ← Headless draft: spec validation, BCC-self, the draft document (no COM)
│   ├── send.py                  ← Headless read-back + guarded send: fingerprint, approval check (no COM)
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
│   │   ├── client.py           ← Outlook COM isolation (OutlookClient): read path + batch surface
│   │   ├── drafts.py           ← Draft surface OutlookClient inherits (DraftSurface)
│   │   ├── process.py          ← Outlook process: is_running / ensure_running
│   │   ├── mapi.py             ← Shared constants and COM-free helpers
│   │   ├── stores.py           ← Which Outlook store a mailbox lives in (by address)
│   │   └── sending.py          ← The one COM Send call (send verb only)
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
├── main_batch.py                ← Headless entry: plan / apply / revert / renumber / draft / read / send / mailboxes (JSON)
├── launch_archive.bat           ← Runs pythonw main_archive.py (no console)
├── launch_scan.bat              ← Runs pythonw main_scan.py (no console)
├── run-scan-nightly.bat         ← Scheduled job: headless scan in the foreground, propagates the exit code
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

### Nightly scan

`run-scan-nightly.bat` is the scheduled-job launcher. It `cd`s to the repo, sets `PYTHONUTF8=1` and `PYTHONUNBUFFERED=1`, runs `main_scan.py --no-ui` **in the foreground** with the repo `.venv` Python, and exits with the scan's exit code — no `start`, no `pythonw`, no `pause`. When stdout is not a terminal (a captured job log), progress is printed one line per tick instead of being rewritten in place with `\r`.

It runs in app-launcher as the job `email-archiver-scan-nightly` (schedule `none`, alert on failure), chained from both `on_success` and `on_failure` of the nightly fleet backup job, so the index is refreshed right after the backup whatever the backup's outcome. Chain hops run one at a time, which is what keeps the nightly scan from overlapping anything — there is no scan lock.

| Exit code | Meaning |
|---|---|
| `0` | Scan complete. Per-file read errors (for example cloud-only placeholders) are counted in the `Done — …` summary line, not a failure |
| `1` | Unexpected crash (uncaught exception), or the launcher could not find the repo or its `.venv` |
| `2` | Config missing or unloadable, or no archive roots configured |
| `3` | A configured archive root does not exist. The roots that are present are still indexed, but the purge is skipped, so index rows under the missing root are left untouched; each missing root is logged by path |

The Stream Deck `launch_scan.bat` is unchanged and still opens the progress window.

---

## Verification

The project ships a pytest suite (`tests/`) covering the filename fitter, sequencing, the date-prefix toggle, the Explorer picker, the `is_running()` regression, the Message-ID column and its migration, renumbering (ordering, both name forms, split numbers, the dry run, the index), the headless scan's exit codes and its missing-root purge guard, and every batch verb (including `draft`, `read` and `send`) end to end against a fake Outlook client. Run it before declaring any change done:

```powershell
& .\.venv\Scripts\python.exe -m pytest tests/
```

`pytest` is a dev-only dependency — see `requirements.txt`.

---

## Batch mode (headless)

`main_batch.py` is the archiver's headless face: eight verbs — three that file the **whole Inbox** in one run instead of one selected mail at a time, one that repairs a folder's numbering, one that opens an unsent draft, two that read a draft back and send it only while it is still what was approved, and one that lists the mailboxes it can act on — each printing exactly one JSON document on stdout. It exists so another local app can drive the archiver as a subprocess — the archiver stays the sole owner of Outlook COM, the suggestion engine and the naming rules, and the caller only decides *which folder* each mail goes to.

```powershell
& .\.venv\Scripts\python.exe main_batch.py plan --candidates 5 > plan.json
& .\.venv\Scripts\python.exe main_batch.py plan --search "<subject words>" --since 2026-09-15 --candidates 0
& .\.venv\Scripts\python.exe main_batch.py apply --decisions decisions.json
& .\.venv\Scripts\python.exe main_batch.py revert --items revert.json
& .\.venv\Scripts\python.exe main_batch.py renumber --folder "<a folder>" --dry-run
& .\.venv\Scripts\python.exe main_batch.py draft --spec spec.json
& .\.venv\Scripts\python.exe main_batch.py read --entry-id <entry_id>
& .\.venv\Scripts\python.exe main_batch.py send --entry-id <entry_id> --expect-hash <sha256>
& .\.venv\Scripts\python.exe main_batch.py mailboxes
```

| Verb | Reads | Does |
|---|---|---|
| `plan` | nothing | Starts Outlook if it is closed, enumerates the Inbox, and returns every mail with its metadata and the top `--candidates` folder suggestions (default 10), each carrying its own `date_prefix`. `--message-id` / `--since` / `--search` / `--ref` narrow it to specific mail — see [Filing specific mail](#filing-specific-mail). Read-only. |
| `apply` | `[{message_id, folder_path, date_prefix}]` | Archives each mail into `folder_path` (which must be inside `archive.root_paths`), then moves it to the Outlook `Archive` folder and tags it with the category. A mail already in the index is finished rather than filed again — see [Retrying a mail that was written but never moved](#retrying-a-mail-that-was-written-but-never-moved). |
| `revert` | `[{message_id, files}]` | Deletes exactly the listed files, removes their index rows, removes the category and moves the mail back to the Inbox. |
| `renumber` | nothing | Re-sequences one folder into sent-date order and prints the old → new map. Touches no Outlook and no COM — see [Renumbering a folder](#renumbering-a-folder). |
| `draft` | `{to, cc, bcc, subject, body_text \| body_html, attachments, ref, display, reply_to, reply_all}` | Creates a filled Outlook draft that blind-copies your own address, saves it to Drafts and opens it for you to review; with `reply_to`, a real threaded reply. **Never sends.** See [Drafting a mail](#drafting-a-mail). |
| `read` | nothing | Reads one unsent draft back as it is stored, with its fingerprint. Writes nothing, shows nothing. See [Sending an approved draft](#sending-an-approved-draft). |
| `send` | `--expect-hash` (and optionally `--expect-part`, `--expect-to`) | Sends that one draft, only if its live fingerprint still matches. Refused with `approval_mismatch` naming what differs otherwise. |
| `mailboxes` | nothing | Lists the mailbox registry with each mailbox's `outlook_status`. Never starts Outlook. See [Mailboxes](#mailboxes). |

An `apply` result can be handed straight back to `revert` — the object with its `results` list is accepted as-is, no reshaping needed.

### Filing specific mail

A caller that already knows which mail it wants to file, and where, doesn't need a ranked plan of the whole Inbox. `plan` takes filters, all optional and combined with AND. With none of them the document is exactly the full-Inbox plan, with no `filters` key.

| Flag | Selects |
|---|---|
| `--message-id <id>` | That mail, looked up by Internet Message-ID (angle brackets optional) instead of enumerating the Inbox. Repeatable. An id found nowhere in the Inbox is listed under `skipped` with `reason: "not_in_inbox"` and its `message_id`. |
| `--since <YYYY-MM-DD[THH:MM]>` | Mail received at or after that local time. Narrowed server-side with `Items.Restrict` (the DASL literal is written in UTC, which is how Outlook compares it), then checked exactly per mail. A store that rejects the filter falls back to walking the Inbox, and the log says which path ran. |
| `--search <text>` | Mail whose subject, sender, recipients or body preview contains the text, case-insensitively. Repeatable: every term must match. |
| `--ref <token>` | Mail whose `X-Archive-Ref` header equals the token, the one `draft` stamps. Headers are read only for mail that passed the cheaper filters. The header survives sending on an IMAP mailbox (verified), but that depends on the account, so a caller keeps a fallback match (`--search` on the subject, `--since` the send time). |
| `--candidates 0` | No folder ranking: each mail comes back with `candidates: []`, for a caller that already knows the destination. |
| `--folder inbox\|sent` | Where to read from; `inbox` is the default. See [Filing from Sent Items](#filing-from-sent-items). Any other value is `bad_input`. |

A filtered document adds `filters: {message_ids, since, search, ref}` (what was applied), and `counts.inbox` counts the Inbox mails that matched. Each mail has the same shape as in a full plan.

#### Filing from Sent Items

A sent mail's self-BCC copy does not always reach the Inbox (a large attachment, a suppressed self-copy), and the sent mail itself, carrying the same `X-Archive-Ref`, sits in Sent Items. A caller that wants that copy asks for it explicitly:

```powershell
& .\.venv\Scripts\python.exe main_batch.py plan --folder sent --ref <token> --since 2026-09-30T11:13 --candidates 0
# decisions.json: [{"message_id": "<from the plan>", "folder_path": "<a folder under an archive root>", "date_prefix": "auto", "source": "sent"}]
```

- `plan --folder sent` reads Sent Items; `--since` compares the **sent** time (`PR_CLIENT_SUBMIT_TIME`), and `--message-id` is looked up in Sent Items (an id not found is skipped with `reason: "not_in_sent"`). The document's `filters` gains `folder: "sent"`, `counts` carries `sent` where an Inbox plan carries `inbox`, and each mail reports `in_inbox: false`. An Inbox plan's document is unchanged.
- An `apply` decision with `"source": "sent"` looks the mail up in Sent Items (`not_in_sent` when it is not there), archives and tags it as usual, and does **not** move it: Sent Items is the record of what went out. The result reports `moved: false` and `move_via: "kept_in_sent"`. Once the scan has indexed the file, `plan` marks the mail `already_archived` and a second apply reuses it (`reused: true`) instead of filing a second copy; until then the mail, still in Sent Items, would be filed again, so a caller records the filing and rescans. A `source` other than `inbox` / `sent` is a per-mail `bad_decision`, never read as the Inbox.
- `revert` does not undo a Sent Items filing's move, because there was none: it deletes the listed files and reports the mail `not_in_archive_folder`.

Two `apply` / `revert` additions serve the same caller:

- **`"date_prefix": "auto"`** on a decision resolves the naming form for its `folder_path` exactly as `plan` resolves a candidate's. Under `naming.date_prefix: auto` that is inferred from what the folder already holds, so an empty or new folder gets the undated form and a folder of dated files gets the dated one. Under a fixed `true` / `false` it is that value. `true` / `false` on a decision behave as they always have.
- **`--category <name>`** tags filed mail with that category instead of `outlook.category` (e.g. `"Archived by life-os"`). The document's `category` reports the one used. Pass the same flag to `revert` so it removes the right one.

```powershell
& .\.venv\Scripts\python.exe main_batch.py plan --ref <token> --candidates 0 > one.json
# decisions.json: [{"message_id": "<from one.json>", "folder_path": "<a folder under an archive root>", "date_prefix": "auto"}]
& .\.venv\Scripts\python.exe main_batch.py apply --decisions decisions.json --category "Archived by life-os" > applied.json
& .\.venv\Scripts\python.exe main_batch.py revert --items applied.json --category "Archived by life-os"
```

### Identity: the Message-ID, not the EntryID

Outlook rewrites a mail's `EntryID` when it is moved between folders, which is exactly what `apply` does to every mail it files. So the identity that ties the three verbs together is the **Internet Message-ID** (MAPI `PR_INTERNET_MESSAGE_ID`), stored without its angle brackets. The scanner records it for every `.msg` it indexes, which is what lets `plan` mark a mail as `already_archived` instead of offering to file it a second time. A mail carrying no Message-ID (rare — drafts, some system mail) is reported under `skipped` with `reason: "no_message_id"` and left alone, never guessed at.

### Exit codes and error handling

| Exit | Meaning |
|---|---|
| `0` | The run completed and stdout carries its document. Individual mails may still have failed — each result has its own `error` with a `code`. One failing mail never aborts the run. |
| `2` | The run could not start, or `renumber` could not finish its one folder. stdout carries `{"error": {"code", "message"}}` instead of results; the code is one of `config_missing` (no config, **or** one that could not be loaded — a malformed `config.yaml` lands here too, with the parser error in `message`), `bad_input`, `outlook_unavailable`, `com_unavailable`, `self_address_unresolved` (`draft` only), `draft_not_found` / `draft_not_editable` (`draft --update`, `read`, `send`), `draft_body_unmarked` (`draft --update` only), `approval_mismatch` / `recipient_unreadable` (`send` only; `approval_mismatch` names the parts under `error.differs`), `mailbox_registry_invalid` (`config/mailboxes.json` is present but malformed), `mailbox_not_in_outlook` (a registry mailbox has no store in the running Outlook — never answered with the default store), `mailbox_unknown` (a `--mailbox` name the registry does not know; checked before Outlook is touched), `account_not_in_outlook` (`draft`: the mailbox's store is there but no Outlook account sends from its address), `sending_account_mismatch` (`draft`: Outlook would not set the mailbox's account, so the draft was discarded unsaved; `send`: the stored draft does not send as the mailbox's account, so nothing was sent). A folder outside `archive.root_paths` is a `bad_input`, and so is a renumber that stopped part-way: no map is printed, because a map the disk may not match is worse than none. |

Per-mail `error.code` values: `bad_decision` (the entry had no `message_id` or `folder_path`, or its `folder_path` does not resolve inside `archive.root_paths`, with a message starting `outside_archive_roots`), `not_in_inbox`, `not_in_archive_folder`, `archive_failed`, `move_failed`, `message_changed`, `category_failed`. The last three are deliberately distinct: after a `move_failed` the mail is still in the Inbox, after a `category_failed` it is already filed and only *looks* untouched in Outlook, and a [`message_changed`](#retrying-a-mail-that-was-written-but-never-moved) is a `move_failed` that an operator clears in seconds — usually by closing an open mail window.

`apply` fills in a result's `files` **before** it moves the mail, so a mail that was written to disk but failed to move is still fully revertible — that is why a result can carry `ok: false` and a non-empty `files` at the same time.

Every document carries `schema_version`, so a consumer can refuse a shape it does not understand rather than read a field that silently moved.

### Retrying a mail that was written but never moved

A mail can end up with its files on disk and its place in the Inbox kept: `SaveAs` leaves the in-memory Outlook item flagged as modified often enough that the `Move` right after it is refused with *"the operation cannot be performed because the message has been changed"* (MAPI `0x80040109`). That mail fails its move with its files already written, and every later `plan` reports it `already_archived` — its `.msg` is in the index — so without a retry path it could never leave the Inbox.

Two things make it retryable:

- **`apply` moves a reference it re-acquires from the store by EntryID**, not the one it just wrote to disk, and on a `0x80040109` from that reference too it saves it once and retries the move. Which of the paths finished the move is logged and reported per result as `move_via`: `refetched` (the re-acquired reference was enough), `saved_retry` (it took a `Save()` and one more attempt), `original` (the re-acquire failed, so the original reference was moved) or `folder_switch` (see below). A move that fails for any *other* reason is still a plain `move_failed`, never retried.
- **`apply` accepts a decision for a mail `plan` reported `already_archived`** and finishes it: it writes nothing, moves the mail to the `Archive` folder, tags it, and reports the existing file as `files` with `reused: true` and an empty `sequence_number` (none was allocated). This is what a consumer offers as *retry* on such a mail. The one exception is an index row whose file is gone — no longer proof of anything, so that mail is archived for real instead.

A third defence covers a mail Outlook's own window is holding (issue #111). A mail that is the Explorer's selected, reading-pane item can refuse every write, a `Move` included, until the Explorer lets go of it — the same refusal, cleared by selecting another folder and then the message again, with no restart. So when the move is still refused after the save **and** the active Explorer is showing that mail's own folder, `apply` switches the Explorer to another folder and straight back, then retries the move once; `move_via` is `folder_switch` and the run logs it at info level. The Explorer is never touched unless the move was refused first, nor when it is showing another folder, and it is always put back on the original folder, even if the switch fails. (`ClearSelection` and `AddToSelection` are not used: they fail in conversation view, Outlook's default.) Your global Outlook settings are never changed.

> **Not yet proven against a live selection hold.** The unit tests drive a fake Outlook, and the switch itself works against a real one, but nobody has yet watched it free a mail that was actually held — conversation view stops a script selecting a mail, so it has to be done by hand. The check: open a scratch folder in Outlook, click one mail so it shows in the reading pane, then try to move it from a script and again after switching folders. If a held mail keeps failing after `apply`'s own switch, it will say so as `message_changed`, with the restart advice.

One case is beyond all of these defences and gets its **own error code**, `message_changed`: when something in the running Outlook is holding the item, every write to it is refused while that lasts — a re-acquired reference and a `Save()` are refused exactly like the original. It is separate from `move_failed` because it is the one move failure that is cleared in seconds with nothing actually wrong, and a consumer offering a retry needs to tell the two apart.

Usually the thing holding the item is **an open mail window** — an Outlook inspector holds its item for the life of the window, so both defences fail by construction. So `apply` asks Outlook whether that mail is in fact open and says which case this is, rather than guessing:

| what the check found | what the message says to do |
|---|---|
| the mail is open in a window | **close that window**, then apply the same decision again |
| no window holds it, and `apply` already switched the Explorer away and back | **restart Outlook**, then apply the same decision again — re-applying before the restart is not expected to help |
| no window holds it, and the Explorer was not on the mail's folder or could not be switched | **select another folder in Outlook, then the message again**, and apply the same decision again; restart Outlook only if that does not clear it |
| the check could not be completed | close any window showing it and re-apply, or select another folder then the message again; restart Outlook only if neither clears it |

The third row is reported as *not determined* and never as "no window": a check that could not run is not evidence the window is closed, and it would point you at a restart you do not need. Either way the files are already on disk and the mail is still in the Inbox, so re-applying the same decision writes nothing.

The no-window case was not root-caused when issue #84 first saw it; the Explorer hold above is the likely cause (issue #111). A live occurrence ruled out an open window, the modified flag `SaveAs` leaves behind (re-applies that wrote nothing were refused on their very first move) and anything transient (three re-applies over 40 seconds, each refused, the `Save()` included) — so `apply` does not retry it on a timer, and once the folder switch has been tried the message says plainly that re-applying before a restart will not help. Restarting Outlook is left to you; `apply` never does it.

**Don't delete a stuck mail by hand from the Inbox.** Its `.msg` is already written and indexed, so it becomes the mail's only copy: uncategorized, with nothing in Outlook pointing at it, and the next `plan` no longer lists the mail at all. If you do, the result's `files` is the list to keep (or to pass to `revert` if the archive copy is not wanted either).

`schema_version` is unchanged at `1`: `message_changed` is a new value for an existing `error.code` key, so a consumer that does not know it still reads the document exactly as before and still sees a mail that failed.

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

### Drafting a mail

`draft` turns a written email into an Outlook draft the user reads and sends by hand:

```json
{
  "to": ["someone@example.org"], "cc": [], "bcc": [],
  "subject": "…",
  "body_text": "…",
  "attachments": ["C:/…/file.pdf"],
  "ref": "opaque-caller-token",
  "display": true
}
```

- **The spec is checked before Outlook starts.** At least one `to`, a non-empty `subject`, exactly one of `body_text` / `body_html`, every attachment an existing file, and no unknown keys (a misspelt `atachments` would otherwise drop the file silently). Anything else is `bad_input` in milliseconds.
- **Your own address is always on the BCC line**, exactly once. The copy that lands back in the Inbox is what gets filed afterwards, so it is taken from `outlook.self_address` when set, else from the default sending account's SMTP address (the account owning the default store, read from `Account.SmtpAddress` — never through `GetExchangeUser`, which can raise the security prompt). When neither resolves the run exits 2 with `self_address_unresolved` and **no draft is created**. With a mailbox registry, the BCC address is the mailbox's own `address` instead (see [Acting on a named mailbox](#acting-on-a-named-mailbox)).
- **Body and signature.** `body_text` is escaped into minimal HTML with its paragraphs and line breaks kept; `body_html` is used as written. With `display: true` (the default) the compose window is opened first — that is when Outlook inserts the account's default signature — and the body is then put at the top, above it. With `display: false` nothing is shown and no signature is added.
- **Saved, never sent.** The finished draft is saved, so it sits in Drafts even if the window is closed. No code path calls `Send()`.
- **`ref`** is stamped as an `X-Archive-Ref` Internet header on the draft, so a caller can recognise its own copy when it arrives. `ref_header` reports `stamped` or `not_stamped`, with `ref_header_reason` saying why (`no ref given`, or the store's refusal). Whether the header survives sending depends on the account, so a caller keeps a fallback match (subject, recipients, sent time).

The document:

```json
{
  "verb": "draft", "schema_version": 1, "generated_at": "…",
  "entry_id": "…", "subject": "…",
  "to": ["…"], "cc": [], "bcc": ["…", "<your address>"],
  "attachments": ["C:/…/file.pdf"],
  "ref": "opaque-caller-token", "ref_header": "stamped", "ref_header_reason": "",
  "displayed": true, "updated": false, "created_at": "2026-09-15T10:00:00+02:00"
}
```

#### Replying to a mail

Add `reply_to` to make the draft a **real Outlook reply** instead of a fresh mail: it carries the thread link (In-Reply-To), the original quoted under the body, a `Re:` subject and the original's recipients, so Outlook threads it under the original.

```json
{
  "reply_to": {"message_id": "<id of an Inbox mail>"},
  "reply_all": false,
  "body_text": "…"
}
```

- **The original** is `{"message_id": …}` (an Inbox mail, found by its Message-ID) or `{"msg_path": "C:/…/saved.msg"}` (an archived `.msg`, absolute, opened with `Namespace.OpenSharedItem` and closed again without saving). Exactly one of the two. `reply_all: true` asks for Reply All (recipients and CC kept); it needs `reply_to`.
- **`to`, `cc` and `subject` become optional** and default to what Outlook computed. Any that you give overrides it. `bcc` (your own address) is unchanged.
- **The same steps as any draft**, in the same order: `X-Archive-Ref` stamped, the compose window displayed first (that is when the signature appears), then your body written **above** the signature and the quote, then saved. The quote is Outlook's and is never rewritten. Without a readable body from Outlook the run stops rather than write a body that would drop the quote.
- **Refused before anything is saved**, exit 2: `reply_source_not_found` (no Inbox mail with that Message-ID, or the `.msg` cannot be opened or is not a mail) and `reply_unavailable` (the original opened but Outlook would not reply to it — for example a `.msg` with no account to reply from; or the original is a mail you sent and its own recipients cannot be read, see the next bullet).
- **A reply to a mail you sent goes to the people you wrote to.** Outlook's `Reply()` answers the original's sender, which on a mail you sent (a saved sent `.msg`, or one from Sent Items) is you. When no `to` is given and Outlook's reply is addressed to your own address (`outlook.self_address`, else the default sending account), the To is taken from the original's To line instead, and the CC from its CC for Reply All only. If those recipients cannot be read the run refuses with `reply_unavailable` rather than leave a draft addressed to you; give `to` to override. Reply All is already addressed correctly by Outlook and is left as it is.
- **The document reports the reply**, in the style of `ref_header`: `reply_to` (what was asked), `replied_to_message_id` (the original's own id), the `to` / `cc` / `subject` the saved draft carries, and `thread_header` — `set` when the saved draft holds the original as its In-Reply-To (read back from the draft, not assumed), else `not_set` with `thread_header_reason`. `recipients_from` says who the reply was addressed from: `sender` (Outlook's own reply), `original_recipients` (the original was sent by you) or `caller` (you gave `to`). A new mail reports `not_a_reply`.
- **`send` is unchanged; `read` gained one field.** The fingerprint's `body` part is the whole stored HTML, so the quoted original is bound by the approval; `read` reports the whole body under `body.html`, the caller's own text under `body.region_text` and what sits below it (the signature and the quoted original) under `body.after_region_html`, so a preview shown for approval can include the quote.
- **Outlook rewrites the HTML of a reply it displays.** With the compose window open, the body is stored the way Outlook's editor re-serialized it: the `<!--/archive-draft-body-->` closing comment is gone and each paragraph is padded with an empty `<o:p></o:p>`. The region is still found (it ends at the `</div>` that closes its own opening tag), `body.region_text` ignores the padding, and `body.after_region_html` is how a caller gets the quote without looking for the comment. A caller must not look for the comment in `body.html` of a reply.
- **`draft --update <entry_id>` on a reply** re-fills only the marked region; the quote below it is left alone. `reply_to` in the update spec is accepted but not re-applied (the draft keeps the thread link it was created with), and a `to` / `cc` / `subject` you leave out stays as the draft has it (reported as `null`, `thread_header: "unchanged"`).
- The real In-Reply-To on the *sent* copy is Outlook's and can only be seen after a send; check it by hand on a mail to yourself.

#### Updating the same draft

```powershell
& .\.venv\Scripts\python.exe main_batch.py draft --update <entry_id> --spec spec.json
```

A caller iterating on one mail ("make it shorter") edits the draft it already made instead of leaving a near-duplicate in Drafts. `<entry_id>` is the `entry_id` the first `draft` returned; the spec is the full new spec, validated exactly as above.

- **Same item, re-filled.** Recipients (your own address still blind-copied exactly once), subject, attachments (every existing one removed, the spec's added) and the `X-Archive-Ref` header are set again, and the item is saved. With `display: true` it is shown again afterwards. **Never sent.**
- **An open compose window on that draft is closed first, saving it.** The window holds its own copy: left open it would still show the old text, and Send pressed there would send that. The update then works on what the window saved.
- **Only the body the tool wrote is replaced.** `draft` wraps the caller's body in `<div id="archive-draft-body">…</div><!--/archive-draft-body-->`; an update swaps that region and keeps everything after it, which is where the signature sits. Text the user typed into the draft by hand inside that region is overwritten, as is any hand edit to the recipients or subject.
- **Refused, and the item left untouched**, each with its own code and exit 2:
  - `draft_not_found` — the EntryID no longer resolves (the draft was deleted, or sent and moved);
  - `draft_not_editable` — the item was sent, or is not in the Drafts folder;
  - `draft_body_unmarked` — no marked body region: a draft created before `--update` existed, or one whose opening `<div id="archive-draft-body">` is gone. (A draft whose closing comment Outlook dropped, which is every displayed reply, still has its region and is updated.) Guessing where the body ends could eat the signature or the user's own text.
- The document is the one above with `updated: true` and `updated_at` in place of `created_at`.

#### Sending an approved draft

```powershell
& .\.venv\Scripts\python.exe main_batch.py read --entry-id <entry_id>
& .\.venv\Scripts\python.exe main_batch.py send --entry-id <entry_id> --expect-hash <sha256> `
    --expect-part body=<sha256> [--expect-part to=<sha256> ...] [--expect-to <address> ...]
```

A caller shows the user a draft, the user approves it, and only then does the caller send. The approval is bound to the item actually in Drafts, not to what the caller wrote, so an edit made by hand in Outlook after the approval stops the send.

- **`read`** reports the draft as stored: `to` / `cc` / `bcc` read from its recipients, `subject`, `body.html` (the whole `HTMLBody`, signature included), `body.region_html` (the marked region `draft` wrote, `null` when the markers are gone) and `body.region_text` (the plain text that region was made from, `null` when it is not exactly what `draft` writes for a `body_text`; the empty `<o:p></o:p>` Outlook pads a reply's paragraphs with is not held against it), `body.after_region_html` (what follows the region: the signature and, on a reply, the quoted original; `null` when the region is gone), and `attachments` with their byte size and sha256. The attachments are saved to a temp folder to be hashed, which is removed afterwards. It also reports `fingerprint`: a sha256 per part (`to`, `cc`, `bcc`, `subject`, `body`, `attachments`) and one `hash` over them. Recipients are compared as a sorted, case-folded set.
- **`send`** closes any window open on the draft, saving it, so nothing typed there is lost and anything typed counts as a change. It then re-reads the item and recomputes the fingerprint **immediately before `Send()`, on the same item it sends**. If the hash differs, or the To line is not exactly the `--expect-to` addresses, the run exits 2 with `approval_mismatch`. `error.differs` names the parts that no longer match the `--expect-part` hashes. Nothing is sent.
- **One item, by EntryID.** No lookup by subject or search, and no recipient is resolved or added. A recipient with no readable address refuses the send as `recipient_unreadable`. A draft that is gone, already sent or outside Drafts is `draft_not_found` / `draft_not_editable`.
- **`draft` has no send path.** The one COM `Send` call lives in `email_archiver/outlook/sending.py`, reached only by this verb, and a test pins that nowhere else calls it.

```json
{
  "verb": "send", "schema_version": 1, "generated_at": "…",
  "entry_id": "…", "subject": "…", "to": ["…"], "cc": [], "bcc": ["you@example.com"],
  "attachments": ["letter.pdf"], "fingerprint": "<sha256>", "sent": true, "sent_at": "…"
}
```

### Safety

- `renumber` renames **only** inside the folder it is given, and only when that folder resolves inside `archive.root_paths`. It never reads or writes file contents, lists the folder once, and renames only the files whose number actually changes.
- Renames run in two phases through a same-length `~XX` placeholder, because closing a gap gives a file the name of the file next to it. The placeholder is the same length as the number it replaces, so a path that fits Windows' `MAX_PATH` today still fits mid-rename. A run that stops part-way logs exactly which files are still under a placeholder.
- The index follows the same two phases: `emails.file_path` is UNIQUE, and two mails on one thread share a subject, so a reorder routinely hands one of them the exact path the other still holds. A row left on a name a rename is about to take, whose own file is gone, is dropped and counted separately as `index_rows_dropped` — "the index followed the renames" and "the index was also wrong" are two different facts.
- `apply` writes **only** inside `archive.root_paths`. A decision whose `folder_path` resolves anywhere else (a sibling folder, a `..` traversal, another drive, a folder whose name merely starts with a root's) is refused as a per-mail `bad_decision` before Outlook is asked anything. Nothing is written, the mail stays in the Inbox, and the next decision is still attempted. A new folder *inside* a root is created as before.
- `revert` deletes **only** the files it is given, and only those that resolve inside `archive.root_paths`. Anything else is refused per file with a reason (`outside_archive_roots`, `unresolvable_path`) and left on disk.
- A file already gone is reported as `missing`, not as an error — running a revert twice is not a failure to explain.
- Deleting a `.msg` also removes its row from the index in the same step (`index_rows_removed` in the result), so `plan` offers the mail again right away instead of waiting for the next full scan to notice the file is gone. A file whose delete failed keeps its row — the file is still there, so the row is still correct.
- Outlook closed at `plan` time is **started** (the registered `outlook.exe`, visible, exactly as your own shortcut would) and waited for, bounded by `--start-timeout` (60 s by default). A still-unreachable Outlook is a loud exit 2, never an empty Inbox nobody read.
- An `outlook.exe` that *this run* started and that never publishes its COM object is **ended again** before the exit-2 document is printed, so an unattended run that times out — a scheduled task has no desktop for Outlook to show a profile or password prompt on — does not leave an invisible Outlook behind holding your mail profile and OST, one more every night. Only the process this run started is ever ended, through the handle it still holds: an Outlook that was already running is attached to and never touched, and nothing is ever ended by image name. The `message` on the exit-2 document says which happened — terminated, already exited, or could not be terminated.
- Batch mode is meant to be spawned with a timeout. A COM modal — the address-book security prompt, a profile chooser — then blocks *this* process, which the caller can kill, and never the caller.
- The `date_prefix` flag on an `apply` decision is per mail and comes from the caller. Batch mode deliberately does **not** re-read the global `naming.date_prefix` toggle itself for `apply` — the right form depends on the destination folder, so each `plan` candidate already carries the `date_prefix` that folder would get under the configured mode (inferred per folder when `naming.date_prefix: auto`, the fixed config value otherwise); the caller normally just hands that value straight back on the decision for the folder it picked.

### Configuration

```yaml
outlook:
  archive_folder: "Archive"           # created under the mailbox root if missing
  category: "Archived by task-os"     # stamped by apply, removed by revert (--category overrides)
  self_address: "you@example.com"     # draft: blind-copied on every draft
```

All three keys are optional. `archive_folder` and `category` fall back to the values above; `self_address` falls back to the default sending account's SMTP address. The folder is matched case-insensitively so a mailbox that already has one is used rather than duplicated. A registry mailbox's own `archive_folder` overrides `outlook.archive_folder` (see [Mailboxes](#mailboxes)).

### Mailboxes

Every Outlook verb acts on **one mailbox**, and every folder it touches (Inbox, Sent, Drafts, the archive folder) is taken from that mailbox's own Outlook store, never from Outlook's default folders. So a profile holding several mailboxes cannot be acted on in the wrong one, even if its default data file changes.

The mailboxes are declared in `config/mailboxes.json`. It is **machine-local and gitignored**, because it holds personal addresses and this repo is public. Copy the tracked template to start one:

```powershell
copy config\mailboxes.sample.json config\mailboxes.json
```

```json
{
  "schema_version": 1,
  "default": "owner",
  "mailboxes": {
    "owner":  {"address": "owner@example.com",  "display_name": "Owner Name",  "aliases": ["me"], "archive_folder": "Archive"},
    "second": {"address": "second@example.com", "display_name": "Second Name", "aliases": [],     "archive_folder": "Archive"}
  }
}
```

- Each key is the mailbox's alias. `aliases` are extra names for it. Every name is matched case-insensitively and must pick exactly one mailbox. `default` is the one a verb uses when it is not told otherwise. `display_name`, `aliases` and `archive_folder` are optional. An unknown key is refused rather than ignored, so a misspelt `archive_folder` cannot silently file into `Archive`.
- A mailbox is found by its `address`: first through the Outlook account whose SMTP address it is (that account's delivery store), else through the one store named after it, which is how Outlook names an IMAP store. No match exits 2 with `mailbox_not_in_outlook`, and so do two stores with that name and no account to tell them apart. There is no fall back to the default store.
- **No file is not an error.** The registry is then one mailbox, `default`, on Outlook's default store. That is exactly what every verb did before the registry existed, and the log says so at info. A file that is present but malformed exits 2 with `mailbox_registry_invalid` before Outlook is touched.
- Every result document carries `"mailbox": "<alias>"`, and each run logs `ℹ️ Resolved mailbox <alias> → store <name>`.

`main_batch.py mailboxes` is how a consumer (task-os, life-os) learns the mailbox list. It never reads the file. It prints the registry with each mailbox's `outlook_status`: `present` (its store is in the running Outlook), `missing` (Outlook answered and the mailbox is not in it) or `unknown` (Outlook is not running, or the check could not run). It only attaches to a running Outlook and never starts one, so with Outlook closed every status is `unknown`, never `present`. A non-`present` entry says why in `outlook_status_reason`. With no registry file, `registry` is `"synthesized"` and the one entry's `address` is the default store's account, read from Outlook.

#### Acting on a named mailbox

`--mailbox <alias>` on `plan`, `apply`, `revert`, `draft`, `read` and `send` picks the mailbox (any alias or extra alias, any case). Without it, the registry default is used, so every existing caller is unchanged. An unknown name exits 2 with `mailbox_unknown` before Outlook is touched.

```powershell
& .\.venv\Scripts\python.exe main_batch.py plan --mailbox second --candidates 0
& .\.venv\Scripts\python.exe main_batch.py apply --mailbox second --decisions decisions.json
& .\.venv\Scripts\python.exe main_batch.py draft --mailbox second --spec spec.json
```

- **Filing.** `plan --mailbox X` reads X's Inbox (or its Sent folder with `--folder sent`). `apply --mailbox X` writes the `.msg` under the same archive roots as any other mail, then moves the mail into X's **own** archive folder. A mail is never moved from one store into another. `revert --mailbox X` looks for the mail in X's archive folder and moves it back to X's Inbox. On Gmail over IMAP, the archive folder is the `[Gmail]/Archive` label, and moving a mail out of `INBOX` is what archives it.
- **Drafting as X.** The draft's `SendUsingAccount` is set to the Outlook account whose SMTP address is X's, right after the item is created (before it is displayed, so the signature is that account's). The account is then read back. If the draft would send as anyone else, it is discarded unsaved (`sending_account_mismatch`). A reply sets the account the same way rather than trusting Outlook to infer it. The BCC copy goes to X's own address, so it lands in X's Inbox to be filed. After saving, the draft is checked to be in X's Drafts and moved there if Outlook put it elsewhere. With no account for X's address, the run exits 2 with `account_not_in_outlook` and nothing is created. The `draft` document reports `from_address`, so the caller's preview can show it.
- **Reading and sending as X.** `read --mailbox X` and `send --mailbox X` accept only a draft in X's Drafts. `send` also requires the draft's sending account to be X's. A draft whose account was changed by hand, or is unset, is refused with `sending_account_mismatch` and nothing is sent. The account is not part of the fingerprint, so an approval taken before this check existed still matches. `read` and `send` report `from_address`.
- The From display name is the account's *Your name* field in Outlook. Set it in Outlook before the first real send from a newly added account.

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
| Archive root missing | Present roots are indexed, purge step is skipped — rows under an unmounted or unsynced root are never deleted; the headless scan exits `3` |

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


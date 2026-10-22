# Archive, scan and suggestions

## What it does

Two commands, each launchable from a Stream Deck button or the command line:

### 1. Scan Archive
Walks the entire OneDrive archive from a configured root folder, opens every `.msg` file it finds, extracts metadata (subject, sender, recipients, date, body preview), and stores it in a local SQLite database with a full-text search index. Subsequent scans are incremental — only new or modified files are processed.

### 2. Archive Email
Connects to the running Outlook instance, reads the currently selected email, queries the database to find the most relevant project folders, and presents ranked suggestions. You click a folder, confirm, and the email is saved as a `.msg` file plus all attachments — all consistently numbered — in under two seconds. The window closes immediately after saving (no confirmation popup); the log file records what was saved. Exchange senders are resolved to their SMTP address through `GetExchangeUser()`, and the dialog runs its work on a background thread so it never blocks on the database or Outlook.

When none of the suggestions is right there are two escapes, both in the dialog's bottom bar:

| Button | What it does |
|---|---|
| **Browse folder…** | Opens the native folder picker, starting at the first configured archive root |
| **Explorer folder** | Archives straight into the folder shown in the **foremost open File Explorer window** — the one you looked at last. No picker, no typing. |

`Explorer folder` is for the case where the right destination is already open on screen. It only considers real filesystem folders (virtual shell locations like *This PC* or *Quick Access* are skipped), it is **not** limited to the configured archive roots, and if no Explorer window is showing a real folder it says so rather than guessing.

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

## Incremental scanning behaviour

| Scenario | What happens |
|---|---|
| File unchanged | `mtime` matches DB → skipped instantly (no file open) |
| New file | Not in DB → parsed and inserted |
| File modified | `mtime` differs → re-parsed and updated |
| File deleted | After a complete scan, `DELETE WHERE file_path NOT IN (all found paths)` purges stale entries |
| Scan cancelled | Purge step is skipped — safe, no phantom deletions |
| Archive root missing | Present roots are indexed, purge step is skipped — rows under an unmounted or unsynced root are never deleted; the headless scan exits `3` |

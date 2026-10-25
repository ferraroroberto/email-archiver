# Email Archiver

A modular Python application that indexes and archives Outlook emails into a structured OneDrive folder system. A SQLite full-text index of the archive drives an instant-startup suggestion dialog designed for Stream Deck buttons, and a headless JSON batch CLI lets another local app drive the same archiver as a subprocess.

Two commands, each launchable from a Stream Deck button or the command line:

- **Scan Archive** walks the archive root, opens every `.msg` file, and stores its metadata in a local SQLite database with a full-text index. Later scans are incremental.
- **Archive Email** reads the email selected in Outlook, ranks the most relevant project folders, and saves the email plus its attachments, consistently numbered, in under two seconds.

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

All other defaults are sensible out of the box. The one knob worth knowing about is `naming.date_prefix` — the shipped example sets `auto`, but an omitted key falls back to `false` (never dated) — see [File naming](docs/naming.md).

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

# The same four tools behind one entry point: archive | scan | ui | batch
& .\.venv\Scripts\python.exe -m email_archiver scan --no-ui
```

The `main_*.py` scripts are thin wrappers over `python -m email_archiver <subcommand>`, kept at their paths for the launchers, the nightly job and the app that spawns `main_batch.py`.

## Verification

The project ships a pytest suite (`tests/`) covering the filename fitter, sequencing, the date-prefix toggle, the Explorer picker, the Message-ID column and its migration, renumbering (ordering, both name forms, split numbers, the dry run, the index), the headless scan's exit codes and its missing-root purge guard, and every batch verb (including `draft`, `read` and `send`) end to end against a fake Outlook client. Run it before declaring any change done:

```powershell
& .\scripts\verify-before-ship.ps1
```

That runs `pytest tests/` and then the static ratchet: a duplicate block (jscpd via `npx`, so Node is needed) or an unused import/variable (ruff) that is not in `.static-ratchet-baseline.json` fails the gate. Existing debt is baselined, not fixed, so the count can only go down; a detector that cannot run fails as NOT CONFIRMED rather than passing. The same gate then runs the docs, structure, instruction, doc-refs, dead-code, conventions and comment ratchets, each with its own baseline file at the repo root; the comment ratchet fails a `#NNN` issue reference added to a code comment.

`pytest` and `ruff` are dev-only dependencies — see `requirements.txt`.

---

## Reference

Durable reference lives in `docs/`; each line below is the entry point for one topic.

| Topic | Doc |
|---|---|
| What Scan and Archive do, the three-stage folder ranking, incremental scan rules | [docs/archive-and-scan.md](docs/archive-and-scan.md) |
| `NNN - name` file naming, the optional `YYYY-MM-DD - ` prefix, `naming.date_prefix` | [docs/naming.md](docs/naming.md) |
| Repository layout | [docs/project-structure.md](docs/project-structure.md) |
| Where new code goes, the declared root exceptions, the structure ratchet | [docs/layout.md](docs/layout.md) |
| Scheduled headless scan and its exit codes | [docs/nightly-scan.md](docs/nightly-scan.md) |
| SQLite schema, the follow-up flag, the Internet Message-ID | [docs/database.md](docs/database.md) |
| `main_batch.py` verbs, filing specific mail, Sent Items, identity, exit codes | [docs/batch-mode.md](docs/batch-mode.md) |
| Retrying a mail that was written but never moved | [docs/batch-retry.md](docs/batch-retry.md) |
| Renumbering a folder into sent-date order | [docs/renumber.md](docs/renumber.md) |
| `draft`, replies, `--update`, `read` and the guarded `send` | [docs/drafting-and-sending.md](docs/drafting-and-sending.md) |
| What batch mode will and will not touch | [docs/batch-safety.md](docs/batch-safety.md) |
| `outlook:` config keys, the mailbox registry, `--mailbox` | [docs/mailboxes.md](docs/mailboxes.md) |
| Live search of a mailbox's Outlook folders (`find`) | [docs/find.md](docs/find.md) |
| Internal module structure (Mermaid) | [docs/architecture.mmd](docs/architecture.mmd) |

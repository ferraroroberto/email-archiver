# Layout

Where new code goes. The shape is `pipeline-cli` (see `shape` in `.fleet.toml`); the structure ratchet (`scripts/structure_ratchet.py`) fails a root folder outside that shape and any file, function or directory that grows past its baseline in `.structure-ratchet-baseline.json`. A module-by-module walk of the package is in [project-structure.md](project-structure.md).

## Top-level folders

| Folder | Holds |
|---|---|
| `email_archiver/` | The package: all logic. Declared in `extra_roots` — see below. |
| `tests/` | The pytest suite. It never sends mail or touches a mailbox. |
| `scripts/` | The pre-ship gate and the vendored ratchets with their gate wrappers. |
| `config/` | The config and mailbox-registry templates; the live copies are git-ignored. |
| `data/` | The SQLite index, created on first scan; only a `.gitkeep` is tracked. |
| `logs/` | The log file's folder. Declared in `extra_roots` — see below. |
| `docs/` | Durable reference, indexed in the README. |

The root also holds the four `main_*.py` scripts and the `.bat` launchers. Those are files, not folders, so the layout check does not see them.

## The package

| Path | Holds |
|---|---|
| `email_archiver/cli/` | Command-line entry points: `__init__.py` is the `python -m email_archiver <subcommand>` dispatcher; `archive`, `scan`, `launcher` and `batch` are the four tools it routes to. |
| `email_archiver/ui/` | The Tkinter windows and the native folder picker. |
| `email_archiver/outlook/` | Everything that touches Outlook over COM, kept out of the headless modules. |
| `email_archiver/database/` | The SQLite schema and every SQL query. |
| `email_archiver/scanner/` | The incremental `.msg` indexer. |
| `email_archiver/archiver/` | Saving an email into the archive tree. |
| `email_archiver/engine/` | The three-stage folder ranking. |
| `email_archiver/*.py` | The headless batch verbs (`batch`, `dismiss`, `draft`, `find`, `send`, `renumber`) and the shared helpers (`config`, `paths`, `text`, `explorer`). |

## Entry points

`python -m email_archiver {archive,scan,ui,batch}` is the one entry point. Each subcommand passes the rest of the command line to the tool that owns it, so that tool's parser, help and exit codes are unchanged. The root `main_archive.py`, `main_scan.py`, `main_ui.py` and `main_batch.py` stay as thin wrappers over the same functions, because the launchers, the nightly job and the app that spawns `main_batch.py` call them by path. A new tool is a module in `email_archiver/cli/` and a row in `SUBCOMMANDS`; logic does not go in a root script.

## Declared exceptions

- `email_archiver/` — the `pipeline-cli` shape expects the package under `src/`. It sits at the root, and every launcher, the nightly job and the tests import it by that name, so it is declared in `extra_roots` rather than moved.
- `logs/` — the log file handler opens its file there, so the folder has to exist on a fresh clone. A tracked `.gitkeep` guarantees that, which makes it a tracked root folder; it is declared in `extra_roots` for that reason.

## Baselined offenders

None: `.structure-ratchet-baseline.json` is empty, so every file, function and directory sits under its limit and a new offender fails the gate. Split a growing module along its seam instead: `email_archiver/batch.py` keeps the verbs, `batch_move.py` the move and its error, `batch_common.py` the document envelope; the batch and draft tests share `tests/batch/batch_support.py` and `tests/draft_support.py`.

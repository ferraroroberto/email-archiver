# Nightly scan

`run-scan-nightly.bat` is the scheduled-job launcher. It `cd`s to the repo, sets `PYTHONUTF8=1` and `PYTHONUNBUFFERED=1`, runs `main_scan.py --no-ui` **in the foreground** with the repo `.venv` Python, and exits with the scan's exit code — no `start`, no `pythonw`, no `pause`. When stdout is not a terminal (a captured job log), progress is printed one line per tick instead of being rewritten in place with `\r`.

It runs in app-launcher as the job `email-archiver-scan-nightly` (schedule `none`, alert on failure), chained from both `on_success` and `on_failure` of the nightly fleet backup job, so the index is refreshed right after the backup whatever the backup's outcome. Chain hops run one at a time, which is what keeps the nightly scan from overlapping anything — there is no scan lock.

| Exit code | Meaning |
|---|---|
| `0` | Scan complete. Per-file read errors (for example cloud-only placeholders) are counted in the `Done — …` summary line, not a failure |
| `1` | Unexpected crash (uncaught exception), folder-explorer's index run failed (see below), or the launcher could not find the repo or its `.venv` |
| `2` | Config missing or unloadable, or no archive roots configured |
| `3` | A configured archive root does not exist. The roots that are present are still indexed, but the purge is skipped, so index rows under the missing root are left untouched; each missing root is logged by path |
| `4` | Busy: folder-explorer owns the index and another of its email runs held its lock past the ten-minute wait; nothing was indexed |

## When folder-explorer owns the index

If `database.path` points at folder-explorer's own email index (what its `python -m app.cli email config --json` reports as `db`), `main_scan.py --no-ui` does not scan here: it runs folder-explorer's `email index --fresh`, so that index keeps one writer. Folder-explorer is the checkout named by `FOLDER_EXPLORER_HOME`, else a sibling `folder-explorer` folder. In every other case the scan runs in this repo as described above.

On this path the exit code is `0` when folder-explorer's run indexed, `4` when it was busy, and `1` for anything else: its run failed, ran past its time limit, or could not start. Exit codes `2` and `3` come only from the scan in this repo. App-launcher alerts on any non-zero code, so this is the live fallback path it watches.

# Nightly scan

`run-scan-nightly.bat` is the scheduled-job launcher. It `cd`s to the repo, sets `PYTHONUTF8=1` and `PYTHONUNBUFFERED=1`, runs `main_scan.py --no-ui` **in the foreground** with the repo `.venv` Python, and exits with the scan's exit code — no `start`, no `pythonw`, no `pause`. When stdout is not a terminal (a captured job log), progress is printed one line per tick instead of being rewritten in place with `\r`.

It runs in app-launcher as the job `email-archiver-scan-nightly` (schedule `none`, alert on failure), chained from both `on_success` and `on_failure` of the nightly fleet backup job, so the index is refreshed right after the backup whatever the backup's outcome. Chain hops run one at a time, which is what keeps the nightly scan from overlapping anything — there is no scan lock.

| Exit code | Meaning |
|---|---|
| `0` | Scan complete. Per-file read errors (for example cloud-only placeholders) are counted in the `Done — …` summary line, not a failure |
| `1` | Unexpected crash (uncaught exception), or the launcher could not find the repo or its `.venv` |
| `2` | Config missing or unloadable, or no archive roots configured |
| `3` | A configured archive root does not exist. The roots that are present are still indexed, but the purge is skipped, so index rows under the missing root are left untouched; each missing root is logged by path |

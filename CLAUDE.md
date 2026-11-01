# Project Instructions

Canonical instructions for AI coding agents working in this repository. Claude Code reads this file directly as project memory. Other agents (Cursor, Codex, etc.) reach it via the one-line `AGENTS.md` pointer.

## This repository
Modular Python app that indexes and archives Outlook emails into a structured OneDrive folder system.
See `README.md` for setup and usage; layout and per-feature reference are indexed there and live in `docs/`.

## Internal architecture

[`docs/architecture.mmd`](docs/architecture.mmd) is a hand-authored Mermaid diagram of this repo's own internal structure — the entry points, the `src/email_archiver/` modules (scanner, database, outlook, engine, archiver, ui), and how a scan/archive email flows between them. Update it in the same PR as any material structural change (a module added/moved/renamed, a new entry point, a data-flow change) — same anti-staleness contract as a `.fleet.toml` `description` field.

## Verification

```powershell
& .\scripts\verify-before-ship.ps1
```

The gate is `pytest tests/` followed by eight ratchets. Each is a vendored-byte-for-byte (never edited here) `scripts/<kind>_ratchet.py`, configured in `.fleet.toml` `[<kind>_ratchet]` with baseline `.<kind>-ratchet-baseline.json`, and is called only through `python -m scripts.<kind>_ratchet_gate`, where exit 3 (a detector could not run) is NOT CONFIRMED, never a pass or a skip. Shrink a baseline in the PR that removes debt (`--update-baseline`); adding to it needs `--allow-additions` and shows in review. Procedure for each: project-scaffolding `docs/<kind>-ratchet.md`.

- **Static** (reads the whole tree, never path-routed): a duplicate block (jscpd via `npx`, so **Node is a prerequisite**), an unused import/variable (ruff `F401,F841`) or a ruff bug (`B`) or security (`S`) finding (`S113` strict) not in the baseline.
- **Docs**: `README.md` past 40 KiB or its baselined size, or `README.md` / `docs/**/*.md` gaining a history marker (an issue reference, a dated heading or a changelog heading). Docs state the present-tense why; history lives in git and the PR.
- **Structure** (`.fleet.toml` `shape = "pipeline-cli"` plus `extra_roots`): a file past 1,000 lines, a function 100 lines or a directory 40 modules that is new or above its baseline entry, or a root folder outside the declared shape. The baseline is empty: split a module that nears a limit along its seam (`batch.py`, `batch_move.py`, `batch_common.py`; the shared test doubles in `tests/batch/batch_support.py` and `tests/draft_support.py`). Where code goes, and why `logs/` is a declared root: `docs/layout.md`. Entry points are subcommands in `src/email_archiver/cli/`; the root `main_*.py` stay thin wrappers.
- **Instruction**: the weekly prompt audit's deterministic rules over `CLAUDE.md` and `AGENTS.md`; fails when a rule count rises above its baseline entry or a file passes its size cap. The baseline is empty: fix a new hit by rewording, not by raising it.
- **Doc-refs**: `README.md`, `CLAUDE.md` or `docs/**/*.md` naming a repo path, `python -m` module or relative link that does not exist in the tracked tree. The baseline is empty, and no port is checked because `.fleet.toml` declares none. Point a stale reference at what exists, or list an intentional pointer in `allow_refs`.
- **Dead-code**: a new function, class or method with no reference anywhere in the tracked Python tree, matched by name, so a helper nobody calls is refused. The baseline is empty. Delete the definition or call it; list a framework hook called by convention in `allow` or `entry_decorators`.
- **Conventions** (`ast` rules): `no-window` (a subprocess spawn passes `creationflags=`), `logging-not-print` (no `print` in `src/` outside the CLI entry points) and `no-drive-path` (no drive-letter path literal). The baseline is empty. Fix a hit, or exempt one call with a reasoned `# <rule>-exempt: <why>` marker. `secret-compare` and `git-wrapper` stay off because the repo has neither helper.
- **Comment**: `#NNN` issue references inside product-code comments, counted per file; tests and vendored scripts are out of scope. The baseline is empty. A comment states the present-tense why, so reword it in words instead of citing an issue; history lives in the issue, the PR and `git log`.

`pytest` and `ruff` are dev-only dependencies (declared in `requirements.txt`). Don't claim a change works without running this gate.

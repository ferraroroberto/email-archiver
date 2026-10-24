# Project Instructions

Canonical instructions for AI coding agents working in this repository. Claude Code reads this file directly as project memory. Other agents (Cursor, Codex, etc.) reach it via the one-line `AGENTS.md` pointer.

## This repository
Modular Python app that indexes and archives Outlook emails into a structured OneDrive folder system.
See `README.md` for setup and usage; layout and per-feature reference are indexed there and live in `docs/`.

## Internal architecture

[`docs/architecture.mmd`](docs/architecture.mmd) is a hand-authored Mermaid diagram of this repo's own internal structure — the entry points, the `email_archiver/` modules (scanner, database, outlook, engine, archiver, ui), and how a scan/archive email flows between them. Update it in the same PR as any material structural change (a module added/moved/renamed, a new entry point, a data-flow change) — same anti-staleness contract as a `.fleet.toml` `description` field.

## Verification

```powershell
& .\scripts\verify-before-ship.ps1
```

The gate is `pytest tests/` followed by the **static ratchet** (`scripts/static_ratchet.py`, vendored byte-for-byte from project-scaffolding and never edited here; `.fleet.toml` `[static_ratchet]`, baseline `.static-ratchet-baseline.json`). The ratchet fails on a duplicate block (jscpd via `npx`, so **Node is a prerequisite**) or an unused import/variable (ruff `F401,F841`) that is not in the baseline; it reads the whole tree, so it is never path-routed. The gate calls it only through `python -m scripts.static_ratchet_gate`, which fails exit 3 (a detector could not run) as NOT CONFIRMED, never a pass or a skip. Shrink the baseline in the PR that removes debt (`--update-baseline`); adding to it needs `--allow-additions` and shows in review.

The **docs ratchet** follows (`scripts/docs_ratchet.py`, also vendored byte-for-byte; `.fleet.toml` `[docs_ratchet]`, baseline `.docs-ratchet-baseline.json`, called through `python -m scripts.docs_ratchet_gate` with the same exit-3 rule). It fails when `README.md` passes 40 KiB or its baselined size, or when `README.md` or `docs/**/*.md` gains a history marker: an issue reference, a dated heading or a changelog heading. Docs state the present-tense why; history lives in git and the PR. Procedure: project-scaffolding `docs/docs-ratchet.md`.

The **structure ratchet** follows (`scripts/structure_ratchet.py`, vendored byte-for-byte; `.fleet.toml` `shape = "pipeline-cli"` plus `extra_roots`, `[structure_ratchet]`, baseline `.structure-ratchet-baseline.json`, called through `python -m scripts.structure_ratchet_gate` with the same exit-3 rule). It fails when a file passes 1,000 lines, a function 100 lines or a directory 40 modules and is new or above its baseline entry, or when a root folder falls outside the declared shape. Baselined offenders (`email_archiver/batch.py`, the two big test files, seven long functions) sit at their real size and cannot grow: new code goes in a new module, never in them. Where code goes, and why `email_archiver/` and `logs/` are declared roots: `docs/layout.md`. Entry points are subcommands in `email_archiver/cli/`; the root `main_*.py` stay thin wrappers.

The **instruction ratchet** follows (`scripts/instruction_ratchet.py`, vendored byte-for-byte; `.fleet.toml` `[instruction_ratchet]`, baseline `.instruction-ratchet-baseline.json`, called through `python -m scripts.instruction_ratchet_gate` with the same exit-3 rule). It runs the weekly prompt audit's deterministic rules over `CLAUDE.md` and `AGENTS.md` and fails when a rule count rises above its baseline entry or a file passes its size cap. The baseline is empty: fix a new hit by rewording, not by raising it. Procedure: project-scaffolding `docs/instruction-ratchet.md`.

`pytest` and `ruff` are dev-only dependencies (declared in `requirements.txt`). Don't claim a change works without running this gate.

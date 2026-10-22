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

`pytest` and `ruff` are dev-only dependencies (declared in `requirements.txt`). Don't claim a change works without running this gate.

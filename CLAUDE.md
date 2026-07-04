# Project Instructions

Canonical instructions for AI coding agents working in this repository. Claude Code reads this file directly as project memory. Other agents (Cursor, Codex, etc.) reach it via the one-line `AGENTS.md` pointer.

## This repository
Modular Python app that indexes and archives Outlook emails into a structured OneDrive folder system.
See `README.md` for setup, layout, and usage.

## Internal architecture

[`docs/architecture.mmd`](docs/architecture.mmd) is a hand-authored Mermaid diagram of this repo's own internal structure — the entry points, the `email_archiver/` modules (scanner, database, outlook, engine, archiver, ui), and how a scan/archive email flows between them. Update it in the same PR as any material structural change (a module added/moved/renamed, a new entry point, a data-flow change) — same anti-staleness contract as a `.fleet.toml` `description` field.

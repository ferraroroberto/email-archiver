# File naming convention

When an email is archived in a folder, files are named `NNN - <name>`, using ` - ` (space-dash-space) as the separator:

```
023 - Project_Alpha_meeting_notes.msg
023 - invoice.pdf
023 - signed_contract.docx
```

- `NNN` is a zero-padded 3-digit sequence number, derived from the highest existing prefix in that destination folder. It groups a bundle — an email and its attachments always share one `NNN` — and increments per folder.
- The email gets `NNN - sanitized_subject.msg`. Unicode letters (accents, `ñ`, `ç`, `·`) are kept as written; only what Windows forbids in a name is replaced (`/ \ :` become `-`; `< > " | ? *` and control characters become `_`), and trailing dots and spaces are dropped
- Each real attachment gets `NNN - original_filename.ext`
- Existing archived files are never renamed — this only affects what gets written going forward
- Embedded images (inline in HTML body) are skipped automatically; real attachments (e.g. PDFs) are saved even when the client sets a ContentId
- The filename is dynamically shortened with a trailing `...` if the destination folder is deep enough that the full path would otherwise exceed Windows' 260-char `MAX_PATH` limit (e.g. `042 - Long_subject_starts_here....msg`)

## Optional date prefix

The email's sent date can be prefixed to every name in the bundle, for folders shared with a document archive that files everything as `YYYY-MM-DD - <slug>`:

```
2026-03-14 - 023 - Project_Alpha_meeting_notes.msg
2026-03-14 - 023 - invoice.pdf
2026-03-14 - 023 - signed_contract.docx
```

`naming.date_prefix` in `config/config.yaml` controls the default form. The shipped `config.example.yaml` sets it to `auto`; omitting the key entirely (or a `config.yaml` that predates this toggle) falls back to `false` — never dated:

```yaml
naming:
  date_prefix: auto   # shipped example's setting; true → always dated, false → never dated (also the fallback when the key is absent)
```

- **`auto` (the shipped example's setting)** — infer the form per destination folder from what it already holds: majority of existing numbered files wins (dated vs. undated), falling back to the undated form when the folder is empty, has no numbered files, or ties. This is what makes an unattended batch run (see [Batch mode](batch-mode.md#batch-mode-headless), which has no checkbox) file into a mixed archive correctly without any per-mail configuration — a folder that already files everything as `YYYY-MM-DD - NNN - <name>` (for instance a tree shared with a dated document archive) keeps getting the dated form, the rest keep the undated form.
- **`true` / `false`** — force that form for every folder, ignoring its contents.
- **Per archive** — the **Date prefix (YYYY-MM-DD)** checkbox in the archive dialog's bottom bar starts pre-ticked to whatever the highlighted suggestion's own folder would get (inferred when the config is `auto`, the fixed config value otherwise), and can still be flipped by hand before clicking `Archive`, `Browse folder…` or `Explorer folder`; nothing is written to the config either way. Until you flip it, that tick is only a preview: `Browse folder…` and `Explorer folder` resolve the form for the folder you actually pick, never for the card you last hovered.

Precedence when a decision is not left to inference: an explicit per-archive choice (the checkbox once flipped by hand, or a batch `apply` decision's `date_prefix`) always wins over the config, `auto` inference is next, and the config's fixed boolean is the last resort.

- `YYYY-MM-DD` is the email's **sent** date (`SentOn`, falling back to `ReceivedTime`), normalized to local time — not the date it was archived. Attachments inherit the parent email's date, so a bundle stays contiguous.
- `NNN` keeps exactly the same meaning and derivation with the prefix on
- If the email's sent date cannot be resolved, that archive falls back to the undated form and logs why, rather than inventing a placeholder date
- Sequence allocation recognizes **both** the `NNN - ` and the `YYYY-MM-DD - NNN - ` forms no matter which one is being written, so a folder holding a mix of both still picks the correct next number and never collides — the toggle is safe to flip at any time, and inference reads that same single folder listing
- Path shortening accounts for the 13 extra characters only when the prefix is actually applied

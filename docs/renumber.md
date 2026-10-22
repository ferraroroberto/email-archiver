# Renumbering a folder

The `NNN` prefix is meant to be browsable: opening a project folder and reading down the list should be reading the thread in the order it happened. Two ordinary batch operations break that, and neither is a bug in how a single mail is archived — `revert` deletes a bundle and leaves a hole, and `apply` always allocates `max + 1`, so a mail filed into a folder after a correction takes the highest number even when it is older than everything already there.

`renumber` is the repair:

```powershell
& .\.venv\Scripts\python.exe main_batch.py renumber --folder "<a folder>" --dry-run
& .\.venv\Scripts\python.exe main_batch.py renumber --folder "<a folder>"
```

`--dry-run` computes and prints exactly the same map without renaming anything, so a run is always previewable. The same repair can be chained onto the two verbs that cause the problem — `apply --renumber` renumbers every destination folder the run wrote into, `revert --renumber` closes the gap in every folder the run deleted from — once per folder, after the verb's own work. Without the flag both verbs behave exactly as before and neither reports anything new.

**The ordering rule**, in one place:

- **Sent date decides.** Bundles are ordered by the mail's sent date, taken from the index and read out of the `.msg` itself when the index has no row for it (a file archived since the last scan). A bundle whose date cannot be established at all keeps its current position rather than being swept to one end.
- **The unit is a bundle** — a sequence number and every file carrying it. An email never parts company with its attachments. A number that carries no `.msg` at all (an attachment whose mail was deleted by hand, a document filed into the sequence) still holds a slot, because leaving it behind would park it on a number that now belongs to a different mail.
- **Numbers are contiguous from the folder's lowest existing one**, not from `001`. A folder that starts at `079` because it continues another folder's sequence keeps starting at `079`.
- **Each file keeps its own name form** — `NNN - ` stays undated, `YYYY-MM-DD - NNN - ` keeps its date. A folder holding a mix keeps the mix.
- **Two mails that ended up on one number are split by date**, and their attachments are matched to the right one by asking each `.msg` what it carries. An attachment that still cannot be placed follows the earlier mail and the map says so, as `attachments_placed`.
- **Files with no sequence prefix are left alone** and reported under `skipped`.
- **A folder outside `archive.root_paths` is refused**, untouched.

**The consumer must heal its stored paths from the map.** A renumber renames files another app may be holding paths to; the archiver updates its own index and nothing else. The map is reported under `renumbered`, keyed by folder, in all three verbs that produce one:

```json
"renumbered": {
  "<folder>": [
    {
      "from": "<folder>\\003 - subject.msg",
      "to":   "<folder>\\002 - subject.msg",
      "message_id": "abc123@mail.example",
      "attachments": [["<folder>\\003 - report.pdf", "<folder>\\002 - report.pdf"]]
    }
  ]
}
```

Only bundles that actually changed appear. `from`/`to` are `null` for a bundle that has no `.msg` (its files are all in `attachments`), so a consumer healing `.msg` paths simply has nothing to do for it.

**This includes the same document's own `files` lists.** An `apply --renumber` result reports each mail's `files` under the names it was *written* with, and the renumber that ran afterwards may have moved some of them — so hand-an-`apply`-result-straight-to-`revert` needs the map applied to those paths first when the flag was used. Without `--renumber` the `files` are final, as they always were. A folder that could **not** be renumbered is never an empty map in there — an empty map means "already in order" — it is listed under `renumber_refused` with its reason. The key is additive: `schema_version` is unchanged, and a consumer that does not know about it reads the rest of the document exactly as before.

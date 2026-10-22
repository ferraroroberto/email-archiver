# Searching live mail

The index (`data/emails.db`) knows only mail already archived to disk, and `plan --search` reads only the Inbox or Sent. A mailbox's history lives in its Outlook folders, so `find` searches them live:

```powershell
& .\.venv\Scripts\python.exe main_batch.py find --mailbox second --query "<words>" [--since 2026-01-01] [--folders inbox,archive,sent] [--limit 25]
```

- **Read-only.** It moves, saves, deletes, drafts and sends nothing. A mailbox with no archive folder is reported with `via: "absent"` for that folder. The folder is never created.
- **Every word must match**, case-insensitively, in the subject, sender (name or address), To / CC, or body: the same semantics as the index search. Each folder is narrowed server-side with `Items.Restrict` (a DASL `LIKE` filter, synchronous, no `AdvancedSearch` callback) and sorted newest first. A store that rejects the filter is walked instead, matching each mail's subject, sender, recipients and body preview, and that folder reports `via: "walk"`.
- **Folders.** The default is the Inbox, the mailbox's `archive_folder` and Sent. `--folders` narrows them. `--since` compares the received time (the sent time in Sent). `--limit` (default 25) caps the hits, newest first across all folders.
- **Each hit** carries `message_id`, `subject`, `sender`, `recipients`, `date` (sent time), `folder` (the Outlook folder's name), `entry_id`, `body_preview` and `already_archived` (the `.msg` path when its Message-ID is in the index, else `null`). The document adds per-folder `matched` counts under `folders`, `counts.matched` / `counts.returned`, and `truncated` when more matched than were returned.
- **Coverage limit.** The document says `coverage: "outlook_folders"` and gives a `coverage_note`, rather than implying completeness. With Gmail over IMAP (root folder path `[Gmail]`, All Mail not synced), mail archived from Gmail web or phone without the `[Gmail]/Archive` label lives only in All Mail, which is invisible to Outlook, so `find` cannot see it. A complete search would need the Gmail API, which is out of scope here.

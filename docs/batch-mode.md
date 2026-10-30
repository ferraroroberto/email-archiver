# Batch mode (headless)

`main_batch.py` is the archiver's headless face: eleven verbs — three that file the **whole Inbox** in one run instead of one selected mail at a time, one that moves Inbox mail into the archive folder without keeping a copy, one that repairs a folder's numbering, one that opens an unsent draft, two that read a draft back and send it only while it is still what was approved, one that lists the mailboxes it can act on, one that searches a mailbox's live folders, and one that asks Outlook to run Send/Receive — each printing exactly one JSON document on stdout. It exists so another local app can drive the archiver as a subprocess — the archiver stays the sole owner of Outlook COM, the suggestion engine and the naming rules, and the caller only decides *which folder* each mail goes to.

```powershell
& .\.venv\Scripts\python.exe main_batch.py plan --candidates 5 > plan.json
& .\.venv\Scripts\python.exe main_batch.py plan --search "<subject words>" --since 2026-09-15 --candidates 0
& .\.venv\Scripts\python.exe main_batch.py apply --decisions decisions.json
& .\.venv\Scripts\python.exe main_batch.py revert --items revert.json
& .\.venv\Scripts\python.exe main_batch.py dismiss --message-id <id> --message-id <id>
& .\.venv\Scripts\python.exe main_batch.py renumber --folder "<a folder>" --dry-run
& .\.venv\Scripts\python.exe main_batch.py draft --spec spec.json
& .\.venv\Scripts\python.exe main_batch.py read --entry-id <entry_id>
& .\.venv\Scripts\python.exe main_batch.py send --entry-id <entry_id> --expect-hash <sha256>
& .\.venv\Scripts\python.exe main_batch.py mailboxes
& .\.venv\Scripts\python.exe main_batch.py find --mailbox <alias> --query "<words>"
& .\.venv\Scripts\python.exe main_batch.py sync [--mailbox <alias>] [--timeout 120]
```

| Verb | Reads | Does |
|---|---|---|
| `plan` | nothing | Starts Outlook if it is closed, enumerates the Inbox, and returns every mail with its metadata and the top `--candidates` folder suggestions (default 10), each carrying its own `date_prefix`. `--message-id` / `--since` / `--search` / `--ref` narrow it to specific mail — see [Filing specific mail](#filing-specific-mail). Read-only. |
| `apply` | `[{message_id, folder_path, date_prefix}]` | Archives each mail into `folder_path` (which must be inside `archive.root_paths`), then moves it to the Outlook `Archive` folder and tags it with the category. A mail already in the index is finished rather than filed again — see [Retrying a mail that was written but never moved](batch-retry.md#retrying-a-mail-that-was-written-but-never-moved). |
| `revert` | `[{message_id, files}]` | Deletes exactly the listed files, removes their index rows, removes the category and moves the mail back to the Inbox. |
| `dismiss` | `--message-id` (repeatable; optionally `--category`) | Moves each named Inbox mail into the mailbox's archive folder and tags it with `outlook.dismiss_category`, **without saving a `.msg`** — see [Dismissing mail](#dismissing-mail). |
| `renumber` | nothing | Re-sequences one folder into sent-date order and prints the old → new map. Touches no Outlook and no COM — see [Renumbering a folder](renumber.md#renumbering-a-folder). |
| `draft` | `{to, cc, bcc, subject, body_text \| body_html, attachments, ref, display, reply_to, reply_all}` | Creates a filled Outlook draft that blind-copies your own address, saves it to Drafts and opens it for you to review; with `reply_to`, a real threaded reply. **Never sends.** See [Drafting a mail](drafting-and-sending.md#drafting-a-mail). |
| `read` | nothing | Reads one unsent draft back as it is stored, with its fingerprint. Writes nothing, shows nothing. See [Sending an approved draft](drafting-and-sending.md#sending-an-approved-draft). |
| `send` | `--expect-hash` (and optionally `--expect-part`, `--expect-to`) | Sends that one draft, only if its live fingerprint still matches. Refused with `approval_mismatch` naming what differs otherwise. |
| `mailboxes` | nothing | Lists the mailbox registry with each mailbox's `outlook_status`. Never starts Outlook. See [Mailboxes](mailboxes.md#mailboxes). |
| `find` | `--query` (and optionally `--since`, `--folders`, `--limit`) | Searches one mailbox's live Outlook folders (Inbox, archive folder, Sent). **Read-only.** See [Searching live mail](find.md#searching-live-mail). |
| `sync` | nothing (optionally `--mailbox`, `--timeout`) | Asks Outlook to run Send/Receive for the whole profile and reports, per mailbox, whether it was seen to finish. Reads, moves and sends no mail. See [Syncing mail](sync.md#syncing-mail). |

An `apply` result can be handed straight back to `revert` — the object with its `results` list is accepted as-is, no reshaping needed.

## Filing specific mail

A caller that already knows which mail it wants to file, and where, doesn't need a ranked plan of the whole Inbox. `plan` takes filters, all optional and combined with AND. With none of them the document is exactly the full-Inbox plan, with no `filters` key.

| Flag | Selects |
|---|---|
| `--message-id <id>` | That mail, looked up by Internet Message-ID (angle brackets optional) instead of enumerating the Inbox. Repeatable. An id found nowhere in the Inbox is listed under `skipped` with `reason: "not_in_inbox"` and its `message_id`. |
| `--since <YYYY-MM-DD[THH:MM]>` | Mail received at or after that local time. Narrowed server-side with `Items.Restrict` (the DASL literal is written in UTC, which is how Outlook compares it), then checked exactly per mail. A store that rejects the filter falls back to walking the Inbox, and the log says which path ran. |
| `--search <text>` | Mail whose subject, sender, recipients or body preview contains the text, case-insensitively. Repeatable: every term must match. |
| `--ref <token>` | Mail whose `X-Archive-Ref` header equals the token, the one `draft` stamps. Headers are read only for mail that passed the cheaper filters. The header survives sending on an IMAP mailbox (verified), but that depends on the account, so a caller keeps a fallback match (`--search` on the subject, `--since` the send time). |
| `--candidates 0` | No folder ranking: each mail comes back with `candidates: []`, for a caller that already knows the destination. |
| `--folder inbox\|sent` | Where to read from; `inbox` is the default. See [Filing from Sent Items](#filing-from-sent-items). Any other value is `bad_input`. |

A filtered document adds `filters: {message_ids, since, search, ref}` (what was applied), and `counts.inbox` counts the Inbox mails that matched. Each mail has the same shape as in a full plan.

### Filing from Sent Items

A sent mail's self-BCC copy does not always reach the Inbox (a large attachment, a suppressed self-copy), and the sent mail itself, carrying the same `X-Archive-Ref`, sits in Sent Items. A caller that wants that copy asks for it explicitly:

```powershell
& .\.venv\Scripts\python.exe main_batch.py plan --folder sent --ref <token> --since 2026-09-30T11:13 --candidates 0
# decisions.json: [{"message_id": "<from the plan>", "folder_path": "<a folder under an archive root>", "date_prefix": "auto", "source": "sent"}]
```

- `plan --folder sent` reads Sent Items; `--since` compares the **sent** time (`PR_CLIENT_SUBMIT_TIME`), and `--message-id` is looked up in Sent Items (an id not found is skipped with `reason: "not_in_sent"`). The document's `filters` gains `folder: "sent"`, `counts` carries `sent` where an Inbox plan carries `inbox`, and each mail reports `in_inbox: false`. An Inbox plan's document carries none of these.
- An `apply` decision with `"source": "sent"` looks the mail up in Sent Items (`not_in_sent` when it is not there), archives and tags it as usual, and does **not** move it: Sent Items is the record of what went out. The result reports `moved: false` and `move_via: "kept_in_sent"`. Once the scan has indexed the file, `plan` marks the mail `already_archived` and a second apply reuses it (`reused: true`) instead of filing a second copy; until then the mail, still in Sent Items, would be filed again, so a caller records the filing and rescans. A `source` other than `inbox` / `sent` is a per-mail `bad_decision`, never read as the Inbox.
- `revert` does not undo a Sent Items filing's move, because there was none: it deletes the listed files and reports the mail `not_in_archive_folder`.

Two `apply` / `revert` options serve the same caller:

- **`"date_prefix": "auto"`** on a decision resolves the naming form for its `folder_path` exactly as `plan` resolves a candidate's. Under `naming.date_prefix: auto` that is inferred from what the folder already holds, so an empty or new folder gets the undated form and a folder of dated files gets the dated one. Under a fixed `true` / `false` it is that value, and a `true` / `false` on a decision forces that form.
- **`--category <name>`** tags filed mail with that category instead of `outlook.category` (e.g. `"Archived by life-os"`). The document's `category` reports the one used. Pass the same flag to `revert` so it removes the right one.

```powershell
& .\.venv\Scripts\python.exe main_batch.py plan --ref <token> --candidates 0 > one.json
# decisions.json: [{"message_id": "<from one.json>", "folder_path": "<a folder under an archive root>", "date_prefix": "auto"}]
& .\.venv\Scripts\python.exe main_batch.py apply --decisions decisions.json --category "Archived by life-os" > applied.json
& .\.venv\Scripts\python.exe main_batch.py revert --items applied.json --category "Archived by life-os"
```

## Dismissing mail

`dismiss` clears mail the caller does not want kept. It looks each `--message-id` up in the Inbox, moves it into the mailbox's archive folder and tags it with `outlook.dismiss_category` (default `Dismissed by task-os`; `--category` overrides it for the run, and the document's `category` reports the one used). It is a different category from `apply`'s, so a mail in the archive folder with no file behind it is recognisable in Outlook.

It writes **nothing** to disk: no `.msg`, no index row, and no archive-roots check, because there is no destination path. Only an Inbox mail is ever touched — an id that matches nothing in the Inbox is `not_in_inbox`, whatever folder it is in. A Message-ID given twice is moved once. `--mailbox` acts on that mailbox's Inbox and archive folder only; an unknown alias is refused with exit 2 before Outlook is touched.

Each result in `results` carries `message_id`, `folder` (the destination), `ok`, `moved`, `categorized`, `move_via`, `entry_id` and `error`. `error.code` is one of `not_in_inbox`, `lookup_failed` (the Inbox lookup itself raised, so whether the mail is there is unknown), `move_failed` and `message_changed` (the mail is still in the Inbox; the second is the [open-window case](batch-retry.md#retrying-a-mail-that-was-written-but-never-moved)) or `category_failed` (the mail **was** moved and only the tag did not stick). The process exits `1`, with the full document on stdout, when any mail is not both moved and tagged, so a caller checking only the exit code cannot miss it.

A dismissed mail has no copy on disk. Recovering one means moving it back from the archive folder by hand in Outlook.

## Identity: the Message-ID, not the EntryID

Outlook rewrites a mail's `EntryID` when it is moved between folders, which is exactly what `apply` does to every mail it files. So the identity that ties the three verbs together is the **Internet Message-ID** (MAPI `PR_INTERNET_MESSAGE_ID`), stored without its angle brackets. The scanner records it for every `.msg` it indexes, which is what lets `plan` mark a mail as `already_archived` instead of offering to file it a second time. A mail carrying no Message-ID (rare — drafts, some system mail) is reported under `skipped` with `reason: "no_message_id"` and left alone, never guessed at.

## Exit codes and error handling

| Exit | Meaning |
|---|---|
| `0` | The run completed and stdout carries its document. Individual mails may still have failed — each result has its own `error` with a `code`. One failing mail never aborts the run. |
| `1` | `dismiss` only: the run completed and stdout carries its document, but at least one mail was not moved and tagged. Each result has its own `error`. |
| `2` | The run could not start, or `renumber` could not finish its one folder. stdout carries `{"error": {"code", "message"}}` instead of results; the code is one of `config_missing` (no config, **or** one that could not be loaded — a malformed `config.yaml` lands here too, with the parser error in `message`), `bad_input`, `outlook_unavailable`, `com_unavailable`, `self_address_unresolved` (`draft` only), `draft_not_found` / `draft_not_editable` (`draft --update`, `read`, `send`), `draft_body_unmarked` (`draft --update` only), `approval_mismatch` / `recipient_unreadable` (`send` only; `approval_mismatch` names the parts under `error.differs`), `mailbox_registry_invalid` (`config/mailboxes.json` is present but malformed), `mailbox_not_in_outlook` (a registry mailbox has no store in the running Outlook — never answered with the default store), `mailbox_unknown` (a `--mailbox` name the registry does not know; checked before Outlook is touched), `account_not_in_outlook` (`draft`: the mailbox's store is there but no Outlook account sends from its address), `sending_account_mismatch` (`draft`: Outlook would not set the mailbox's account, so the draft was discarded unsaved; `send`: the stored draft does not send as the mailbox's account, so nothing was sent), `run_failed` (Outlook was reached but the walk then hit something unclassified — a locked or corrupt index database, a renumber or disk error — distinct from `com_unavailable`, which is reserved for pywin32 not being installed). A folder outside `archive.root_paths` is a `bad_input`, and so is a renumber that stopped part-way: no map is printed, because a map the disk may not match is worse than none. |
| `3` | `sync` only: deferred, because this is a Windows session with no desktop (session 0). Outlook was not started and nothing was triggered; stdout carries the `desktop` block, `outcome: "deferred_no_desktop"` and an `error` with that code. See [Syncing mail](sync.md#syncing-mail). |

Per-mail `error.code` values: `bad_decision` (the entry had no `message_id` or `folder_path`, or its `folder_path` does not resolve inside `archive.root_paths`, with a message starting `outside_archive_roots`), `not_in_inbox`, `not_in_archive_folder`, `archive_failed`, `move_failed`, `message_changed`, `category_failed`. The last three are deliberately distinct: after a `move_failed` the mail is still in the Inbox, after a `category_failed` it is already filed and only *looks* untouched in Outlook, and a [`message_changed`](batch-retry.md#retrying-a-mail-that-was-written-but-never-moved) is a `move_failed` that an operator clears in seconds — usually by closing an open mail window.

`apply` fills in a result's `files` **before** it moves the mail, so a mail that was written to disk but failed to move is still fully revertible — that is why a result can carry `ok: false` and a non-empty `files` at the same time.

Every document carries `schema_version`, so a consumer can refuse a shape it does not understand rather than read a field that silently moved.

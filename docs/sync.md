# Syncing mail

A mail that is already on the server stays invisible to `find`, `plan` and `apply` until Outlook syncs on its own. `sync` makes Outlook do the Send/Receive (the F9 action) and reports how it went, so a caller can search right after:

```powershell
& .\.venv\Scripts\python.exe main_batch.py sync [--mailbox <alias>] [--timeout 120]
```

- **It spans mailboxes.** Outlook's Send/Receive is global to the profile, so one sync is triggered for all of it and `--mailbox` only narrows which mailboxes are *reported* (default: every mailbox in the registry; with no registry file, the one `default`). The document therefore has no single `mailbox` key; it lists them under `mailboxes`. `trigger.scope` says `"profile"`, and every mailbox that is present shares the one observed outcome rather than implying a per-account observation that does not exist.
- **Outcomes**, per mailbox, with `waited_seconds` and a `reason` for anything but `completed`:

  | `outcome` | Meaning |
  |---|---|
  | `completed` | The end of the sync was observed. |
  | `timed_out` | No end was reported within `--timeout` seconds (default 120, any positive number). The sync may still be running; the verb returned at the deadline. |
  | `unknown` | The end cannot be observed (no usable sync object, events that cannot be attached, an error Outlook reported), the trigger failed, or the mailbox check could not run. **Never folded into `completed`.** |
  | `mailbox_not_in_outlook` | The registry names a mailbox the running Outlook does not hold. That mailbox only: the others still sync. With none present, nothing is triggered. |

- **Exit code `0`** whenever the document was produced: a `timed_out` or `unknown` mailbox is a result, not a failure, like a per-mail error elsewhere. `2` when the run could not start (`bad_input` for a non-positive or non-numeric `--timeout`, `mailbox_unknown`, `mailbox_registry_invalid`, `outlook_unavailable`, `com_unavailable`). `3` when deferred, below.
- **No desktop, no sync.** In a Windows session with no desktop (session 0: a service or a scheduled task with no logged-on user) Outlook has nowhere to show a profile prompt or a progress dialog, so the verb starts nothing, prints the deferral document and exits `3`, before it touches COM or Outlook. The `desktop` block (`session_id`, `status` of `interactive`, `session_0` or `unknown`) is in every `sync` document; a session state that cannot be read is `unknown` and the run proceeds.
- **Outlook closed** is started first, as `plan` and `find` do (bounded by `--start-timeout`). It is never killed or restarted to force a sync.
- **Reads nothing.** It moves, saves, drafts and sends no mail. The one COM call that starts a sync lives in `src/email_archiver/outlook/syncing.py`, reached only by this verb, and a test pins that nowhere else references it.
- **A progress dialog may take focus** on the desktop while it runs.
- **How the end is observed:** each of Outlook's sync groups raises a `SyncEnd` event, which the verb waits for while pumping COM messages. Folder item counts are deliberately not used: a count that stops changing is not proof the sync ended. If Outlook exposes no usable sync group, `Namespace.SendAndReceive` is the trigger and the outcome is `unknown`.
- **What `completed` does not prove.** Outlook may raise `SyncEnd` for an IMAP account's group before "Synchronizing subscribed folders" has finished, so a caller that expects a specific mail keeps its own retry as the safety net; that loop lives in the consumer, not here.

```json
{
  "verb": "sync", "schema_version": 1, "generated_at": "…",
  "desktop": {"session_id": 1, "status": "interactive"},
  "timeout_seconds": 120.0,
  "trigger": {"method": "SyncObject.Start", "scope": "profile", "signal": "sync_object", "started": true},
  "mailboxes": [
    {"alias": "owner", "address": "owner@example.com", "outcome": "completed", "waited_seconds": 14.25, "reason": ""}
  ]
}
```

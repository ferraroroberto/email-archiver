# Mailboxes and configuration

## Configuration

```yaml
outlook:
  archive_folder: "Archive"           # created under the mailbox root if missing
  category: "Archived by task-os"     # stamped by apply, removed by revert (--category overrides)
  self_address: "you@example.com"     # draft: blind-copied on every draft
```

All three keys are optional. `archive_folder` and `category` fall back to the values above; `self_address` falls back to the default sending account's SMTP address. The folder is matched case-insensitively so a mailbox that already has one is used rather than duplicated. A registry mailbox's own `archive_folder` overrides `outlook.archive_folder` (see [Mailboxes](#mailboxes)).

## Mailboxes

Every Outlook verb acts on **one mailbox**, and every folder it touches (Inbox, Sent, Drafts, the archive folder) is taken from that mailbox's own Outlook store, never from Outlook's default folders. So a profile holding several mailboxes cannot be acted on in the wrong one, even if its default data file changes.

The mailboxes are declared in `config/mailboxes.json`. It is **machine-local and gitignored**, because it holds personal addresses and this repo is public. Copy the tracked template to start one:

```powershell
copy config\mailboxes.sample.json config\mailboxes.json
```

```json
{
  "schema_version": 1,
  "default": "owner",
  "mailboxes": {
    "owner":  {"address": "owner@example.com",  "display_name": "Owner Name",  "aliases": ["me"], "archive_folder": "Archive"},
    "second": {"address": "second@example.com", "display_name": "Second Name", "aliases": [],     "archive_folder": "Archive"}
  }
}
```

- Each key is the mailbox's alias. `aliases` are extra names for it. Every name is matched case-insensitively and must pick exactly one mailbox. `default` is the one a verb uses when it is not told otherwise. `display_name`, `aliases` and `archive_folder` are optional. An unknown key is refused rather than ignored, so a misspelt `archive_folder` cannot silently file into `Archive`.
- A mailbox is found by its `address`: first through the Outlook account whose SMTP address it is (that account's delivery store), else through the one store named after it, which is how Outlook names an IMAP store. No match exits 2 with `mailbox_not_in_outlook`, and so do two stores with that name and no account to tell them apart. There is no fall back to the default store.
- **No file is not an error.** The registry is then one mailbox, `default`, on Outlook's default store. The log says so at info. A file that is present but malformed exits 2 with `mailbox_registry_invalid` before Outlook is touched.
- Every result document carries `"mailbox": "<alias>"`, and each run logs `ℹ️ Resolved mailbox <alias> → store <name>`.

`main_batch.py mailboxes` is how a consumer (task-os, life-os) learns the mailbox list. It never reads the file. It prints the registry with each mailbox's `outlook_status`: `present` (its store is in the running Outlook), `missing` (Outlook answered and the mailbox is not in it) or `unknown` (Outlook is not running, or the check could not run). It only attaches to a running Outlook and never starts one, so with Outlook closed every status is `unknown`, never `present`. A non-`present` entry says why in `outlook_status_reason`. With no registry file, `registry` is `"synthesized"` and the one entry's `address` is the default store's account, read from Outlook.

### Acting on a named mailbox

`--mailbox <alias>` on `plan`, `apply`, `revert`, `draft`, `read` and `send` picks the mailbox (any alias or extra alias, any case). Without it, the registry default is used. An unknown name exits 2 with `mailbox_unknown` before Outlook is touched.

```powershell
& .\.venv\Scripts\python.exe main_batch.py plan --mailbox second --candidates 0
& .\.venv\Scripts\python.exe main_batch.py apply --mailbox second --decisions decisions.json
& .\.venv\Scripts\python.exe main_batch.py draft --mailbox second --spec spec.json
```

- **Filing.** `plan --mailbox X` reads X's Inbox (or its Sent folder with `--folder sent`). `apply --mailbox X` writes the `.msg` under the same archive roots as any other mail, then moves the mail into X's **own** archive folder. A mail is never moved from one store into another. `revert --mailbox X` looks for the mail in X's archive folder and moves it back to X's Inbox. On Gmail over IMAP, the archive folder is the `[Gmail]/Archive` label, and moving a mail out of `INBOX` is what archives it.
- **Drafting as X.** The draft's `SendUsingAccount` is set to the Outlook account whose SMTP address is X's, right after the item is created (before it is displayed, so the signature is that account's). The account is then read back. If the draft would send as anyone else, it is discarded unsaved (`sending_account_mismatch`). A reply sets the account the same way rather than trusting Outlook to infer it. The BCC copy goes to X's own address, so it lands in X's Inbox to be filed. After saving, the draft is checked to be in X's Drafts and moved there if Outlook put it elsewhere. With no account for X's address, the run exits 2 with `account_not_in_outlook` and nothing is created. The `draft` document reports `from_address`, so the caller's preview can show it.
- **Reading and sending as X.** `read --mailbox X` and `send --mailbox X` accept only a draft in X's Drafts. `send` also requires the draft's sending account to be X's. A draft whose account was changed by hand, or is unset, is refused with `sending_account_mismatch` and nothing is sent. The account is not part of the fingerprint, so the account check never changes an approval's hash. `read` and `send` report `from_address`.
- The From display name is the account's *Your name* field in Outlook. Set it in Outlook before the first real send from a newly added account.

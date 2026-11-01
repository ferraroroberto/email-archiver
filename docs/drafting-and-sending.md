# Drafting a mail

`draft` turns a written email into an Outlook draft the user reads and sends by hand:

```json
{
  "to": ["someone@example.org"], "cc": [], "bcc": [],
  "subject": "…",
  "body_text": "…",
  "attachments": ["C:/…/file.pdf"],
  "ref": "opaque-caller-token",
  "display": true
}
```

- **The spec is checked before Outlook starts.** At least one `to`, a non-empty `subject`, exactly one of `body_text` / `body_html`, every attachment an existing file, and no unknown keys (a misspelt `atachments` would otherwise drop the file silently). Anything else is `bad_input` in milliseconds.
- **Your own address is always on the BCC line**, exactly once. The copy that lands back in the Inbox is what gets filed afterwards, so it is taken from `outlook.self_address` when set, else from the default sending account's SMTP address (the account owning the default store, read from `Account.SmtpAddress` — never through `GetExchangeUser`, which can raise the security prompt). When neither resolves the run exits 2 with `self_address_unresolved` and **no draft is created**. With a mailbox registry, the BCC address is the mailbox's own `address` instead (see [Acting on a named mailbox](mailboxes.md#acting-on-a-named-mailbox)).
- **Body and signature.** `body_text` is escaped into minimal HTML with its paragraphs and line breaks kept; `body_html` is used as written. With `display: true` (the default) the compose window is opened first — that is when Outlook inserts the account's default signature — and the body is then put at the top, above it. With `display: false` nothing is shown and no signature is added.
- **Saved, never sent.** The finished draft is saved, so it sits in Drafts even if the window is closed. No code path calls `Send()`.
- **`ref`** is stamped as an `X-Archive-Ref` Internet header on the draft, so a caller can recognise its own copy when it arrives. `ref_header` reports `stamped` or `not_stamped`, with `ref_header_reason` saying why (`no ref given`, or the store's refusal). Whether the header survives sending depends on the account, so a caller keeps a fallback match (subject, recipients, sent time).

The document:

```json
{
  "verb": "draft", "schema_version": 1, "generated_at": "…",
  "entry_id": "…", "subject": "…",
  "to": ["…"], "cc": [], "bcc": ["…", "<your address>"],
  "attachments": ["C:/…/file.pdf"],
  "ref": "opaque-caller-token", "ref_header": "stamped", "ref_header_reason": "",
  "displayed": true, "updated": false, "created_at": "2026-09-15T10:00:00+02:00"
}
```

## Replying to a mail

Add `reply_to` to make the draft a **real Outlook reply** instead of a fresh mail: it carries the thread link (In-Reply-To), the original quoted under the body, a `Re:` subject and the original's recipients, so Outlook threads it under the original.

```json
{
  "reply_to": {"message_id": "<id of an Inbox mail>"},
  "reply_all": false,
  "body_text": "…"
}
```

- **The original** is `{"message_id": …}` (an Inbox mail, found by its Message-ID) or `{"msg_path": "C:/…/saved.msg"}` (an archived `.msg`, absolute, opened with `Namespace.OpenSharedItem` and closed again without saving). Exactly one of the two. `reply_all: true` asks for Reply All (recipients and CC kept); it needs `reply_to`.
- **`to`, `cc` and `subject` become optional** and default to what Outlook computed. Any that you give overrides it. `bcc` (your own address) is set as for any draft.
- **The same steps as any draft**, in the same order: `X-Archive-Ref` stamped, the compose window displayed first (that is when the signature appears), then your body written **above** the signature and the quote, then saved. The quote is Outlook's and is never rewritten. Without a readable body from Outlook the run stops rather than write a body that would drop the quote.
- **Refused before anything is saved**, exit 2: `reply_source_not_found` (no Inbox mail with that Message-ID, or the `.msg` cannot be opened or is not a mail) and `reply_unavailable` (the original opened but Outlook would not reply to it — for example a `.msg` with no account to reply from; or the original is a mail you sent and its own recipients cannot be read, see the next bullet).
- **A reply to a mail you sent goes to the people you wrote to.** Outlook's `Reply()` answers the original's sender, which on a mail you sent (a saved sent `.msg`, or one from Sent Items) is you. When no `to` is given and Outlook's reply is addressed to your own address (`outlook.self_address`, else the default sending account), the To is taken from the original's To line instead, and the CC from its CC for Reply All only. If those recipients cannot be read the run refuses with `reply_unavailable` rather than leave a draft addressed to you; give `to` to override. Reply All is already addressed correctly by Outlook and is left as it is.
- **The document reports the reply**, in the style of `ref_header`: `reply_to` (what was asked), `replied_to_message_id` (the original's own id), the `to` / `cc` / `subject` the saved draft carries, and `thread_header` — `set` when the saved draft holds the original as its In-Reply-To (read back from the draft, not assumed), else `not_set` with `thread_header_reason`. `recipients_from` says who the reply was addressed from: `sender` (Outlook's own reply), `original_recipients` (the original was sent by you) or `caller` (you gave `to`). A new mail reports `not_a_reply`.
- **`send` and `read` work on a reply as on any draft.** The fingerprint's `body` part is the whole stored HTML, so the quoted original is bound by the approval; `read` reports the whole body under `body.html`, the caller's own text under `body.region_text` and what sits below it (the signature and the quoted original) under `body.after_region_html`, so a preview shown for approval can include the quote.
- **Outlook rewrites the HTML of a reply it displays.** With the compose window open, the body is stored the way Outlook's editor re-serialized it: the `<!--/archive-draft-body-->` closing comment is gone and each paragraph is padded with an empty `<o:p></o:p>`. The region is still found (it ends at the `</div>` that closes its own opening tag), `body.region_text` ignores the padding, and `body.after_region_html` is how a caller gets the quote without looking for the comment. A caller must not look for the comment in `body.html` of a reply.
- **`draft --update <entry_id>` on a reply** re-fills only the marked region; the quote below it is left alone. `reply_to` in the update spec is accepted but not re-applied (the draft keeps the thread link it was created with), and a `to` / `cc` / `subject` you leave out stays as the draft has it (reported as `null`, `thread_header: "unchanged"`).
- The real In-Reply-To on the *sent* copy is Outlook's and can only be seen after a send; check it by hand on a mail to yourself.

## Updating the same draft

```powershell
& .\.venv\Scripts\python.exe main_batch.py draft --update <entry_id> --spec spec.json
```

A caller iterating on one mail ("make it shorter") edits the draft it already made instead of leaving a near-duplicate in Drafts. `<entry_id>` is the `entry_id` the first `draft` returned; the spec is the full new spec, validated exactly as above.

- **Same item, re-filled.** Recipients (your own address still blind-copied exactly once), subject, attachments (every existing one removed, the spec's added) and the `X-Archive-Ref` header are set again, and the item is saved. With `display: true` it is shown again afterwards. **Never sent.**
- **An open compose window on that draft is closed first, saving it.** The window holds its own copy: left open it would still show the old text, and Send pressed there would send that. The update then works on what the window saved.
- **Only the body the tool wrote is replaced.** `draft` wraps the caller's body in `<div id="archive-draft-body">…</div><!--/archive-draft-body-->`; an update swaps that region and keeps everything after it, which is where the signature sits. Text the user typed into the draft by hand inside that region is overwritten, as is any hand edit to the recipients or subject.
- **Refused, and the item left untouched**, each with its own code and exit 2:
  - `draft_not_found` — the EntryID no longer resolves (the draft was deleted, or sent and moved);
  - `draft_not_editable` — the item was sent, or is not in the Drafts folder;
  - `draft_body_unmarked` — no marked body region: a draft created before `--update` existed, or one whose opening `<div id="archive-draft-body">` is gone. (A draft whose closing comment Outlook dropped, which is every displayed reply, still has its region and is updated.) Guessing where the body ends could eat the signature or the user's own text.
- The document is the one above with `updated: true` and `updated_at` in place of `created_at`.

## Sending an approved draft

```powershell
& .\.venv\Scripts\python.exe main_batch.py read --entry-id <entry_id>
& .\.venv\Scripts\python.exe main_batch.py send --entry-id <entry_id> --expect-hash <sha256> `
    --expect-part body=<sha256> [--expect-part to=<sha256> ...] [--expect-to <address> ...]
```

A caller shows the user a draft, the user approves it, and only then does the caller send. The approval is bound to the item actually in Drafts, not to what the caller wrote, so an edit made by hand in Outlook after the approval stops the send.

- **`read`** reports the draft as stored: `to` / `cc` / `bcc` read from its recipients, `subject`, `body.html` (the whole `HTMLBody`, signature included), `body.region_html` (the marked region `draft` wrote, `null` when the markers are gone) and `body.region_text` (the plain text that region was made from, `null` when it is not exactly what `draft` writes for a `body_text`; the empty `<o:p></o:p>` Outlook pads a reply's paragraphs with is not held against it), `body.after_region_html` (what follows the region: the signature and, on a reply, the quoted original; `null` when the region is gone), and `attachments` with their `kind`, byte size and sha256. The attachments are saved to a temp folder to be hashed, which is removed afterwards. A saved mail attached as an item (`kind: "embedded"`) is the exception: Outlook writes it out afresh on every save, so its bytes never hash the same twice. Its sha256 is taken over what it says instead (subject, sender, sent time, To, CC, plain body, and its own attachments the same way, three levels deep), and its `size_bytes` is Outlook's stored size, which reads 0 until the draft is saved, so it is informational and left out of the fingerprint: the content sha256 alone binds it. A file (`kind: "file"`) keeps its byte hash. It also reports `fingerprint`: a sha256 per part (`to`, `cc`, `bcc`, `subject`, `body`, `attachments`) and one `hash` over them. Recipients are compared as a sorted, case-folded set.
- **`send`** closes any window open on the draft, saving it, so nothing typed there is lost and anything typed counts as a change. It then re-reads the item and recomputes the fingerprint **immediately before `Send()`, on the same item it sends**. If the hash differs, or the To line is not exactly the `--expect-to` addresses, the run exits 2 with `approval_mismatch`. `error.differs` names the parts that no longer match the `--expect-part` hashes. Nothing is sent.
- **One item, by EntryID.** No lookup by subject or search, and no recipient is resolved or added. A recipient with no readable address refuses the send as `recipient_unreadable`. A draft that is gone, already sent or outside Drafts is `draft_not_found` / `draft_not_editable`.
- **`draft` has no send path.** The one COM `Send` call lives in `src/email_archiver/outlook/sending.py`, reached only by this verb, and a test pins that nowhere else calls it.

```json
{
  "verb": "send", "schema_version": 1, "generated_at": "…",
  "entry_id": "…", "subject": "…", "to": ["…"], "cc": [], "bcc": ["you@example.com"],
  "attachments": ["letter.pdf"], "fingerprint": "<sha256>", "sent": true, "sent_at": "…"
}
```

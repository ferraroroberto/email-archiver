# Filing state

`apply` and `dismiss` walk each mail through look-up, write, move and tag. Where that walk ended is one declared state, `FilingState` in `src/email_archiver/filing_state.py`, and the three result keys that report it (`ok`, `moved`, `categorized`) are computed from it. The method is the fleet's implicit-state-machine convention; the machine is declared in `.fleet.toml` as `filing-state`.

## Why it is declared

The result keys `ok`, `moved`, `categorized`, `reused`, `move_via` and `error.code` are scattered facts, and a caller that infers the mode from them can read a combination that contradicts the documented meaning of `move_failed` ("the mail is still in the Inbox"): a move that landed, followed by a failure reading the mail's new EntryID, would read as `moved: true` together with `error.code: move_failed`. So the walk collects what it observed, `derive_filing_state` maps that to one state, and `settle` writes the three keys from it. Nothing else sets them.

## The states

| State | Meaning | `ok` | `moved` | `categorized` | `error.code` |
|---|---|---|---|---|---|
| `refused` | The decision could not be acted on; nothing was looked up or written. | no | no | no | `bad_decision` |
| `not_found` | The look-up ran and the mail is not in the folder it was asked for. | no | no | no | `not_in_inbox`, `not_in_sent` |
| `lookup_failed` | The look-up itself raised, so whether the mail is there is not known. | no | no | no | `lookup_failed` (`dismiss`), `not_in_inbox` or `not_in_sent` (`apply`) |
| `archive_failed` | The files could not be written; the mail is untouched. `apply` only. | no | no | no | `archive_failed` |
| `move_failed` | The move did not land: the mail is still in the Inbox. `apply` has written its files by then, `dismiss` writes none. | no | no | no | `move_failed`, `message_changed` |
| `moved_untagged` | The move landed and the tag did not stick. | no | yes | no | `category_failed` |
| `filed` | Moved and tagged. | yes | yes | yes | none |
| `kept_in_sent_untagged` | A Sent Items mail was archived, is never moved, and the tag did not stick. | no | no | no | `category_failed` |
| `filed_kept_in_sent` | A Sent Items mail archived and tagged where it is. | yes | no | yes | none |

`apply` reports a look-up that raised under the same code as a mail that is not there, so a caller of the document cannot tell the two apart; the state can, and `dismiss` already gives the raise its own code. The emitted codes are unchanged.

## The observation

`FilingFacts` holds what a run observed about one mail: whether the decision was refused, what the look-up answered (`Lookup`), whether the files failed to write, whether the mail moved, whether it was kept in Sent Items and whether the tag stuck. `derive_filing_state` is the only function that maps those to a state, and it raises `ValueError` for a combination the walk cannot reach (a refused decision that was looked up, a mail tagged before it was filed, a mail both moved and kept in Sent Items, a missing mail that was archived). A new fact joins `FilingFacts` and the derive function, never a reader.

A move that landed stays moved when the read of its new EntryID raises: `entry_id` stays empty, the failure is logged and the run continues to the tag step. A mail that was kept in Sent Items has no move to fail, so the same read cannot turn it into `move_failed` either.

## Who reads it

- `_apply_one` in `src/email_archiver/batch.py` and `_dismiss_one` in `src/email_archiver/dismiss.py` collect the facts and call `settle` once per exit.
- The move and tag steps both verbs share are `move_and_tag` in `src/email_archiver/batch_move.py`. It returns what it observed and holds no state of its own.

The emitted documents keep `schema_version` 1 and every existing key and value; the state itself is not a key. `revert` shares this state set but not yet the machine, and `plan`, `find` and `apply` still read "already archived" from the index row separately.

## Tests

`tests/batch/test_filing_state.py` is a table with one row per reachable observation and its state and result keys, and a table of the combinations the derive function refuses. `tests/batch/test_batch_filing_state.py` drives the landed-move-then-unreadable-id case through `apply`, `dismiss` and a Sent Items mail.

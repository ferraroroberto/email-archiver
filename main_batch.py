"""
Headless batch entry point – plan / apply / revert / renumber / draft / read / send, over JSON.

Meant to be spawned as a subprocess by another local app (task-os, life-os),
never used interactively: every verb prints exactly one JSON document on
stdout and logs to the usual log file and to stderr. Only `draft` opens a
window — the compose window of the draft it creates, which it never sends.
Only `send` sends, and only a draft whose live fingerprint still matches the
one the caller approved (`read` reports it).

    python main_batch.py plan --candidates 5 > plan.json
    python main_batch.py plan --since 2026-09-15 --search invoice --candidates 0
    python main_batch.py plan --message-id <id> [--message-id <id> ...]
    python main_batch.py plan --ref <token> --candidates 0
    python main_batch.py apply --decisions decisions.json [--renumber] [--category <name>]
    python main_batch.py revert --items revert.json [--renumber] [--category <name>]
    python main_batch.py renumber --folder "<a folder>" [--dry-run]
    python main_batch.py draft --spec spec.json
    python main_batch.py draft --update <entry_id> --spec spec.json
    python main_batch.py read --entry-id <entry_id>
    python main_batch.py send --entry-id <entry_id> --expect-hash <sha256> \
        [--expect-part body=<sha256> ...] [--expect-to <address> ...]

Exit codes:
    0  the run completed and stdout carries its document — individual mails may
       still have failed, each with its own `error` inside the document
    2  the run could not start: no config, unreadable input, a folder outside
       the archive roots, Outlook unreachable, (draft) no self address to
       blind-copy, (draft --update / read / send) a draft that is gone, sent,
       outside Drafts or (update) has no marked body region, or (send) a draft
       that is not what was approved — `approval_mismatch`, naming the parts
       under `error.differs` — or has a recipient with no readable address.
       stdout carries `{"error": {"code", "message"}}` instead

Why a separate process rather than a library call: the Inbox verbs drive
Outlook over COM, and a COM modal (the address-book security prompt, a profile
chooser) blocks whatever thread it is raised on. Spawning this with a timeout
means such a prompt hangs *this* process, which the caller can kill, and never
the caller. ``renumber`` is the exception that proves the rule: it reads a
folder and the index and touches no COM at all, so it never starts Outlook.
"""
from __future__ import annotations

import argparse
import json
import logging
import re
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

# Ensure project root is on the path regardless of cwd
sys.path.insert(0, str(Path(__file__).resolve().parent))

from email_archiver import batch, draft, send
from email_archiver.config import load_config, setup_logging
from email_archiver.outlook.client import OutlookClient
from email_archiver.outlook.drafts import DraftUpdateError
from email_archiver.outlook.mapi import OutlookUnavailableError
from email_archiver.outlook.process import DEFAULT_START_TIMEOUT_SECONDS
from email_archiver.text import normalize_message_id

logger = logging.getLogger(__name__)

EXIT_OK = 0
EXIT_CANNOT_START = 2

ERROR_CONFIG_MISSING = "config_missing"
ERROR_BAD_INPUT = "bad_input"
ERROR_OUTLOOK_UNAVAILABLE = "outlook_unavailable"
ERROR_COM_UNAVAILABLE = "com_unavailable"
ERROR_SELF_ADDRESS_UNRESOLVED = "self_address_unresolved"


def _emit(document: dict[str, Any]) -> None:
    """Print the one JSON document this process produces.

    stdout is reconfigured to UTF-8 first: under a pipe Python falls back to the
    console code page, and a single accented subject would otherwise raise
    UnicodeEncodeError and kill a run that had already moved mail.
    """
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except (AttributeError, OSError):  # pragma: no cover - defensive
        pass
    json.dump(document, sys.stdout, ensure_ascii=False, indent=2)
    sys.stdout.write("\n")
    sys.stdout.flush()


def _fail(verb: str, code: str, message: str, **details: Any) -> int:
    logger.error("%s: %s", code, message)
    _emit(batch.error_document(verb, code, message, **details))
    return EXIT_CANNOT_START


def _load_input(path: str) -> list[dict[str, Any]]:
    """Read a decisions / items file into a list of dicts.

    Accepts either a bare JSON list or an object carrying one under
    ``decisions`` / ``items`` / ``results`` — the last so an `apply` document
    can be handed straight back to `revert` without being reshaped.
    """
    with open(path, encoding="utf-8") as fh:
        data = json.load(fh)
    if isinstance(data, dict):
        for key in ("decisions", "items", "results"):
            if isinstance(data.get(key), list):
                data = data[key]
                break
        else:
            raise ValueError(
                "expected a JSON list, or an object with a decisions/items/"
                "results list"
            )
    if not isinstance(data, list):
        raise ValueError("expected a JSON list of objects")
    bad = [i for i, entry in enumerate(data) if not isinstance(entry, dict)]
    if bad:
        raise ValueError(f"entries at index {bad[:5]} are not JSON objects")
    return data


def _plan_filters(args: argparse.Namespace) -> batch.PlanFilters:
    """``plan``'s filter flags, validated before Outlook is touched.

    Raises:
        ValueError: a flag's value is unusable; the message names the flag.
    """
    message_ids: list[str] = []
    for raw in args.message_ids:
        message_id = normalize_message_id(raw)
        if not message_id:
            raise ValueError("--message-id must not be blank")
        if message_id not in message_ids:
            message_ids.append(message_id)

    since = None
    if args.since is not None:
        try:
            since = datetime.fromisoformat(args.since.strip())
        except ValueError:
            raise ValueError(
                f"--since {args.since!r} is not YYYY-MM-DD or YYYY-MM-DDTHH:MM"
            ) from None
        if since.tzinfo is not None:
            raise ValueError("--since is local time; leave the UTC offset off")

    search = [term.strip() for term in args.search]
    if not all(search):
        raise ValueError("--search must not be blank")

    ref = args.ref.strip() if args.ref is not None else None
    if ref == "":
        raise ValueError("--ref must not be blank")

    return batch.PlanFilters(
        message_ids=tuple(message_ids), since=since, search=tuple(search), ref=ref
    )


_SHA256 = re.compile(r"[0-9a-f]{64}")


def _send_expectations(args: argparse.Namespace) -> tuple[str, dict[str, str], list[str] | None]:
    """``send``'s ``--expect-*`` flags, validated before Outlook is touched.

    Raises:
        ValueError: a flag's value is unusable; the message names the flag.
    """
    expect_hash = args.expect_hash.strip().lower()
    if not _SHA256.fullmatch(expect_hash):
        raise ValueError("--expect-hash must be a sha256 hex digest")
    parts: dict[str, str] = {}
    for raw in args.expect_parts:
        name, _, digest = raw.partition("=")
        name, digest = name.strip(), digest.strip().lower()
        if name not in send.PARTS or not _SHA256.fullmatch(digest):
            raise ValueError(
                f"--expect-part {raw!r} is not NAME=SHA256 with NAME one of {', '.join(send.PARTS)}"
            )
        parts[name] = digest
    expect_to = None
    if args.expect_to:
        expect_to = [address.strip() for address in args.expect_to]
        if not all(expect_to):
            raise ValueError("--expect-to must not be blank")
    return expect_hash, parts, expect_to


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="main_batch.py",
        description="Plan, apply or revert archiving of the whole Outlook Inbox, "
                    "or open an unsent draft.",
    )
    parser.add_argument(
        "--start-timeout", type=float, default=DEFAULT_START_TIMEOUT_SECONDS,
        help="Seconds to wait for a freshly started Outlook to answer over COM "
             "(default: %(default)s).",
    )
    sub = parser.add_subparsers(dest="verb", required=True)

    p_plan = sub.add_parser(
        "plan", help="List Inbox mails (all, or the filtered ones) with ranked folders."
    )
    p_plan.add_argument(
        "--candidates", type=int, default=batch.DEFAULT_CANDIDATES,
        help="Ranked folder candidates per mail; 0 lists the mails with no "
             "ranking (default: %(default)s).",
    )
    p_plan.add_argument(
        "--message-id", action="append", default=[], dest="message_ids",
        help="Plan only this mail, looked up by Internet Message-ID instead of "
             "enumerating the Inbox. Repeatable.",
    )
    p_plan.add_argument(
        "--since",
        help="Only mail received at or after this local time: YYYY-MM-DD or "
             "YYYY-MM-DDTHH:MM.",
    )
    p_plan.add_argument(
        "--search", action="append", default=[],
        help="Only mail whose subject, sender, recipients or body preview "
             "contains this text (case-insensitive). Repeatable; all must match.",
    )
    p_plan.add_argument(
        "--ref",
        help="Only mail whose X-Archive-Ref header (stamped by `draft`) equals "
             "this token.",
    )

    p_apply = sub.add_parser("apply", help="Archive, move and tag decided mails.")
    p_apply.add_argument(
        "--decisions", required=True,
        help='JSON file: [{message_id, folder_path, date_prefix: true|false|"auto"}, ...]',
    )
    p_apply.add_argument(
        "--renumber", action="store_true",
        help="Afterwards, renumber every destination folder this run wrote "
             "into and report the old-to-new map under `renumbered`.",
    )
    p_apply.add_argument(
        "--category",
        help="Tag filed mail with this Outlook category instead of outlook.category.",
    )

    p_revert = sub.add_parser("revert", help="Undo an apply: delete files, move back.")
    p_revert.add_argument(
        "--items", required=True,
        help="JSON file: [{message_id, files: [...]}, ...]",
    )
    p_revert.add_argument(
        "--renumber", action="store_true",
        help="Afterwards, close the gaps in every folder this run deleted from "
             "and report the old-to-new map under `renumbered`.",
    )
    p_revert.add_argument(
        "--category",
        help="Remove this Outlook category instead of outlook.category — the "
             "one the apply used.",
    )

    p_renumber = sub.add_parser(
        "renumber",
        help="Renumber one folder into date order and print the old-to-new map.",
    )
    p_renumber.add_argument(
        "--folder", required=True,
        help="The folder to renumber. Must be inside archive.root_paths.",
    )
    p_renumber.add_argument(
        "--dry-run", action="store_true",
        help="Report the same map without renaming anything.",
    )

    p_draft = sub.add_parser(
        "draft",
        help="Open a filled, unsent Outlook draft that blind-copies your own address.",
    )
    p_draft.add_argument(
        "--spec", required=True,
        help="JSON file: {to, cc, bcc, subject, body_text | body_html, "
             "attachments, ref, display}",
    )
    p_draft.add_argument(
        "--update", metavar="ENTRY_ID",
        help="Re-fill this existing unsent draft in place instead of creating one. "
             "Refused, untouched, when it is gone, sent, outside Drafts, or has no "
             "marked body region.",
    )

    p_read = sub.add_parser(
        "read",
        help="Read an unsent draft back as stored, with its fingerprint. Writes nothing.",
    )
    p_read.add_argument("--entry-id", required=True, help="The draft's EntryID.")

    p_send = sub.add_parser(
        "send",
        help="Send one unsent draft, only if its live fingerprint matches the approved one.",
    )
    p_send.add_argument("--entry-id", required=True, help="The draft's EntryID.")
    p_send.add_argument(
        "--expect-hash", required=True,
        help="The approved fingerprint hash, as `read` reported it.",
    )
    p_send.add_argument(
        "--expect-part", action="append", default=[], dest="expect_parts",
        metavar="NAME=SHA256",
        help="An approved per-part hash, so a refusal can name what differs. Repeatable.",
    )
    p_send.add_argument(
        "--expect-to", action="append", default=[],
        help="An address the To line must hold; given at all, the To line must be "
             "exactly these. Repeatable.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = _build_parser().parse_args(argv)
    verb: str = args.verb

    try:
        cfg = load_config()
    except Exception as exc:
        # Deliberately broad: this process's whole contract is "exit 2 with a
        # JSON error document when the run cannot start". A malformed YAML, an
        # unreadable path, a missing section — anything that stops the config
        # loading has to come out as that document, not as a traceback on
        # stderr with nothing on stdout, which is what a spawning caller would
        # see as a crash it cannot classify.
        return _fail(verb, ERROR_CONFIG_MISSING, f"{type(exc).__name__}: {exc}")
    setup_logging(cfg)

    # Read the caller's input before touching Outlook: a malformed file should
    # fail in milliseconds, not after a 60-second Outlook start.
    payload: list[dict[str, Any]] = []
    if verb in ("apply", "revert"):
        source = args.decisions if verb == "apply" else args.items
        try:
            payload = _load_input(source)
        except (OSError, ValueError, json.JSONDecodeError) as exc:
            return _fail(verb, ERROR_BAD_INPUT, f"{source}: {exc}")

    spec: draft.DraftSpec | None = None
    if verb == "draft":
        try:
            spec = draft.load_spec(args.spec)
        except draft.SpecError as exc:
            return _fail(verb, ERROR_BAD_INPUT, str(exc))
        if args.update is not None and not args.update.strip():
            return _fail(verb, ERROR_BAD_INPUT, "--update must not be blank")

    expectations: tuple[str, dict[str, str], list[str] | None] = ("", {}, None)
    if verb in ("read", "send"):
        if not args.entry_id.strip():
            return _fail(verb, ERROR_BAD_INPUT, "--entry-id must not be blank")
    if verb == "send":
        try:
            expectations = _send_expectations(args)
        except ValueError as exc:
            return _fail(verb, ERROR_BAD_INPUT, str(exc))

    if args.start_timeout <= 0:
        return _fail(verb, ERROR_BAD_INPUT, "--start-timeout must be positive")
    filters = batch.PlanFilters()
    if verb == "plan":
        if args.candidates < 0:
            return _fail(verb, ERROR_BAD_INPUT, "--candidates must be 0 or more")
        try:
            filters = _plan_filters(args)
        except ValueError as exc:
            return _fail(verb, ERROR_BAD_INPUT, str(exc))
    category: str | None = None
    if verb in ("apply", "revert") and args.category is not None:
        category = args.category.strip()
        if not category:
            return _fail(verb, ERROR_BAD_INPUT, "--category must not be blank")

    if verb == "renumber":
        # No Outlook, no COM: this verb only reads a folder and the index.
        document = batch.renumber(cfg, args.folder, dry_run=args.dry_run)
        if document["renumber_refused"]:
            refusal = document["renumber_refused"][0]
            return _fail(
                verb, ERROR_BAD_INPUT,
                f"{refusal['folder_path']}: {refusal['reason']}",
            )
        _emit(document)
        return EXIT_OK

    try:
        import pythoncom  # noqa: PLC0415
    except ImportError as exc:
        return _fail(
            verb, ERROR_COM_UNAVAILABLE,
            f"pywin32 is not available, so Outlook cannot be reached: {exc}",
        )

    # Single-threaded apartment on the one thread this process has. Every COM
    # call below happens here; nothing is handed to a worker.
    pythoncom.CoInitialize()
    try:
        client = OutlookClient()
        try:
            client.ensure_running(timeout=args.start_timeout)
        except OutlookUnavailableError as exc:
            return _fail(verb, ERROR_OUTLOOK_UNAVAILABLE, str(exc))

        try:
            if verb == "plan":
                document = batch.plan(
                    client, cfg, candidates=args.candidates, filters=filters
                )
            elif verb == "apply":
                document = batch.apply(
                    client, cfg, payload, renumber=args.renumber, category=category
                )
            elif verb == "draft":
                self_address = draft.resolve_self_address(cfg, client)
                if self_address is None:
                    # Before any draft exists: one without the BCC copy would
                    # never reach the Inbox, so it could never be filed.
                    return _fail(
                        verb, ERROR_SELF_ADDRESS_UNRESOLVED,
                        "No address to blind-copy: set outlook.self_address, or "
                        "make sure Outlook's default account reports an SMTP address.",
                    )
                if args.update is None:
                    document = draft.create(client, spec, self_address)
                else:
                    document = draft.update(client, args.update.strip(), spec, self_address)
            elif verb == "read":
                document = send.read_document(client.read_draft(args.entry_id.strip()))
            elif verb == "send":
                document = send.send(client, args.entry_id.strip(), *expectations)
            else:
                document = batch.revert(
                    client, cfg, payload, renumber=args.renumber, category=category
                )
        except DraftUpdateError as exc:
            # Raised before the item was changed: the caller learns which of
            # gone / not editable / unmarked it is, never a generic failure.
            return _fail(verb, exc.code, str(exc))
        except send.SendRefused as exc:
            # Raised before Send(): nothing went out, and the caller learns
            # which parts no longer match what was approved.
            return _fail(verb, exc.code, str(exc), differs=exc.differs)
        except OutlookUnavailableError as exc:
            # Outlook was up when ensure_running() checked but quit or went
            # unreachable partway through the walk (client._namespace()).
            return _fail(verb, ERROR_OUTLOOK_UNAVAILABLE, str(exc))
        except Exception as exc:
            # Deliberately broad, mirroring the config-load guard above: once
            # Outlook has been reached, any other failure during the walk (a
            # locked or corrupt database, a COM call raising something other
            # than OutlookUnavailableError) still has to come out as this
            # process's one JSON document, not as a traceback on stderr with
            # nothing on stdout, which is what a spawning caller would see as
            # a crash it cannot classify.
            return _fail(verb, ERROR_COM_UNAVAILABLE, f"{type(exc).__name__}: {exc}")
    finally:
        pythoncom.CoUninitialize()

    _emit(document)
    return EXIT_OK


if __name__ == "__main__":
    sys.exit(main())

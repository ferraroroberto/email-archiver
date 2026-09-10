"""
Headless batch entry point – plan / apply / revert / renumber, over JSON.

Meant to be spawned as a subprocess by another local app (task-os), never used
interactively: every verb prints exactly one JSON document on stdout, logs to
the usual log file and to stderr, and never opens a window.

    python main_batch.py plan --candidates 5 > plan.json
    python main_batch.py apply --decisions decisions.json [--renumber]
    python main_batch.py revert --items revert.json [--renumber]
    python main_batch.py renumber --folder "<a folder>" [--dry-run]

Exit codes:
    0  the run completed and stdout carries its document — individual mails may
       still have failed, each with its own `error` inside the document
    2  the run could not start: no config, unreadable input, a folder outside
       the archive roots, or Outlook unreachable. stdout carries
       `{"error": {"code", "message"}}` instead

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
import sys
from pathlib import Path
from typing import Any

# Ensure project root is on the path regardless of cwd
sys.path.insert(0, str(Path(__file__).resolve().parent))

from email_archiver import batch
from email_archiver.config import load_config, setup_logging
from email_archiver.outlook.client import (
    DEFAULT_START_TIMEOUT_SECONDS,
    OutlookClient,
    OutlookUnavailableError,
)

logger = logging.getLogger(__name__)

EXIT_OK = 0
EXIT_CANNOT_START = 2

ERROR_CONFIG_MISSING = "config_missing"
ERROR_BAD_INPUT = "bad_input"
ERROR_OUTLOOK_UNAVAILABLE = "outlook_unavailable"
ERROR_COM_UNAVAILABLE = "com_unavailable"


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


def _fail(verb: str, code: str, message: str) -> int:
    logger.error("%s: %s", code, message)
    _emit(batch.error_document(verb, code, message))
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


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="main_batch.py",
        description="Plan, apply or revert archiving of the whole Outlook Inbox.",
    )
    parser.add_argument(
        "--start-timeout", type=float, default=DEFAULT_START_TIMEOUT_SECONDS,
        help="Seconds to wait for a freshly started Outlook to answer over COM "
             "(default: %(default)s).",
    )
    sub = parser.add_subparsers(dest="verb", required=True)

    p_plan = sub.add_parser("plan", help="List every Inbox mail with ranked folders.")
    p_plan.add_argument(
        "--candidates", type=int, default=batch.DEFAULT_CANDIDATES,
        help="Ranked folder candidates per mail (default: %(default)s).",
    )

    p_apply = sub.add_parser("apply", help="Archive, move and tag decided mails.")
    p_apply.add_argument(
        "--decisions", required=True,
        help="JSON file: [{message_id, folder_path, date_prefix}, ...]",
    )
    p_apply.add_argument(
        "--renumber", action="store_true",
        help="Afterwards, renumber every destination folder this run wrote "
             "into and report the old-to-new map under `renumbered`.",
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

    if args.start_timeout <= 0:
        return _fail(verb, ERROR_BAD_INPUT, "--start-timeout must be positive")
    if verb == "plan" and args.candidates < 1:
        return _fail(verb, ERROR_BAD_INPUT, "--candidates must be at least 1")

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
                document = batch.plan(client, cfg, candidates=args.candidates)
            elif verb == "apply":
                document = batch.apply(client, cfg, payload, renumber=args.renumber)
            else:
                document = batch.revert(client, cfg, payload, renumber=args.renumber)
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

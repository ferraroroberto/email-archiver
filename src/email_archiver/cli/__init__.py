"""The one entry point: ``python -m email_archiver <subcommand> [args]``.

Each subcommand hands the rest of the command line to the tool that owns it, so
that tool's own parser, help text and exit codes are unchanged. The root
``main_*.py`` scripts stay as thin wrappers over the same functions, because a
scheduled job or a spawning app calls them by path.
"""
from __future__ import annotations

import sys
from collections.abc import Callable
from importlib import import_module

# subcommand -> (module under this package, one-line help)
SUBCOMMANDS = {
    "archive": ("archive", "Archive the selected Outlook email (Tkinter dialog)."),
    "scan": ("scan", "Scan and index the archive (--no-ui for headless)."),
    "ui": ("launcher", "Full launcher window with both buttons."),
    "batch": ("batch", "Headless JSON verbs: plan, apply, revert, draft, send, find, ..."),
}

_USAGE = "usage: python -m email_archiver {%s} [args ...]" % ",".join(SUBCOMMANDS)


def _exit_code(result: int | None) -> int:
    """`main` returns an exit code, or None for a window that closed normally."""
    return 0 if result is None else result


def main(argv: list[str] | None = None) -> int:
    args = sys.argv[1:] if argv is None else argv
    if not args or args[0] in ("-h", "--help"):
        print(_USAGE)
        for name, (_, helptext) in SUBCOMMANDS.items():
            print(f"  {name:<8}{helptext}")
        return 0 if args else 2
    if args[0] not in SUBCOMMANDS:
        print(f"unknown subcommand {args[0]!r}\n{_USAGE}", file=sys.stderr)
        return 2
    module = import_module(f"{__name__}.{SUBCOMMANDS[args[0]][0]}")
    run: Callable[..., int | None] = module.main
    # archive and ui take no argv; scan and batch parse their own.
    return _exit_code(run(args[1:]) if args[0] in ("scan", "batch") else run())

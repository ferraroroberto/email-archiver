"""Headless batch entry point: plan / apply / revert / renumber / draft / read / send / mailboxes / find.

    python main_batch.py plan --candidates 5 > plan.json

Kept at this path because another local app spawns it by path; the verbs, their
JSON documents and exit codes are documented in `email_archiver.cli.batch`.
"""
import sys
from pathlib import Path

# Ensure project root is on the path regardless of cwd
sys.path.insert(0, str(Path(__file__).resolve().parent))

from email_archiver.cli.batch import main

if __name__ == "__main__":
    sys.exit(main())

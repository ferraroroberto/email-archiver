"""Shared test doubles and helpers.

Every one of these existed as a byte-for-byte (or near-identical) copy in two
or more test modules before issue #63 hoisted them here.
"""
from __future__ import annotations

import sqlite3
import sys
from datetime import datetime
from pathlib import Path

import pytest

# The suite imports the package from src/ itself; only test_cli goes through the
# repo-root email_archiver.py shim, so a broken shim fails that test alone.
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from email_archiver import config  # noqa: E402


@pytest.fixture(autouse=True)
def _no_machine_local_mailbox_registry(tmp_path, monkeypatch):
    """Every test sees no config/mailboxes.json unless it writes its own:
    the live file is machine-local and must not change what a test does."""
    monkeypatch.setattr(config, "MAILBOXES_FILE", tmp_path / "no-mailboxes.json")


class _FakeAttachment:
    """A real (non-inline) attachment: no MAPI properties, so the archiver's
    inline detection falls through both branches and saves it."""

    def __init__(self, filename):
        self.FileName = filename

    @property
    def PropertyAccessor(self):
        raise AttributeError("no PropertyAccessor on this fake")

    def SaveAsFile(self, path):  # noqa: N802 - COM-shaped API
        with open(path, "w", encoding="utf-8") as fh:
            fh.write("att")


class _FakeAttachments:
    def __init__(self, items=()):
        self._items = list(items)

    def __iter__(self):
        return iter(self._items)


class _FakeMailItem:
    """Enough of a MailItem for the archiver: dates, SaveAs, attachments."""

    def __init__(
        self, sent_on=datetime(2026, 3, 14, 9, 30), attachments=(), entry_id="entry-1"
    ):
        self.SentOn = sent_on
        self.EntryID = entry_id
        self.ReceivedTime = None
        self.Attachments = _FakeAttachments(
            _FakeAttachment(n) for n in attachments
        )

    def SaveAs(self, path, fmt):  # noqa: N802 - COM-shaped API
        with open(path, "w", encoding="utf-8") as fh:
            fh.write("msg")


class _FakeMsg:
    """Stands in for ``extract_msg.Message``.

    ``messageId`` is a property on the real class and raises AttributeError
    when the message type does not carry one, which is why the scanner has a
    proptag fallback at all.
    """

    def __init__(self, message_id=..., props: dict | None = None) -> None:
        self._props = props or {}
        if message_id is not ...:
            self.messageId = message_id  # noqa: N815 - extract_msg's spelling

    def getPropertyVal(self, key: str):  # noqa: N802 - extract_msg's spelling
        return self._props.get(key)


def _columns(conn: sqlite3.Connection) -> set[str]:
    """The current column names of the `emails` table."""
    return {r["name"] for r in conn.execute("PRAGMA table_info(emails)")}

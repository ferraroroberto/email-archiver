"""Shared test doubles and helpers.

Every one of these existed as a byte-for-byte (or near-identical) copy in two
or more test modules before issue #63 hoisted them here.
"""
from __future__ import annotations

import sqlite3
from datetime import datetime


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

    def __init__(self, sent_on=datetime(2026, 3, 14, 9, 30), attachments=()):
        self.SentOn = sent_on
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

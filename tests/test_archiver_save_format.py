"""The archived ``.msg`` is saved in Outlook's Unicode format.

``olMSG`` (3) is the ANSI format and writes ``?`` for any text outside the system
code page into the permanent copy; ``olMSGUnicode`` (9) keeps it.
"""
from __future__ import annotations

from email_archiver.archiver.archiver import EmailArchiver

from .conftest import _FakeMailItem

OL_MSG_UNICODE = 9


def test_the_msg_is_saved_as_olmsgunicode(tmp_path):
    formats: list[int] = []

    class _Recording(_FakeMailItem):
        def SaveAs(self, path, fmt):  # noqa: N802 - COM-shaped API
            formats.append(fmt)
            super().SaveAs(path, fmt)

    cfg = {"path": {"max_length": 255}, "naming": {"date_prefix": False}}
    EmailArchiver(cfg).archive(_Recording(), str(tmp_path), "日本語 Ελληνικά")

    assert formats == [OL_MSG_UNICODE]

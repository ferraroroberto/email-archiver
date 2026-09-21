"""The archive dialog holds one COM apartment for its lifetime (issue #88).

``_do_archive`` re-acquires the Outlook selection on the main thread, so COM
must be initialised there. ``run()`` pairs ``CoInitialize`` with
``CoUninitialize`` around ``mainloop`` — released even when the loop raises —
instead of the click handler initialising on every archive and never releasing.
"""
from __future__ import annotations

import sys
import types

import pytest

from email_archiver.ui import app


class _RaisingRoot:
    def mainloop(self) -> None:
        raise RuntimeError("the dialog blew up")


def test_run_releases_the_apartment_it_took_even_when_mainloop_raises(monkeypatch):
    calls: list[str] = []
    fake = types.SimpleNamespace(
        CoInitialize=lambda: calls.append("init"),
        CoUninitialize=lambda: calls.append("uninit"),
    )
    monkeypatch.setitem(sys.modules, "pythoncom", fake)
    dialog = app.ArchiveDialog.__new__(app.ArchiveDialog)
    dialog._root = _RaisingRoot()

    with pytest.raises(RuntimeError):
        dialog.run()

    assert calls == ["init", "uninit"]

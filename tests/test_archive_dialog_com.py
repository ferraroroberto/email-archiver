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


def test_load_worker_reads_outlook_inside_its_own_apartment(monkeypatch):
    """The background worker is not the main thread: it needs its own apartment (#97)."""
    calls: list[str] = []
    fake = types.SimpleNamespace(
        CoInitialize=lambda: calls.append("init"),
        CoUninitialize=lambda: calls.append("uninit"),
    )
    monkeypatch.setitem(sys.modules, "pythoncom", fake)

    class _Client:
        def get_selected_email(self):
            calls.append("read")
            return "email"

    monkeypatch.setattr(app, "OutlookClient", _Client)

    assert app.ArchiveDialog._read_selected_email() == "email"
    assert calls == ["init", "read", "uninit"]


def test_load_worker_releases_its_apartment_when_the_read_raises(monkeypatch):
    calls: list[str] = []
    fake = types.SimpleNamespace(
        CoInitialize=lambda: calls.append("init"),
        CoUninitialize=lambda: calls.append("uninit"),
    )
    monkeypatch.setitem(sys.modules, "pythoncom", fake)

    class _Client:
        def get_selected_email(self):
            raise RuntimeError("boom")

    monkeypatch.setattr(app, "OutlookClient", _Client)

    with pytest.raises(RuntimeError):
        app.ArchiveDialog._read_selected_email()

    assert calls == ["init", "uninit"]

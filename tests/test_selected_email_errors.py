"""A failed Outlook read is not "no email selected" (issue #99).

``get_selected_email`` returns ``None`` only for an empty selection; every other
reason there is no mail to return raises ``SelectedEmailError`` with the reason,
and the archive dialog shows that reason instead of asking the user to select a
mail that already is selected.
"""
from __future__ import annotations

import sys
import types

import pytest

from email_archiver.outlook import client as client_mod
from email_archiver.outlook.client import OutlookClient
from email_archiver.outlook.mapi import OL_CLASS_MAIL_ITEM, SelectedEmailError
from email_archiver.ui import app

from .conftest import _FakeMailItem


@pytest.fixture
def outlook(monkeypatch):
    """A client whose Outlook is running and whose selection is ``state['item']``."""
    monkeypatch.setitem(sys.modules, "win32com", types.ModuleType("win32com"))
    monkeypatch.setitem(sys.modules, "win32com.client", types.ModuleType("client"))
    state: dict = {"running": True, "item": None, "raises": None}

    monkeypatch.setattr(OutlookClient, "is_running", lambda self: state["running"])

    def _selected():
        if state["raises"] is not None:
            raise state["raises"]
        return state["item"]

    monkeypatch.setattr(client_mod, "get_selected_mail_item", _selected)
    return state


def test_pywin32_missing_is_reported(monkeypatch):
    monkeypatch.setitem(sys.modules, "win32com", None)
    monkeypatch.setitem(sys.modules, "win32com.client", None)
    with pytest.raises(SelectedEmailError, match="pywin32"):
        OutlookClient().get_selected_email()


def test_outlook_not_running_is_reported(outlook):
    outlook["running"] = False
    with pytest.raises(SelectedEmailError, match="not running"):
        OutlookClient().get_selected_email()


def test_com_failure_carries_the_error_text(outlook):
    outlook["raises"] = RuntimeError("CoInitialize has not been called.")
    with pytest.raises(SelectedEmailError, match="CoInitialize has not been called"):
        OutlookClient().get_selected_email()


def test_empty_selection_is_none_not_an_error(outlook):
    assert OutlookClient().get_selected_email() is None


def test_non_mail_item_is_reported(outlook):
    outlook["item"] = types.SimpleNamespace(Class=OL_CLASS_MAIL_ITEM + 1)
    with pytest.raises(SelectedEmailError, match="not an email"):
        OutlookClient().get_selected_email()


def test_unreadable_metadata_is_reported(outlook):
    class _Item:
        Class = OL_CLASS_MAIL_ITEM

        @property
        def Subject(self):
            raise RuntimeError("store denied")

    outlook["item"] = _Item()
    with pytest.raises(SelectedEmailError, match="store denied"):
        OutlookClient().get_selected_email()


class _Root:
    def after(self, _ms, fn):
        fn()


def _worker_dialog(monkeypatch, read):
    shown: list[str] = []
    dialog = app.ArchiveDialog.__new__(app.ArchiveDialog)
    dialog._root = _Root()
    dialog._show_error = shown.append
    monkeypatch.setattr(app.ArchiveDialog, "_read_selected_email", staticmethod(read))
    return dialog, shown


def test_worker_shows_the_reason_for_a_failed_read(monkeypatch):
    def read():
        raise SelectedEmailError("Cannot reach Outlook: boom")

    dialog, shown = _worker_dialog(monkeypatch, read)
    dialog._load_worker()
    assert shown == ["Cannot reach Outlook: boom"]


def test_worker_asks_for_a_selection_only_when_nothing_is_selected(monkeypatch):
    dialog, shown = _worker_dialog(monkeypatch, lambda: None)
    dialog._load_worker()
    assert len(shown) == 1 and "No email selected" in shown[0]


# ---------------------------------------- the Archive click vs. the selection ---


class _SelectionRoot:
    destroyed = False

    def destroy(self) -> None:
        self.destroyed = True


class _ShownEmail:
    subject = "First mail"
    entry_id = "entry-first"


class _NoDatePrefix:
    def get(self) -> bool:
        return False


@pytest.fixture
def archive_dialog(monkeypatch):
    """A dialog showing ``First mail``; the Outlook selection is set per test."""
    dialog = app.ArchiveDialog.__new__(app.ArchiveDialog)
    dialog._cfg = {"path": {"max_length": 255}, "naming": {"date_prefix": "off"}}
    dialog._email = _ShownEmail()
    dialog._root = _SelectionRoot()
    dialog._date_prefix_var = _NoDatePrefix()
    dialog._date_prefix_set_by_hand = False
    errors: list = []
    monkeypatch.setattr(app.messagebox, "showerror", lambda *a, **k: errors.append(a))
    return dialog, errors


def test_a_selection_moved_in_outlook_is_refused_and_nothing_is_written(
    archive_dialog, monkeypatch, tmp_path
):
    dialog, errors = archive_dialog
    other = _FakeMailItem(entry_id="entry-second")
    monkeypatch.setattr(app, "get_selected_mail_item", lambda: other)

    dialog._do_archive(str(tmp_path))

    assert list(tmp_path.iterdir()) == []
    assert len(errors) == 1 and "Re-select" in errors[0][1]
    assert dialog._root.destroyed is False


def test_the_same_mail_still_selected_is_archived(archive_dialog, monkeypatch, tmp_path):
    dialog, errors = archive_dialog
    same = _FakeMailItem(entry_id="entry-first")
    monkeypatch.setattr(app, "get_selected_mail_item", lambda: same)

    dialog._do_archive(str(tmp_path))

    assert errors == []
    assert [p.name.endswith("001 - First mail.msg") for p in tmp_path.iterdir()] == [True]
    assert dialog._root.destroyed is True


def test_the_entry_id_of_the_selected_mail_is_read(outlook):
    outlook["item"] = types.SimpleNamespace(
        Class=OL_CLASS_MAIL_ITEM,
        Subject="Hello",
        EntryID="entry-xyz",
        SenderEmailAddress="a@example.invalid",
        SenderEmailType="SMTP",
        Recipients=[],
        SentOn=None,
        ReceivedTime=None,
    )

    assert OutlookClient().get_selected_email().entry_id == "entry-xyz"

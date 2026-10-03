"""
Tests for OutlookClient.release_folder_hold() — the folder switch behind
``move_via: folder_switch`` (issue #111).

An item the Explorer shows as its selected / reading-pane mail can refuse every
write, a ``Move`` included, until the Explorer lets go of it. Selecting another
folder and the message again frees it, so the client switches the Explorer away
and back. What matters here is the user's view: it is only moved when it is
showing the mail's own folder, it is always put back — even when the switch
fails — and the three answers (switched / not applicable / not determined)
stay distinct. Nothing here reaches a real Outlook.
"""
from __future__ import annotations

import pytest

from email_archiver.outlook import client as client_module
from email_archiver.outlook import process
from email_archiver.outlook.client import OutlookClient
from email_archiver.outlook.mapi import OL_FOLDER_DRAFTS, OL_FOLDER_OUTBOX


class _Folder:
    def __init__(self, entry_id: str) -> None:
        self.EntryID = entry_id


class _Item:
    def __init__(self, parent: _Folder | None) -> None:
        self.Parent = parent


class _Explorer:
    """Records every folder it is pointed at, so a test can read the round trip."""

    def __init__(self, current: _Folder, fail_on: str | None = None) -> None:
        self._current = current
        self._fail_on = fail_on
        self.visited: list[str] = []

    @property
    def CurrentFolder(self) -> _Folder:  # noqa: N802 - COM's spelling
        return self._current

    @CurrentFolder.setter
    def CurrentFolder(self, folder: _Folder) -> None:  # noqa: N802
        if self._fail_on == folder.EntryID:
            raise RuntimeError("Outlook would not change folder")
        self.visited.append(folder.EntryID)
        self._current = folder


class _Namespace:
    def __init__(self, defaults: dict[int, _Folder]) -> None:
        self._defaults = defaults

    def GetDefaultFolder(self, which: int) -> _Folder:  # noqa: N802
        return self._defaults[which]


class _App:
    def __init__(self, explorer: _Explorer | None, namespace: _Namespace) -> None:
        self._explorer = explorer
        self._namespace = namespace

    def ActiveExplorer(self) -> _Explorer | None:  # noqa: N802
        return self._explorer

    def GetNamespace(self, name: str) -> _Namespace:  # noqa: N802
        return self._namespace


INBOX = _Folder("inbox")
OUTBOX = _Folder("outbox")
DRAFTS = _Folder("drafts")


def _with_app(monkeypatch, explorer: _Explorer | None) -> None:
    namespace = _Namespace({OL_FOLDER_OUTBOX: OUTBOX, OL_FOLDER_DRAFTS: DRAFTS})
    monkeypatch.setattr(
        process, "get_active_application", lambda: _App(explorer, namespace)
    )
    monkeypatch.setattr(client_module, "EXPLORER_SETTLE_SECONDS", 0)


def test_switches_away_and_back_when_the_explorer_shows_the_mails_folder(monkeypatch):
    explorer = _Explorer(INBOX)
    _with_app(monkeypatch, explorer)

    assert OutlookClient().release_folder_hold(_Item(INBOX)) is True

    assert explorer.visited == ["outbox", "inbox"]
    assert explorer.CurrentFolder is INBOX


def test_leaves_an_explorer_showing_another_folder_alone(monkeypatch):
    """Another folder is not holding the mail, and moving the user's view for
    nothing is a cost: nothing is done and the answer is a plain no."""
    explorer = _Explorer(_Folder("elsewhere"))
    _with_app(monkeypatch, explorer)

    assert OutlookClient().release_folder_hold(_Item(INBOX)) is False

    assert explorer.visited == []


def test_no_explorer_means_nothing_holds_the_mail_there(monkeypatch):
    _with_app(monkeypatch, None)

    assert OutlookClient().release_folder_hold(_Item(INBOX)) is False


def test_an_unreachable_outlook_is_not_determined(monkeypatch):
    monkeypatch.setattr(process, "get_active_application", lambda: None)

    assert OutlookClient().release_folder_hold(_Item(INBOX)) is None


@pytest.mark.parametrize(
    "item",
    [_Item(None), _Item(_Folder(""))],
    ids=["no_parent", "no_entry_id"],
)
def test_a_mail_whose_folder_cannot_be_read_is_not_determined(monkeypatch, item):
    """Not knowing which folder the mail is in is not a reason to move the
    view — and not evidence the Explorer is elsewhere either."""
    explorer = _Explorer(INBOX)
    _with_app(monkeypatch, explorer)

    assert OutlookClient().release_folder_hold(item) is None
    assert explorer.visited == []


def test_the_original_folder_is_restored_when_the_switch_away_fails(monkeypatch):
    """The folder is put back even on failure, and the failed switch is
    reported as not determined — it is not evidence the hold survives one."""
    explorer = _Explorer(INBOX, fail_on="outbox")
    _with_app(monkeypatch, explorer)

    assert OutlookClient().release_folder_hold(_Item(INBOX)) is None

    assert explorer.CurrentFolder is INBOX
    assert explorer.visited == ["inbox"]


def test_a_failed_restore_does_not_raise_and_is_logged(monkeypatch, caplog):
    """A restore Outlook refuses cannot be retried into working, and the hold
    was released by then — so the move still gets its retry, with the failure
    on the log for whoever finds the view on the wrong folder."""
    explorer = _Explorer(INBOX, fail_on="inbox")
    _with_app(monkeypatch, explorer)

    with caplog.at_level("ERROR"):
        assert OutlookClient().release_folder_hold(_Item(INBOX)) is True

    assert "put Outlook back" in caplog.text


def test_parks_in_drafts_when_the_mail_lives_in_the_outbox(monkeypatch):
    """The folder parked in must be a different one from the mail's own."""
    explorer = _Explorer(OUTBOX)
    _with_app(monkeypatch, explorer)

    assert OutlookClient().release_folder_hold(_Item(OUTBOX)) is True

    assert explorer.visited == ["drafts", "outbox"]

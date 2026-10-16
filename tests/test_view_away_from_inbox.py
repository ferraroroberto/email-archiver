"""
Tests for OutlookClient.view_away_from_inbox() — the folder switch behind
``move_via: folder_switch`` (issues #111, #130).

An item the Explorer shows as its selected / reading-pane mail can refuse every
write, a ``Move`` included, until the Explorer lets go of it. Selecting another
folder frees it, and coming back to the Inbox selects it again — so the client
parks the Explorer elsewhere for the length of a ``with`` block, and batch mode
moves the mail inside it. What matters here is the user's view: it is only
moved when it is showing the Inbox, it is always put back — even when the body
raises — and the three answers (parked / not applicable / not determined) stay
distinct. Nothing here reaches a real Outlook.
"""
from __future__ import annotations

import pytest

from email_archiver.outlook import client as client_module
from email_archiver.outlook import process
from email_archiver.outlook.client import OutlookClient
from email_archiver.outlook.mapi import OL_FOLDER_DRAFTS, OL_FOLDER_INBOX, OL_FOLDER_OUTBOX


class _Folder:
    def __init__(self, entry_id: str) -> None:
        self.EntryID = entry_id


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

    @property
    def DefaultStore(self):  # noqa: N802
        # The default store's folders are this namespace's (issue #108).
        return self

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


def _with_app(monkeypatch, explorer: _Explorer | None, outbox: _Folder = OUTBOX) -> None:
    namespace = _Namespace({
        OL_FOLDER_INBOX: INBOX, OL_FOLDER_OUTBOX: outbox, OL_FOLDER_DRAFTS: DRAFTS,
    })
    monkeypatch.setattr(
        process, "get_active_application", lambda: _App(explorer, namespace)
    )
    monkeypatch.setattr(client_module, "EXPLORER_SETTLE_SECONDS", 0)


def test_parks_the_view_for_the_block_and_puts_it_back(monkeypatch):
    """The move is made inside the block, so the view must still be away there
    — coming back first is what re-selected the mail (issue #130)."""
    explorer = _Explorer(INBOX)
    _with_app(monkeypatch, explorer)

    with OutlookClient().view_away_from_inbox() as parked:
        assert parked is True
        assert explorer.CurrentFolder is OUTBOX

    assert explorer.visited == ["outbox", "inbox"]
    assert explorer.CurrentFolder is INBOX


def test_the_view_is_put_back_when_the_block_raises(monkeypatch):
    explorer = _Explorer(INBOX)
    _with_app(monkeypatch, explorer)

    with pytest.raises(RuntimeError):
        with OutlookClient().view_away_from_inbox():
            raise RuntimeError("the store refused the move")

    assert explorer.CurrentFolder is INBOX


def test_leaves_an_explorer_showing_another_folder_alone(monkeypatch):
    """Another folder is not holding the mail, and moving the user's view for
    nothing is a cost: nothing is done and the answer is a plain no."""
    explorer = _Explorer(_Folder("elsewhere"))
    _with_app(monkeypatch, explorer)

    with OutlookClient().view_away_from_inbox() as parked:
        assert parked is False

    assert explorer.visited == []


def test_no_explorer_means_nothing_holds_the_mail_there(monkeypatch):
    _with_app(monkeypatch, None)

    with OutlookClient().view_away_from_inbox() as parked:
        assert parked is False


def test_an_unreachable_outlook_is_not_determined(monkeypatch):
    monkeypatch.setattr(process, "get_active_application", lambda: None)

    with OutlookClient().view_away_from_inbox() as parked:
        assert parked is None


def test_an_explorer_whose_folder_cannot_be_read_is_not_determined(monkeypatch):
    """Not knowing which folder is on screen is not a reason to move the view
    — and not evidence the Explorer is elsewhere either."""
    explorer = _Explorer(_Folder(""))
    _with_app(monkeypatch, explorer)

    with OutlookClient().view_away_from_inbox() as parked:
        assert parked is None

    assert explorer.visited == []


def test_a_failed_switch_away_is_not_determined_and_moves_nothing(monkeypatch):
    """A switch that did not happen is reported as not determined — it is not
    evidence the hold survives one — and the view is left where it was."""
    explorer = _Explorer(INBOX, fail_on="outbox")
    _with_app(monkeypatch, explorer)

    with OutlookClient().view_away_from_inbox() as parked:
        assert parked is None

    assert explorer.CurrentFolder is INBOX
    assert explorer.visited == []


def test_a_failed_restore_does_not_raise_and_is_logged(monkeypatch, caplog):
    """A restore Outlook refuses cannot be retried into working, and the move
    was already made — so it is not raised over the result, but logged at
    error level for whoever finds the view on the wrong folder."""
    explorer = _Explorer(INBOX, fail_on="inbox")
    _with_app(monkeypatch, explorer)

    with caplog.at_level("ERROR"):
        with OutlookClient().view_away_from_inbox() as parked:
            assert parked is True

    assert "put Outlook back" in caplog.text


def test_parks_in_drafts_when_the_mailbox_has_no_outbox_of_its_own(monkeypatch):
    """The folder parked in must be a different one from the Inbox."""
    explorer = _Explorer(INBOX)
    _with_app(monkeypatch, explorer, outbox=INBOX)

    with OutlookClient().view_away_from_inbox() as parked:
        assert parked is True

    assert explorer.visited == ["drafts", "inbox"]

"""
Tests for OutlookClient.is_open_in_inspector() — the check behind the
`message_changed` error code (issue #76).

An Outlook inspector holds its item for the life of the window and refuses
every write to it, which is how a mail survives both of batch mode's move
defences. The check exists to name that cause in the result, so what matters
here is that it answers three states and never two: open, not open, and *not
determined*. A window that could not be read must come back `None` — reporting
it as "no window" sends an operator off to restart Outlook when closing one
window is the actual fix.
"""
from __future__ import annotations

import pytest

from email_archiver.outlook import client as client_mod
from email_archiver.outlook.client import OutlookClient


class _FakeItem:
    def __init__(self, entry_id: str) -> None:
        self.EntryID = entry_id


class _FakeUnreadable:
    """A COM object whose every property read raises, as a disconnected one
    does. ``_safe_com`` is what is supposed to absorb this."""

    @property
    def EntryID(self) -> str:  # noqa: N802 - COM's spelling
        raise RuntimeError("the object is disconnected")


class _FakeInspector:
    def __init__(self, item, raises: bool = False) -> None:
        self._item = item
        self._raises = raises

    @property
    def CurrentItem(self):  # noqa: N802 - COM's spelling
        if self._raises:
            raise RuntimeError("the window would not hand over its item")
        return self._item


class _FakeInspectors:
    def __init__(self, inspectors: list, count_raises: bool = False) -> None:
        self._inspectors = inspectors
        self._count_raises = count_raises

    @property
    def Count(self) -> int:  # noqa: N802 - COM's spelling
        if self._count_raises:
            raise RuntimeError("Outlook would not list its windows")
        return len(self._inspectors)

    def Item(self, index: int):  # noqa: N802 - COM's spelling
        return self._inspectors[index - 1]


class _FakeApplication:
    def __init__(self, inspectors: _FakeInspectors) -> None:
        self.Inspectors = inspectors


@pytest.fixture
def client() -> OutlookClient:
    return OutlookClient()


def _with_app(monkeypatch, app) -> None:
    monkeypatch.setattr(client_mod, "_get_active_application", lambda: app)


def _app(*inspectors, count_raises: bool = False) -> _FakeApplication:
    return _FakeApplication(_FakeInspectors(list(inspectors), count_raises))


def test_finds_the_mail_open_in_a_window(monkeypatch, client):
    """The case the error message leads with: this mail is the one on screen."""
    _with_app(monkeypatch, _app(
        _FakeInspector(_FakeItem("other-mail")),
        _FakeInspector(_FakeItem("the-mail")),
    ))
    assert client.is_open_in_inspector(_FakeItem("the-mail")) is True


def test_reports_no_window_when_every_inspector_was_read(monkeypatch, client):
    """Only a complete read earns a ``False`` — that answer is what sends the
    operator to a restart, so it has to be a finding, not a default."""
    _with_app(monkeypatch, _app(_FakeInspector(_FakeItem("other-mail"))))
    assert client.is_open_in_inspector(_FakeItem("the-mail")) is False


def test_no_windows_open_at_all_is_a_clean_no(monkeypatch, client):
    _with_app(monkeypatch, _app())
    assert client.is_open_in_inspector(_FakeItem("the-mail")) is False


def test_an_unsaved_compose_window_is_not_a_gap(monkeypatch, client):
    """A compose window has no EntryID, so it cannot be the mail being filed.
    Skipping it must not downgrade the answer to "not determined"."""
    _with_app(monkeypatch, _app(
        _FakeInspector(_FakeUnreadable()),          # an unsaved draft
        _FakeInspector(_FakeItem("other-mail")),
    ))
    assert client.is_open_in_inspector(_FakeItem("the-mail")) is False


def test_a_window_that_will_not_be_read_makes_the_answer_undetermined(
    monkeypatch, client
):
    """One unreadable window means "none of them holds it" is a guess. The
    honest answer is ``None``, never ``False``."""
    _with_app(monkeypatch, _app(
        _FakeInspector(None, raises=True),
        _FakeInspector(_FakeItem("other-mail")),
    ))
    assert client.is_open_in_inspector(_FakeItem("the-mail")) is None


def test_outlook_that_will_not_list_its_windows_is_undetermined(monkeypatch, client):
    _with_app(monkeypatch, _app(count_raises=True))
    assert client.is_open_in_inspector(_FakeItem("the-mail")) is None


def test_outlook_unreachable_is_undetermined(monkeypatch, client):
    _with_app(monkeypatch, None)
    assert client.is_open_in_inspector(_FakeItem("the-mail")) is None


def test_a_mail_with_no_readable_entry_id_is_undetermined(monkeypatch, client):
    """Nothing to compare against, so nothing can be concluded."""
    _with_app(monkeypatch, _app(_FakeInspector(_FakeItem("the-mail"))))
    assert client.is_open_in_inspector(_FakeUnreadable()) is None

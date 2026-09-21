"""The archive dialog's date-prefix checkbox against ``naming.date_prefix: auto``.

The checkbox is only a preview until the user flips it: a hover over a
suggestion card sets it for *that* card's folder, and the seed value belongs to
no folder at all. Archiving through *Browse folder…* or *Explorer folder* must
therefore infer the form from the folder actually chosen, not carry the preview
into it (issue #89). No Tk window is built: the dialog's state is set by hand
and ``_do_archive`` runs against a fake mail item.
"""
from __future__ import annotations

import pytest

from email_archiver.ui import app

from .conftest import _FakeMailItem


class _FakeVar:
    def __init__(self, value: bool) -> None:
        self._value = value

    def get(self) -> bool:
        return self._value

    def set(self, value: bool) -> None:
        self._value = value


class _FakeRoot:
    def destroy(self) -> None:
        pass


class _FakeEmail:
    subject = "Next one"


@pytest.fixture
def dialog(monkeypatch):
    cfg = {"path": {"max_length": 255}, "naming": {"date_prefix": "auto"}}
    d = app.ArchiveDialog.__new__(app.ArchiveDialog)
    d._cfg = cfg
    d._email = _FakeEmail()
    d._root = _FakeRoot()
    d._date_prefix_var = _FakeVar(app.get_date_prefix_enabled(cfg))
    d._date_prefix_set_by_hand = False
    monkeypatch.setattr(app, "get_selected_mail_item", lambda: _FakeMailItem())
    monkeypatch.setattr(app.messagebox, "showerror", lambda *a, **k: pytest.fail(a))
    return d


def _dated(folder):
    folder.mkdir()
    (folder / "2026-01-01 - 001 - earlier.msg").write_text("x")
    (folder / "2026-01-02 - 002 - earlier2.msg").write_text("x")
    return folder


def _undated(folder):
    folder.mkdir()
    (folder / "001 - earlier.msg").write_text("x")
    return folder


def _written(folder):
    return sorted(p.name for p in folder.iterdir())


def test_a_hover_preview_is_not_carried_into_a_browsed_undated_folder(dialog, tmp_path):
    """Hover a dated suggestion, then file through Browse into an undated one:
    the undated folder keeps its own form."""
    dialog._on_suggestion_highlighted(str(_dated(tmp_path / "dated")))
    assert dialog._date_prefix_var.get() is True

    target = _undated(tmp_path / "undated")
    dialog._do_archive(str(target))

    assert "002 - Next one.msg" in _written(target)


def test_the_seed_value_is_not_carried_into_a_dated_explorer_folder(dialog, tmp_path):
    """No card hovered at all: the seed (False under auto) must not strip the
    date from a folder that files everything dated."""
    target = _dated(tmp_path / "dated")
    dialog._do_archive(str(target))

    assert "2026-03-14 - 003 - Next one.msg" in _written(target)


def test_a_checkbox_flipped_by_hand_still_overrides_auto(dialog, tmp_path):
    """The user's own choice stays an override, as before."""
    target = _undated(tmp_path / "undated")
    dialog._date_prefix_var.set(True)
    dialog._on_date_prefix_toggled()
    dialog._do_archive(str(target))

    assert "2026-03-14 - 002 - Next one.msg" in _written(target)

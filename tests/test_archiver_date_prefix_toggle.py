"""Tests for the naming.date_prefix toggle end-to-end through EmailArchiver."""
from __future__ import annotations

from datetime import datetime

from email_archiver.archiver.archiver import EmailArchiver
from email_archiver.config import get_date_prefix_enabled


class _FakeAttachments:
    def __iter__(self):
        return iter(())


class _FakeMailItem:
    """Enough of a MailItem for the archiver: dates, SaveAs, no attachments."""

    def __init__(self, sent_on=datetime(2026, 3, 14, 9, 30)):
        self.SentOn = sent_on
        self.ReceivedTime = None
        self.Attachments = _FakeAttachments()

    def SaveAs(self, path, fmt):  # noqa: N802 - COM-shaped API
        with open(path, "w", encoding="utf-8") as fh:
            fh.write("msg")


def _cfg(**naming):
    cfg = {"path": {"max_length": 255}}
    if naming:
        cfg["naming"] = naming
    return cfg


# ------------------------------------------------------- config accessor ----

def test_absent_naming_section_defaults_to_off():
    assert get_date_prefix_enabled({}) is False


def test_absent_key_in_present_section_defaults_to_off():
    assert get_date_prefix_enabled({"naming": {}}) is False


def test_explicit_true_is_on():
    assert get_date_prefix_enabled({"naming": {"date_prefix": True}}) is True


def test_explicit_false_is_off():
    assert get_date_prefix_enabled({"naming": {"date_prefix": False}}) is False


# ------------------------------------------------------------- archiving ----

def test_default_config_writes_the_undated_form(tmp_path):
    result = EmailArchiver(_cfg()).archive(
        _FakeMailItem(), str(tmp_path), "Project Alpha meeting notes"
    )
    assert result.sequence_number == "001"
    assert result.email_path.endswith("001 - Project Alpha meeting notes.msg")


def test_toggle_on_writes_the_dated_form(tmp_path):
    result = EmailArchiver(_cfg(date_prefix=True)).archive(
        _FakeMailItem(), str(tmp_path), "Project Alpha meeting notes"
    )
    assert result.email_path.endswith(
        "2026-03-14 - 001 - Project Alpha meeting notes.msg"
    )


def test_toggle_on_with_no_resolvable_date_falls_back_to_undated(tmp_path, caplog):
    with caplog.at_level("WARNING"):
        result = EmailArchiver(_cfg(date_prefix=True)).archive(
            _FakeMailItem(sent_on=None), str(tmp_path), "No date here"
        )
    assert result.email_path.endswith("001 - No date here.msg")
    assert any("sent date" in rec.message for rec in caplog.records)


def test_toggle_off_does_not_even_read_the_sent_date(tmp_path):
    # The date must not be consulted when the toggle is off: an item whose date
    # properties explode would otherwise fail an archive that never needed them.
    class _ExplodingDates:
        Attachments = _FakeAttachments()

        @property
        def SentOn(self):  # noqa: N802 - COM-shaped API
            raise AssertionError("SentOn must not be read when the toggle is off")

        @property
        def ReceivedTime(self):  # noqa: N802 - COM-shaped API
            raise AssertionError("ReceivedTime must not be read when the toggle is off")

        def SaveAs(self, path, fmt):  # noqa: N802 - COM-shaped API
            with open(path, "w", encoding="utf-8") as fh:
                fh.write("msg")

    result = EmailArchiver(_cfg()).archive(_ExplodingDates(), str(tmp_path), "Quiet")
    assert result.email_path.endswith("001 - Quiet.msg")


def test_explicit_override_beats_the_config_in_both_directions(tmp_path):
    # The dialog's checkbox passes its state through this override, so a ticked
    # box must win over a config saying off, and an unticked box over one on.
    on = EmailArchiver(_cfg(), date_prefix=True).archive(
        _FakeMailItem(), str(tmp_path / "a"), "Ticked"
    )
    assert on.email_path.endswith("2026-03-14 - 001 - Ticked.msg")

    off = EmailArchiver(_cfg(date_prefix=True), date_prefix=False).archive(
        _FakeMailItem(), str(tmp_path / "b"), "Unticked"
    )
    assert off.email_path.endswith("001 - Unticked.msg")


def test_override_none_defers_to_the_config(tmp_path):
    result = EmailArchiver(_cfg(date_prefix=True), date_prefix=None).archive(
        _FakeMailItem(), str(tmp_path), "Deferred"
    )
    assert result.email_path.endswith("2026-03-14 - 001 - Deferred.msg")


def test_toggle_off_still_continues_a_dated_folders_sequence(tmp_path):
    # Flipping the toggle off in a folder already holding dated files must not
    # restart the sequence at 001 and collide.
    (tmp_path / "2026-03-14 - 007 - earlier.msg").write_text("x")
    result = EmailArchiver(_cfg()).archive(_FakeMailItem(), str(tmp_path), "Next one")
    assert result.email_path.endswith("008 - Next one.msg")


def test_toggle_on_still_continues_an_undated_folders_sequence(tmp_path):
    (tmp_path / "012 - earlier.msg").write_text("x")
    result = EmailArchiver(_cfg(date_prefix=True)).archive(
        _FakeMailItem(), str(tmp_path), "Next one"
    )
    assert result.email_path.endswith("2026-03-14 - 013 - Next one.msg")

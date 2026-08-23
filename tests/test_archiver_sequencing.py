"""Tests for sequence-number derivation and sent-date resolution."""
from __future__ import annotations

from datetime import datetime

from email_archiver.archiver.archiver import (
    _get_sent_date_prefix,
    get_next_sequence_number,
)


class _FakeMailItem:
    """Stand-in for the Outlook COM MailItem's date properties."""

    def __init__(self, sent_on=None, received_time=None, raise_on_sent_on=False):
        self._sent_on = sent_on
        self._received_time = received_time
        self._raise_on_sent_on = raise_on_sent_on

    @property
    def SentOn(self):
        if self._raise_on_sent_on:
            raise AttributeError("SentOn not available")
        return self._sent_on

    @property
    def ReceivedTime(self):
        return self._received_time


# ------------------------------------------------------- get_next_sequence_number ---

def test_empty_folder_starts_at_001(tmp_path):
    assert get_next_sequence_number(str(tmp_path)) == "001"


def test_undated_only_folder_picks_max_plus_one(tmp_path):
    (tmp_path / "001 - alpha.msg").write_text("x")
    (tmp_path / "007 - beta.pdf").write_text("x")
    assert get_next_sequence_number(str(tmp_path)) == "008"


def test_dated_only_folder_picks_max_plus_one(tmp_path):
    (tmp_path / "2026-01-01 - 003 - alpha.msg").write_text("x")
    (tmp_path / "2026-03-14 - 012 - beta.pdf").write_text("x")
    assert get_next_sequence_number(str(tmp_path)) == "013"


def test_mixed_undated_and_dated_folder_never_collides(tmp_path):
    (tmp_path / "023 - old_email.msg").write_text("x")
    (tmp_path / "2026-03-14 - 024 - new_email.msg").write_text("x")
    assert get_next_sequence_number(str(tmp_path)) == "025"


def test_dated_prefix_year_is_not_mistaken_for_sequence(tmp_path):
    # A naive "leading digits" scan would read "2026" out of the date prefix
    # as the sequence number. It must not.
    (tmp_path / "2026-03-14 - 001 - only_file.msg").write_text("x")
    assert get_next_sequence_number(str(tmp_path)) == "002"


# ------------------------------------------------------------- _get_sent_date_prefix ---

def test_sent_on_is_preferred():
    item = _FakeMailItem(
        sent_on=datetime(2026, 3, 14, 9, 30),
        received_time=datetime(2026, 3, 15, 8, 0),
    )
    assert _get_sent_date_prefix(item) == "2026-03-14"


def test_falls_back_to_received_time_when_sent_on_unavailable():
    item = _FakeMailItem(
        received_time=datetime(2026, 3, 15, 8, 0),
        raise_on_sent_on=True,
    )
    assert _get_sent_date_prefix(item) == "2026-03-15"


def test_returns_none_and_logs_when_neither_date_resolves(caplog):
    item = _FakeMailItem(sent_on=None, received_time=None)
    with caplog.at_level("WARNING"):
        result = _get_sent_date_prefix(item)
    assert result is None
    assert any("sent date" in rec.message for rec in caplog.records)

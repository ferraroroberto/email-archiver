"""Tests for the naming.date_prefix toggle end-to-end through EmailArchiver."""
from __future__ import annotations

import os

from email_archiver.archiver.archiver import EmailArchiver, resolve_date_prefix_for_folder
from email_archiver.config import get_date_prefix_enabled, get_date_prefix_mode

from .conftest import _FakeAttachments, _FakeMailItem


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


def test_mode_returns_true_false_or_auto():
    assert get_date_prefix_mode({}) is False
    assert get_date_prefix_mode({"naming": {"date_prefix": True}}) is True
    assert get_date_prefix_mode({"naming": {"date_prefix": False}}) is False
    assert get_date_prefix_mode({"naming": {"date_prefix": "auto"}}) == "auto"
    # Case-insensitive, since it comes from hand-edited YAML.
    assert get_date_prefix_mode({"naming": {"date_prefix": "AUTO"}}) == "auto"


def test_enabled_resolves_auto_to_false():
    # get_date_prefix_enabled has no folder to infer from, so "auto" is a
    # starting position of False — the UI resolves the real value per
    # suggestion via resolve_date_prefix_for_folder instead.
    assert get_date_prefix_enabled({"naming": {"date_prefix": "auto"}}) is False


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


def test_the_whole_bundle_shares_one_prefix_with_the_toggle_off(tmp_path):
    result = EmailArchiver(_cfg()).archive(
        _FakeMailItem(attachments=("invoice.pdf", "signed_contract.docx")),
        str(tmp_path),
        "Project Alpha",
    )
    names = sorted(os.path.basename(p) for p in
                   [result.email_path, *result.attachment_paths])
    assert names == [
        "001 - Project Alpha.msg",
        "001 - invoice.pdf",
        "001 - signed_contract.docx",
    ]


def test_the_whole_bundle_shares_one_date_and_seq_with_the_toggle_on(tmp_path):
    result = EmailArchiver(_cfg(date_prefix=True)).archive(
        _FakeMailItem(attachments=("invoice.pdf", "signed_contract.docx")),
        str(tmp_path),
        "Project Alpha",
    )
    names = sorted(os.path.basename(p) for p in
                   [result.email_path, *result.attachment_paths])
    assert names == [
        "2026-03-14 - 001 - Project Alpha.msg",
        "2026-03-14 - 001 - invoice.pdf",
        "2026-03-14 - 001 - signed_contract.docx",
    ]


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


# --------------------------------------------------- naming.date_prefix: auto ----

def test_auto_writes_the_dated_form_into_an_already_dated_folder(tmp_path):
    # Acceptance criterion: auto writes the dated form with no checkbox/
    # explicit-argument interaction at all — the config alone decides.
    (tmp_path / "2026-01-01 - 001 - earlier.msg").write_text("x")
    (tmp_path / "2026-01-02 - 002 - earlier2.msg").write_text("x")
    result = EmailArchiver(_cfg(date_prefix="auto")).archive(
        _FakeMailItem(), str(tmp_path), "Next one"
    )
    assert result.email_path.endswith("2026-03-14 - 003 - Next one.msg")


def test_auto_writes_the_undated_form_into_an_already_undated_folder(tmp_path):
    (tmp_path / "001 - earlier.msg").write_text("x")
    result = EmailArchiver(_cfg(date_prefix="auto")).archive(
        _FakeMailItem(), str(tmp_path), "Next one"
    )
    assert result.email_path.endswith("002 - Next one.msg")


def test_auto_falls_back_to_undated_on_an_empty_folder(tmp_path):
    result = EmailArchiver(_cfg(date_prefix="auto")).archive(
        _FakeMailItem(), str(tmp_path), "First one"
    )
    assert result.email_path.endswith("001 - First one.msg")


def test_auto_falls_back_to_undated_on_a_tied_folder(tmp_path):
    (tmp_path / "001 - a.msg").write_text("x")
    (tmp_path / "2026-01-01 - 002 - b.msg").write_text("x")
    result = EmailArchiver(_cfg(date_prefix="auto")).archive(
        _FakeMailItem(), str(tmp_path), "Next one"
    )
    assert result.email_path.endswith("003 - Next one.msg")


# --------------------------------------------------------------- precedence ----
# explicit constructor override > config "auto" (per-folder inference) >
# config boolean.

def test_explicit_override_beats_auto_inference(tmp_path):
    (tmp_path / "2026-01-01 - 001 - earlier.msg").write_text("x")
    # Folder is majority-dated, but the explicit override says no prefix.
    result = EmailArchiver(_cfg(date_prefix="auto"), date_prefix=False).archive(
        _FakeMailItem(), str(tmp_path), "Overridden"
    )
    assert result.email_path.endswith("002 - Overridden.msg")


def test_explicit_override_true_beats_auto_inference_toward_dated(tmp_path):
    (tmp_path / "001 - earlier.msg").write_text("x")
    # Folder is majority-undated, but the explicit override forces the date.
    result = EmailArchiver(_cfg(date_prefix="auto"), date_prefix=True).archive(
        _FakeMailItem(), str(tmp_path), "Overridden"
    )
    assert result.email_path.endswith("2026-03-14 - 002 - Overridden.msg")


def test_auto_beats_the_config_boolean_it_replaces(tmp_path):
    # Sanity check on precedence order, not just presence: "auto" governs
    # once selected, the (now-irrelevant) DEFAULT_DATE_PREFIX_ENABLED value
    # never leaks in as a silent third form.
    (tmp_path / "2026-01-01 - 001 - earlier.msg").write_text("x")
    result = EmailArchiver(_cfg(date_prefix="auto")).archive(
        _FakeMailItem(), str(tmp_path), "Next one"
    )
    assert "2026-03-14" in result.email_path


# ---------------------------------------------- resolve_date_prefix_for_folder ----
# The dialog (hover pre-tick) and batch plan (per-candidate date_prefix) both
# go through this to answer "what would EmailArchiver pick for this folder,
# absent an explicit override" without constructing an EmailArchiver at all.

def test_resolve_for_folder_infers_in_auto_mode(tmp_path):
    (tmp_path / "2026-01-01 - 001 - a.msg").write_text("x")
    (tmp_path / "2026-01-02 - 002 - b.msg").write_text("x")
    assert resolve_date_prefix_for_folder(_cfg(date_prefix="auto"), str(tmp_path)) is True


def test_resolve_for_folder_falls_back_to_false_on_ambiguous_auto(tmp_path):
    assert resolve_date_prefix_for_folder(_cfg(date_prefix="auto"), str(tmp_path)) is False


def test_resolve_for_folder_ignores_folder_contents_when_not_auto(tmp_path):
    (tmp_path / "2026-01-01 - 001 - a.msg").write_text("x")
    (tmp_path / "2026-01-02 - 002 - b.msg").write_text("x")
    # Folder is majority-dated, but the config forces the undated form.
    assert resolve_date_prefix_for_folder(_cfg(date_prefix=False), str(tmp_path)) is False

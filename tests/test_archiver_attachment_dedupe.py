"""Regression tests for the attachment de-duplication loop (issue #43).

``_fit_filename_to_path``'s truncation branch computes ``stem_budget`` from
the folder/prefix/suffix lengths only -- never from the stem itself. Folding
the disambiguating counter into the *stem* (the old behaviour) therefore
produced a byte-identical truncated filename on every iteration once
truncation kicked in, so ``while os.path.exists(att_path):`` never saw the
path stop existing and spun forever on the Tk main thread. The fix appends
the counter to the *suffix* instead (its width is then reserved out of the
stem budget via the existing ``fixed`` term) and bounds the loop at
``_MAX_DEDUPE_ATTEMPTS``.

Every test here runs the archiver call on a background thread with a bounded
``join(timeout=...)`` so a regression fails fast with a clear assertion
instead of hanging the test run itself.
"""
from __future__ import annotations

import os
import threading

from email_archiver.archiver.archiver import (
    _MAX_DEDUPE_ATTEMPTS,
    _fit_filename_to_path,
    EmailArchiver,
)

from .conftest import _FakeMailItem


def _run_bounded(target, *, timeout=10.0):
    """Run ``target`` on a daemon thread; assert it returned within ``timeout``."""
    t = threading.Thread(target=target, daemon=True)
    t.start()
    t.join(timeout=timeout)
    assert not t.is_alive(), (
        "attachment de-dupe loop did not return within the bounded timeout "
        "-- the collision loop is hanging again (#43)"
    )


def test_attachments_that_truncate_to_the_same_stem_get_distinct_names(tmp_path):
    # Force the truncation branch deterministically by capping max_path just
    # past folder + prefix + suffix, regardless of tmp_path's own length --
    # only a handful of stem chars fit, so two attachments whose stems share
    # a long common prefix collide once truncated.
    folder = str(tmp_path)
    max_path = len(folder) + 1 + len("007 - ") + len(".pdf") + 5
    cfg = {"path": {"max_length": max_path}}

    mail_item = _FakeMailItem(attachments=(
        "Quarterly_report_2026_meeting_notes_Q1.pdf",
        "Quarterly_report_2026_meeting_notes_Q2.pdf",
    ))
    archiver = EmailArchiver(cfg)

    result_holder = {}

    def run():
        result_holder["result"] = archiver.archive(mail_item, folder, "Subject")

    _run_bounded(run)

    result = result_holder["result"]
    names = [os.path.basename(p) for p in result.attachment_paths]
    assert len(names) == 2
    assert len(set(names)) == 2, f"attachments collided on disk: {names}"
    for name in names:
        assert os.path.exists(os.path.join(folder, name))


def test_dedupe_loop_caps_out_and_skips_instead_of_hanging(tmp_path, caplog):
    folder = str(tmp_path)
    seq = "007"
    stem = "report"
    suffix = ".pdf"

    # Pre-occupy the base name and every disambiguated variant up to (and one
    # past) the bounded cap, so the real attempt below is forced to exhaust
    # every attempt and give up rather than ever finding a free name.
    (tmp_path / _fit_filename_to_path(folder, seq, stem, suffix)).write_bytes(b"x")
    for counter in range(2, _MAX_DEDUPE_ATTEMPTS + 2):
        name = _fit_filename_to_path(folder, seq, stem, f"_{counter}{suffix}")
        (tmp_path / name).write_bytes(b"x")

    mail_item = _FakeMailItem(attachments=(f"{stem}{suffix}",))
    archiver = EmailArchiver({"path": {"max_length": 255}})

    result_holder = {}

    def run():
        result_holder["paths"] = archiver._save_attachments(mail_item, folder, seq, None)

    with caplog.at_level("ERROR"):
        _run_bounded(run)

    assert result_holder["paths"] == []
    assert any(
        "non-colliding filename" in rec.message for rec in caplog.records
    ), "expected an error log breadcrumb when the dedupe cap is exceeded"

"""Path-length tests for the archiver filename fitter."""
from __future__ import annotations

from email_archiver.archiver.archiver import (
    _ELLIPSIS,
    _WINDOWS_MAX_PATH,
    _fit_filename_to_path,
)


def _full_len(folder: str, filename: str) -> int:
    # Mirror the os.path.join + sep accounting used by the fitter.
    return len(folder) + 1 + len(filename)


def test_short_path_is_untouched():
    folder = r"C:\Users\me\OneDrive\Archive\Project Alpha"
    name = _fit_filename_to_path(folder, "042", "meeting_notes_q4", ".msg")
    assert name == "042 - meeting_notes_q4.msg"
    assert _full_len(folder, name) <= _WINDOWS_MAX_PATH


def test_long_path_gets_truncated_with_ellipsis():
    # Construct a folder that pushes the full path well past MAX_PATH.
    folder = r"C:\Users\me\OneDrive\Archive\\" + ("nested_subdir\\" * 12) + "Final"
    long_stem = "Quarterly_planning_meeting_with_the_extended_leadership_team_recap_attachments"
    name = _fit_filename_to_path(folder, "007", long_stem, ".msg")

    assert _full_len(folder, name) <= _WINDOWS_MAX_PATH
    assert name.startswith("007 - ")
    assert name.endswith(".msg")
    # The cut portion is signposted with an ellipsis so the user can tell it's been trimmed.
    assert _ELLIPSIS + ".msg" in name
    # And the start of the original subject is preserved.
    assert "Quarterly_planning" in name


def test_truncation_preserves_extension_and_prefix():
    folder = "C:\\" + ("x" * 230)
    name = _fit_filename_to_path(folder, "999", "subject_text_here", ".pdf")
    assert name.startswith("999 - ")
    assert name.endswith(".pdf")
    assert _full_len(folder, name) <= _WINDOWS_MAX_PATH


def test_pathological_folder_does_not_crash():
    # Folder alone already eats the entire budget — we should still return *something*
    # rather than raise; the underlying SaveAs call can then fail loudly if needed.
    folder = "C:\\" + ("x" * 280)
    name = _fit_filename_to_path(folder, "001", "anything", ".msg")
    assert name.startswith("001 - ")
    assert name.endswith(".msg")

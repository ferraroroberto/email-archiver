"""``sanitize_filename`` keeps Unicode letters and strips only what NTFS forbids."""
from __future__ import annotations

import unicodedata

import pytest

from email_archiver.archiver.archiver import (
    _ELLIPSIS,
    _fit_filename_to_path,
    sanitize_filename,
)
from email_archiver.config import DEFAULT_MAX_PATH_LENGTH as _WINDOWS_MAX_PATH

_FORBIDDEN = '<>:"/\\|?*'


def _utf16_len(text: str) -> int:
    # Windows counts MAX_PATH in UTF-16 code units, not code points.
    return len(text.encode("utf-16-le")) // 2


@pytest.mark.parametrize("text", [
    "petición y autorización",
    "canción del año",
    "Barça contra el Atlético",
    "Tramitació de derivació",
    "col·legi Maxil·lofacial",
    "Über Straße, ñandú, façade, naïve",
])
def test_unicode_letters_are_kept(text):
    assert sanitize_filename(text) == text


def test_forbidden_characters_never_survive():
    out = sanitize_filename('a<b>c:d"e/f\\g|h?i*j k\x00l\x1fm\x7fn')
    assert not any(c in out for c in _FORBIDDEN)
    assert not any(ord(c) < 32 or ord(c) == 127 for c in out)


def test_a_slash_becomes_a_visible_dash():
    assert sanitize_filename("informe (03/07/2015)") == "informe (03-07-2015)"
    assert sanitize_filename("a\\b: c") == "a-b- c"


def test_other_forbidden_characters_become_an_underscore():
    assert sanitize_filename("what? <now> *really*") == "what_ _now_ _really"


def test_a_run_of_substituted_characters_collapses_but_a_real_underscore_stays():
    assert sanitize_filename("a??b") == "a_b"
    assert sanitize_filename("snake_case__name") == "snake_case__name"


def test_nfd_and_nfc_spell_one_name_one_way():
    nfc = "canción del año"
    nfd = unicodedata.normalize("NFD", nfc)
    assert nfd != nfc
    assert sanitize_filename(nfd) == sanitize_filename(nfc) == nfc


def test_trailing_dots_and_spaces_are_stripped():
    assert sanitize_filename("informe final. . ") == "informe final"


def test_truncation_never_leaves_a_trailing_space():
    assert sanitize_filename("abcd efgh", max_len=5) == "abcd"


def test_an_empty_result_falls_back_to_email():
    assert sanitize_filename("") == "email"
    assert sanitize_filename(" . ") == "email"


def test_max_len_still_cuts():
    assert len(sanitize_filename("ñ" * 200, max_len=60)) == 60


def test_a_long_emoji_subject_still_fits_the_path_budget():
    # An emoji is two UTF-16 units, so counting code points would let a name
    # run past MAX_PATH even though the fitter reports it fits.
    folder = "C:\\" + ("x" * 150)
    subject = sanitize_filename("🎉" * 40 + "ñ" * 40, max_len=200)
    name = _fit_filename_to_path(folder, "007", subject, ".msg")

    assert _utf16_len(folder) + 1 + _utf16_len(name) <= _WINDOWS_MAX_PATH
    assert name.startswith("007 - ")
    assert name.endswith(_ELLIPSIS + ".msg")
    # Never cut through a surrogate pair.
    name.encode("utf-8")


def test_a_long_accented_subject_is_cut_to_a_valid_path():
    folder = "C:\\" + ("x" * 200)
    name = _fit_filename_to_path(
        folder, "042",
        sanitize_filename("resonancia magnética " * 10, max_len=200), ".msg",
    )
    assert len(folder) + 1 + len(name) <= _WINDOWS_MAX_PATH
    assert name.startswith("042 - resonancia magnética")

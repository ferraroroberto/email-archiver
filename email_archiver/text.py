"""
Shared text-normalisation helpers.

These utilities must stay identical on every path that processes email
subjects — the scanner (FTS index) and the Outlook client (live email)
both call :func:`clean_subject`, so the two sides of the app always
agree on what "the same subject" looks like.
"""
from __future__ import annotations

import re

# Matches common reply/forward prefixes at the start of a subject line.
RE_REPLY_PREFIX = re.compile(r"^\s*(re|rv|fwd?)\s*:?\s*", re.IGNORECASE)

# Matches a trailing `.msg` filename extension.
_RE_MSG_SUFFIX = re.compile(r"\s*\.msg$", re.IGNORECASE)


def clean_subject(raw: str | None, *, strip_msg_suffix: bool = False) -> str:
    """Return a normalised subject string.

    Strips leading Re:/Rv:/Fwd: prefixes (case-insensitive).  When
    *strip_msg_suffix* is ``True`` also removes a trailing ``.msg``
    extension — useful when the subject was derived from a filename.

    Args:
        raw: The raw subject string, or ``None``/empty.
        strip_msg_suffix: Whether to strip a trailing ``.msg`` suffix.

    Returns:
        The cleaned string, never ``None``.
    """
    if not raw:
        return ""
    s = RE_REPLY_PREFIX.sub("", raw.strip())
    if strip_msg_suffix:
        s = _RE_MSG_SUFFIX.sub("", s)
    return s.strip()

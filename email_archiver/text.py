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


def normalize_message_id(raw: str | None) -> str:
    """Return the Internet Message-ID in the one canonical form this app stores.

    Two sources have to agree on it or the batch verbs cannot pair a live mail
    with the file archived from it: Outlook COM reads MAPI
    ``PR_INTERNET_MESSAGE_ID`` (``0x1035001F``) off the live item, while the
    scanner reads the same header back out of the ``.msg`` on disk. Both spell
    it ``<local@domain>`` most of the time, but neither guarantees the angle
    brackets or the surrounding whitespace, so the stored form drops both.

    Case is deliberately preserved: RFC 5322 makes the left-hand side of a
    Message-ID case-sensitive, so folding it would merge ids that are genuinely
    distinct. Returns ``""`` for a missing or empty id -- a real answer ("this
    mail has no Message-ID"), which the caller reports and skips rather than
    treating as an identity.
    """
    if not raw:
        return ""
    s = str(raw).strip()
    if s.startswith("<") and s.endswith(">") and len(s) > 1:
        s = s[1:-1].strip()
    return s


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

"""
The per-mail filter ``plan`` and ``find`` share.

Kept apart from ``email_archiver.batch`` so ``find`` can match a mail without
importing the plan/apply/revert module.
"""
from __future__ import annotations

from collections.abc import Sequence
from datetime import datetime
from typing import Any


def mail_matches(
    mail: Any, terms: Sequence[str], since: datetime | None, sent: bool
) -> bool:
    """Whether ``mail`` carries every term (casefolded, in the subject, sender,
    recipients or body preview) and is dated at or after ``since``.

    The one per-mail rule shared by ``plan --search``/``--since`` and the
    ``find`` folder-walk fallback. ``since`` is checked here even after a
    server-side Restrict: this is the exact comparison, the Restrict only
    narrows the walk."""
    if since is not None:
        # A Sent Items mail's own clock is when it was sent; its received time
        # is whatever the store stamped and means nothing here.
        when = (
            (mail.date_sent or mail.date_received)
            if sent
            else (mail.date_received or mail.date_sent)
        )
        if when is None or when < since:
            return False
    if terms:
        haystack = "\n".join(
            (mail.subject, mail.sender, mail.recipients, mail.body_preview)
        ).casefold()
        if not all(term.casefold() in haystack for term in terms):
            return False
    return True

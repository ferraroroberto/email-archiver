"""
The one place this app sends mail over COM.

Kept out of ``client.py`` and ``drafts.py`` on purpose: ``draft`` must have no
send path, and a test pins that ``draft.py``, ``outlook/client.py``,
``outlook/drafts.py`` and ``main_batch.py`` carry no ``Send`` call. The only caller is :func:`email_archiver.send.send`.
"""
from __future__ import annotations

import logging
from collections.abc import Callable
from typing import Any

from email_archiver.outlook.drafts import DraftSnapshot

logger = logging.getLogger(__name__)


def send_if_approved(
    client: Any, entry_id: str, check: Callable[[DraftSnapshot], None],
) -> DraftSnapshot:
    """Send the draft ``entry_id`` only if ``check`` accepts it as stored now.

    A window open on the draft holds its own copy, and Send pressed there
    would send that; it is closed first, saving, so nothing typed in it is lost
    and anything typed counts as a change. The item is then re-read and
    ``check`` runs on that snapshot, and ``Send`` is called on the very same
    reference. ``check`` raising means nothing was sent.

    Raises:
        DraftUpdateError: the item is gone, sent, or not in Drafts.
        Whatever ``check`` raises.
    """
    client.open_draft(entry_id)  # refuse a non-draft before touching any window
    if client.close_inspectors_of(entry_id):
        logger.info("Closed the window open on draft %s before reading it to send.", entry_id)
    mail = client.open_draft(entry_id)
    snapshot = client.snapshot_draft(mail, entry_id)
    check(snapshot)
    mail.Send()
    return snapshot

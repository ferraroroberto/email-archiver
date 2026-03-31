"""
Suggestion engine: ranks archive folders for an incoming email.

Three-stage scoring pipeline:
─────────────────────────────────────────────────────────────────────────────
Stage 1 – FTS5 BM25 (repository layer)
    Full-text search over every indexed email (subject + sender + recipients
    + body preview). Results are aggregated per folder and normalised to
    [0, 1]. BM25 column weights give subject 10×, sender/recipients 3×, and
    body 1× — so subject matches dominate the ranking.

Stage 2 – Subject thread score (this module)
    Compares the incoming subject against the subjects of recent emails
    already stored in each candidate folder (up to 5 samples). A high score
    means the same conversation thread exists in that folder — the strongest
    signal for "same subject → same folder" routing.
    Uses rapidfuzz token_set_ratio (handles Re:/Fwd: prefixes naturally).

Stage 3 – Folder-name boost (this module)
    Computes fuzzy similarity between the incoming subject and the folder
    *leaf name*. Useful for new projects with few or no prior emails.

Final score = 0.45 × fts_score
            + 0.30 × subject_thread_score
            + 0.25 × folder_name_score

Fallback (empty DB):
    If no FTS results are found (e.g. the DB hasn't been scanned yet), the
    engine returns an empty list and the UI shows a browse button instead.
─────────────────────────────────────────────────────────────────────────────
"""
from __future__ import annotations

import logging
import os
from dataclasses import dataclass
from pathlib import Path

from email_archiver.database.models import get_connection
from email_archiver.database.repository import EmailRepository, FolderSuggestion
from email_archiver.outlook.client import EmailData

logger = logging.getLogger(__name__)

# Scoring weights (must sum to 1.0)
_W_FTS = 0.45
_W_SUBJ = 0.30   # subject thread continuity (incoming subject vs. folder's stored subjects)
_W_NAME = 0.25   # folder leaf-name fuzzy match


@dataclass
class RankedSuggestion:
    folder_path: str
    display_name: str        # last 2 path components for compact display
    score: float             # final blended score [0, 1]
    match_count: int         # number of matching emails in this folder
    sample_subjects: list[str]


def _folder_display_name(path: str) -> str:
    """Return the last two components of a path for display."""
    parts = Path(path).parts
    if len(parts) >= 2:
        return os.path.join(parts[-2], parts[-1])
    return parts[-1] if parts else path


def _fuzzy_folder_score(subject: str, folder_path: str) -> float:
    """
    Compute similarity between the email subject and the folder's leaf name.
    Returns a value in [0, 1].
    """
    if not subject:
        return 0.0
    try:
        from rapidfuzz import fuzz  # type: ignore
        folder_name = Path(folder_path).name
        return fuzz.token_set_ratio(subject, folder_name) / 100.0
    except ImportError:
        # rapidfuzz not installed — skip the name boost
        return 0.0


def _subject_thread_score(incoming_subject: str, sample_subjects: list[str]) -> float:
    """
    Score thread continuity: compare the incoming subject against subjects of
    emails already stored in the candidate folder.

    Uses token_set_ratio so "Re: Project Alpha – Q2" scores high against
    "Project Alpha – Q2" or "Fwd: Project Alpha – Q2".

    Returns the best match score in [0, 1], or 0 if no samples available.
    """
    if not incoming_subject or not sample_subjects:
        return 0.0
    try:
        from rapidfuzz import fuzz  # type: ignore
        best = max(
            fuzz.token_set_ratio(incoming_subject, s) / 100.0
            for s in sample_subjects if s
        )
        return best
    except ImportError:
        return 0.0


class SuggestionEngine:
    """
    Produces ranked folder suggestions for an incoming email.

    Stateless: creates a DB connection per call so it is safe to use from
    any thread (the archive dialog calls it in a background thread).
    """

    def __init__(self, cfg: dict) -> None:
        self._db_path: str = cfg["database"]["path"]
        self._max: int = cfg["suggestion"]["max_suggestions"]
        self._min_score: float = cfg["suggestion"]["min_score"]

    def suggest(self, email: EmailData) -> list[RankedSuggestion]:
        """
        Return up to max_suggestions ranked folders for the given email.
        An empty list is returned when the DB is empty or no match is found.
        """
        conn = get_connection(self._db_path)
        repo = EmailRepository(conn)

        try:
            raw: list[FolderSuggestion] = repo.suggest_folders(
                subject=email.subject,
                sender=email.sender,
                recipients=email.recipients,
                max_results=self._max * 3,   # get extra candidates for re-ranking
                min_score=0.0,               # we apply min_score after blending
            )
        finally:
            conn.close()

        if not raw:
            logger.info("No FTS matches for subject: %r", email.subject)
            return []

        blended: list[RankedSuggestion] = []
        for s in raw:
            name_score = _fuzzy_folder_score(email.subject, s.folder_path)
            thread_score = _subject_thread_score(email.subject, s.sample_subjects)
            final_score = _W_FTS * s.score + _W_SUBJ * thread_score + _W_NAME * name_score

            if final_score < self._min_score:
                continue

            blended.append(
                RankedSuggestion(
                    folder_path=s.folder_path,
                    display_name=_folder_display_name(s.folder_path),
                    score=final_score,
                    match_count=s.match_count,
                    sample_subjects=s.sample_subjects,
                )
            )

        # Sort by final score descending, deduplicate, take top N
        blended.sort(key=lambda x: x.score, reverse=True)
        seen: set[str] = set()
        unique: list[RankedSuggestion] = []
        for s in blended:
            if s.folder_path not in seen:
                seen.add(s.folder_path)
                unique.append(s)
            if len(unique) >= self._max:
                break

        logger.info(
            "Suggestions for %r: %s",
            email.subject,
            [(s.display_name, round(s.score, 3)) for s in unique],
        )
        return unique

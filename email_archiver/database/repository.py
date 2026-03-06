"""
Data-access layer: all SQL queries are centralised here.

All public methods accept / return plain Python dicts or dataclasses so
upper layers never import sqlite3 directly.
"""
from __future__ import annotations

import logging
import re
import sqlite3
from dataclasses import dataclass, field
from datetime import datetime
from typing import Sequence

logger = logging.getLogger(__name__)


# ----------------------------------------------------------------- types ----

@dataclass
class EmailRecord:
    file_path: str
    folder_path: str
    filename: str
    subject: str = ""
    sender: str = ""
    recipients: str = ""
    date_sent: str = ""
    body_preview: str = ""
    file_mtime: float = 0.0
    id: int | None = None


@dataclass
class FolderSuggestion:
    folder_path: str
    score: float
    match_count: int
    sample_subjects: list[str] = field(default_factory=list)


# -------------------------------------------------------------- helpers -----

def _build_fts_query(text: str) -> str:
    """
    Convert a free-text string into a safe FTS5 MATCH expression.

    Strategy: tokenise on word boundaries, discard short tokens, wrap each
    in double-quotes (exact token match), join with OR so any matching word
    contributes to relevance. This avoids FTS5 syntax errors from special
    chars in email subjects.
    """
    tokens = re.findall(r"[A-Za-z0-9\u00C0-\u024F]{3,}", text)
    if not tokens:
        return ""
    # Deduplicate while preserving order
    seen: set[str] = set()
    unique = [t.lower() for t in tokens if t.lower() not in seen and not seen.add(t.lower())]  # type: ignore[func-returns-value]
    return " OR ".join(f'"{t}"' for t in unique[:30])  # cap at 30 tokens


# ------------------------------------------------------------ repository ----

class EmailRepository:
    """
    Thin repository wrapping a SQLite connection.

    Thread-safety note: SQLite connections are NOT thread-safe. If you need
    concurrent access, create one repository per thread (or use a connection
    pool). The scanner runs in a background thread and creates its own repo.
    """

    def __init__(self, conn: sqlite3.Connection) -> None:
        self._conn = conn

    # -------------------------------------------------- write operations ---

    def upsert_email(self, rec: EmailRecord) -> None:
        """Insert or replace an email record. Triggers keep FTS in sync."""
        self._conn.execute(
            """
            INSERT INTO emails
                (file_path, folder_path, filename, subject, sender,
                 recipients, date_sent, body_preview, file_mtime)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(file_path) DO UPDATE SET
                subject      = excluded.subject,
                sender       = excluded.sender,
                recipients   = excluded.recipients,
                date_sent    = excluded.date_sent,
                body_preview = excluded.body_preview,
                file_mtime   = excluded.file_mtime,
                indexed_at   = datetime('now')
            """,
            (
                rec.file_path, rec.folder_path, rec.filename,
                rec.subject, rec.sender, rec.recipients,
                rec.date_sent, rec.body_preview, rec.file_mtime,
            ),
        )

    def upsert_folder(self, folder_path: str) -> None:
        """Upsert folder record; email_count is refreshed separately."""
        self._conn.execute(
            """
            INSERT INTO folders (folder_path, last_updated)
            VALUES (?, datetime('now'))
            ON CONFLICT(folder_path) DO UPDATE SET
                last_updated = datetime('now')
            """,
            (folder_path,),
        )

    def refresh_folder_counts(self) -> None:
        """Recompute email_count for all folders from the emails table."""
        self._conn.execute(
            """
            UPDATE folders SET email_count = (
                SELECT COUNT(*) FROM emails
                WHERE emails.folder_path = folders.folder_path
            )
            """
        )

    def delete_missing_emails(self, known_paths: Sequence[str]) -> int:
        """
        Remove DB entries whose files no longer exist on disk.
        Called at the end of a full scan to purge deleted emails.
        Returns the number of rows deleted.
        """
        if not known_paths:
            return 0
        placeholders = ",".join("?" * len(known_paths))
        cur = self._conn.execute(
            f"DELETE FROM emails WHERE file_path NOT IN ({placeholders})",
            known_paths,
        )
        return cur.rowcount

    def commit(self) -> None:
        self._conn.commit()

    # -------------------------------------------------- read operations ----

    def get_mtime(self, file_path: str) -> float | None:
        """Return stored mtime for a file, or None if not indexed."""
        row = self._conn.execute(
            "SELECT file_mtime FROM emails WHERE file_path = ?", (file_path,)
        ).fetchone()
        return row["file_mtime"] if row else None

    def count_emails(self) -> int:
        row = self._conn.execute("SELECT COUNT(*) FROM emails").fetchone()
        return row[0]

    def count_folders(self) -> int:
        row = self._conn.execute("SELECT COUNT(*) FROM folders").fetchone()
        return row[0]

    def suggest_folders(
        self,
        subject: str,
        sender: str,
        recipients: str,
        max_results: int = 3,
        min_score: float = 0.05,
    ) -> list[FolderSuggestion]:
        """
        Return ranked folder suggestions for an incoming email.

        Why two steps instead of a JOIN:
        SQLite's bm25() auxiliary function and the FTS5 `rank` column can only
        be resolved when the FTS5 virtual table is the *outermost* primary table
        in the query.  Any JOIN — even inside a subquery — breaks this contract
        and raises "unable to use function bm25 in the requested context".

        Solution: query the FTS5 table in isolation (step 1) to get rowids and
        scores, then look up folder_path in a single batched IN query (step 2),
        and aggregate per-folder in Python (step 3).  Clean, fast, zero SQL
        context issues.
        """
        from collections import defaultdict

        query_text = f"{subject} {sender} {recipients}"
        fts_query = _build_fts_query(query_text)

        if not fts_query:
            logger.warning("Empty FTS query for: %r", query_text)
            return []

        # Step 1 – FTS5 search.
        # `rank` is the built-in BM25 column; negative values, more negative
        # = better match.  We invert to positive for intuitive aggregation.
        try:
            fts_rows = self._conn.execute(
                """
                SELECT rowid, rank
                FROM   emails_fts
                WHERE  emails_fts MATCH ?
                ORDER  BY rank
                LIMIT  1000
                """,
                (fts_query,),
            ).fetchall()
        except sqlite3.OperationalError as exc:
            logger.warning("FTS query failed (%s) for query %r", exc, fts_query)
            return []

        if not fts_rows:
            return []

        # Step 2 – Batch-fetch folder_path for matched rowids.
        rowid_score: dict[int, float] = {r[0]: -r[1] for r in fts_rows}
        placeholders = ",".join("?" * len(rowid_score))
        email_rows = self._conn.execute(
            f"SELECT id, folder_path FROM emails WHERE id IN ({placeholders})",
            list(rowid_score.keys()),
        ).fetchall()

        if not email_rows:
            return []

        # Step 3 – Aggregate scores and counts per folder in Python.
        folder_scores: dict[str, float] = defaultdict(float)
        folder_counts: dict[str, int] = defaultdict(int)
        for row in email_rows:
            fp = row["folder_path"]
            folder_scores[fp] += rowid_score.get(row["id"], 0.0)
            folder_counts[fp] += 1

        max_raw = max(folder_scores.values()) or 1.0
        sorted_folders = sorted(
            folder_scores.items(), key=lambda x: x[1], reverse=True
        )[:20]

        suggestions: list[FolderSuggestion] = []
        for fp, raw_score in sorted_folders:
            score = raw_score / max_raw
            if score < min_score:
                continue

            samples = self._conn.execute(
                """
                SELECT DISTINCT subject FROM emails
                WHERE  folder_path = ?
                ORDER  BY date_sent DESC
                LIMIT  2
                """,
                (fp,),
            ).fetchall()
            sample_subjects = [s["subject"] for s in samples if s["subject"]]

            suggestions.append(
                FolderSuggestion(
                    folder_path=fp,
                    score=score,
                    match_count=folder_counts[fp],
                    sample_subjects=sample_subjects,
                )
            )

        return suggestions[:max_results]

    def get_all_folder_paths(self) -> list[str]:
        """Return all indexed folder paths (used by suggestion engine fallback)."""
        rows = self._conn.execute(
            "SELECT folder_path FROM folders ORDER BY email_count DESC"
        ).fetchall()
        return [r["folder_path"] for r in rows]

    def get_last_scan_time(self) -> str:
        """Return ISO datetime of the most recent indexing operation."""
        row = self._conn.execute(
            "SELECT MAX(indexed_at) AS t FROM emails"
        ).fetchone()
        return row["t"] or "Never"

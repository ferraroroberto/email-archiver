"""
Main Tkinter application.

There are two entry modes:
  1. Full launcher (main_ui.py): shows both "Scan" and "Archive" buttons.
  2. Direct archive mode (main_archive.py): opens ArchiveDialog immediately.
  3. Direct scan mode (main_scan.py): runs scan with a progress window.

Stream Deck optimisation:
  - The Tkinter root is created and shown FIRST (< 0.1 s).
  - All heavy work (Outlook COM, DB query) runs in a background Thread.
  - The window title and status label provide immediate feedback.
"""
from __future__ import annotations

import logging
import os
import threading
import tkinter as tk
from tkinter import messagebox, ttk
from typing import Any

from email_archiver.archiver.archiver import EmailArchiver
from email_archiver.engine.suggester import RankedSuggestion, SuggestionEngine
from email_archiver.outlook.client import EmailData, OutlookClient
from email_archiver.ui.dialogs import (
    ProgressDialog,
    browse_folder,
)

logger = logging.getLogger(__name__)

_ACCENT = "#1a6b3c"
_BG = "#f5f5f5"
_CARD_BG = "#ffffff"
_SCORE_GOOD = "#2d6a4f"
_SCORE_MID = "#b5451b"


# ============================================================= helpers ======

def _score_color(score: float) -> str:
    return _SCORE_GOOD if score >= 0.5 else _SCORE_MID


def _pct(score: float) -> str:
    return f"{int(score * 100)}%"


# ======================================================= archive dialog =====

class ArchiveDialog:
    """
    Self-contained window for the 'Archive Email' workflow.

    Opens immediately with a loading spinner, then a background thread
    fetches the Outlook email and DB suggestions and populates the UI.
    """

    def __init__(self, cfg: dict) -> None:
        self._cfg = cfg
        self._email: EmailData | None = None
        self._suggestions: list[RankedSuggestion] = []
        self._chosen_folder: str | None = None

        self._root = tk.Tk()
        self._root.title("Email Archiver – Archive")
        self._root.configure(bg=_BG)
        self._root.resizable(True, False)

        w = cfg["ui"]["window_width"]
        h = cfg["ui"]["window_height"]
        sw = self._root.winfo_screenwidth()
        sh = self._root.winfo_screenheight()
        self._root.geometry(f"{w}x{h}+{(sw - w) // 2}+{(sh - h) // 2}")
        self._root.attributes("-topmost", True)

        self._build_ui()
        self._root.bind("<Escape>", lambda e: self._root.destroy())
        self._root.after(80, self._start_loading)

    # --------------------------------------------------------- UI build ----

    def _build_ui(self) -> None:
        ff = self._cfg["ui"]["font_family"]
        fn = self._cfg["ui"]["font_size_normal"]
        ft = self._cfg["ui"]["font_size_title"]

        # ---- header ----
        hdr = tk.Frame(self._root, bg=_ACCENT, pady=10)
        hdr.pack(fill="x")
        tk.Label(
            hdr, text="📧  Archive Email",
            bg=_ACCENT, fg="white",
            font=(ff, ft, "bold"), padx=16,
        ).pack(side="left")

        # ---- subject bar ----
        subj_frame = tk.Frame(self._root, bg=_BG, pady=6, padx=14)
        subj_frame.pack(fill="x")
        tk.Label(subj_frame, text="Subject:", bg=_BG,
                 font=(ff, fn, "bold")).pack(side="left")
        self._subj_var = tk.StringVar(value="Loading…")
        tk.Label(subj_frame, textvariable=self._subj_var,
                 bg=_BG, font=(ff, fn), fg="#333").pack(side="left", padx=6)

        # ---- sender / date ----
        meta_frame = tk.Frame(self._root, bg=_BG, padx=14)
        meta_frame.pack(fill="x")
        self._meta_var = tk.StringVar(value="")
        tk.Label(meta_frame, textvariable=self._meta_var,
                 bg=_BG, font=(ff, fn - 1), fg="#666").pack(side="left")

        ttk.Separator(self._root).pack(fill="x", pady=6)

        # ---- suggestions section ----
        tk.Label(
            self._root, text="Suggested folders:", bg=_BG,
            font=(ff, fn, "bold"), anchor="w", padx=14,
        ).pack(fill="x")

        self._suggestions_frame = tk.Frame(self._root, bg=_BG, padx=14, pady=4)
        self._suggestions_frame.pack(fill="x")

        self._status_label = tk.Label(
            self._suggestions_frame, text="⏳  Searching…",
            bg=_BG, font=(ff, fn), fg="#666",
        )
        self._status_label.pack(anchor="w")

        ttk.Separator(self._root).pack(fill="x", pady=6)

        # ---- manual browse + actions ----
        bottom = tk.Frame(self._root, bg=_BG, padx=14, pady=8)
        bottom.pack(fill="x")

        tk.Button(
            bottom, text="📂  Browse folder…",
            command=self._on_browse,
            relief="flat", bg="#e0e0e0", padx=8, pady=4,
            font=(ff, fn),
        ).pack(side="left")

        tk.Button(
            bottom, text="✗  Cancel",
            command=self._root.destroy,
            relief="flat", bg="#e0e0e0", padx=8, pady=4,
            font=(ff, fn),
        ).pack(side="right")

    # ------------------------------------------------ background loading ---

    def _start_loading(self) -> None:
        thread = threading.Thread(target=self._load_worker, daemon=True)
        thread.start()

    def _load_worker(self) -> None:
        """Runs in background thread: fetch email + suggestions."""
        try:
            client = OutlookClient()
            email = client.get_selected_email()

            if email is None:
                self._root.after(0, lambda: self._show_error(
                    "No email selected in Outlook.\n\n"
                    "Please select an email and try again."
                ))
                return

            self._email = email
            self._root.after(0, self._show_email_meta)

            engine = SuggestionEngine(self._cfg)
            suggestions = engine.suggest(email)
            self._suggestions = suggestions
            self._root.after(0, self._show_suggestions)

        except Exception as exc:
            logger.exception("Error in archive load worker")
            self._root.after(
                0, lambda: self._show_error(f"Unexpected error:\n{exc}")
            )

    def _show_email_meta(self) -> None:
        if not self._email:
            return
        self._subj_var.set(self._email.subject or "(no subject)")
        date_str = (
            self._email.date_sent.strftime("%d/%m/%Y %H:%M")
            if self._email.date_sent else ""
        )
        self._meta_var.set(
            f"From: {self._email.sender}   |   {date_str}"
        )

    def _show_suggestions(self) -> None:
        # Clear the spinner
        for w in self._suggestions_frame.winfo_children():
            w.destroy()

        ff = self._cfg["ui"]["font_family"]
        fn = self._cfg["ui"]["font_size_normal"]

        if not self._suggestions:
            tk.Label(
                self._suggestions_frame,
                text="No matching folders found. Use 'Browse folder…' to select manually.",
                bg=_BG, font=(ff, fn), fg="#888",
            ).pack(anchor="w", pady=4)
            return

        for i, s in enumerate(self._suggestions, start=1):
            card = tk.Frame(
                self._suggestions_frame, bg=_CARD_BG,
                relief="groove", bd=1, pady=6, padx=10,
            )
            card.pack(fill="x", pady=3)

            # Rank badge
            rank_lbl = tk.Label(
                card, text=f"#{i}", bg=_CARD_BG,
                font=(ff, fn + 2, "bold"), fg=_ACCENT, width=3,
            )
            rank_lbl.grid(row=0, column=0, rowspan=2, sticky="ns", padx=(0, 8))

            # Folder path (last 2 components bold, rest grey)
            path_parts = s.folder_path.replace("\\", "/").split("/")
            short = "/".join(path_parts[-2:]) if len(path_parts) >= 2 else s.folder_path
            rest = "/".join(path_parts[:-2]) + "/" if len(path_parts) > 2 else ""

            path_frame = tk.Frame(card, bg=_CARD_BG)
            path_frame.grid(row=0, column=1, sticky="w")
            if rest:
                tk.Label(path_frame, text=rest, bg=_CARD_BG,
                         font=(ff, fn - 1), fg="#999").pack(side="left")
            tk.Label(path_frame, text=short, bg=_CARD_BG,
                     font=(ff, fn, "bold"), fg="#222").pack(side="left")

            # Score pill + match count
            info_frame = tk.Frame(card, bg=_CARD_BG)
            info_frame.grid(row=1, column=1, sticky="w")
            tk.Label(
                info_frame,
                text=f"  {_pct(s.score)} match  ",
                bg=_score_color(s.score), fg="white",
                font=(ff, fn - 1, "bold"), padx=4, pady=1,
            ).pack(side="left")
            tk.Label(
                info_frame,
                text=f"  {s.match_count} similar email(s)",
                bg=_CARD_BG, fg="#666", font=(ff, fn - 1),
            ).pack(side="left", padx=4)
            if s.sample_subjects:
                sample = s.sample_subjects[0][:70]
                tk.Label(
                    info_frame,
                    text=f'  e.g. "{sample}"',
                    bg=_CARD_BG, fg="#999", font=(ff, fn - 2),
                ).pack(side="left")

            # Action buttons — Archive directly + Open folder
            btn_frame = tk.Frame(card, bg=_CARD_BG)
            btn_frame.grid(row=0, column=2, rowspan=2, padx=(12, 0), sticky="e")

            tk.Button(
                btn_frame, text="✓  Archive",
                command=lambda fp=s.folder_path: self._do_archive(fp),
                relief="flat", bg=_ACCENT, fg="white",
                font=(ff, fn, "bold"), padx=10, pady=3, width=10,
            ).pack(pady=(0, 3))

            tk.Button(
                btn_frame, text="📂  Open",
                command=lambda fp=s.folder_path: os.startfile(fp),
                relief="flat", bg="#e0e0e0", fg="#333",
                font=(ff, fn), padx=10, pady=2, width=10,
            ).pack()

            card.columnconfigure(1, weight=1)

    def _show_error(self, msg: str) -> None:
        for w in self._suggestions_frame.winfo_children():
            w.destroy()
        ff = self._cfg["ui"]["font_family"]
        fn = self._cfg["ui"]["font_size_normal"]
        tk.Label(
            self._suggestions_frame, text=f"⚠  {msg}",
            bg=_BG, fg="#c0392b", font=(ff, fn), wraplength=700,
            justify="left",
        ).pack(anchor="w", pady=4)

    # ------------------------------------------------ user interactions ----

    def _on_browse(self) -> None:
        archive_root = self._cfg["archive"]["root_path"]
        chosen = browse_folder(self._root, initial_dir=archive_root)
        if chosen:
            self._do_archive(chosen)

    def _do_archive(self, folder_path: str) -> None:
        if not self._email:
            messagebox.showerror("Error", "Email data not available.")
            return
        try:
            # COM objects are STA (apartment-threaded). The cached raw_item was
            # obtained in the background worker thread and cannot be used from
            # the main UI thread — doing so produces AttributeError: <unknown>.SaveAs.
            # Re-acquire a fresh reference from the main thread at click time.
            # The user's selection in Outlook won't have changed between
            # seeing suggestions and clicking Archive.
            import pythoncom
            import win32com.client as _win32

            pythoncom.CoInitialize()

            try:
                app = _win32.GetActiveObject("Outlook.Application")
                explorer = app.ActiveExplorer()
                if explorer is None or explorer.Selection.Count == 0:
                    messagebox.showerror("Error", "No email is selected in Outlook.")
                    return
                # Dispatch forces full COM interface resolution → MailItem with SaveAs
                mail_item = _win32.Dispatch(explorer.Selection.Item(1))
            except Exception as exc:
                messagebox.showerror("Outlook error", f"Cannot access Outlook:\n{exc}")
                return

            archiver = EmailArchiver()
            result = archiver.archive(
                mail_item=mail_item,
                folder_path=folder_path,
                subject=self._email.subject,
            )
            logger.info(
                "Archived: %s | %d attachment(s) → %s",
                os.path.basename(result.email_path),
                len(result.attachment_paths),
                folder_path,
            )
            self._root.destroy()
        except Exception as exc:
            logger.exception("Archive failed")
            messagebox.showerror("Archive failed", str(exc))

    def run(self) -> None:
        self._root.mainloop()


# ============================================================ scan window ===

class ScanWindow:
    """
    Minimal launcher for the scanner with live progress feedback.
    Can run headless (no parent) or embedded in the full app.
    """

    def __init__(self, cfg: dict) -> None:
        self._cfg = cfg
        self._stop_flag: list[bool] = [False]

    def run(self) -> None:
        root = tk.Tk()
        root.title("Email Archiver – Scan Archive")
        root.configure(bg=_BG)
        root.resizable(False, False)

        ff = self._cfg["ui"]["font_family"]
        fn = self._cfg["ui"]["font_size_normal"]
        ft = self._cfg["ui"]["font_size_title"]

        hdr = tk.Frame(root, bg=_ACCENT, pady=10)
        hdr.pack(fill="x")
        tk.Label(hdr, text="🗂  Scan Archive", bg=_ACCENT, fg="white",
                 font=(ff, ft, "bold"), padx=16).pack(side="left")

        info = tk.Label(
            root,
            text=f"Archive root:\n{self._cfg['archive']['root_path']}",
            bg=_BG, font=(ff, fn), anchor="w", padx=14, pady=8,
            wraplength=580, justify="left",
        )
        info.pack(fill="x")

        bar = ttk.Progressbar(root, mode="indeterminate", length=560)
        bar.pack(padx=14, pady=4)

        status_var = tk.StringVar(value="Click 'Start Scan' to begin.")
        status_lbl = tk.Label(root, textvariable=status_var,
                              bg=_BG, font=(ff, fn), fg="#555",
                              anchor="w", padx=14, pady=4, wraplength=560)
        status_lbl.pack(fill="x")

        detail_var = tk.StringVar(value="")
        tk.Label(root, textvariable=detail_var, bg=_BG,
                 font=(ff, fn - 1), fg="#999", anchor="w",
                 padx=14, wraplength=560).pack(fill="x")

        btn_frame = tk.Frame(root, bg=_BG)
        btn_frame.pack(pady=10)

        start_btn: list[tk.Button] = []

        def on_progress(current: int, total: int, path: str) -> None:
            if total > 0:
                pct = int(current / total * 100)
                bar["value"] = pct
                status_var.set(f"Indexing {current:,} / {total:,}  ({pct}%)")
            else:
                status_var.set(f"Indexing {current:,} files…")
            detail_var.set(path[-70:] if len(path) > 70 else path)

        def do_scan() -> None:
            start_btn[0].config(state="disabled", text="Scanning…")
            bar.config(mode="determinate")
            self._stop_flag[0] = False

            from email_archiver.scanner.scanner import FolderScanner

            def worker() -> None:
                try:
                    scanner = FolderScanner(self._cfg)
                    stats = scanner.scan(
                        progress_callback=lambda c, t, p: root.after(
                            0, lambda c=c, t=t, p=p: on_progress(c, t, p)
                        ),
                        stop_flag=self._stop_flag,
                    )
                    root.after(0, lambda: on_done(stats))
                except Exception as exc:
                    root.after(0, lambda: on_error(exc))

            threading.Thread(target=worker, daemon=True).start()

        def on_done(stats: Any) -> None:
            deleted_part = f", {stats.deleted:,} deleted" if stats.deleted else ""
            was_cancelled = self._stop_flag[0]
            if was_cancelled:
                icon, label = "⏹", "Cancelled"
                bar["value"] = int(stats.total_found and
                                   (stats.newly_indexed + stats.updated + stats.skipped)
                                   / stats.total_found * 100)
            else:
                icon, label = "✓", "Done"
                bar["value"] = 100
            status_var.set(
                f"{icon}  {label} — {stats.newly_indexed:,} new, "
                f"{stats.updated:,} updated, {stats.skipped:,} skipped"
                f"{deleted_part}, {stats.errors:,} errors  "
                f"({stats.duration_seconds:.1f}s)"
            )
            detail_var.set("")
            start_btn[0].config(state="normal", text="Scan Again")

        def on_error(exc: Exception) -> None:
            status_var.set(f"⚠  Error: {exc}")
            start_btn[0].config(state="normal", text="Retry")

        def on_cancel() -> None:
            self._stop_flag[0] = True
            status_var.set("⏹  Cancelling… finishing current file, please wait.")
            detail_var.set("")

        btn = tk.Button(btn_frame, text="▶  Start Scan",
                        command=do_scan, relief="flat",
                        bg=_ACCENT, fg="white",
                        font=(ff, fn, "bold"), padx=12, pady=5)
        btn.pack(side="left", padx=6)
        start_btn.append(btn)

        tk.Button(btn_frame, text="✗  Cancel", command=on_cancel,
                  relief="flat", bg="#e0e0e0",
                  font=(ff, fn), padx=8, pady=5).pack(side="left", padx=6)

        root.update_idletasks()
        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        rw, rh = 620, 220
        root.geometry(f"{rw}x{rh}+{(sw - rw) // 2}+{(sh - rh) // 2}")
        root.mainloop()


# =========================================================== launcher app ===

class LauncherApp:
    """
    Two-button launcher: 'Scan Archive' and 'Archive Email'.
    Also shows DB stats so the user knows when it was last scanned.
    """

    def __init__(self, cfg: dict) -> None:
        self._cfg = cfg
        self._root = tk.Tk()
        self._root.title("Email Archiver")
        self._root.configure(bg=_BG)
        self._root.resizable(False, False)
        self._build()

    def _build(self) -> None:
        ff = self._cfg["ui"]["font_family"]
        fn = self._cfg["ui"]["font_size_normal"]
        ft = self._cfg["ui"]["font_size_title"]

        # Header
        hdr = tk.Frame(self._root, bg=_ACCENT, pady=12)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📧  Email Archiver", bg=_ACCENT, fg="white",
                 font=(ff, ft + 2, "bold"), padx=20).pack(side="left")

        # DB stats
        self._stats_var = tk.StringVar(value="Loading stats…")
        tk.Label(self._root, textvariable=self._stats_var,
                 bg=_BG, font=(ff, fn - 1), fg="#666",
                 anchor="w", padx=20, pady=8).pack(fill="x")

        ttk.Separator(self._root).pack(fill="x")

        # Buttons
        btn_area = tk.Frame(self._root, bg=_BG, pady=20, padx=20)
        btn_area.pack()

        tk.Button(
            btn_area, text="🗂  Scan Archive",
            command=self._on_scan,
            relief="flat", bg=_ACCENT, fg="white",
            font=(ff, fn + 1, "bold"), padx=20, pady=10, width=18,
        ).grid(row=0, column=0, padx=10)

        tk.Button(
            btn_area, text="📥  Archive Email",
            command=self._on_archive,
            relief="flat", bg="#2c5f8a", fg="white",
            font=(ff, fn + 1, "bold"), padx=20, pady=10, width=18,
        ).grid(row=0, column=1, padx=10)

        # Config path note
        from email_archiver.config import CONFIG_FILE
        tk.Label(
            self._root,
            text=f"Config: {CONFIG_FILE}",
            bg=_BG, fg="#bbb", font=(ff, 8),
            anchor="w", padx=20, pady=4,
        ).pack(fill="x")

        # Load stats async
        self._root.after(200, self._load_stats)

        # Size + centre
        self._root.update_idletasks()
        sw, sh = self._root.winfo_screenwidth(), self._root.winfo_screenheight()
        rw, rh = 480, 230
        self._root.geometry(f"{rw}x{rh}+{(sw - rw) // 2}+{(sh - rh) // 2}")

    def _load_stats(self) -> None:
        try:
            from email_archiver.database.models import get_connection
            from email_archiver.database.repository import EmailRepository
            conn = get_connection(self._cfg["database"]["path"])
            repo = EmailRepository(conn)
            n_emails = repo.count_emails()
            n_folders = repo.count_folders()
            last = repo.get_last_scan_time()
            conn.close()
            self._stats_var.set(
                f"Index: {n_emails:,} emails  •  {n_folders:,} folders  •  Last scan: {last}"
            )
        except Exception:
            self._stats_var.set("Index: not built yet — run 'Scan Archive' first.")

    def _on_scan(self) -> None:
        self._root.withdraw()
        ScanWindow(self._cfg).run()
        self._root.deiconify()
        self._load_stats()

    def _on_archive(self) -> None:
        self._root.withdraw()
        ArchiveDialog(self._cfg).run()
        self._root.deiconify()

    def run(self) -> None:
        self._root.mainloop()

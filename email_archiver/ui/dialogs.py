"""
Reusable Tkinter dialog helpers.

All dialogs are modal (grab_set) and centred on screen.
They block until the user closes them and return a result value.
"""
from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Any


def _centre(win: tk.Toplevel | tk.Tk, w: int, h: int) -> None:
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")


# -------------------------------------------------------- confirm dialog ----

def confirm_archive(
    parent: tk.Misc,
    folder_path: str,
    email_subject: str,
) -> str:
    """
    Show a confirmation dialog before archiving.

    Returns:
        "yes"         – proceed with archiving
        "open_folder" – open the folder in Explorer, do NOT archive
        "no"          – cancel
    """
    result: list[str] = ["no"]

    dlg = tk.Toplevel(parent)
    dlg.title("Confirm Archive")
    dlg.resizable(False, False)
    dlg.grab_set()
    dlg.lift()
    dlg.focus_force()
    _centre(dlg, 780, 180)

    # Message
    msg = (
        f"Archive email to:\n\n"
        f"  {folder_path}\n\n"
        f"Subject: {email_subject[:100]}"
    )
    tk.Label(dlg, text=msg, anchor="w", justify="left",
             wraplength=740, padx=14, pady=10).pack(fill="x")

    # Buttons
    btn_frame = tk.Frame(dlg)
    btn_frame.pack(pady=(0, 12))

    def on_yes():
        result[0] = "yes"
        dlg.destroy()

    def on_open():
        result[0] = "open_folder"
        dlg.destroy()

    def on_no():
        result[0] = "no"
        dlg.destroy()

    tk.Button(btn_frame, text="✓  Archive",  width=14, command=on_yes,
              bg="#2d6a4f", fg="white", relief="flat", padx=6).pack(side="left", padx=5)
    tk.Button(btn_frame, text="📂  Open Folder", width=14, command=on_open,
              relief="flat", padx=6).pack(side="left", padx=5)
    tk.Button(btn_frame, text="✗  Cancel",  width=10, command=on_no,
              relief="flat", padx=6).pack(side="left", padx=5)

    dlg.wait_window()
    return result[0]


# ------------------------------------------------------- browse dialog -----

def browse_folder(parent: tk.Misc, initial_dir: str = "") -> str | None:
    """Open a native folder-picker dialog. Returns the selected path or None."""
    chosen = filedialog.askdirectory(
        parent=parent,
        title="Select archive folder",
        initialdir=initial_dir or os.path.expanduser("~"),
        mustexist=True,
    )
    return chosen or None


# --------------------------------------------------- progress dialog -------

class ProgressDialog:
    """
    Non-blocking progress dialog for long-running operations (scanner).
    Call update() from the worker thread via root.after() or directly.
    """

    def __init__(self, parent: tk.Misc, title: str = "Scanning…") -> None:
        self._dlg = tk.Toplevel(parent)
        self._dlg.title(title)
        self._dlg.resizable(False, False)
        self._dlg.grab_set()
        _centre(self._dlg, 660, 160)

        self._label = tk.Label(
            self._dlg, text="Initialising…", anchor="w",
            padx=14, pady=8, wraplength=630
        )
        self._label.pack(fill="x")

        self._bar = ttk.Progressbar(
            self._dlg, mode="determinate", length=620
        )
        self._bar.pack(padx=14, pady=(0, 8))

        self._detail = tk.Label(
            self._dlg, text="", anchor="w", fg="#666",
            font=("Segoe UI", 8), padx=14, pady=2, wraplength=630
        )
        self._detail.pack(fill="x")

        self._cancelled = False
        btn = tk.Button(
            self._dlg, text="Cancel", command=self._cancel,
            relief="flat", padx=8
        )
        btn.pack(pady=(4, 10))

        self._dlg.protocol("WM_DELETE_WINDOW", self._cancel)

    def update(self, current: int, total: int, detail: str = "") -> None:
        if total > 0:
            pct = int(current / total * 100)
            self._bar["value"] = pct
            self._label.config(text=f"Indexing: {current:,} / {total:,} emails ({pct}%)")
        else:
            self._bar.config(mode="indeterminate")
            self._bar.start(10)
            self._label.config(text=f"Indexing… {current:,} emails processed")

        if detail:
            # Show only last 80 chars of a long path
            self._detail.config(text=detail[-80:] if len(detail) > 80 else detail)
        self._dlg.update_idletasks()

    def finish(self, message: str = "Done!") -> None:
        self._bar.stop()
        self._bar["value"] = 100
        self._label.config(text=message)
        self._detail.config(text="")
        self._dlg.update_idletasks()

    def close(self) -> None:
        self._dlg.destroy()

    def _cancel(self) -> None:
        self._cancelled = True
        self._dlg.destroy()

    @property
    def cancelled(self) -> bool:
        return self._cancelled

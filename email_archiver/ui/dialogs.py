"""
Reusable Tkinter dialog helpers.

All dialogs are modal (grab_set) and centred on screen.
They block until the user closes them and return a result value.
"""
from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog


def _centre(win: tk.Toplevel | tk.Tk, w: int, h: int) -> None:
    win.update_idletasks()
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")


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

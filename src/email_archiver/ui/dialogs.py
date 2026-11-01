"""
Reusable Tkinter dialog helpers.

Exports a single helper, ``browse_folder``, which wraps the native
``filedialog.askdirectory`` call with sensible defaults.
"""
from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog


def browse_folder(parent: tk.Misc, initial_dir: str = "") -> str | None:
    """Open a native folder-picker dialog. Returns the selected path or None."""
    chosen = filedialog.askdirectory(
        parent=parent,
        title="Select archive folder",
        initialdir=initial_dir or os.path.expanduser("~"),
        mustexist=True,
    )
    return chosen or None

"""
Resolve the folder shown in the foremost open Windows File Explorer window.

Used by the archive dialog's "Explorer folder" button: when none of the ranked
suggestions is right but the correct destination is already open on screen,
archiving into it should not require walking a folder picker.

Two layers, deliberately split:

- ``pick_foremost_folder`` is pure — given the eligible windows and the current
  z-order it decides which one wins. It has no COM dependency and is unit
  tested directly.
- ``get_current_explorer_folder`` does the Windows-specific part: enumerate the
  shell's open windows, keep the ones showing a real filesystem folder, read the
  top-level window z-order, and hand both to the pure picker.

"Foremost" means highest in the z-order — the Explorer window the user looked
at last, not an arbitrary one from the shell's collection. The archive dialog
itself is a topmost window but is not an Explorer window, so it never wins.
"""
from __future__ import annotations

import logging
import os
from collections.abc import Sequence

logger = logging.getLogger(__name__)

# Shell windows whose hosting executable is not Explorer (legacy Internet
# Explorer instances still surface in the same collection) are not destinations.
_EXPLORER_EXE = "explorer.exe"


def pick_foremost_folder(
    folders_by_hwnd: dict[int, str],
    z_order: Sequence[int],
) -> str | None:
    """
    Choose the winning folder from the eligible Explorer windows.

    Args:
        folders_by_hwnd: window handle → filesystem folder path, in the order
            the shell enumerated them (roughly oldest window first).
        z_order: top-level window handles, topmost first, as reported by the
            window manager. May contain handles that are not Explorer windows.

    Returns:
        The folder of the highest-z-order eligible window; when none of the
        eligible windows appears in ``z_order`` at all, the last-enumerated
        one (the most recently opened) as a fallback; ``None`` when there are
        no eligible windows.
    """
    if not folders_by_hwnd:
        return None

    for hwnd in z_order:
        folder = folders_by_hwnd.get(hwnd)
        if folder is not None:
            return folder

    # No eligible window was found in the z-order (minimised on some Windows
    # builds, or the enumeration raced a window closing). Falling back to the
    # newest window is closer to "the one the user just opened" than the oldest.
    return list(folders_by_hwnd.values())[-1]


def _enumerate_explorer_folders() -> dict[int, str]:
    """Map window handle → folder path for every open File Explorer window
    that is showing a real filesystem folder.

    Virtual shell locations (This PC, Control Panel, Quick Access, Recycle Bin)
    have no usable path and are skipped, as are non-Explorer shell windows.
    """
    import win32com.client  # noqa: PLC0415 - Windows-only, imported at use site

    folders: dict[int, str] = {}
    for window in win32com.client.Dispatch("Shell.Application").Windows():
        try:
            if not str(window.FullName or "").lower().endswith(_EXPLORER_EXE):
                continue
            hwnd = int(window.HWND)
            path = str(window.Document.Folder.Self.Path or "")
        except Exception as exc:  # a window can close mid-enumeration
            logger.debug("Skipping a shell window: %s", exc)
            continue

        if path and os.path.isdir(path):
            folders[hwnd] = path

    return folders


def _top_level_z_order() -> list[int]:
    """Top-level window handles, topmost first."""
    import win32gui  # noqa: PLC0415 - Windows-only, imported at use site

    handles: list[int] = []
    # EnumWindows walks top-level windows in z-order, topmost first.
    win32gui.EnumWindows(lambda hwnd, acc: acc.append(hwnd), handles)
    return handles


def get_current_explorer_folder() -> str | None:
    """
    Return the filesystem path shown in the foremost open File Explorer window.

    Returns ``None`` — and logs why — when no Explorer window is showing a real
    folder, or when the shell cannot be reached at all. The caller is expected
    to tell the user rather than pick a destination of its own.
    """
    try:
        import pythoncom  # noqa: PLC0415 - Windows-only, imported at use site

        # The Tk callback runs on the main thread, which may not have been
        # COM-initialised yet. CoInitialize is reference-counted and safe to
        # call again on a thread that already has an apartment.
        pythoncom.CoInitialize()
    except Exception as exc:
        logger.warning("Cannot initialise COM to read Explorer windows: %s", exc)
        return None

    try:
        folders = _enumerate_explorer_folders()
    except Exception as exc:
        logger.warning("Cannot enumerate Explorer windows: %s", exc)
        return None

    if not folders:
        logger.info("No open File Explorer window is showing a real folder")
        return None

    try:
        z_order = _top_level_z_order()
    except Exception as exc:
        # Without a z-order the newest-window fallback still gives an answer.
        logger.warning("Cannot read the window z-order: %s", exc)
        z_order = []

    chosen = pick_foremost_folder(folders, z_order)
    logger.info("Foremost Explorer folder: %s", chosen)
    return chosen

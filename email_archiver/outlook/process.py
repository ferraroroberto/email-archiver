"""
The Outlook process: whether it runs, and starting it for batch mode.

``OutlookClient.is_running()`` checks the process list without starting
Outlook, which matters for the dialog's fast launch. ``ensure_running()`` is its
deliberate opposite, used only by batch mode, which has no user to open Outlook
for it. It also owns the lifetime of what it starts: an outlook.exe that never
publishes its COM object is terminated by the handle ``ensure_running()`` holds,
never by image name, so an Outlook it merely attached to is out of reach.
"""
from __future__ import annotations

import logging
import time
from typing import Any

from email_archiver.outlook.mapi import OutlookUnavailableError

logger = logging.getLogger(__name__)

# How long ensure_running() waits for a freshly launched Outlook to publish its
# COM object. Outlook's first start on a cold profile is genuinely slow.
DEFAULT_START_TIMEOUT_SECONDS = 60.0
_POLL_INTERVAL_SECONDS = 1.0
# How long the teardown waits for a terminated Outlook to actually go away
# before giving up and saying so in the error it raises.
_TERMINATE_WAIT_SECONDS = 10.0


def tasklist_has_image(image_name: str) -> bool:
    """
    Return True if `image_name` (e.g. "OUTLOOK.EXE") appears in the Windows
    process list, using the `tasklist` console tool.

    A failed query is *not* the same fact as "the process is not running", so
    every failure is logged before falling back to False — otherwise a broken
    query is indistinguishable from a quiet machine.
    """
    import subprocess  # noqa: PLC0415
    import sys  # noqa: PLC0415

    is_windows = sys.platform == "win32"
    try:
        result = subprocess.run(
            ["tasklist", "/FI", f"IMAGENAME eq {image_name}", "/NH"],
            capture_output=True, timeout=5,
            # tasklist writes the OEM code page, not the parent's locale.
            # text=True decodes with the ambient locale, which under
            # PYTHONUTF8=1 is UTF-8: the OEM bytes then fail to decode, stdout
            # comes back None, and the result silently reads as "not running".
            encoding="oem" if is_windows else "utf-8",
            errors="replace",
            creationflags=subprocess.CREATE_NO_WINDOW if is_windows else 0,
        )
    except (OSError, subprocess.SubprocessError) as exc:
        logger.warning(
            "Could not run tasklist to check for %s (%s: %s) - "
            "treating as not running.", image_name, type(exc).__name__, exc,
        )
        return False

    if result.returncode != 0:
        logger.warning(
            "tasklist exited %s while checking for %s (%s) - "
            "treating as not running.",
            result.returncode, image_name, (result.stderr or "").strip(),
        )
        return False

    if not (result.stdout or "").strip():
        logger.warning(
            "tasklist returned no output while checking for %s - cannot tell "
            "whether it is running; treating as not running.", image_name,
        )
        return False

    return image_name.upper() in result.stdout.upper()


def _outlook_executable() -> str | None:
    """Path to the registered outlook.exe, or None when it cannot be found.

    Reads the ``App Paths`` registry entry Windows itself uses to resolve
    ``outlook.exe``, falling back to a PATH lookup. Deliberately not
    ``Dispatch("Outlook.Application")``: Dispatch starts a *hidden* instance
    that behaves differently (no explorer, add-ins in a different state) and
    that the user cannot see or interact with.
    """
    import shutil  # noqa: PLC0415
    import sys  # noqa: PLC0415

    if sys.platform == "win32":
        try:
            import winreg  # noqa: PLC0415

            for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                try:
                    with winreg.OpenKey(
                        hive,
                        r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths"
                        r"\OUTLOOK.EXE",
                    ) as key:
                        path, _ = winreg.QueryValueEx(key, "")
                        if path:
                            return str(path).strip('"')
                except OSError:
                    continue
        except ImportError:  # pragma: no cover - winreg ships with CPython
            pass

    return shutil.which("outlook.exe")


def get_active_application() -> Any | None:
    """Return the running Outlook Application, or None if none is published.

    ``GetActiveObject`` only ever attaches to an Outlook the user (or we)
    started; unlike ``Dispatch`` it never starts one, which is what makes it
    safe to call in a poll loop.
    """
    try:
        import win32com.client  # noqa: PLC0415

        return win32com.client.GetActiveObject("Outlook.Application")
    except Exception:
        return None


def _spawn_outlook(exe: str) -> Any:
    """Start ``exe`` and return the handle that owns the started process.

    A seam on purpose: the teardown in ``ensure_running()`` is unit-tested
    against a fake spawner, because the only Outlook on this machine is the
    user's own and no test may be able to reach it.
    """
    import subprocess  # noqa: PLC0415
    import sys  # noqa: PLC0415

    # Deliberately WITHOUT CREATE_NO_WINDOW: this is the one spawn in the
    # project whose window is meant to be visible — a hidden Outlook is
    # exactly what ensure_running() exists to avoid.
    return subprocess.Popen(  # noqa: S603
        [exe],
        creationflags=(
            subprocess.CREATE_NEW_PROCESS_GROUP
            if sys.platform == "win32"
            else 0
        ),
    )


def _terminate_spawned_outlook(proc: Any) -> str:
    """End the Outlook *this run started*, and report what happened to it.

    Takes the handle returned by ``_spawn_outlook`` and nothing else. Ending it
    through that handle rather than by image name or a PID lookup is the whole
    point: an Outlook that was already running — the user's own, on their
    own desktop — is not reachable from here, and a live handle keeps its
    PID reserved, so the call cannot land on a reused PID either.

    Returns one sentence for the ``OutlookUnavailableError`` message, so the
    ``outlook_unavailable`` error document records whether the process was
    cleaned up or is still out there.
    """
    pid = getattr(proc, "pid", None)
    try:
        if proc.poll() is not None:
            logger.info("The outlook.exe started (PID %s) had already exited.", pid)
            return f"The outlook.exe it started (PID {pid}) had already exited."
        proc.terminate()
        proc.wait(timeout=_TERMINATE_WAIT_SECONDS)
    except Exception as exc:
        logger.warning(
            "Could not terminate the outlook.exe started (PID %s): %s", pid, exc
        )
        return (
            f"The outlook.exe it started (PID {pid}) could NOT be terminated "
            f"({type(exc).__name__}: {exc}) and may still be running."
        )
    logger.warning("Terminated the outlook.exe this run started (PID %s).", pid)
    return f"The outlook.exe it started (PID {pid}) was terminated."


def ensure_running(timeout: float = DEFAULT_START_TIMEOUT_SECONDS) -> Any:
    """Return the Outlook Application object, starting Outlook if needed.

    The opposite of ``OutlookClient.is_running()``, and used only by batch
    mode, which runs unattended and has nobody to open Outlook for it. Starts the
    *registered* ``outlook.exe`` (visible, with its explorer, exactly as
    the user's own shortcut would) and then polls ``GetActiveObject`` until
    the COM object appears or ``timeout`` elapses.

    When the wait times out, the process this call started is terminated
    before the error is raised — see ``_terminate_spawned_outlook``. An
    Outlook that was already running is only ever attached to, never
    started and never terminated.

    Raises:
        OutlookUnavailableError: Outlook could not be started or never
            published its COM object inside the timeout. A loud failure on
            purpose — every batch verb needs Outlook, so continuing would
            report an empty Inbox nobody ever read. The message says what
            became of the process this call started.
    """
    app = get_active_application()
    if app is not None:
        return app

    exe = _outlook_executable()
    if exe is None:
        raise OutlookUnavailableError(
            "Outlook is not running and outlook.exe could not be located "
            "(no App Paths registry entry and not on PATH)."
        )

    logger.info("Outlook is not running; starting %s", exe)
    try:
        proc = _spawn_outlook(exe)
    except OSError as exc:
        raise OutlookUnavailableError(
            f"Could not start Outlook ({exe}): {exc}"
        ) from exc

    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        time.sleep(_POLL_INTERVAL_SECONDS)
        app = get_active_application()
        if app is not None:
            logger.info("Outlook is up.")
            return app

    # Nothing else owns this process's lifetime. Left running it outlives
    # the run — on a scheduled unattended run, an invisible orphan
    # holding the profile and OST against the user's own Outlook, one more
    # every time the wait times out (#78).
    teardown = _terminate_spawned_outlook(proc)
    raise OutlookUnavailableError(
        f"Outlook was started but did not publish its COM object within "
        f"{timeout:.0f}s. It may be showing a profile or password prompt, "
        f"or have no desktop to show one on. {teardown}"
    )

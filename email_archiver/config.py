"""
Configuration loader.

Resolves all relative paths against the project root so that the app
can be launched from any working directory (e.g. via Stream Deck).
"""
from __future__ import annotations

import logging
from pathlib import Path
from typing import Any

import yaml

# Project root = the 'archiver/' directory that contains this package
PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_FILE = PROJECT_ROOT / "config" / "config.yaml"

# Windows MAX_PATH is 260 (including the terminating NUL → 259 usable chars).
# We stay a few chars under to leave headroom for the OS, COM marshalling, and
# any internal use of long-path prefixes. This single budget is consumed by
# BOTH the scanner (skips over-long existing paths) and the archiver (shortens
# filenames so the path it writes never overflows). Used when config omits the
# key so older config.yaml files keep working.
DEFAULT_MAX_PATH_LENGTH = 255

_config: dict[str, Any] | None = None


def load_config() -> dict[str, Any]:
    """Load and cache the YAML config. Safe to call multiple times."""
    global _config
    if _config is not None:
        return _config

    if not CONFIG_FILE.exists():
        raise FileNotFoundError(
            f"Config file not found: {CONFIG_FILE}\n"
            "Copy config/config.yaml and set archive.root_paths."
        )

    with open(CONFIG_FILE, encoding="utf-8") as fh:
        _config = yaml.safe_load(fh)

    _resolve_paths(_config)
    return _config


def get_archive_roots(cfg: dict[str, Any]) -> list[str]:
    """Return the configured archive root paths.

    Reads the canonical ``archive.root_paths`` list, falling back to the
    legacy singular ``archive.root_path`` key when ``root_paths`` is absent.
    This migration shim lives here and nowhere else — every caller (scanner,
    UI, headless scan) goes through this function so the legacy-key handling
    is defined exactly once. Falsy/empty entries are filtered out.
    """
    archive = cfg["archive"]
    raw_paths = archive.get("root_paths")
    if not raw_paths:
        raw_paths = [archive.get("root_path")]
    return [p for p in raw_paths if p]


def get_max_path_length(cfg: dict[str, Any]) -> int:
    """Return the unified maximum path-length budget (in chars).

    Reads the canonical ``path.max_length`` key, falling back to
    ``DEFAULT_MAX_PATH_LENGTH`` when the section/key is absent so older
    ``config.yaml`` files (and the legacy ``scanning.max_path_length`` layout)
    keep working. Both the scanner and the archiver go through this function so
    the budget is defined in exactly one place — never split across a config
    value and a module constant again.
    """
    path_cfg = cfg.get("path") or {}
    value = path_cfg.get("max_length")
    if value is None:
        return DEFAULT_MAX_PATH_LENGTH
    return int(value)


def _resolve_paths(cfg: dict[str, Any]) -> None:
    """Convert relative paths in the config to absolute paths."""
    db_path = Path(cfg["database"]["path"])
    if not db_path.is_absolute():
        cfg["database"]["path"] = str(PROJECT_ROOT / db_path)

    log_path = Path(cfg["logging"]["file"])
    if not log_path.is_absolute():
        cfg["logging"]["file"] = str(PROJECT_ROOT / log_path)

    # Ensure directories exist
    Path(cfg["database"]["path"]).parent.mkdir(parents=True, exist_ok=True)
    Path(cfg["logging"]["file"]).parent.mkdir(parents=True, exist_ok=True)


def setup_logging(cfg: dict[str, Any] | None = None) -> None:
    """Configure root logger from config. Call once at startup."""
    if cfg is None:
        cfg = load_config()

    log_cfg = cfg["logging"]
    level = getattr(logging, log_cfg["level"].upper(), logging.INFO)

    fmt = "%(asctime)s [%(levelname)s] %(name)s – %(message)s"
    datefmt = "%Y-%m-%d %H:%M:%S"

    handlers: list[logging.Handler] = [logging.StreamHandler()]
    try:
        handlers.append(logging.FileHandler(log_cfg["file"], encoding="utf-8"))
    except OSError as exc:
        logging.warning("Cannot open log file %s: %s", log_cfg["file"], exc)

    logging.basicConfig(level=level, format=fmt, datefmt=datefmt, handlers=handlers)

    # extract_msg is very chatty about missing MAPI streams and encoding
    # fallbacks – these are normal for real-world .msg files, suppress them.
    logging.getLogger("extract_msg").setLevel(logging.ERROR)

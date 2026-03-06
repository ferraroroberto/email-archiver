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

_config: dict[str, Any] | None = None


def load_config() -> dict[str, Any]:
    """Load and cache the YAML config. Safe to call multiple times."""
    global _config
    if _config is not None:
        return _config

    if not CONFIG_FILE.exists():
        raise FileNotFoundError(
            f"Config file not found: {CONFIG_FILE}\n"
            "Copy config/config.yaml and set archive.root_path."
        )

    with open(CONFIG_FILE, encoding="utf-8") as fh:
        _config = yaml.safe_load(fh)

    _resolve_paths(_config)
    return _config


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

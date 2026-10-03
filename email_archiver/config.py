"""
Configuration loader.

Resolves all relative paths against the project root so that the app
can be launched from any working directory (e.g. via Stream Deck).
"""
from __future__ import annotations

import json
import logging
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

logger = logging.getLogger(__name__)

# Project root = the 'archiver/' directory that contains this package
PROJECT_ROOT = Path(__file__).resolve().parent.parent
CONFIG_FILE = PROJECT_ROOT / "config" / "config.yaml"
# Machine-local and gitignored: it holds personal addresses and this repo is
# public. config/mailboxes.sample.json is the tracked template.
MAILBOXES_FILE = PROJECT_ROOT / "config" / "mailboxes.json"
MAILBOXES_SCHEMA_VERSION = 1
# The alias of the one mailbox synthesized when there is no registry file.
SYNTHESIZED_ALIAS = "default"

# Windows MAX_PATH is 260 (including the terminating NUL → 259 usable chars).
# We stay a few chars under to leave headroom for the OS, COM marshalling, and
# any internal use of long-path prefixes. This single budget is consumed by
# BOTH the scanner (skips over-long existing paths) and the archiver (shortens
# filenames so the path it writes never overflows). Used when config omits the
# key so older config.yaml files keep working.
DEFAULT_MAX_PATH_LENGTH = 255

# Archived filenames carry the sequence prefix only (``NNN - name.ext``) unless
# the sent-date prefix is explicitly switched on. Off by default: the shorter
# form leaves more MAX_PATH headroom in deep folders, and the date is already
# inside the .msg. See ``get_date_prefix_mode`` / ``get_date_prefix_enabled``.
DEFAULT_DATE_PREFIX_ENABLED = False

# The third ``naming.date_prefix`` value (alongside ``true``/``false``): infer
# the form per destination folder from what it already holds. See
# ``email_archiver.archiver.archiver.infer_date_prefix``.
DATE_PREFIX_AUTO = "auto"

# Where batch `apply` parks a mail once it is on disk, and the category it
# stamps on it. The folder is looked up by name directly under the mailbox's
# root and created if missing; the category makes it obvious in Outlook which
# mails a batch run touched, and is what `revert` removes again.
DEFAULT_ARCHIVE_FOLDER = "Archive"
DEFAULT_ARCHIVED_CATEGORY = "Archived by task-os"

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


def get_date_prefix_mode(cfg: dict[str, Any]) -> bool | str:
    """Return the raw ``naming.date_prefix`` setting: ``True``, ``False``, or
    the literal ``"auto"`` (case-insensitive in the YAML, normalised here).

    Falls back to ``DEFAULT_DATE_PREFIX_ENABLED`` (``False``) when the
    section or key is absent, so a ``config.yaml`` written before the toggle
    existed keeps the default sequence-only naming. This is the one place
    the raw key is read — the archiver goes through it (or the narrower
    ``get_date_prefix_enabled`` below) rather than reaching into the config
    dict itself. ``"auto"`` means "infer per destination folder"; resolving
    that inference is the archiver's job, not this accessor's.
    """
    naming_cfg = cfg.get("naming") or {}
    value = naming_cfg.get("date_prefix")
    if value is None:
        return DEFAULT_DATE_PREFIX_ENABLED
    if isinstance(value, str) and value.strip().lower() == DATE_PREFIX_AUTO:
        return DATE_PREFIX_AUTO
    return bool(value)


def get_date_prefix_enabled(cfg: dict[str, Any]) -> bool:
    """Return whether archived filenames get the email's sent date prefixed,
    as a plain boolean starting position — ``"auto"`` resolves to ``False``
    here since no destination folder is known yet.

    Used where only a boolean makes sense (the dialog checkbox's initial
    position before any suggestion is highlighted); a caller that needs to
    honour ``"auto"`` by inferring from a folder should use
    ``get_date_prefix_mode`` together with
    ``email_archiver.archiver.archiver.resolve_date_prefix_for_folder``
    instead.
    """
    mode = get_date_prefix_mode(cfg)
    return mode if isinstance(mode, bool) else False


def get_outlook_archive_folder(cfg: dict[str, Any]) -> str:
    """Return the Outlook folder batch ``apply`` moves a filed mail into.

    Reads ``outlook.archive_folder``, falling back to
    ``DEFAULT_ARCHIVE_FOLDER``. Like the other accessors here this is the one
    place the key is read, so the batch layer never reaches into the config
    dict itself. A blank value falls back rather than resolving to the mailbox
    root, which would "move" the mail somewhere no one expects.
    """
    outlook_cfg = cfg.get("outlook") or {}
    value = outlook_cfg.get("archive_folder")
    return str(value).strip() if value and str(value).strip() else DEFAULT_ARCHIVE_FOLDER


def get_outlook_category(cfg: dict[str, Any]) -> str:
    """Return the Outlook category batch ``apply`` stamps on a filed mail.

    Reads ``outlook.category``, falling back to
    ``DEFAULT_ARCHIVED_CATEGORY``. ``revert`` removes exactly this category, so
    changing it between an apply and its revert leaves the old one in place --
    the files and the folder move still revert correctly.
    """
    outlook_cfg = cfg.get("outlook") or {}
    value = outlook_cfg.get("category")
    return (
        str(value).strip() if value and str(value).strip() else DEFAULT_ARCHIVED_CATEGORY
    )


def get_outlook_self_address(cfg: dict[str, Any]) -> str | None:
    """Return the configured ``outlook.self_address``, or ``None`` when unset.

    The address batch ``draft`` blind-copies on every draft. Optional: when it
    is absent the draft verb falls back to the default sending account's SMTP
    address, read from Outlook (see ``email_archiver.draft.resolve_self_address``).
    A blank value counts as unset rather than as an address.
    """
    outlook_cfg = cfg.get("outlook") or {}
    value = outlook_cfg.get("self_address")
    return str(value).strip() if value and str(value).strip() else None


# ------------------------------------------------------------- mailboxes ---

class MailboxRegistryError(ValueError):
    """``config/mailboxes.json`` exists but is not a usable registry."""


class MailboxUnknownError(LookupError):
    """A mailbox name that is neither an alias nor an extra alias in the registry."""


_MAILBOX_KEYS = {"address", "display_name", "aliases", "archive_folder"}
_REGISTRY_KEYS = {"schema_version", "default", "mailboxes"}


@dataclass(frozen=True)
class Mailbox:
    """One mailbox in the registry.

    ``synthesized`` is true only for the single mailbox made up when there is
    no registry file: its ``address`` is empty and the Outlook client resolves
    it to the profile's default store, which is today's behaviour. A registry
    mailbox is always resolved by its ``address`` and never by the default store.
    """

    alias: str
    address: str = ""
    display_name: str = ""
    aliases: tuple[str, ...] = ()
    archive_folder: str | None = None   # None = outlook.archive_folder
    synthesized: bool = False


@dataclass(frozen=True)
class MailboxRegistry:
    """Every mailbox this install knows, and which one a verb uses by default."""

    default: str
    mailboxes: dict[str, Mailbox] = field(default_factory=dict)
    synthesized: bool = False

    def select(self, name: str | None = None) -> Mailbox:
        """The mailbox called ``name`` (an alias or extra alias, any case),
        or the default one for ``None``.

        Raises:
            MailboxUnknownError: ``name`` matches nothing in the registry.
        """
        if name is None:
            return self.mailboxes[self.default]
        wanted = name.strip().casefold()
        for mailbox in self.mailboxes.values():
            if wanted == mailbox.alias.casefold() or wanted in (a.casefold() for a in mailbox.aliases):
                return mailbox
        raise MailboxUnknownError(
            f"no mailbox called {name!r} in the registry; known: {', '.join(self.mailboxes)}"
        )


def _synthesized_registry() -> MailboxRegistry:
    mailbox = Mailbox(alias=SYNTHESIZED_ALIAS, synthesized=True)
    return MailboxRegistry(
        default=SYNTHESIZED_ALIAS, mailboxes={SYNTHESIZED_ALIAS: mailbox}, synthesized=True,
    )


def _optional_str(entry: dict[str, Any], key: str, where: str) -> str:
    value = entry.get(key, "")
    if not isinstance(value, str):
        raise MailboxRegistryError(f"{where}: {key} must be a string")
    return value.strip()


def _parse_mailbox(alias: str, entry: Any) -> Mailbox:
    where = f"mailboxes.{alias}"
    if not isinstance(entry, dict):
        raise MailboxRegistryError(f"{where} must be an object")
    unknown = sorted(set(entry) - _MAILBOX_KEYS)
    if unknown:
        # A misspelt archive_folder would otherwise file into "Archive" silently.
        raise MailboxRegistryError(f"{where}: unknown key(s) {', '.join(unknown)}")
    address = _optional_str(entry, "address", where)
    if "@" not in address:
        raise MailboxRegistryError(f"{where}: address must be an e-mail address")
    aliases = entry.get("aliases", [])
    if not isinstance(aliases, list) or not all(isinstance(a, str) and a.strip() for a in aliases):
        raise MailboxRegistryError(f"{where}: aliases must be a list of non-blank strings")
    archive_folder = _optional_str(entry, "archive_folder", where)
    if "archive_folder" in entry and not archive_folder:
        raise MailboxRegistryError(f"{where}: archive_folder must not be blank")
    return Mailbox(
        alias=alias,
        address=address,
        display_name=_optional_str(entry, "display_name", where),
        aliases=tuple(a.strip() for a in aliases),
        archive_folder=archive_folder or None,
    )


def parse_mailbox_registry(data: Any) -> MailboxRegistry:
    """Validate a registry document into a :class:`MailboxRegistry`.

    Raises:
        MailboxRegistryError: the document is not a usable registry; the
            message names the offending key.
    """
    if not isinstance(data, dict):
        raise MailboxRegistryError("the registry must be a JSON object")
    unknown = sorted(set(data) - _REGISTRY_KEYS)
    if unknown:
        raise MailboxRegistryError(f"unknown top-level key(s) {', '.join(unknown)}")
    if data.get("schema_version") != MAILBOXES_SCHEMA_VERSION:
        raise MailboxRegistryError(
            f"schema_version must be {MAILBOXES_SCHEMA_VERSION}, got {data.get('schema_version')!r}"
        )
    raw = data.get("mailboxes")
    if not isinstance(raw, dict) or not raw:
        raise MailboxRegistryError("mailboxes must be a non-empty object")

    mailboxes: dict[str, Mailbox] = {}
    names: set[str] = set()
    addresses: set[str] = set()
    for alias, entry in raw.items():
        if not alias.strip():
            raise MailboxRegistryError("a mailbox alias must not be blank")
        mailbox = _parse_mailbox(alias.strip(), entry)
        # Every name must pick exactly one mailbox, or --mailbox could act on
        # the wrong one.
        for name in (mailbox.alias, *mailbox.aliases):
            if name.casefold() in names:
                raise MailboxRegistryError(f"the name {name!r} is used by more than one mailbox")
            names.add(name.casefold())
        if mailbox.address.casefold() in addresses:
            raise MailboxRegistryError(f"mailboxes.{alias}: address is used by another mailbox")
        addresses.add(mailbox.address.casefold())
        mailboxes[mailbox.alias] = mailbox

    default = data.get("default")
    if not isinstance(default, str) or default.strip() not in mailboxes:
        raise MailboxRegistryError(f"default must name one of: {', '.join(mailboxes)}")
    return MailboxRegistry(default=default.strip(), mailboxes=mailboxes)


def load_mailbox_registry(path: Path | None = None) -> MailboxRegistry:
    """Load ``config/mailboxes.json``, or synthesize today's single mailbox.

    No file is not an error: the registry is one mailbox resolved to Outlook's
    default store, which is what every verb did before the registry existed.
    A file that is present but unreadable or malformed *is* an error — falling
    back to the default store then could act on the wrong mailbox.

    Raises:
        MailboxRegistryError: the file exists but is not a usable registry.
    """
    path = path or MAILBOXES_FILE
    if not path.exists():
        logger.info(
            "ℹ️ No %s; using Outlook's default store as the only mailbox.", path.name
        )
        return _synthesized_registry()
    try:
        with open(path, encoding="utf-8") as fh:
            data = json.load(fh)
    except (OSError, json.JSONDecodeError) as exc:
        raise MailboxRegistryError(f"{path}: {type(exc).__name__}: {exc}") from exc
    try:
        return parse_mailbox_registry(data)
    except MailboxRegistryError as exc:
        raise MailboxRegistryError(f"{path}: {exc}") from exc


def get_mailbox_archive_folder(cfg: dict[str, Any], mailbox: Mailbox | None) -> str:
    """The Outlook folder ``apply`` files into for ``mailbox``: its own
    ``archive_folder`` when the registry sets one, else ``outlook.archive_folder``."""
    if mailbox is not None and mailbox.archive_folder:
        return mailbox.archive_folder
    return get_outlook_archive_folder(cfg)


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

    # Records carry archive paths and mail subjects, accented as often as not.
    # Under a captured stream Python writes stderr in the ANSI code page, so a
    # caller decoding it as UTF-8 (task-os does, without setting PYTHONUTF8)
    # gets U+FFFD and \uXXXX escapes in the failure detail it shows. Fixed
    # here, the one call every entry point makes, so it depends on no caller
    # remembering the env var (issue #87). pythonw has no stderr at all.
    if sys.stderr is not None:
        try:
            sys.stderr.reconfigure(encoding="utf-8", errors="backslashreplace")
        except (AttributeError, OSError):  # a stream that cannot be reconfigured
            pass

    handlers: list[logging.Handler] = [logging.StreamHandler()]
    try:
        handlers.append(logging.FileHandler(log_cfg["file"], encoding="utf-8"))
    except OSError as exc:
        logging.warning("Cannot open log file %s: %s", log_cfg["file"], exc)

    logging.basicConfig(level=level, format=fmt, datefmt=datefmt, handlers=handlers)

    # extract_msg is very chatty about missing MAPI streams and encoding
    # fallbacks – these are normal for real-world .msg files, suppress them.
    logging.getLogger("extract_msg").setLevel(logging.ERROR)

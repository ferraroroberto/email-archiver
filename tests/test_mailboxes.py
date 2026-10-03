"""Tests for the mailbox registry and per-store resolution (issue #108).

Every Outlook folder is taken from the store of one resolved mailbox, matched
by address; Outlook's default store is used only when there is no registry
file at all. A registry mailbox with no store is an error, never a quiet fall
back to the default store. Everything here is synthetic.
"""
from __future__ import annotations

import json
import sys
import types
from pathlib import Path

import pytest

import main_batch
from email_archiver import batch, config
from email_archiver.config import (
    Mailbox,
    MailboxRegistryError,
    MailboxUnknownError,
    get_mailbox_archive_folder,
    load_mailbox_registry,
    parse_mailbox_registry,
)
from email_archiver.outlook import mapi, process
from email_archiver.outlook.client import OutlookClient
from email_archiver.outlook.mapi import MailboxNotInOutlookError
from email_archiver.outlook.stores import find_store, resolve_store
from tests.test_batch import (  # noqa: F401 - fixtures are used by name
    _mail,
    archive_root,
    cfg,
)
from tests.test_batch_targeted import TargetedFakeClient

REPO_ROOT = Path(__file__).resolve().parent.parent
OWNER = "owner@example.invalid"
SECOND = "second@example.invalid"


def _registry_doc(**overrides) -> dict:
    doc = {
        "schema_version": 1,
        "default": "owner",
        "mailboxes": {
            "owner": {"address": OWNER, "display_name": "Owner", "aliases": ["me"]},
            "second": {"address": SECOND, "archive_folder": "Filed"},
        },
    }
    doc.update(overrides)
    return doc


# --------------------------------------------------------------- registry ---

def test_no_registry_file_is_one_mailbox_on_the_default_store(tmp_path):
    registry = load_mailbox_registry(tmp_path / "absent.json")

    assert registry.synthesized
    mailbox = registry.select(None)
    assert (mailbox.alias, mailbox.address, mailbox.synthesized) == ("default", "", True)


def test_a_registry_selects_by_alias_extra_alias_and_any_case():
    registry = parse_mailbox_registry(_registry_doc())

    assert registry.select(None).alias == "owner"
    assert registry.select("ME").alias == "owner"
    assert registry.select(" Second ").address == SECOND
    with pytest.raises(MailboxUnknownError):
        registry.select("nobody")


def test_the_tracked_sample_is_a_valid_registry():
    with open(REPO_ROOT / "config" / "mailboxes.sample.json", encoding="utf-8") as fh:
        registry = parse_mailbox_registry(json.load(fh))
    assert set(registry.mailboxes) == {"owner", "second"}


def test_the_live_registry_is_gitignored():
    ignored = (REPO_ROOT / ".gitignore").read_text(encoding="utf-8").splitlines()
    assert "config/mailboxes.json" in ignored


@pytest.mark.parametrize("doc", [
    [],
    _registry_doc(schema_version=2),
    _registry_doc(default="nobody"),
    _registry_doc(mailboxes={}),
    _registry_doc(extra=True),
    _registry_doc(mailboxes={"owner": {"address": "not-an-address"}}),
    _registry_doc(mailboxes={"owner": {"address": OWNER, "archive_foldr": "X"}}),
    _registry_doc(mailboxes={"owner": {"address": OWNER, "archive_folder": " "}}),
    _registry_doc(mailboxes={"owner": {"address": OWNER, "aliases": "me"}}),
    _registry_doc(mailboxes={
        "owner": {"address": OWNER}, "second": {"address": SECOND, "aliases": ["OWNER"]},
    }),
    _registry_doc(mailboxes={"owner": {"address": OWNER}, "second": {"address": OWNER.upper()}}),
])
def test_a_malformed_registry_is_refused(doc):
    with pytest.raises(MailboxRegistryError):
        parse_mailbox_registry(doc)


def test_an_unreadable_registry_file_is_refused_not_ignored(tmp_path):
    path = tmp_path / "mailboxes.json"
    path.write_text("{ not json", encoding="utf-8")
    with pytest.raises(MailboxRegistryError):
        load_mailbox_registry(path)


def test_a_mailbox_archive_folder_overrides_the_config(cfg):
    registry = parse_mailbox_registry(_registry_doc())

    assert get_mailbox_archive_folder(cfg, registry.select("second")) == "Filed"
    assert get_mailbox_archive_folder(cfg, registry.select("owner")) == "Archive"
    assert get_mailbox_archive_folder(cfg, None) == "Archive"


# --------------------------------------------------------- store resolution ---

class _Folder:
    def __init__(self, name: str, entry_id: str = "") -> None:
        self.Name = name
        self.EntryID = entry_id or name
        self.Items = types.SimpleNamespace(Count=0)


class _Folders:
    def __init__(self) -> None:
        self.items: list[_Folder] = []

    @property
    def Count(self) -> int:  # noqa: N802 - COM's spelling
        return len(self.items)

    def Item(self, i: int) -> _Folder:  # noqa: N802 - COM's spelling
        return self.items[i - 1]

    def Add(self, name: str) -> _Folder:  # noqa: N802 - COM's spelling
        folder = _Folder(name)
        self.items.append(folder)
        return folder


class _Store:
    def __init__(self, display_name: str, store_id: str) -> None:
        self.DisplayName = display_name
        self.StoreID = store_id
        self.root = types.SimpleNamespace(Name=display_name, Folders=_Folders())

    def GetDefaultFolder(self, kind: int) -> _Folder:  # noqa: N802 - COM's spelling
        return _Folder(f"{self.StoreID}:{kind}")

    def GetRootFolder(self) -> types.SimpleNamespace:  # noqa: N802 - COM's spelling
        return self.root


class _Collection:
    def __init__(self, items: list) -> None:
        self._items = items
        self.Count = len(items)

    def Item(self, i: int):  # noqa: N802 - COM's spelling
        return self._items[i - 1]


class _Namespace:
    """Two stores, one per address. Reading ``DefaultStore`` is recorded: a
    registry mailbox must never be answered with it."""

    def __init__(self, accounts: list, stores: list[_Store]) -> None:
        self.Accounts = _Collection(accounts)
        self.Stores = _Collection(stores)
        self.default_store_reads = 0
        self._default = stores[0] if stores else None

    @property
    def DefaultStore(self) -> _Store:  # noqa: N802 - COM's spelling
        self.default_store_reads += 1
        return self._default


def _account(smtp: str, store: _Store) -> types.SimpleNamespace:
    return types.SimpleNamespace(SmtpAddress=smtp, DeliveryStore=store)


OWNER_STORE = _Store(OWNER, "store-owner")
SECOND_STORE = _Store(SECOND, "store-second")


def test_a_mailbox_resolves_through_the_account_that_owns_its_address():
    namespace = _Namespace(
        [_account(OWNER, OWNER_STORE), _account(SECOND.upper(), SECOND_STORE)],
        [OWNER_STORE, _Store("misleading", "x"), SECOND_STORE],
    )
    assert find_store(namespace, SECOND) is SECOND_STORE
    assert namespace.default_store_reads == 0


def test_a_store_named_after_the_address_is_found_when_no_account_lists_it():
    # Outlook's Accounts collection can lag behind a newly added account.
    namespace = _Namespace([_account(OWNER, OWNER_STORE)], [OWNER_STORE, SECOND_STORE])
    assert find_store(namespace, SECOND) is SECOND_STORE


def test_two_stores_named_after_one_address_are_refused_not_guessed():
    namespace = _Namespace([], [_Store(SECOND, "a"), _Store(SECOND, "b")])
    with pytest.raises(MailboxNotInOutlookError):
        find_store(namespace, SECOND)


def test_a_registry_mailbox_with_no_store_never_falls_back_to_the_default_store():
    namespace = _Namespace([_account(OWNER, OWNER_STORE)], [OWNER_STORE])
    with pytest.raises(MailboxNotInOutlookError):
        resolve_store(namespace, Mailbox(alias="second", address=SECOND))
    assert namespace.default_store_reads == 0


def test_only_the_synthesized_mailbox_uses_the_default_store():
    namespace = _Namespace([], [OWNER_STORE])
    assert resolve_store(namespace, Mailbox(alias="default", synthesized=True)) is OWNER_STORE
    assert resolve_store(namespace, None) is OWNER_STORE


def _client(monkeypatch, mailbox: Mailbox | None, namespace: _Namespace) -> OutlookClient:
    client = OutlookClient(mailbox)
    monkeypatch.setattr(client, "_namespace", lambda: namespace)
    return client


def test_the_client_takes_every_folder_from_its_mailbox_store(monkeypatch):
    namespace = _Namespace(
        [_account(OWNER, OWNER_STORE), _account(SECOND, SECOND_STORE)],
        [OWNER_STORE, SECOND_STORE],
    )
    client = _client(monkeypatch, Mailbox(alias="second", address=SECOND), namespace)

    assert client._inbox().Name == f"store-second:{mapi.OL_FOLDER_INBOX}"
    assert client._sent_items().Name == f"store-second:{mapi.OL_FOLDER_SENT_MAIL}"
    filed = client._named_folder("Filed")
    assert SECOND_STORE.root.Folders.items == [filed]
    assert OWNER_STORE.root.Folders.items == []
    assert namespace.default_store_reads == 0


def test_the_client_refuses_to_walk_an_inbox_its_mailbox_does_not_have(monkeypatch):
    namespace = _Namespace([_account(OWNER, OWNER_STORE)], [OWNER_STORE])
    client = _client(monkeypatch, Mailbox(alias="second", address=SECOND), namespace)

    with pytest.raises(MailboxNotInOutlookError):
        list(client.iter_inbox())
    assert namespace.default_store_reads == 0


def test_the_draft_folder_check_uses_the_mailbox_store(monkeypatch):
    draft_parent = SECOND_STORE.GetDefaultFolder(mapi.OL_FOLDER_DRAFTS)
    mail = types.SimpleNamespace(Sent=False, Parent=draft_parent)
    namespace = _Namespace([_account(SECOND, SECOND_STORE)], [OWNER_STORE, SECOND_STORE])
    namespace.GetItemFromID = lambda entry_id: mail
    client = _client(monkeypatch, Mailbox(alias="second", address=SECOND), namespace)

    assert client.open_draft("entry-1") is mail


# ------------------------------------------------------------ mailboxes doc ---

class _StatusClient:
    def __init__(self, mailbox: Mailbox, outcome) -> None:
        self.mailbox = mailbox
        self._outcome = outcome

    def store(self):
        if isinstance(self._outcome, Exception):
            raise self._outcome
        return object()

    def default_account_smtp(self) -> str:
        return OWNER


def test_the_mailboxes_document_reports_present_missing_and_unknown():
    registry = parse_mailbox_registry(_registry_doc(mailboxes={
        "owner": {"address": OWNER},
        "second": {"address": SECOND},
        "third": {"address": "third@example.invalid"},
    }))
    outcomes = {
        "owner": None,
        "second": MailboxNotInOutlookError("no store"),
        "third": RuntimeError("COM went away"),
    }

    doc = batch.mailboxes(registry, lambda mb: _StatusClient(mb, outcomes[mb.alias]))

    statuses = {e["alias"]: e["outlook_status"] for e in doc["mailboxes"]}
    assert statuses == {"owner": "present", "second": "missing", "third": "unknown"}
    assert doc["registry"] == "file" and doc["default"] == "owner"


def test_with_outlook_not_running_every_status_is_unknown_never_present():
    registry = parse_mailbox_registry(_registry_doc())

    doc = batch.mailboxes(registry, None)

    assert {e["outlook_status"] for e in doc["mailboxes"]} == {"unknown"}


def test_the_synthesized_mailbox_reports_the_default_account_address():
    doc = batch.mailboxes(load_mailbox_registry(Path("absent.json")),
                          lambda mb: _StatusClient(mb, None))

    assert doc["registry"] == "synthesized"
    assert doc["mailboxes"][0]["address"] == OWNER


# --------------------------------------------------------- process contract ---

@pytest.fixture
def batch_process(cfg, tmp_path, monkeypatch, capsys):
    """``main_batch.main`` with a fake pythoncom and a registry file in tmp_path."""
    monkeypatch.setattr(main_batch, "load_config", lambda: cfg)
    monkeypatch.setattr(main_batch, "setup_logging", lambda _cfg: None)
    monkeypatch.setitem(sys.modules, "pythoncom", types.SimpleNamespace(
        CoInitialize=lambda: None, CoUninitialize=lambda: None,
    ))
    registry_path = tmp_path / "mailboxes.json"
    monkeypatch.setattr(config, "MAILBOXES_FILE", registry_path)
    built: list = []

    def run(argv: list[str], registry=None, client_factory=None) -> tuple[int, dict, list]:
        if registry is not None:
            registry_path.write_text(
                registry if isinstance(registry, str) else json.dumps(registry), encoding="utf-8",
            )

        def _factory(mailbox=None):
            client = client_factory(mailbox)
            built.append(client)
            return client

        monkeypatch.setattr(main_batch, "OutlookClient", _factory)
        code = main_batch.main(argv)
        return code, json.loads(capsys.readouterr().out), built

    return run


def test_a_mailbox_with_no_store_exits_2_and_never_plans(batch_process, monkeypatch):
    namespace = _Namespace([_account(SECOND, SECOND_STORE)], [SECOND_STORE])

    def _real_client(mailbox):
        client = OutlookClient(mailbox)
        monkeypatch.setattr(client, "_namespace", lambda: namespace)
        monkeypatch.setattr(client, "ensure_running", lambda timeout: None)
        return client

    def _no_plan(*args, **kwargs):
        raise AssertionError("plan must not run on an unresolved mailbox")

    monkeypatch.setattr(batch, "plan", _no_plan)

    code, doc, _ = batch_process(["plan"], registry=_registry_doc(), client_factory=_real_client)

    assert code == main_batch.EXIT_CANNOT_START
    assert doc["error"]["code"] == "mailbox_not_in_outlook"
    assert namespace.default_store_reads == 0


def test_a_malformed_registry_exits_2_before_outlook_is_touched(batch_process):
    code, doc, built = batch_process(
        ["plan"], registry="{ not json", client_factory=lambda mb: pytest.fail("Outlook touched"),
    )

    assert code == main_batch.EXIT_CANNOT_START
    assert doc["error"]["code"] == "mailbox_registry_invalid"
    assert built == []


@pytest.mark.parametrize("registry, alias", [(None, "default"), (_registry_doc(), "owner")])
def test_every_result_document_carries_its_mailbox(batch_process, registry, alias):
    code, doc, built = batch_process(
        ["plan", "--candidates", "0"], registry=registry,
        client_factory=lambda mb: TargetedFakeClient([_mail("m-1", "Synthetic subject")]),
    )

    assert code == main_batch.EXIT_OK
    assert doc["mailbox"] == alias
    assert len(doc["mails"]) == 1


def test_apply_files_into_the_mailbox_archive_folder(batch_process, archive_root, tmp_path):
    decisions = tmp_path / "decisions.json"
    decisions.write_text(json.dumps([
        {"message_id": "m-1", "folder_path": str(archive_root), "date_prefix": False},
    ]), encoding="utf-8")
    registry = _registry_doc(default="second")

    code, doc, built = batch_process(
        ["apply", "--decisions", str(decisions)], registry=registry,
        client_factory=lambda mb: TargetedFakeClient([_mail("m-1", "Synthetic subject")]),
    )

    assert code == main_batch.EXIT_OK
    assert (doc["mailbox"], doc["archive_folder"]) == ("second", "Filed")
    assert built[0].moves == [("m-1", "Filed")]


def test_the_mailboxes_verb_reports_unknown_when_outlook_is_not_running(batch_process, monkeypatch):
    monkeypatch.setattr(process, "get_active_application", lambda: None)

    code, doc, built = batch_process(
        ["mailboxes"], registry=_registry_doc(),
        client_factory=lambda mb: pytest.fail("no client without Outlook"),
    )

    assert code == main_batch.EXIT_OK
    assert [e["outlook_status"] for e in doc["mailboxes"]] == ["unknown", "unknown"]
    assert built == []

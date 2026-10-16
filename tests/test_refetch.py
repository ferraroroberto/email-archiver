"""
Tests for OutlookClient.item_ids() and refetch() — how ``apply`` lets go of a
mail and fetches it again from plain strings (issue #130).

The re-fetch must name the store the mail came from: ``GetItemFromID`` with no
StoreID leaves Outlook to find the mail in whichever store it opens first.
Nothing here reaches a real Outlook.
"""
from __future__ import annotations

import sys
import types

import pytest

from email_archiver.outlook import process
from email_archiver.outlook.client import OutlookClient


class _Store:
    StoreID = "store-0001"


class _Mail:
    EntryID = "entry-0001"
    Parent = _Store()


class _Unreadable:
    """A COM object whose every property read raises, as a disconnected one does."""

    def __getattr__(self, name: str):
        raise RuntimeError("the object is disconnected")


class _Namespace:
    def __init__(self) -> None:
        self.calls: list[tuple[str, ...]] = []

    def GetItemFromID(self, *ids: str):  # noqa: N802 - COM's spelling
        self.calls.append(ids)
        return _Mail()


@pytest.fixture
def namespace(monkeypatch) -> _Namespace:
    namespace = _Namespace()
    app = types.SimpleNamespace(GetNamespace=lambda name: namespace)
    monkeypatch.setattr(process, "get_active_application", lambda: app)
    # Dispatch wraps a raw COM object; there is none here to wrap.
    fake_client = types.SimpleNamespace(Dispatch=lambda obj: obj)
    monkeypatch.setitem(sys.modules, "win32com", types.SimpleNamespace(client=fake_client))
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)
    return namespace


def test_item_ids_reads_the_entry_and_store_ids():
    assert OutlookClient().item_ids(_Mail()) == ("entry-0001", "store-0001")


def test_item_ids_of_an_unreadable_mail_are_empty():
    assert OutlookClient().item_ids(_Unreadable()) == ("", "")


def test_refetch_names_the_store(namespace):
    assert OutlookClient().refetch("entry-0001", "store-0001") is not None
    assert namespace.calls == [("entry-0001", "store-0001")]


def test_refetch_without_a_store_id_still_asks_by_entry_id(namespace):
    assert OutlookClient().refetch("entry-0001", "") is not None
    assert namespace.calls == [("entry-0001",)]


def test_refetch_without_an_entry_id_asks_nothing(namespace):
    assert OutlookClient().refetch("", "store-0001") is None
    assert namespace.calls == []

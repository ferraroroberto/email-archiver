"""Tests for targeted batch mode (issue #71).

A second caller (life-os) files one specific mail, or a handful matching a
topic, into a folder it already knows. These tests pin what that caller relies
on, on top of the full-Inbox contract ``test_batch.py`` pins:

- ``plan`` filters (``message_ids`` / ``since`` / ``search`` / ``ref``) narrow
  the result without enumerating more of the Inbox than they must, and no
  filter leaves the document exactly as it was;
- ``--candidates 0`` lists mails with no ranking;
- ``"date_prefix": "auto"`` resolves the form against the destination folder;
- ``--category`` overrides the configured category on ``apply`` and ``revert``;
- ``apply`` refuses a ``folder_path`` outside the archive roots, per mail, with
  nothing written;
- task-os's decision shape (boolean ``date_prefix``, no category) still files.

Everything here is synthetic: no real folder name, address or subject.
"""
from __future__ import annotations

import json
import sys
import types
from datetime import datetime, timedelta, timezone
from pathlib import Path

import pytest

import main_batch
from email_archiver import batch
from email_archiver.outlook.client import (
    InboxMail,
    header_value,
    received_since_filter,
)
from tests.test_batch import (  # noqa: F401 - fixtures are used by name
    FakeOutlookClient,
    _mail,
    _seed_index,
    archive_root,
    cfg,
)


class TargetedFakeClient(FakeOutlookClient):
    """The batch fake plus the targeted-plan surface, recording what ran.

    ``iter_inbox_received_since`` deliberately returns the *whole* Inbox — the
    widest a minute-precise server-side Restrict could ever be — so a test only
    passes if ``batch.plan`` applies the exact date check itself.
    """

    def __init__(self, inbox, **kw) -> None:
        super().__init__(inbox, **kw)
        self.calls: list[str] = []
        self.ref_reads: list[str] = []

    def ensure_running(self, timeout: float = 0) -> None:
        return None

    def iter_inbox(self, preview_len: int = 500):
        self.calls.append("iter_inbox")
        yield from super().iter_inbox(preview_len)

    def iter_inbox_received_since(self, since: datetime, preview_len: int = 500):
        self.calls.append("iter_inbox_received_since")
        for item in list(self.folders[None]):
            yield self.read_mail(item, preview_len)

    def find_by_message_id(self, message_id: str, folder_name: str | None = None):
        self.calls.append("find_by_message_id")
        return super().find_by_message_id(message_id, folder_name)

    def read_mail(self, item, preview_len: int = 0) -> InboxMail:
        mail = super().read_mail(item, preview_len)
        mail.date_received = getattr(item, "received", None)
        return mail

    def archive_ref(self, item) -> str:
        self.ref_reads.append(item.message_id)
        return header_value(getattr(item, "headers", ""), "X-Archive-Ref")


def _received(message_id: str, subject: str, received: datetime, **kw):
    item = _mail(message_id, subject, **kw)
    item.received = received
    return item


# ------------------------------------------------------------------- plan ---

def test_plan_by_message_id_looks_the_mail_up_without_enumerating(cfg):
    client = TargetedFakeClient([
        _mail("a@example.invalid", "Wanted"),
        _mail("b@example.invalid", "Not wanted"),
    ])

    doc = batch.plan(client, cfg, candidates=0, filters=batch.PlanFilters(
        message_ids=("a@example.invalid", "gone@example.invalid"),
    ))

    assert "iter_inbox" not in client.calls
    assert "iter_inbox_received_since" not in client.calls
    assert [m["message_id"] for m in doc["mails"]] == ["a@example.invalid"]
    assert doc["mails"][0]["candidates"] == []
    assert doc["skipped"] == [{
        "entry_id": "", "subject": "", "reason": batch.SKIP_NOT_IN_INBOX,
        "message_id": "gone@example.invalid",
    }]
    assert doc["counts"] == {
        "inbox": 1, "planned": 1, "already_archived": 0, "skipped": 1,
    }
    assert doc["filters"] == {
        "message_ids": ["a@example.invalid", "gone@example.invalid"],
        "since": None, "search": [], "ref": None,
    }


def test_plan_since_keeps_only_mail_received_at_or_after_it(cfg):
    boundary = datetime(2026, 9, 15, 8, 0)
    client = TargetedFakeClient([
        _received("old@example.invalid", "Yesterday", boundary - timedelta(minutes=1)),
        _received("edge@example.invalid", "On the dot", boundary),
        _received("new@example.invalid", "Later", boundary + timedelta(hours=3)),
    ])

    doc = batch.plan(client, cfg, candidates=0,
                     filters=batch.PlanFilters(since=boundary))

    assert client.calls == ["iter_inbox_received_since"]
    assert [m["message_id"] for m in doc["mails"]] == [
        "edge@example.invalid", "new@example.invalid",
    ]
    assert doc["filters"]["since"] == "2026-09-15T08:00"


def test_plan_search_matches_every_term_case_insensitively_across_fields(cfg):
    client = TargetedFakeClient([
        _mail("a@example.invalid", "Quarterly INVOICE", sender="billing@vendor.invalid"),
        _mail("b@example.invalid", "Invoice", sender="someone@else.invalid"),
        _mail("c@example.invalid", "Unrelated", sender="billing@vendor.invalid"),
    ])

    doc = batch.plan(client, cfg, candidates=0, filters=batch.PlanFilters(
        search=("invoice", "VENDOR"),
    ))

    assert [m["message_id"] for m in doc["mails"]] == ["a@example.invalid"]


def test_plan_ref_reads_headers_only_for_mail_that_passed_the_cheaper_filters(cfg):
    wanted = _mail("a@example.invalid", "Filing test")
    wanted.headers = "Received: from x\r\nX-Archive-Ref: tok-123\r\nSubject: Filing test\r\n"
    decoy = _mail("b@example.invalid", "Filing test")
    decoy.headers = "X-Archive-Ref: tok-999\r\n"
    elsewhere = _mail("c@example.invalid", "Something else")
    elsewhere.headers = "X-Archive-Ref: tok-123\r\n"
    client = TargetedFakeClient([wanted, decoy, elsewhere])

    doc = batch.plan(client, cfg, candidates=0, filters=batch.PlanFilters(
        search=("filing test",), ref="tok-123",
    ))

    assert [m["message_id"] for m in doc["mails"]] == ["a@example.invalid"]
    assert client.ref_reads == ["a@example.invalid", "b@example.invalid"]


def test_plan_without_filters_has_no_filters_key(cfg):
    client = TargetedFakeClient([_mail("a@example.invalid", "Anything")])

    doc = batch.plan(client, cfg)

    assert "filters" not in doc
    assert client.calls == ["iter_inbox"]


def test_plan_with_zero_candidates_skips_the_ranking_even_with_a_seeded_index(
    cfg, archive_root
):
    _seed_index(cfg, archive_root, {"Project Alpha": ["Project Alpha kickoff"]})
    client = TargetedFakeClient([_mail("a@example.invalid", "Project Alpha update")])

    assert batch.plan(client, cfg)["mails"][0]["candidates"], "sanity: ranks by default"
    doc = batch.plan(client, cfg, candidates=0)

    assert doc["mails"][0]["candidates"] == []
    assert doc["counts"]["planned"] == 1


# ------------------------------------------------------------ date prefix ---

def _dated_folder(archive_root: Path, name: str) -> Path:
    folder = archive_root / name
    folder.mkdir()
    for n in (1, 2):
        (folder / f"2026-01-0{n} - 00{n} - earlier.msg").write_text("x", encoding="utf-8")
    return folder


def test_apply_auto_date_prefix_follows_the_destination_folder(cfg, archive_root):
    cfg["naming"]["date_prefix"] = "auto"
    dated = _dated_folder(archive_root, "Dated")
    fresh = archive_root / "Brand new"
    client = TargetedFakeClient([
        _mail("a@example.invalid", "Into dated", sent=datetime(2026, 3, 14, 9, 0)),
        _mail("b@example.invalid", "Into new", sent=datetime(2026, 3, 14, 9, 0)),
    ])

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dated), "date_prefix": "auto"},
        {"message_id": "b@example.invalid", "folder_path": str(fresh), "date_prefix": "auto"},
    ])

    into_dated, into_new = (Path(r["files"][0]).name for r in doc["results"])
    assert into_dated == "2026-03-14 - 003 - Into dated.msg"
    assert into_new == "001 - Into new.msg"


def test_apply_auto_date_prefix_gives_what_plan_would_under_a_fixed_config(
    cfg, archive_root
):
    """``"auto"`` means "the form plan reports for this folder" — under a fixed
    ``naming.date_prefix: false`` that is undated, whatever the folder holds."""
    assert cfg["naming"]["date_prefix"] is False
    dated = _dated_folder(archive_root, "Dated")
    client = TargetedFakeClient([_mail("a@example.invalid", "Fixed config")])

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dated), "date_prefix": "auto"},
    ])

    assert Path(doc["results"][0]["files"][0]).name == "003 - Fixed config.msg"


# --------------------------------------------------------------- category ---

def test_apply_and_revert_use_the_category_the_caller_names(cfg, archive_root):
    item = _mail("a@example.invalid", "Tagged differently")
    item.Categories = "Personal"
    client = TargetedFakeClient([item])

    applied = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(archive_root / "X")},
    ], category="Archived by life-os")

    assert applied["category"] == "Archived by life-os"
    assert item.Categories == "Personal; Archived by life-os"

    reverted = batch.revert(client, cfg, applied["results"], category="Archived by life-os")

    assert reverted["category"] == "Archived by life-os"
    assert reverted["results"][0]["ok"] is True
    assert item.Categories == "Personal"


# ------------------------------------------------------------------ guard ---

@pytest.mark.parametrize("where", ["sibling", "traversal", "prefix_twin", "other_drive"])
def test_apply_refuses_a_folder_outside_the_roots_and_writes_nothing(
    cfg, archive_root, tmp_path, where
):
    outside = {
        "sibling": tmp_path / "not-the-archive",
        "traversal": archive_root / ".." / "escaped",
        "prefix_twin": Path(f"{archive_root}-other"),
        "other_drive": Path("Q:/somewhere/else" if sys.platform == "win32" else "/elsewhere"),
    }[where]
    refused = _mail("a@example.invalid", "Keep me out")
    fine = _mail("b@example.invalid", "Inside")
    client = TargetedFakeClient([refused, fine])
    before = sorted(p for p in tmp_path.rglob("*") if p.name != "emails.db")

    doc = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(outside), "date_prefix": "auto"},
        {"message_id": "b@example.invalid", "folder_path": str(archive_root / "Inside")},
    ])

    bad, good = doc["results"]
    assert bad["ok"] is False
    assert bad["error"]["code"] == batch.ERROR_BAD_DECISION
    assert bad["error"]["message"].startswith(batch.REFUSED_OUTSIDE_ROOTS)
    assert bad["files"] == [] and bad["moved"] is False
    assert not Path(outside).exists()
    assert refused.saved_as == []
    assert refused in client.folders[None], "a refused mail stays in the Inbox"
    assert good["ok"] is True, "one refused decision must not stop the next"
    written_outside = [
        p for p in tmp_path.rglob("*")
        if p.name != "emails.db" and p not in before and not p.is_relative_to(archive_root)
    ]
    assert written_outside == []


# ------------------------------------------------------ task-os compat -----

def test_task_os_decision_shape_still_files_and_reverts(cfg, archive_root):
    """task-os's archive flow (``src/archive_batch.py``): boolean date_prefix
    handed back from a plan candidate, no category override, schema 1."""
    client = TargetedFakeClient([_mail("a@example.invalid", "Task-os shaped")])
    dest = archive_root / "Project Alpha"

    applied = batch.apply(client, cfg, [
        {"message_id": "a@example.invalid", "folder_path": str(dest), "date_prefix": False},
    ], renumber=True)

    assert applied["schema_version"] == 1
    assert applied["category"] == "Filed by batch"
    assert applied["results"][0]["ok"] is True
    assert client.folders["Archive"][0].Categories == "Filed by batch"

    reverted = batch.revert(client, cfg, applied["results"], renumber=True)
    assert reverted["schema_version"] == 1
    assert reverted["results"][0]["ok"] is True
    assert client.folders["Archive"] == []


# --------------------------------------------------------- pure helpers ----

def test_header_value_is_case_insensitive_and_unfolds_continuations():
    headers = "Subject: hi\r\nx-archive-ref: tok\r\n  -continued\r\nTo: a@example.invalid\r\n"
    assert header_value(headers, "X-Archive-Ref") == "tok -continued"
    assert header_value(headers, "Cc") == ""
    assert header_value(None, "X-Archive-Ref") == ""


def test_the_received_since_filter_is_written_in_utc():
    since = datetime(2026, 9, 15, 0, 30, tzinfo=timezone(timedelta(hours=2)))
    assert received_since_filter(since) == (
        "@SQL=\"urn:schemas:httpmail:datereceived\" >= '2026-09-14 22:30'"
    )


# ------------------------------------------------------ client surface -----

class _ComAccessor:
    def __init__(self, props: dict[str, str]) -> None:
        self._props = props

    def GetProperty(self, key: str):  # noqa: N802 - COM's spelling
        if key in self._props:
            return self._props[key]
        raise RuntimeError("property not found")


class _ComItems:
    """An ``Items`` collection: 1-based ``Item``, ``Count`` and ``Restrict``."""

    def __init__(self, items, *, restrict_raises: bool = False) -> None:
        self._items = list(items)
        self.restrict_raises = restrict_raises
        self.filters: list[str] = []

    @property
    def Count(self) -> int:  # noqa: N802 - COM's spelling
        return len(self._items)

    def Item(self, i: int):  # noqa: N802 - COM's spelling
        return self._items[i - 1]

    def Restrict(self, flt: str):  # noqa: N802 - COM's spelling
        if self.restrict_raises:
            raise RuntimeError("the store rejected the filter")
        self.filters.append(flt)
        return _ComItems(self._items[:1])


def _outlook_client_over(monkeypatch, items: _ComItems):
    import win32com.client

    from email_archiver.outlook import client as client_mod

    monkeypatch.setattr(win32com.client, "Dispatch", lambda obj: obj)
    client = client_mod.OutlookClient()
    monkeypatch.setattr(client, "_inbox", lambda: types.SimpleNamespace(Items=items))
    monkeypatch.setattr(client, "_read_mail", lambda item, _len: item)
    return client


def _com_mail(name: str):
    return types.SimpleNamespace(Class=43, name=name)


def test_the_client_restricts_by_received_date_server_side(monkeypatch):
    items = _ComItems([_com_mail("first"), _com_mail("second")])
    client = _outlook_client_over(monkeypatch, items)

    got = list(client.iter_inbox_received_since(datetime(2026, 9, 15, 8, 0)))

    assert [m.name for m in got] == ["first"]
    assert items.filters == [received_since_filter(datetime(2026, 9, 15, 8, 0))]


def test_the_client_walks_the_inbox_when_the_store_rejects_the_filter(monkeypatch):
    items = _ComItems([_com_mail("first"), _com_mail("second")], restrict_raises=True)
    client = _outlook_client_over(monkeypatch, items)

    got = list(client.iter_inbox_received_since(datetime(2026, 9, 15, 8, 0)))

    assert [m.name for m in got] == ["first", "second"]


def test_archive_ref_prefers_the_transport_headers_then_the_named_property():
    from email_archiver.outlook import client as client_mod

    client = client_mod.OutlookClient()
    transport = types.SimpleNamespace(PropertyAccessor=_ComAccessor({
        client_mod.DASL_TRANSPORT_HEADERS: "X-Archive-Ref: from-headers\r\n",
        client_mod.DASL_X_ARCHIVE_REF: "from-named",
    }))
    named_only = types.SimpleNamespace(PropertyAccessor=_ComAccessor({
        client_mod.DASL_X_ARCHIVE_REF: " from-named ",
    }))
    neither = types.SimpleNamespace(PropertyAccessor=_ComAccessor({}))

    assert client.archive_ref(transport) == "from-headers"
    assert client.archive_ref(named_only) == "from-named"
    assert client.archive_ref(neither) == ""


# ------------------------------------------------------ process contract ---

@pytest.fixture
def plan_process(cfg, monkeypatch, capsys):
    """``main_batch.main`` over the fake client, recording every one built."""
    monkeypatch.setattr(main_batch, "load_config", lambda: cfg)
    monkeypatch.setattr(main_batch, "setup_logging", lambda _cfg: None)
    monkeypatch.setitem(sys.modules, "pythoncom", types.SimpleNamespace(
        CoInitialize=lambda: None, CoUninitialize=lambda: None,
    ))
    clients: list[TargetedFakeClient] = []

    def run(argv: list[str], inbox=()) -> tuple[int, dict, list[TargetedFakeClient]]:
        def _factory() -> TargetedFakeClient:
            fake = TargetedFakeClient(list(inbox))
            clients.append(fake)
            return fake

        monkeypatch.setattr(main_batch, "OutlookClient", _factory)
        code = main_batch.main(argv)
        return code, json.loads(capsys.readouterr().out), clients

    return run


@pytest.mark.parametrize("argv", [
    ["plan", "--since", "15/09/2026"],
    ["plan", "--since", "2026-09-15T08:00+02:00"],
    ["plan", "--candidates", "-1"],
    ["plan", "--search", "  "],
    ["plan", "--ref", ""],
    ["plan", "--message-id", "<>"],
])
def test_a_bad_plan_filter_exits_2_without_starting_outlook(plan_process, argv):
    code, doc, clients = plan_process(argv)

    assert code == main_batch.EXIT_CANNOT_START
    assert doc["error"]["code"] == main_batch.ERROR_BAD_INPUT
    assert clients == []


def test_a_blank_category_exits_2_without_starting_outlook(plan_process, tmp_path):
    decisions = tmp_path / "decisions.json"
    decisions.write_text("[]", encoding="utf-8")

    code, doc, clients = plan_process(["apply", "--decisions", str(decisions), "--category", " "])

    assert code == main_batch.EXIT_CANNOT_START
    assert clients == []


def test_the_cli_passes_filters_and_zero_candidates_through(plan_process):
    code, doc, clients = plan_process(
        ["plan", "--message-id", "<a@example.invalid>", "--candidates", "0",
         "--search", "wanted"],
        inbox=[_mail("a@example.invalid", "Wanted"), _mail("b@example.invalid", "Wanted")],
    )

    assert code == main_batch.EXIT_OK
    assert [m["message_id"] for m in doc["mails"]] == ["a@example.invalid"]
    assert doc["filters"]["message_ids"] == ["a@example.invalid"]
    assert doc["filters"]["search"] == ["wanted"]
    assert "iter_inbox" not in clients[0].calls

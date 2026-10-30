"""The fake COM objects, fixtures and helpers the draft tests share.

Every address, path and subject here is synthetic.
"""
from __future__ import annotations

import json
import sys
import types
from pathlib import Path

import pytest

from email_archiver.cli import batch as main_batch

from email_archiver.outlook.drafts import (
    CreatedDraft,
    DraftUpdateError,
    ReplySourceError,
)


REPO_ROOT = Path(__file__).resolve().parents[1]
SELF = "me@example.invalid"


@pytest.fixture
def attachment(tmp_path) -> Path:
    path = tmp_path / "note.txt"
    path.write_text("synthetic attachment", encoding="utf-8")
    return path


def _spec(**overrides) -> dict:
    spec = {"to": ["someone@example.invalid"], "subject": "A subject", "body_text": "Hello"}
    spec.update(overrides)
    return spec


# --------------------------------------------------------------- fake COM ---

class _FakePropertyAccessor:
    def __init__(self, calls: list, refuse: bool) -> None:
        self._calls = calls
        self._refuse = refuse

    def SetProperty(self, name: str, value: str) -> None:  # noqa: N802 - COM's spelling
        if self._refuse:
            raise RuntimeError("the store refused the property")
        self._calls.append(("SetProperty", name, value))


class _FakeComAttachments:
    def __init__(self, calls: list) -> None:
        self._calls = calls
        self.paths: list[str] = []

    @property
    def Count(self) -> int:  # noqa: N802 - COM's spelling
        return len(self.paths)

    def Add(self, path: str) -> None:  # noqa: N802 - COM's spelling
        self._calls.append(("Attachments.Add", path))
        self.paths.append(path)

    def Remove(self, index: int) -> None:  # noqa: N802 - COM's spelling
        self._calls.append(("Attachments.Remove", index))
        del self.paths[index - 1]


class _FakeComMail:
    """A new MailItem that logs every call, in order, and every body write.

    ``Display`` puts a signature into ``HTMLBody``, the way Outlook does when it
    opens the compose window for an account with a default signature.
    """

    SIGNATURE_HTML = '<html><body lang=EN><div id="_MailAutoSig">-- sig</div></body></html>'

    def __init__(self, refuse_property: bool = False) -> None:
        self.calls: list = []
        self.To = self.CC = self.BCC = self.Subject = ""
        self._html = ""
        self.EntryID = "draft-entry-1"
        self.Attachments = _FakeComAttachments(self.calls)
        self.PropertyAccessor = _FakePropertyAccessor(self.calls, refuse_property)

    @property
    def HTMLBody(self) -> str:  # noqa: N802 - COM's spelling
        return self._html

    @HTMLBody.setter
    def HTMLBody(self, value: str) -> None:  # noqa: N802 - COM's spelling
        self.calls.append(("HTMLBody=",))
        self._html = value

    def Display(self, modal: bool) -> None:  # noqa: N802 - COM's spelling
        self.calls.append(("Display", modal))
        self._html = self.SIGNATURE_HTML

    def Save(self) -> None:  # noqa: N802 - COM's spelling
        self.calls.append(("Save",))

    def Send(self) -> None:  # noqa: N802 - COM's spelling
        self.calls.append(("Send",))


class _FakeApplication:
    def __init__(self, mail: _FakeComMail) -> None:
        self.mail = mail
        self.created: list[int] = []

    def CreateItem(self, kind: int) -> _FakeComMail:  # noqa: N802 - COM's spelling
        self.created.append(kind)
        return self.mail


class FakeDraftClient:
    """Stands in for OutlookClient's draft surface."""

    def __init__(self, account_smtp: str = SELF, ref_stamped: bool = True) -> None:
        self.account_smtp = account_smtp
        self.ref_stamped = ref_stamped
        self.account_lookups = 0
        self.drafts: list[dict] = []
        self.updates: list[dict] = []

    def ensure_running(self, timeout: float = 60.0) -> None:
        return None

    def store(self):
        # The real client resolves its mailbox's store; nothing to resolve here.
        return self

    def default_account_smtp(self) -> str:
        self.account_lookups += 1
        return self.account_smtp

    refuse_create: ReplySourceError | None = None

    def create_draft(self, **kwargs) -> CreatedDraft:
        if self.refuse_create is not None:
            raise self.refuse_create
        self.drafts.append(kwargs)
        return CreatedDraft(
            entry_id="draft-entry-1",
            ref_stamped=self.ref_stamped and bool(kwargs["ref"]),
            ref_reason="" if self.ref_stamped else "SetProperty refused: synthetic",
            displayed=kwargs["display"],
        )

    refuse_update: DraftUpdateError | None = None

    def update_draft(self, **kwargs) -> CreatedDraft:
        if self.refuse_update is not None:
            raise self.refuse_update
        self.updates.append(kwargs)
        return CreatedDraft(entry_id=kwargs["entry_id"], ref_stamped=bool(kwargs["ref"]),
                            displayed=kwargs["display"])


@pytest.fixture
def batch_process(tmp_path, monkeypatch, capsys):
    """``main_batch.main`` with a temp config, no log file, a stub pythoncom
    and an Outlook client that records whether it was ever constructed."""
    cfg = {"outlook": {}}
    monkeypatch.setattr(main_batch, "load_config", lambda: cfg)
    monkeypatch.setattr(main_batch, "setup_logging", lambda _cfg: None)
    monkeypatch.setitem(sys.modules, "pythoncom", types.SimpleNamespace(
        CoInitialize=lambda: None, CoUninitialize=lambda: None,
    ))
    clients: list[FakeDraftClient] = []

    def run(
        spec: dict, account_smtp: str = SELF, extra: tuple[str, ...] = (),
        refuse_update: DraftUpdateError | None = None,
        refuse_create: ReplySourceError | None = None,
    ) -> tuple[int, dict, list[FakeDraftClient]]:
        def _factory(mailbox=None) -> FakeDraftClient:
            fake = FakeDraftClient(account_smtp=account_smtp)
            fake.refuse_update = refuse_update
            fake.refuse_create = refuse_create
            clients.append(fake)
            return fake

        monkeypatch.setattr(main_batch, "OutlookClient", _factory)
        spec_path = tmp_path / "spec.json"
        spec_path.write_text(json.dumps(spec), encoding="utf-8")
        code = main_batch.main(["draft", "--spec", str(spec_path), *extra])
        return code, json.loads(capsys.readouterr().out), clients

    return run

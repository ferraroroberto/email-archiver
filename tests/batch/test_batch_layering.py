"""The document-building verbs share one small base and never import each other.

``batch_common`` says so in its docstring; this makes the claim fail a build
instead of waiting for the next audit to find a verb reaching through ``batch``.
"""
from __future__ import annotations

import ast
from pathlib import Path

import pytest

PACKAGE = Path(__file__).resolve().parents[2] / "src" / "email_archiver"

# Every module that builds on the shared base. ``batch`` is the one verb module
# they must not reach through.
BASE_USERS = ["batch_common", "batch_move", "dismiss", "draft", "find", "mail_filter", "send", "sync"]


def _imports_batch(source: str) -> list[str]:
    hits: list[str] = []
    for node in ast.walk(ast.parse(source)):
        if isinstance(node, ast.ImportFrom):
            if node.module == "email_archiver.batch":
                hits.append(f"from {node.module} import ...")
            elif node.module == "email_archiver":
                hits.extend(
                    f"from email_archiver import {a.name}" for a in node.names if a.name == "batch"
                )
        elif isinstance(node, ast.Import):
            hits.extend(f"import {a.name}" for a in node.names if a.name == "email_archiver.batch")
    return hits


@pytest.mark.parametrize("module", BASE_USERS)
def test_the_module_does_not_import_the_batch_verb_module(module):
    source = (PACKAGE / f"{module}.py").read_text(encoding="utf-8")
    assert _imports_batch(source) == []


def test_the_check_sees_each_way_of_importing_batch():
    assert _imports_batch("from email_archiver.batch import SCHEMA_VERSION")
    assert _imports_batch("from email_archiver import batch")
    assert _imports_batch("import email_archiver.batch")
    assert not _imports_batch("from email_archiver.batch_common import header")

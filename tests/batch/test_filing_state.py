"""The per-mail filing state: one row per reachable observation, and the
combinations the walk cannot reach are refused."""
from __future__ import annotations

import pytest

from email_archiver.filing_state import (
    FilingFacts,
    FilingState,
    Lookup,
    derive_filing_state,
    settle,
)

FOUND = Lookup.FOUND

# facts -> (state, ok, moved, categorized): every reachable end of a run.
REACHABLE = [
    (FilingFacts(refused=True), FilingState.REFUSED, (False, False, False)),
    (FilingFacts(lookup=Lookup.MISSING), FilingState.NOT_FOUND, (False, False, False)),
    (FilingFacts(lookup=Lookup.RAISED), FilingState.LOOKUP_FAILED, (False, False, False)),
    (FilingFacts(lookup=FOUND, archive_failed=True), FilingState.ARCHIVE_FAILED, (False, False, False)),
    (FilingFacts(lookup=FOUND), FilingState.MOVE_FAILED, (False, False, False)),
    (FilingFacts(lookup=FOUND, moved=True), FilingState.MOVED_UNTAGGED, (False, True, False)),
    (FilingFacts(lookup=FOUND, moved=True, categorized=True), FilingState.FILED, (True, True, True)),
    (
        FilingFacts(lookup=FOUND, kept_in_sent=True),
        FilingState.KEPT_IN_SENT_UNTAGGED, (False, False, False),
    ),
    (
        FilingFacts(lookup=FOUND, kept_in_sent=True, categorized=True),
        FilingState.FILED_KEPT_IN_SENT, (True, False, True),
    ),
]

IMPOSSIBLE = [
    pytest.param(FilingFacts(), id="nothing-observed"),
    pytest.param(FilingFacts(refused=True, lookup=FOUND), id="refused-but-looked-up"),
    pytest.param(FilingFacts(refused=True, moved=True), id="refused-but-moved"),
    pytest.param(FilingFacts(lookup=Lookup.MISSING, moved=True), id="missing-but-moved"),
    pytest.param(FilingFacts(lookup=Lookup.RAISED, archive_failed=True), id="raised-but-archived"),
    pytest.param(FilingFacts(lookup=FOUND, archive_failed=True, moved=True), id="not-archived-but-moved"),
    pytest.param(FilingFacts(lookup=FOUND, moved=True, kept_in_sent=True), id="moved-and-kept"),
    pytest.param(FilingFacts(lookup=FOUND, categorized=True), id="tagged-before-filed"),
]


@pytest.mark.parametrize(("facts", "state", "flags"), REACHABLE)
def test_each_reachable_observation_derives_its_state_and_result_flags(facts, state, flags):
    result = {"ok": None, "moved": None, "categorized": None}

    assert derive_filing_state(facts) is state
    assert settle(result, facts) is state
    assert (result["ok"], result["moved"], result["categorized"]) == flags


def test_every_state_is_reached_by_some_observation():
    assert {state for _, state, _ in REACHABLE} == set(FilingState)


@pytest.mark.parametrize("facts", IMPOSSIBLE)
def test_an_impossible_combination_is_refused(facts):
    # state-machine-proof: filing-state
    result = {"ok": "untouched", "moved": "untouched", "categorized": "untouched"}

    with pytest.raises(ValueError):
        derive_filing_state(facts)
    with pytest.raises(ValueError):
        settle(result, facts)
    assert set(result.values()) == {"untouched"}

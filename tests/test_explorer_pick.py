"""Tests for the Explorer-window pick — the pure half, no COM required."""
from __future__ import annotations

from email_archiver.explorer import pick_foremost_folder


def test_no_eligible_windows_returns_none():
    assert pick_foremost_folder({}, [111, 222]) is None


def test_single_window_wins():
    assert pick_foremost_folder({222: r"E:\archive\alpha"}, [999, 222]) == r"E:\archive\alpha"


def test_highest_z_order_window_wins_not_the_first_enumerated():
    # The shell enumerates roughly oldest-first; the user's foremost window is
    # the one highest in the z-order, which here is the *second* enumerated.
    folders = {111: r"E:\archive\old", 222: r"E:\archive\current"}
    z_order = [222, 111]
    assert pick_foremost_folder(folders, z_order) == r"E:\archive\current"


def test_non_explorer_windows_in_the_z_order_are_ignored():
    # The archive dialog itself is topmost but is not an Explorer window, so it
    # never appears in the eligible mapping and must not block the pick.
    folders = {111: r"E:\archive\alpha"}
    z_order = [900, 901, 111]
    assert pick_foremost_folder(folders, z_order) == r"E:\archive\alpha"


def test_falls_back_to_newest_window_when_z_order_has_no_match():
    # A minimised or racing window can be missing from the z-order snapshot;
    # the last-enumerated (newest) window is the closest stand-in.
    folders = {111: r"E:\archive\old", 222: r"E:\archive\newest"}
    assert pick_foremost_folder(folders, []) == r"E:\archive\newest"

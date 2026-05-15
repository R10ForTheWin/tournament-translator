#!/usr/bin/env python3
"""
Golden test suite for Tournament Translator parsers.

Run from the project root:
    python3 tests/test_parsers.py

Each fixture has an expected minimum game count and format. A test fails if:
  - Fewer games than expected are parsed
  - Wrong parser selected (format mismatch)
  - Duplicate game_ids within a sheet
  - Any team slot fails structural validation

Add a new fixture row when a new tournament format is encountered.
"""
from __future__ import annotations
import os, sys, re
from collections import Counter
from datetime import date, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from parsers.detect import load_and_parse
from parsers.validate import _valid_slot
from app import _expand_bracket_games, team_matches

FIXTURES_DIR = os.path.join(ROOT, "Tournaments Excels")

# ── Fixture definitions ──────────────────────────────────────────────────────
# (filename, expected_format, min_games, max_games)
FIXTURES = [
    ("2025 NEWPORT SPRING INVITE.xlsx",                                        "A",  160,  200),
    ("2026 KAP7 INTERNATIONAL.xlsx",                                           "A",  580,  650),
    ("2026 TURBO OC CUP.xlsx",                                                 "A",  280,  340),
    ("2026 KAP7 Futures WPL - Southern California - Presented by BIWPA.xlsx",  "B", 1200, 1400),
]

# ── Teams to spot-check after expansion ─────────────────────────────────────
# (filename, sheet, team, min_direct, max_direct, max_extras)
TEAM_CHECKS = [
    # Trojan Gold 16U Boys — the canonical WPL regression canary.
    # Local fixture is an older snapshot (~4 weekends), so direct count varies.
    # The critical guard is extras ≤ 5: more than that means cross-weekend
    # pool contamination has returned (was 31 before the fix).
    (
        "2026 KAP7 Futures WPL - Southern California - Presented by BIWPA.xlsx",
        "16u Boys",
        "trojan gold",
        8, 20,   # min/max direct (varies by snapshot date)
        5,       # max extras — the real regression guard
    ),
]

# ── Helpers ──────────────────────────────────────────────────────────────────

PASS = "\033[32mPASS\033[0m"
FAIL = "\033[31mFAIL\033[0m"

def _check(label: str, condition: bool, detail: str = "") -> bool:
    status = PASS if condition else FAIL
    suffix = f"  ({detail})" if detail else ""
    print(f"  [{status}] {label}{suffix}")
    return condition


# ── Tests ────────────────────────────────────────────────────────────────────

def test_fixture(fname: str, fmt: str, min_g: int, max_g: int) -> int:
    path = os.path.join(FIXTURES_DIR, fname)
    failures = 0
    print(f"\n{fname}")

    if not os.path.exists(path):
        print(f"  [SKIP] file not found: {path}")
        return 0

    games = load_and_parse(path)

    # Game count
    ok = _check(f"game count {len(games)} in [{min_g}, {max_g}]",
                min_g <= len(games) <= max_g,
                f"got {len(games)}")
    if not ok:
        failures += 1

    # Format
    formats = {g.get("format") for g in games} - {None}
    ok = _check(f"parser format is {fmt!r}", fmt in formats, f"got {formats}")
    if not ok:
        failures += 1

    # Per-sheet duplicate game_ids
    by_sheet: dict[str, list] = {}
    for g in games:
        by_sheet.setdefault(g.get("sheet", ""), []).append(g["game_id"])
    dup_sheets = []
    for sh, ids in by_sheet.items():
        dups = [gid for gid, n in Counter(ids).items() if n > 1]
        if dups:
            dup_sheets.append(f"{sh}: {dups[:3]}")
    ok = _check("no duplicate game_ids per sheet", not dup_sheets,
                "; ".join(dup_sheets))
    if not ok:
        failures += 1

    # Slot validity (sample first 200 games)
    bad_slots = []
    for g in games[:200]:
        for field in ("white_team", "dark_team"):
            slot = str(g.get(field) or "").strip()
            if not _valid_slot(slot):
                bad_slots.append(f"{g['game_id']}.{field}={slot!r}")
    ok = _check("team slots structurally valid (first 200 games)",
                not bad_slots, "; ".join(bad_slots[:3]))
    if not ok:
        failures += 1

    # At least one scored game (sanity: we're reading real data)
    scored = sum(1 for g in games if g.get("white_score") is not None)
    ok = _check(f"at least one scored game", scored > 0, f"{scored} scored")
    if not ok:
        failures += 1

    return failures


def test_team_expansion(fname: str, sheet: str, team: str,
                         min_d: int, max_d: int, max_extras: int) -> int:
    path = os.path.join(FIXTURES_DIR, fname)
    failures = 0
    label = f"{team!r} in {sheet!r}"
    print(f"\n{label}")

    if not os.path.exists(path):
        print(f"  [SKIP] file not found: {path}")
        return 0

    games = load_and_parse(path)
    div_games = [g for g in games if g["sheet"] == sheet]
    direct = sorted(
        [g for g in div_games
         if team_matches(g["white_team"], team) or team_matches(g["dark_team"], team)],
        key=lambda g: (g.get("date") or date.min, g.get("time") or datetime.min.time()),
    )
    extras = _expand_bracket_games(team, direct, div_games)

    ok = _check(f"direct game count {len(direct)} in [{min_d}, {max_d}]",
                min_d <= len(direct) <= max_d, f"got {len(direct)}")
    if not ok:
        failures += 1

    ok = _check(f"extras ≤ {max_extras} (cross-weekend contamination guard)",
                len(extras) <= max_extras, f"got {len(extras)}")
    if not ok:
        failures += 1

    # No non-placeholder extra should have a date before the team's first direct game
    if direct and extras:
        first_date = min(g["date"] for g in direct if g.get("date"))
        early = [g["game_id"] for g in extras
                 if not g.get("placeholder") and g.get("date") and g["date"] < first_date]
        ok = _check("no real extra predates team's first direct game",
                    not early, str(early))
        if not ok:
            failures += 1

    return failures


# ── Main ─────────────────────────────────────────────────────────────────────

def main():
    total_failures = 0

    print("=" * 60)
    print("Parser fixture tests")
    print("=" * 60)
    for args in FIXTURES:
        total_failures += test_fixture(*args)

    print("\n" + "=" * 60)
    print("Team expansion tests")
    print("=" * 60)
    for args in TEAM_CHECKS:
        total_failures += test_team_expansion(*args)

    print("\n" + "=" * 60)
    if total_failures == 0:
        print(f"\033[32mAll tests passed.\033[0m")
    else:
        print(f"\033[31m{total_failures} test(s) FAILED.\033[0m")
    print("=" * 60)
    sys.exit(0 if total_failures == 0 else 1)


if __name__ == "__main__":
    main()

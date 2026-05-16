#!/usr/bin/env python3
"""
Golden test suite for Tournament Translator parsers.

Run from the project root:
    python3 tests/test_parsers.py

Each fixture has an expected minimum game count and format. A test fails if:
  - Fewer/more games than expected are parsed
  - Wrong parser selected (format mismatch)
  - Duplicate game_ids within a sheet
  - Any team slot fails structural validation

IMPORTANT: After fixing any bug, add (or tighten) a test here so the bug
can never silently return. The test suite is the long-term memory of the
system — every entry encodes a lesson learned the hard way.
"""
from __future__ import annotations
import os, sys, re
from collections import Counter
from datetime import date, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from parsers.detect import load_and_parse
from parsers.validate import _valid_slot
from app import (
    _expand_bracket_games, _build_wpl_game_tree,
    team_matches, describe_slot, _SLOT_LIKE_RE,
)

FIXTURES_DIR = os.path.join(ROOT, "Tournaments Excels")

WPL_FILE = "2026 KAP7 Futures WPL - Southern California - Presented by BIWPA.xlsx"

# ── Fixture definitions ──────────────────────────────────────────────────────
# (filename, expected_format, min_games, max_games)
# Game count range should be wide enough to survive minor organizer edits
# but tight enough to catch a format switch or wholesale data loss.
FIXTURES = [
    ("2025 NEWPORT SPRING INVITE.xlsx",   "A",  160,  200),
    ("2026 KAP7 INTERNATIONAL.xlsx",      "A",  580,  650),
    ("2026 TURBO OC CUP.xlsx",            "A",  280,  340),
    # WPL Futures — all 5 weekends.  Updated range reflects full-season fixture
    # loaded after Weekend 5 championship (was 1200-1400 for 3-weekend snapshot).
    (WPL_FILE,                            "B", 1800, 2200),
]

# ── Cross-weekend contamination canary ───────────────────────────────────────
# (filename, sheet, team, min_direct, max_direct, max_extras)
# max_extras is the contamination guard: if expansion pulls in games from the
# wrong weekend, extras spike.  The canary value tracks the KNOWN-GOOD state;
# tighten it whenever the all-weekends fixture is refreshed.
TEAM_CHECKS = [
    # Trojan Gold, all weekends in fixture.  With championship data loaded,
    # extras = 3 (the 3 Sunday -B placeholder games for pool H placement).
    # If extras ever exceed 6 cross-weekend contamination has returned.
    (WPL_FILE, "16u Boys", "trojan gold", 14, 22, 6),
]

# ── Championship-weekend specific checks ─────────────────────────────────────
# These check the behavior we care about most on game day:
#   - Team sees the right number of direct games this weekend
#   - Bracket expansion finds Sunday games
#   - Tree builder returns a non-empty staircase with Sunday nodes
#   - Opponent labels are resolved (no raw slot strings leaked through)
#
# Schema: (sheet, team, anchor_date, min_direct, max_direct,
#           min_extras, min_tree_nodes, min_sun_nodes)
#
# anchor_date — Saturday of championship weekend; scopes checks to that weekend.
# min_extras  — lower bound prevents regression to "only 1 game showing."
# min_sun_nodes — must see at least this many Sunday nodes in the bracket.
CHAMPIONSHIP_CHECKS = [
    # Pool-slot seeded team (H1). 2 Sat pool games + 3 Sun placement extras.
    # Tree: root → pool game → 3×Sunday placeholder = 5 nodes.
    ("16u Boys", "trojan gold",     date(2026, 5, 16),  2, 2,  3, 5, 3),

    # Seed-number prelim team.  1 direct (prelim) + 8 extras (pool + placement).
    # Tree: prelim → pool → Sunday pool → 3×placement = 6 nodes.
    # Previously showed only 1 game — this is the regression canary for that bug.
    ("16u Boys", "trojan cardinal", date(2026, 5, 16),  1, 1,  6, 5, 3),

    # Pool-slot seeded team in a different pool (E1). Same structure as Gold.
    ("16u Boys", "imperial",        date(2026, 5, 16),  2, 2,  3, 4, 2),

    # Pool-slot seeded team in pool G (another seed-1 team).
    ("16u Boys", "socal",           date(2026, 5, 16),  2, 2,  3, 4, 2),
]


# ── Helpers ──────────────────────────────────────────────────────────────────

PASS = "\033[32mPASS\033[0m"
FAIL = "\033[31mFAIL\033[0m"

def _check(label: str, condition: bool, detail: str = "") -> bool:
    status = PASS if condition else FAIL
    suffix = f"  ({detail})" if detail else ""
    print(f"  [{status}] {label}{suffix}")
    return condition


# ── Tests ─────────────────────────────────────────────────────────────────────

def test_fixture(fname: str, fmt: str, min_g: int, max_g: int) -> int:
    path = os.path.join(FIXTURES_DIR, fname)
    failures = 0
    print(f"\n{fname}")

    if not os.path.exists(path):
        print(f"  [SKIP] file not found: {path}")
        return 0

    games = load_and_parse(path)

    ok = _check(f"game count {len(games)} in [{min_g}, {max_g}]",
                min_g <= len(games) <= max_g, f"got {len(games)}")
    if not ok: failures += 1

    formats = {g.get("format") for g in games} - {None}
    ok = _check(f"parser format is {fmt!r}", fmt in formats, f"got {formats}")
    if not ok: failures += 1

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
    if not ok: failures += 1

    bad_slots = []
    for g in games[:200]:
        for field in ("white_team", "dark_team"):
            slot = str(g.get(field) or "").strip()
            if not _valid_slot(slot):
                bad_slots.append(f"{g['game_id']}.{field}={slot!r}")
    ok = _check("team slots structurally valid (first 200 games)",
                not bad_slots, "; ".join(bad_slots[:3]))
    if not ok: failures += 1

    scored = sum(1 for g in games if g.get("white_score") is not None)
    ok = _check(f"at least one scored game", scored > 0, f"{scored} scored")
    if not ok: failures += 1

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
    if not ok: failures += 1

    ok = _check(f"extras ≤ {max_extras} (cross-weekend contamination guard)",
                len(extras) <= max_extras, f"got {len(extras)}")
    if not ok: failures += 1

    if direct and extras:
        first_date = min(g["date"] for g in direct if g.get("date"))
        early = [g["game_id"] for g in extras
                 if not g.get("placeholder") and g.get("date") and g["date"] < first_date]
        ok = _check("no real extra predates team's first direct game",
                    not early, str(early))
        if not ok: failures += 1

    return failures


def test_championship_team(sheet: str, team: str, anchor: date,
                            min_direct: int, max_direct: int,
                            min_extras: int, min_tree: int,
                            min_sun: int) -> int:
    """Check a team's championship-weekend bracket and expansion.

    This is the test that would have caught every game-day bug we've fixed:
    - Too few direct games  (organizer format change, parser regression)
    - Too few extras        (expansion not following WIN GM # or finish slots)
    - Empty tree            (seed-number root not handled)
    - No Sunday nodes       (Sunday games missing from sheet or dedup-dropped)
    - Unresolved slots      (describe_slot returning raw "1stD-" strings)
    - Stale opponent names  (cross-weekend contamination in standings lookup)
    """
    path = os.path.join(FIXTURES_DIR, WPL_FILE)
    failures = 0
    label = f"{team!r} | {sheet} championship {anchor}"
    print(f"\n{label}")

    if not os.path.exists(path):
        print(f"  [SKIP] file not found: {path}")
        return 0

    games     = load_and_parse(path)
    div_games = [g for g in games if g["sheet"] == sheet]

    # Direct games this weekend only
    direct = sorted(
        [g for g in div_games
         if (team_matches(g["white_team"], team) or team_matches(g["dark_team"], team))
         and g.get("date") and abs((g["date"] - anchor).days) <= 1],
        key=lambda g: (g.get("date") or date.min, g.get("time") or datetime.min.time()),
    )

    ok = _check(f"direct championship games in [{min_direct}, {max_direct}]",
                min_direct <= len(direct) <= max_direct, f"got {len(direct)}")
    if not ok: failures += 1

    # Expansion
    extras = _expand_bracket_games(team, direct, div_games)
    ok = _check(f"extras ≥ {min_extras} (bracket expansion reaches Sunday)",
                len(extras) >= min_extras, f"got {len(extras)}")
    if not ok: failures += 1

    # Tree builder
    tree = _build_wpl_game_tree(team, div_games, anchor_date=anchor)
    ok = _check(f"tree has ≥ {min_tree} nodes",
                len(tree) >= min_tree, f"got {len(tree)}")
    if not ok: failures += 1

    sun_nodes = [n for n in tree if n.get("date") and n["date"].weekday() == 6]
    ok = _check(f"tree has ≥ {min_sun} Sunday node(s)",
                len(sun_nodes) >= min_sun, f"got {len(sun_nodes)}")
    if not ok: failures += 1

    # Unresolved slot strings (describe_slot contamination check)
    bad_opps = []
    for n in tree:
        wt, dt = n["white_team"], n["dark_team"]
        opp_slot = dt if team.split()[-1].upper() in wt.upper() else wt
        opp = describe_slot(opp_slot, div_games, ref_date=anchor)
        if _SLOT_LIKE_RE.match(opp):
            bad_opps.append(f"{n['game_id']}: {opp!r}")
    ok = _check("no unresolved slot strings in opponent labels",
                not bad_opps, "; ".join(bad_opps[:3]))
    if not ok: failures += 1

    # Placeholder discipline: WIN-GM-inferred games must not be marked real
    # before the prelim has been played.
    inferred_real = []
    for n in tree:
        wt, dt = n["white_team"], n["dark_team"]
        is_inferred = (
            not team_matches(wt, team) and not team_matches(dt, team)
            and re.search(r'\bWIN\s+GM\s+#', wt + dt, re.IGNORECASE)
        )
        if is_inferred and not n.get("placeholder"):
            inferred_real.append(n["game_id"])
    ok = _check("no inferred WIN-GM games marked as confirmed (placeholder=False)",
                not inferred_real, str(inferred_real))
    if not ok: failures += 1

    return failures


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    total_failures = 0

    print("=" * 60)
    print("Parser fixture tests")
    print("=" * 60)
    for args in FIXTURES:
        total_failures += test_fixture(*args)

    print("\n" + "=" * 60)
    print("Team expansion tests (all-weekends contamination guard)")
    print("=" * 60)
    for args in TEAM_CHECKS:
        total_failures += test_team_expansion(*args)

    print("\n" + "=" * 60)
    print("Championship weekend tests")
    print("=" * 60)
    for args in CHAMPIONSHIP_CHECKS:
        total_failures += test_championship_team(*args)

    print("\n" + "=" * 60)
    if total_failures == 0:
        print(f"\033[32mAll tests passed.\033[0m")
    else:
        print(f"\033[31m{total_failures} test(s) FAILED.\033[0m")
    print("=" * 60)
    sys.exit(0 if total_failures == 0 else 1)


if __name__ == "__main__":
    main()

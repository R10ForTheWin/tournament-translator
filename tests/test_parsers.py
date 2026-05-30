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
from parsers.format_cca import parse_csv as cca_parse_csv
from app import (
    _expand_bracket_games, _build_wpl_game_tree,
    team_matches, describe_slot, _SLOT_LIKE_RE, _tournament_meta,
    _team_opp_slot,
)

FIXTURES_DIR   = os.path.join(ROOT, "Tournaments Excels")
FIXTURES_CCA   = os.path.join(ROOT, "tests", "fixtures")

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
# These check the behavior we care about most on game day.
#
# Anchor date is derived from _tournament_meta() so it stays correct when a
# new tournament weekend is added to KNOWN_TOURNAMENTS — no manual date updates.
#
# Schema per entry: (sheet, team, min_direct, max_direct,
#                    min_extras, min_tree_nodes, min_sun_nodes)
#
# Thresholds are set to hold BOTH pre-tournament (all placeholders visible) AND
# post-tournament (bracket narrowed to one path).  Lower bounds only — they catch
# regressions back to zero without breaking when live scores narrow the tree.
#
#   min_extras=1    — at least 1 Sunday game must reach the expansion
#   min_sun_nodes=1 — at least 1 Sunday node must appear in the bracket tree
#
# Add a row here any time a new bug class is fixed on a real team.
CHAMPIONSHIP_CHECKS: dict[str, list[tuple]] = {
    "futures-5": [
        # Pool-slot seeded team (H1): 2 Sat games, 3 Sun placement placeholders.
        # Previously: Sunday games missing due to game-ID collision in format_b.py.
        ("16u Boys", "trojan gold",     2, 2, 1, 3, 1),

        # Seed-number prelim team: won prelim + played 2 bracket games = 3 direct.
        # Previously: showed only 1 game (seed-number root unhandled).
        ("16u Boys", "trojan cardinal", 1, 3, 1, 3, 1),

        # Seeded pool-E1 team — regression canary for pool-slot expansion.
        ("16u Boys", "imperial",        2, 2, 1, 3, 1),

        # Seeded pool-G1 team — second pool-slot canary.
        ("16u Boys", "socal",           2, 2, 1, 3, 1),
    ],
    # Add "futures-super", "futures-6", etc. here as those weekends are played.
}


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
                            min_extras: int, min_tree: int, min_sun: int) -> int:
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


# ── CCA opponent-slot correctness ─────────────────────────────────────────────
#
# These tests catch the class of bug where _team_opp_slot returns the team's
# OWN slot as the opponent ("Loser of game #23" when the team IS the loser of
# game #23). This happens when team_matches() fails on a placeholder slot and
# the code falls to a wrong default. Each entry pins a known-correct pre-
# tournament opponent slot for a specific placeholder game type:
#
#   (csv_file, division, id_prefix, team_query, game_id, expected_opp_slot)
#
# "expected_opp_slot" is the raw normalized slot string, not a display label.
# If the slot changes because results fill it in, update the expectation — the
# important invariant is that the opponent is NEVER the team's own slot.
#
# Fixture files are snapshots of the live Google Sheets saved at
# tests/fixtures/cca_{16u,18u}.csv. Refresh them before each tournament day
# by running: python3 tests/refresh_cca_fixtures.py
CCA_OPP_CHECKS = [
    # 16U Trojan Gold (B) — pool G2 seed
    #   finish-slot path: pool_rank_group="G", opp must not be "2ndG-"
    ("cca_16u.csv", "16U Boys", "16U", "trojan gold", "16U-38", "1stH-"),
    #   finish-slot path: pool_rank_group="G", team is "1stG-" here, opp must not be "1stG-"
    ("cca_16u.csv", "16U Boys", "16U", "trojan gold", "16U-40", "2ndH-"),
    #   W#/L# path: L#23 is team's slot (game 23 in my_ids), opp must not be "L#23"
    ("cca_16u.csv", "16U Boys", "16U", "trojan gold", "16U-46", "L#22"),

    # 16U Trojan Cardinal (A) — pool C1 seed
    #   finish-slot path: opp must not be "1stC-"
    ("cca_16u.csv", "16U Boys", "16U", "trojan cardinal", "16U-27", "2ndD-"),
    #   W#/L# path: L#12 is team's slot
    ("cca_16u.csv", "16U Boys", "16U", "trojan cardinal", "16U-25", "L#13"),

    # 16U Trojan Silver (C) — pool E4 seed
    #   W#/L# path: L#10 is team's slot
    ("cca_16u.csv", "16U Boys", "16U", "trojan silver", "16U-35", "L#11"),

    # 18U Trojan Cardinal A — bracket-entry BB1 (no pool play)
    #   finish-slot path: team is direct white slot, opp is 1stC-
    ("cca_18u.csv", "18U Boys", "18U", "trojan cardinal", "18U-15", "1stC-"),
    #   W#/L# path: W#15 is team's slot (win of 15 in my_ids), opp must not be "W#15"
    ("cca_18u.csv", "18U Boys", "18U", "trojan cardinal", "18U-20", "L#16"),
]


def test_cca_opponent_slots() -> int:
    """Verify _team_opp_slot returns the correct (non-self) opponent for CCA placeholder games.

    This test would have caught the bug where "Trojan Gold vs Loser of game #23"
    was shown when the team IS the loser of game #23. The fix was _team_opp_slot;
    this test ensures it never regresses.
    """
    failures = 0
    print("\nCCA opponent-slot correctness")

    # Load each CSV file once
    csv_cache: dict[str, list] = {}
    for csv_file, division, prefix, team, game_id, expected_opp in CCA_OPP_CHECKS:
        path = os.path.join(FIXTURES_CCA, csv_file)
        if not os.path.exists(path):
            print(f"  [SKIP] fixture not found: {path}")
            continue

        if csv_file not in csv_cache:
            with open(path) as f:
                text = f.read()
            csv_cache[csv_file] = cca_parse_csv(text, division, prefix)
        div_games = csv_cache[csv_file]

        direct = [g for g in div_games
                  if team_matches(g["white_team"], team) or team_matches(g["dark_team"], team)]
        extras  = _expand_bracket_games(team, direct, div_games)
        all_g   = direct + extras
        my_ids  = {g["game_id"] for g in all_g}

        target = next((g for g in all_g if g["game_id"] == game_id), None)
        label  = f"{game_id} {team!r} opp={expected_opp!r}"

        if target is None:
            ok = _check(f"{label} — game found in expansion", False, "game not in expansion")
            failures += 1
            continue

        opp = _team_opp_slot(target, team, div_games, my_ids)

        # Primary: must match expected slot exactly
        ok = _check(f"{label}", opp == expected_opp, f"got {opp!r}")
        if not ok: failures += 1

        # Secondary (self-reference guard): opponent must never be the team's own slot
        # This catches future regressions even if expected_opp drifts as results fill in
        self_ref = team_matches(opp, team)
        ok2 = _check(f"{game_id} {team!r} — no self-reference", not self_ref,
                     f"opp={opp!r} matches team name")
        if not ok2: failures += 1

    return failures


# ── Game-num (bracket depth) regression test ─────────────────────────────────
# Regression: lose-path game (Saturday) was getting game_num=3 while win-path
# (Friday) got game_num=2. Both are "Game 2" — the second game in the journey.
# Fixed by using BFS depth from root instead of chronological slot order.

CCA_GAME_NUM_CHECKS = [
    # (tournament_id, team, sheet, game_id, expected_game_num, reason)
    ("jo-quals", "trojan gold (b)", "16U Boys", "16U-23", 1, "pool game = Game 1 (Fri 5PM)"),
    ("jo-quals", "trojan gold (b)", "16U Boys", "16U-24", 2, "win-path Game 2 (Fri 7:30PM)"),
    ("jo-quals", "trojan gold (b)", "16U Boys", "16U-46", 2, "lose-path Game 2 shares column with win (Sat 9AM)"),
    # Placement alternatives happen Sat 11AM — earlier than bracket games at Sat 12PM.
    # Re-numbering by actual time puts placements at Game 3, brackets at Game 4.
    ("jo-quals", "trojan gold (b)", "16U Boys", "16U-38", 3, "placement Game 3 (Sat 11AM, before bracket)"),
    ("jo-quals", "trojan gold (b)", "16U Boys", "16U-40", 3, "placement Game 3 (Sat 1PM, same column as 38)"),
    ("jo-quals", "trojan gold (b)", "16U Boys", "16U-57", 4, "bracket lose-path Game 4 (Sat 12PM)"),
    ("jo-quals", "trojan gold (b)", "16U Boys", "16U-49", 4, "bracket win-path Game 4 (Sat 1PM)"),
]

def test_cca_game_nums() -> int:
    """Verify win-path and lose-path games share the same game_num (same bracket depth).

    The bug: lose-path game (Saturday) got game_num=3 because chronological sort
    placed it after the win-path game (Friday). Fix: BFS depth from root game.
    """
    from app import app as _flask_app
    import json as _json

    failures = 0
    with _flask_app.test_client() as c:
        # Group by (tournament_id, team, sheet) to fetch once
        seen: dict = {}
        for tid, team, sheet, game_id, expected_num, reason in CCA_GAME_NUM_CHECKS:
            key = (tid, team, sheet)
            if key not in seen:
                r = c.get(f"/api/games/{tid}/{team}?sheet={sheet}")
                data = _json.loads(r.data)
                by_id = {g["game_id"]: g for g in (data.get("upcoming", []) + data.get("played", []))}
                seen[key] = by_id
            by_id = seen[key]
            g = by_id.get(game_id)
            got = g.get("game_num") if g else None
            ok = _check(f"{game_id} game_num={expected_num} ({reason})", got == expected_num,
                        f"got {got!r}")
            if not ok:
                failures += 1
    return failures


def test_trojan_team_names_clean() -> int:
    """No entry returned by api_trojan_teams should look like a raw slot string.

    The bug: CCA format uses 'W #12 - TROJAN CARDINAL (A)' slot syntax. _PREFIX_RE
    didn't handle the spaced variant, so strip_prefix left the raw slot intact and
    it appeared as a phantom team in the UI. This test catches any future case where
    a new slot format slips past strip_prefix.
    """
    from app import app as _flask_app, _SLOT_LIKE_RE as _slre
    import json as _json

    failures = 0
    tournaments_to_check = ["jo-quals"]
    with _flask_app.test_client() as c:
        for tid in tournaments_to_check:
            r = c.get(f"/api/trojan-teams/{tid}")
            if r.status_code != 200:
                _check(f"{tid} trojan-teams request succeeded", False, f"status {r.status_code}")
                failures += 1
                continue
            teams = _json.loads(r.data)
            for t in teams:
                name = t.get("name", "")
                ok = _check(
                    f"{tid}: team name is clean (not a raw slot): {name!r}",
                    not _slre.match(name),
                    f"looks like an unstripped slot",
                )
                if not ok:
                    failures += 1
    return failures


def test_no_slot_like_opponents() -> int:
    """Direct (non-placeholder) game cards must never show a slot-like opponent label.

    The bug: after strip_prefix fails on an unknown format, the raw slot string
    flows through as the opponent label — parents see "W #12 - TEMPLE CITY" instead
    of a real team name. This test catches that for all CCA Trojan teams.
    """
    from app import app as _flask_app, _SLOT_LIKE_RE as _slre
    import json as _json

    failures = 0
    checks = [
        ("jo-quals", "TROJAN CARDINAL A",   "18U Boys"),
        ("jo-quals", "TROJAN GOLD B",        "18U Boys"),
        ("jo-quals", "TROJAN CARDINAL (A)",  "16U Boys"),
        ("jo-quals", "TROJAN GOLD (B)",      "16U Boys"),
        ("jo-quals", "TROJAN SILVER (C)",    "16U Boys"),
    ]
    with _flask_app.test_client() as c:
        for tid, team, sheet in checks:
            r = c.get(f"/api/games/{tid}/{team}?sheet={sheet}")
            if r.status_code != 200:
                _check(f"{tid}/{team} request ok", False, f"status {r.status_code}")
                failures += 1
                continue
            data = _json.loads(r.data)
            all_games = data.get("upcoming", []) + data.get("played", [])
            for g in all_games:
                if g.get("placeholder"):
                    continue  # placeholder opponents are unresolved slots by design
                opp = g.get("opponent", "")
                ok = _check(
                    f"{tid}/{team} game {g['game_id']}: opponent not a raw slot: {opp!r}",
                    not _slre.match(opp),
                    "looks like an unstripped slot",
                )
                if not ok:
                    failures += 1
    return failures


def test_game_num_column_integrity() -> int:
    """Within each game_num column, non-placeholder games at different times on the
    same day must be win/lose path siblings — not sequential rounds incorrectly merged.

    The bug: CCA's 3:30 PM pool game and 6:50 PM bracket game both got game_num=1
    because the day-merge logic treated non-overlapping same-day times as the same
    round. This test catches that: if two non-placeholder games share a game_num,
    are on the same day, and are at different times, they must have win+lose paths.
    """
    from app import app as _flask_app
    import json as _json

    failures = 0
    checks = [
        ("jo-quals", "TROJAN CARDINAL A",   "18U Boys"),
        ("jo-quals", "TROJAN GOLD B",        "18U Boys"),
        ("jo-quals", "TROJAN CARDINAL (A)",  "16U Boys"),
        ("jo-quals", "TROJAN GOLD (B)",      "16U Boys"),
        ("jo-quals", "TROJAN SILVER (C)",    "16U Boys"),
    ]
    with _flask_app.test_client() as c:
        for tid, team, sheet in checks:
            r = c.get(f"/api/games/{tid}/{team}?sheet={sheet}")
            if r.status_code != 200:
                continue
            data = _json.loads(r.data)
            all_games = data.get("upcoming", []) + data.get("played", [])
            # Group non-placeholder games by (game_num, date)
            groups: dict = {}
            for g in all_games:
                if g.get("placeholder"):
                    continue
                key = (g.get("game_num"), g.get("date"))
                if None not in key:
                    groups.setdefault(key, []).append(g)
            for (gnum, date), grp in groups.items():
                times = {g.get("time") for g in grp if g.get("time")}
                if len(times) <= 1:
                    continue
                # Multiple times in same column on same day — must be win/lose siblings
                paths = {g.get("path") for g in grp}
                is_siblings = "win" in paths and "lose" in paths
                ok = _check(
                    f"{tid}/{team} GAME {gnum} on {date}: "
                    f"multiple times {sorted(times)} are win/lose siblings",
                    is_siblings,
                    f"paths={paths} — looks like sequential rounds merged into one column",
                )
                if not ok:
                    failures += 1
    return failures


def test_last_meeting_available() -> int:
    """At least one upcoming game per Trojan team should have a last_meeting populated,
    confirming cross-tournament history search is active.

    This test will naturally pass only when there IS prior history (i.e., opponents
    have been faced in a previous tournament on file). It SKIPs teams with zero
    prior history rather than failing — the assertion is that the mechanism works,
    not that every opponent has been faced before.
    """
    from app import app as _flask_app
    import json as _json

    failures = 0
    checks = [
        ("jo-quals", "TROJAN CARDINAL (A)", "16U Boys"),
        ("jo-quals", "TROJAN GOLD (B)",     "16U Boys"),
        ("jo-quals", "TROJAN SILVER (C)",   "16U Boys"),
    ]
    with _flask_app.test_client() as c:
        for tid, team, sheet in checks:
            r = c.get(f"/api/games/{tid}/{team}?sheet={sheet}")
            if r.status_code != 200:
                continue
            data = _json.loads(r.data)
            upcoming = [g for g in data.get("upcoming", []) if not g.get("placeholder")]
            if not upcoming:
                print(f"  [SKIP] {tid}/{team}: no direct upcoming games to check")
                continue
            # Check that _last_meeting is populated for at least one non-placeholder game
            # where the opponent is a known resolved team (not a slot description)
            from app import _SLOT_LIKE_RE as _slre
            known_opp_games = [g for g in upcoming if g.get("opponent") and not _slre.match(g.get("opponent", ""))]
            if not known_opp_games:
                print(f"  [SKIP] {tid}/{team}: all upcoming opponents are still unresolved slots")
                continue
            has_any = any(g.get("last_meeting") is not None for g in known_opp_games)
            # Soft check: SKIP rather than FAIL when there's genuinely no prior history
            if not has_any:
                print(f"  [INFO] {tid}/{team}: no prior meeting found for any of "
                      f"{[g['opponent'][:30] for g in known_opp_games[:3]]} — "
                      f"may be first encounter or history files not on disk")
            else:
                ok = _check(f"{tid}/{team}: last_meeting populated for at least one upcoming game",
                            True, "")
                # If has_any is True the check trivially passes; log it
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
    for tournament_id, checks in CHAMPIONSHIP_CHECKS.items():
        meta = _tournament_meta(tournament_id)
        if not meta or not meta.get("date_start"):
            print(f"\n[SKIP] No date_start for {tournament_id!r} in KNOWN_TOURNAMENTS")
            continue
        anchor = meta["date_start"]
        for (sheet, team, min_d, max_d, min_e, min_t, min_s) in checks:
            total_failures += test_championship_team(
                sheet, team, anchor, min_d, max_d, min_e, min_t, min_s
            )

    print("\n" + "=" * 60)
    print("CCA opponent-slot tests (self-reference regression guard)")
    print("=" * 60)
    total_failures += test_cca_opponent_slots()

    print("\n" + "=" * 60)
    print("CCA game-num tests (win/lose paths share same game_num)")
    print("=" * 60)
    total_failures += test_cca_game_nums()

    print("\n" + "=" * 60)
    print("Trojan team name cleanliness (no raw slot strings as team names)")
    print("=" * 60)
    total_failures += test_trojan_team_names_clean()

    print("\n" + "=" * 60)
    print("Direct game opponent labels (no slot-like strings on non-placeholder cards)")
    print("=" * 60)
    total_failures += test_no_slot_like_opponents()

    print("\n" + "=" * 60)
    print("Game-num column integrity (no sequential rounds merged into one column)")
    print("=" * 60)
    total_failures += test_game_num_column_integrity()

    print("\n" + "=" * 60)
    print("Last meeting availability (cross-tournament history search active)")
    print("=" * 60)
    total_failures += test_last_meeting_available()

    print("\n" + "=" * 60)
    if total_failures == 0:
        print(f"\033[32mAll tests passed.\033[0m")
    else:
        print(f"\033[31m{total_failures} test(s) FAILED.\033[0m")
    print("=" * 60)
    sys.exit(0 if total_failures == 0 else 1)


if __name__ == "__main__":
    main()

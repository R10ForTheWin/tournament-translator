#!/usr/bin/env python3
"""
Round-by-round replay of REAL 2025 Junior Olympics results against the live
app API -- not just a static "does it parse" check, but a full progression
through the actual tournament, checkpoint by checkpoint, verifying each
round's results, opponent resolution, and self-consistency as the app would
have shown them live, then a final check that the app's own inferred
placement matches the real recorded final rank.

Why real data instead of synthetic scenarios: see feedback memory
"round-by-round-simulation" -- simulating win/loss progression through the
real API (not just checking a fully-played or fully-unplayed snapshot)
already caught 3 real bugs during Quiksilver Cup prep, including one that
silently inverted a win into a loss. Using real historical results instead
of hand-authored ones means the exact real slot syntax, comment
conventions, and edge cases the organizer's spreadsheet actually contains
get exercised, not just what a synthetic scenario author thought to cover.

Source fixture: Tournaments Excels/2025_NJO_Public_Sched_S2.xlsx (real 2025
Junior Olympics data). Most boys (M) sheets in this export are schedule-only
(zero played games) -- only "16U_M_CHAMP-41 teams" has real results among
boys sheets. To get real-result coverage of the pool-letter bracket format
(matching this year's Trojan Gold 18U_M_Invite 24 / Trojan 12U_M_Classic_53
division), this script also replays a fully-played girls/coed sheet using
the identical slot syntax -- the parser and downstream code do not
distinguish by gender, only by slot format, so this still exercises the
same code paths this year's real Trojan games will hit.

Injection strategy: parse the real fixture once via the real parser (ground
truth), then monkeypatch app.find_excel / app.load_and_parse so each
checkpoint serves a masked copy of that SAME parsed data -- games scheduled
after the checkpoint get played=False/scores=None, exactly recreating what
the sheet would have looked like live at that point. Everything downstream
of that point (expansion, tree building, confidence validation, describe_slot,
predictor) runs through the real, unmodified production code via the real
Flask API route -- only the raw-Excel-read step is bypassed, and that layer
already has its own dedicated fixture-based tests in test_parsers.py.

Usage:
    python3 tests/round_by_round_njo2025.py
"""
from __future__ import annotations
import os
import re
import sys
from collections import Counter
from datetime import date, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

FIXTURE = os.path.join(ROOT, "Tournaments Excels", "2025_NJO_Public_Sched_S2.xlsx")
FAKE_TID = "test-njo2025-replay"

FAILURES = []


def check(label: str, cond: bool, detail: str = "") -> bool:
    status = "PASS" if cond else "FAIL"
    suffix = f"  ({detail})" if detail else ""
    print(f"  [{status}] {label}{suffix}")
    if not cond:
        FAILURES.append(label)
    return cond


def pick_our_team(sheet_games: list) -> str | None:
    """Pick the real team name (not a raw slot reference) with the most
    game appearances in this sheet -- gives the longest real trajectory to
    replay round by round."""
    import app as app_module
    counts: Counter = Counter()
    for g in sheet_games:
        for slot in (g["white_team"], g["dark_team"]):
            name = app_module.strip_prefix(slot).strip()
            if name and not app_module._SLOT_LIKE_RE.match(name):
                counts[name] += 1
    if not counts:
        return None
    return counts.most_common(1)[0][0]


def checkpoints(sheet_games: list) -> list:
    """Unique (date, time) pairs in chronological order -- one checkpoint
    per real scheduled round."""
    pairs = sorted({(g.get("date"), g.get("time")) for g in sheet_games
                     if g.get("date") is not None}, key=lambda p: (p[0], p[1] or datetime.min.time()))
    return pairs


def mask_after(sheet_games: list, cutoff) -> list:
    """Copy of sheet_games where any game scheduled strictly after cutoff
    (date, time) is reset to unplayed -- recreating the live state of the
    sheet as of that checkpoint."""
    cd, ct = cutoff
    out = []
    for g in sheet_games:
        gd, gt = g.get("date"), g.get("time") or datetime.min.time()
        after = gd is not None and (gd, gt) > (cd, ct or datetime.min.time())
        if after:
            g2 = dict(g)
            g2["played"] = False
            g2["white_score"] = None
            g2["dark_score"] = None
            out.append(g2)
        else:
            out.append(g)
    return out


def ground_truth_result(g: dict, team: str) -> str | None:
    import app as app_module
    if not g.get("played") or g.get("white_score") is None or g.get("dark_score") is None:
        return None
    return app_module._result_str(g, team)


def replay_sheet(sheet_name: str, label: str) -> int:
    print(f"\n{'=' * 60}\n{label}: {sheet_name}\n{'=' * 60}")
    import app as app_module
    from parsers.detect import load_and_parse as real_load_and_parse

    all_games = real_load_and_parse(FIXTURE)
    sheet_games = [g for g in all_games if g["sheet"] == sheet_name]
    if not sheet_games:
        check(f"{sheet_name}: found games in fixture", False, "0 games")
        return 1

    team = pick_our_team(sheet_games)
    if not team:
        check(f"{sheet_name}: found a real team to replay", False)
        return 1
    print(f"  replaying as: {team!r} ({len(sheet_games)} games in sheet)")

    cps = checkpoints(sheet_games)
    check(f"{sheet_name}: has scheduled rounds to replay", len(cps) > 0, f"{len(cps)} checkpoints")
    if not cps:
        return 1

    # Ground truth: real result for every one of THIS team's played games,
    # keyed by game_id, computed once against the fully-played real data.
    truth: dict = {}
    for g in sheet_games:
        if app_module.team_matches(g["white_team"], team) or app_module.team_matches(g["dark_team"], team):
            r = ground_truth_result(g, team)
            if r is not None:
                truth[g["game_id"]] = r

    failures = 0
    _orig_find_excel = app_module.find_excel
    _orig_load_and_parse = app_module.load_and_parse
    _orig_njo = set(app_module._NJO_TOURNAMENTS)
    app_module._NJO_TOURNAMENTS.add(FAKE_TID)

    seen_wrong_result = False
    seen_self_ref = False
    seen_leaked_slot = False
    seen_crash = False

    try:
        with app_module.app.test_client() as c:
            for i, cp in enumerate(cps):
                masked = {**{s: [g for g in all_games if g["sheet"] == s] for s in {g["sheet"] for g in all_games} if s != sheet_name}}
                masked[sheet_name] = mask_after(sheet_games, cp)
                flat = [g for gs in masked.values() for g in gs]

                app_module.find_excel = lambda tid, _flat=flat: "dummy"
                app_module.load_and_parse = lambda excel, _flat=flat: list(_flat)

                r = c.get(f"/api/games/{FAKE_TID}/{team}?sheet={sheet_name}")
                if r.status_code != 200:
                    check(f"round {i+1}/{len(cps)} ({cp[0]}): request succeeds", False,
                          f"status {r.status_code}")
                    failures += 1
                    seen_crash = True
                    continue

                data = r.get_json()
                played = data.get("played") or []

                for pg in played:
                    gid = pg.get("game_id")
                    if gid not in truth:
                        continue
                    got = pg.get("result")
                    if got != truth[gid]:
                        check(f"round {i+1} game {gid}: result matches real outcome", False,
                              f"app said {got!r}, real result was {truth[gid]!r}")
                        failures += 1
                        seen_wrong_result = True

                for pg in played + (data.get("upcoming") or []):
                    opp = pg.get("opponent") or ""
                    if opp and app_module.team_matches(opp, team):
                        check(f"round {i+1} game {pg.get('game_id')}: opponent is not self", False,
                              f"opponent={opp!r}")
                        failures += 1
                        seen_self_ref = True
                    if (opp and not pg.get("placeholder")
                            and app_module._SLOT_LIKE_RE.match(opp)
                            and not app_module._POOL_PREVIEW_RE.match(opp)):
                        check(f"round {i+1} game {pg.get('game_id')}: opponent not a raw slot leak", False,
                              f"opponent={opp!r}")
                        failures += 1
                        seen_leaked_slot = True

                conf = data.get("bracket_confidence")
                if conf == "red" and i == len(cps) - 1:
                    check(f"final round: bracket confidence is not red", False,
                          f"warnings={data.get('bracket_warnings')}")
                    failures += 1

            # Final check: computed placement vs. real recorded final rank.
            final_r = c.get(f"/api/games/{FAKE_TID}/{team}?sheet={sheet_name}")
            final_data = final_r.get_json()
            app_placement = final_data.get("placement")

            real_played = [dict(g, result=ground_truth_result(g, team)) for g in sheet_games
                            if (app_module.team_matches(g["white_team"], team)
                                or app_module.team_matches(g["dark_team"], team))
                            and g.get("played") and g.get("white_score") is not None]
            real_played.sort(key=lambda g: (g.get("date") or date.min, g.get("time") or datetime.min.time()))
            real_placement = app_module._infer_placement(real_played) if real_played else None

            if real_placement:
                ok = check(f"{sheet_name}: final placement matches real recorded rank",
                           app_placement == real_placement,
                           f"app said {app_placement!r}, real recorded rank was {real_placement!r}")
                if not ok:
                    failures += 1
            else:
                print(f"  [SKIP] {sheet_name}: no ordinal placement comment found in real data to check against")
    finally:
        app_module.find_excel = _orig_find_excel
        app_module.load_and_parse = _orig_load_and_parse
        app_module._NJO_TOURNAMENTS.clear()
        app_module._NJO_TOURNAMENTS.update(_orig_njo)

    print(f"  -- summary: wrong_result={seen_wrong_result} self_ref={seen_self_ref} "
          f"leaked_slot={seen_leaked_slot} crash={seen_crash}")
    return failures


if __name__ == "__main__":
    if not os.path.exists(FIXTURE):
        print(f"Fixture not found: {FIXTURE}")
        sys.exit(2)

    # Champ tier: seeded single-elimination bracket, matches this year's
    # Trojan Cardinal 16U/18U Champ divisions. Only boys (M) sheet in this
    # fixture with real results.
    replay_sheet("16U_M_CHAMP-41 teams", "Champ / seeded elimination bracket")

    # Classic/Invite tier: pool-letter pool+bracket hybrid, matches this
    # year's Trojan Gold 18U_M_Invite 24 / Trojan 12U_M_Classic_53. No boys
    # sheet in this fixture has real results in this format -- using a
    # fully-played sheet with the identical slot syntax instead (parser and
    # downstream logic are format-driven, not gender-driven).
    replay_sheet("10U_C_Classic 24", "Classic/Invite / pool-letter bracket")

    # Broader sweep, 2026-07-19: additional fully-played real sheets across
    # other tiers/genders to check for the same bug class (or a new one)
    # elsewhere in the fixture -- not needed for 2026 Trojan-format coverage
    # specifically, but the tied-game false-attribution bug found via the two
    # sheets above was previously invisible to every other check in this
    # codebase, so broader real-result coverage is cheap insurance before
    # calling the app tournament-ready.
    for sheet, label in [
        ("16U_F_Champ",       "extra sweep: Champ tier, girls"),
        ("18U_F_Classic 47",  "extra sweep: Classic tier, mixed seed+pool slots"),
        ("10U_C_Champ 50",    "extra sweep: Champ tier, coed"),
        ("14U_F_Classic_35",  "extra sweep: Classic tier, girls"),
        ("12U_F_Classic 14",  "extra sweep: Classic tier, girls, smaller division"),
    ]:
        replay_sheet(sheet, label)

    # FAILURES (the check() global) is the single source of truth for pass/fail
    # -- not each replay_sheet()'s own returned count, which drifted out of
    # sync with it once before (see the final-placement check history in git
    # blame) and silently under-reported.
    print(f"\n{'=' * 60}")
    if FAILURES:
        print(f"{len(FAILURES)} check(s) FAILED")
        sys.exit(1)
    print("All checks passed")

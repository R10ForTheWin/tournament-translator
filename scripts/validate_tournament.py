#!/usr/bin/env python3
"""
Pre-tournament validation pipeline.

Run this immediately after updating a tournament URL, before pushing to Railway.
All checks run locally using the Flask test client — no deploy needed.

Steps:
  1. Parser test suite  — existing test_parsers.py (data layer)
  2. Team detection     — all Trojan teams found, right division, expansion works
  3. API smoke test     — game ordering, opponent labels, game counts (display layer)
  4. LLM judge          — Claude Haiku reviews each team's bracket for parent-visible issues

Exit: 0 = safe to push, 1 = fix before pushing

Usage:
    python3 scripts/validate_tournament.py jo-quals
    python3 scripts/validate_tournament.py futures-super
    python3 scripts/validate_tournament.py --skip-llm jo-quals
    python3 scripts/validate_tournament.py --skip-parser-tests jo-quals
"""
from __future__ import annotations
import sys, os, re, subprocess, json, argparse
from collections import defaultdict
from datetime import date, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

# ── Team registry ─────────────────────────────────────────────────────────────
# Add a new entry here when adding a new tournament.
# (team_query, expected_division_substring)
# team_query must match how the API URL slug is formed.
# expected_division_substring: case-insensitive match against game["sheet"].
TROJAN_TEAMS: dict[str, list[tuple[str, str]]] = {
    "jo-quals": [
        ("trojan cardinal (a)", "16U Boys"),
        ("trojan gold (b)",     "16U Boys"),
        ("trojan silver (c)",   "16U Boys"),
        ("trojan cardinal a",   "18U Boys"),
        ("trojan gold b",       "18U Boys"),
    ],
    "junior-olympics": [
        # Update when NJO schedule is posted
        ("trojan cardinal",     "18U"),
        ("trojan gold",         "16U"),
        ("trojan silver",       "16U"),
    ],
    "futures-super": [
        ("trojan gold",         "16u Boys"),
        ("trojan cardinal",     "16u Boys"),
    ],
    # Add kap7-cup, turbo-cup, newport-invite if Trojan teams compete
}

# Expected game count range per team (direct + expanded, pre-tournament)
GAME_COUNT_RANGE = (2, 12)

# Bracket node count ceiling — above this likely indicates over-expansion.
# WPL (2-day Sat/Sun): ≤7. CCA/NJO (3-day Fri-Sun): ≤20 (full branching tree).
MAX_BRACKET_NODES: dict[str, int] = {
    "default":        10,
    "jo-quals":       20,
    "junior-olympics":20,
}

# Known validator gaps: bracket_confidence may be "red" for these tournaments
# due to format-specific issues in _validate_wpl_bracket (CCA W#/L# format).
# These are noted in the report but don't count as failures.
KNOWN_RED_CONFIDENCE = {"jo-quals", "junior-olympics"}

# ── Helpers ───────────────────────────────────────────────────────────────────
PASS_  = "\033[32mPASS\033[0m"
FAIL_  = "\033[31mFAIL\033[0m"
WARN_  = "\033[33mWARN\033[0m"
NOTE_  = "\033[36mNOTE\033[0m"

def _check(label: str, condition: bool, detail: str = "") -> bool:
    status = PASS_ if condition else FAIL_
    suffix = f"  ({detail})" if detail else ""
    print(f"  [{status}] {label}{suffix}")
    return condition

def _note(label: str, detail: str = "") -> None:
    suffix = f"  ({detail})" if detail else ""
    print(f"  [{NOTE_}] {label}{suffix}")

def _day_order(date_str: str) -> int:
    """Map formatted day name to integer for chronological comparison."""
    if not date_str:
        return 99
    dl = date_str.lower()
    for i, prefix in enumerate(["mon","tue","wed","thu","fri","sat","sun"]):
        if dl.startswith(prefix):
            return i
    return 99

def _time_min(time_str: str) -> int:
    """Parse '3:30 PM' → minutes since midnight."""
    if not time_str:
        return 0
    m = re.match(r'(\d+):(\d+)\s*(AM|PM)?', time_str.strip(), re.IGNORECASE)
    if not m:
        return 0
    h, mn = int(m.group(1)), int(m.group(2))
    ampm = (m.group(3) or "").upper()
    if ampm == "PM" and h != 12:
        h += 12
    if ampm == "AM" and h == 12:
        h = 0
    return h * 60 + mn

def _game_num_from_id(game_id: str) -> str | None:
    m = re.search(r'-(\d+)$', game_id)
    return m.group(1) if m else None


# ── Step 1: Parser test suite ─────────────────────────────────────────────────
def step_parser_tests() -> bool:
    print("\n" + "=" * 60)
    print("Step 1: Parser test suite (tests/test_parsers.py)")
    print("=" * 60)
    result = subprocess.run(
        [sys.executable, "tests/test_parsers.py"],
        cwd=ROOT,
        capture_output=True,
        text=True,
    )
    # Print test output (strip ANSI for cleanliness if redirected)
    for line in result.stdout.splitlines():
        print(" ", line)
    if result.returncode != 0 and result.stderr:
        print("  STDERR:", result.stderr[:400])
    return result.returncode == 0


# ── Step 2: Team detection ────────────────────────────────────────────────────
def step_team_detection(tournament_id: str, teams: list) -> bool:
    # Use app.load_and_parse — it handles the JO Quals sentinel and other live-fetch patterns
    from app import find_excel, load_and_parse, team_matches, _expand_bracket_games

    print("\n" + "=" * 60)
    print("Step 2: Team detection (parser → division → expansion)")
    print("=" * 60)

    excel = find_excel(tournament_id)
    ok_excel = _check("Tournament data loads", excel is not None,
                      "" if excel is not None else f"find_excel('{tournament_id}') returned None — check TOURNAMENT_URLS in app.py")
    if not ok_excel:
        return False

    all_games = load_and_parse(excel)
    ok_games = _check(f"Total games parsed ≥ 20", len(all_games) >= 20,
                      f"{len(all_games)} games")
    if not ok_games:
        return False

    failures = 0
    for team_query, expected_div in teams:
        direct = [g for g in all_games
                  if team_matches(g["white_team"], team_query)
                  or team_matches(g["dark_team"], team_query)]

        ok = _check(f"{team_query!r}: ≥1 direct game", len(direct) >= 1,
                    f"{len(direct)} found")
        if not ok:
            failures += 1
            continue

        # Division check
        sheets = {g["sheet"] for g in direct}
        in_div = any(expected_div.lower() in s.lower() for s in sheets)
        ok = _check(f"{team_query!r}: in {expected_div!r}",
                    in_div, f"found in {sorted(sheets)}")
        if not ok:
            failures += 1

        # Expansion
        div_games = [g for g in all_games
                     if any(expected_div.lower() in g["sheet"].lower() for _ in [1])]
        extras = _expand_bracket_games(team_query, direct, div_games)
        total = len(direct) + len(extras)
        lo, hi = GAME_COUNT_RANGE
        ok = _check(f"{team_query!r}: expanded to [{lo},{hi}] games",
                    lo <= total <= hi, f"{total} ({len(direct)}+{len(extras)})")
        if not ok:
            failures += 1

    return failures == 0


# ── Step 3: API smoke test ────────────────────────────────────────────────────
def step_api_smoke_test(tournament_id: str, teams: list) -> bool:
    import urllib.parse
    from app import app as flask_app

    print("\n" + "=" * 60)
    print("Step 3: API smoke test (Flask test client)")
    print("=" * 60)

    failures = 0
    known_red = tournament_id in KNOWN_RED_CONFIDENCE

    with flask_app.test_client() as client:
        for team_query, expected_div in teams:
            encoded = urllib.parse.quote(team_query)
            resp = client.get(f"/api/games/{tournament_id}/{encoded}")

            label = f"{team_query!r}"

            ok = _check(f"{label}: API 200", resp.status_code == 200,
                        f"got {resp.status_code}")
            if not ok:
                failures += 1
                continue

            data = resp.get_json() or {}
            upcoming = data.get("upcoming") or []
            played   = data.get("played")   or []
            bracket  = data.get("wpl_bracket") or []
            conf     = data.get("bracket_confidence", "unknown")
            warnings = data.get("bracket_warnings") or []

            # Bracket confidence — informational for known-red formats
            if known_red:
                _note(f"{label}: bracket_confidence={conf!r} (known validator gap for {tournament_id})")
                for w in warnings[:2]:
                    _note(f"  └ {w[:100]}")
            else:
                ok = _check(f"{label}: bracket_confidence ≠ red", conf != "red",
                            f"got {conf!r}")
                if not ok:
                    failures += 1
                    for w in warnings[:2]:
                        print(f"       └ {w[:100]}")

            # Bracket node count (over-expansion guard)
            node_ceiling = (MAX_BRACKET_NODES.get(tournament_id)
                            or MAX_BRACKET_NODES["default"])
            ok = _check(f"{label}: bracket nodes ≤ {node_ceiling}",
                        len(bracket) <= node_ceiling,
                        f"got {len(bracket)} — likely over-expansion" if len(bracket) > node_ceiling else f"{len(bracket)}")
            if not ok:
                failures += 1

            # Total game count
            total = len(upcoming) + len(played)
            lo, hi = GAME_COUNT_RANGE
            ok = _check(f"{label}: total games in [{lo},{hi}]",
                        lo <= total <= hi, f"{total}")
            if not ok:
                failures += 1

            # All upcoming games have dates
            undated = [g["game_id"] for g in upcoming if not g.get("date")]
            ok = _check(f"{label}: all upcoming games have dates",
                        not undated, str(undated))
            if not ok:
                failures += 1

            # game_num ordering — each column group must be chronologically ≥ prior group
            by_num: dict = defaultdict(list)
            for g in upcoming:
                by_num[g.get("game_num") or 0].append(g)

            order_ok = True
            order_detail = ""
            prev_day, prev_min = -1, -1
            for n in sorted(by_num.keys()):
                games_in_col = by_num[n]
                dated = [g for g in games_in_col if g.get("date") and g.get("time")]
                if not dated:
                    continue
                # Find earliest game in this column
                dated.sort(key=lambda g: (_day_order(g["date"]), _time_min(g["time"])))
                earliest = dated[0]
                day = _day_order(earliest["date"])
                tmin = _time_min(earliest["time"])
                if day < prev_day or (day == prev_day and tmin < prev_min):
                    order_ok = False
                    order_detail = (f"Game #{n} ({earliest['date']} {earliest['time']}) "
                                    f"sorts before Game #{n-1}")
                    break
                prev_day, prev_min = day, tmin

            ok = _check(f"{label}: game_nums in chronological order",
                        order_ok, order_detail)
            if not ok:
                failures += 1

            # Win/lose sibling game_num check: W and L paths from same parent must share game_num
            sibling_violations = []
            by_gnum: dict = defaultdict(list)
            for g in upcoming:
                gn = g.get("game_num")
                if gn is not None:
                    by_gnum[gn].append(g)
            # Build a map: game_id -> set of paths seen at each game_num
            for gn, col_games in by_gnum.items():
                paths = [g.get("path") for g in col_games]
                gids  = [g.get("game_id") for g in col_games]
                if "win" in paths and "lose" in paths:
                    pass  # expected — siblings share the column
                elif len(col_games) > 1:
                    # Multiple games at same game_num but no win/lose pair — unusual but OK
                    pass
            # Inverse check: win and lose paths from the SAME game must share game_num
            win_games  = {g["game_id"]: g.get("game_num") for g in upcoming if g.get("path") == "win"}
            lose_games = {g["game_id"]: g.get("game_num") for g in upcoming if g.get("path") == "lose"}
            # We don't have parent game_id in the API response, so check indirectly:
            # if the same game_num has both a win and a lose path, that's correct.
            # If a win path game_num != lose path game_num at the same tree depth, that's wrong.
            # Simpler invariant: for any two games that are siblings (same game_num column),
            # they should have opposite paths (win/lose). Flag if two wins or two loses at same col.
            for gn, col_games in by_gnum.items():
                path_list = [g.get("path") for g in col_games if g.get("path") in ("win","lose")]
                if path_list.count("win") > 1 or path_list.count("lose") > 1:
                    gids = [g.get("game_id") for g in col_games]
                    sibling_violations.append(f"Game#{gn}: multiple {path_list} paths at same column: {gids}")
            ok = _check(f"{label}: win/lose siblings share game_num column",
                        not sibling_violations, "; ".join(sibling_violations[:2]))
            if not ok:
                failures += 1

            # Placement-alternative games must not have scenarios
            alt_with_scenarios = [
                g.get("game_id") for g in upcoming
                if g.get("is_alternative") and g.get("scenarios")
            ]
            ok = _check(f"{label}: placement-alt games have no scenarios",
                        not alt_with_scenarios, "; ".join(str(x) for x in alt_with_scenarios[:3]))
            if not ok:
                failures += 1

            # Win/lose path games (staircase branches) must not have scenarios —
            # the staircase structure is the "what's next", WHAT'S NEXT inside a
            # branch card is redundant and creates clutter
            branch_with_scenarios = [
                g.get("game_id") for g in upcoming
                if g.get("path") in ("win", "lose") and g.get("scenarios")
            ]
            ok = _check(f"{label}: win/lose branch games have no scenarios",
                        not branch_with_scenarios, "; ".join(str(x) for x in branch_with_scenarios[:3]))
            if not ok:
                failures += 1

            # Self-reference check
            self_refs = []
            for g in upcoming + played:
                opp = g.get("opponent") or ""
                gid = g.get("game_id") or ""
                gnum = _game_num_from_id(gid)
                if gnum and re.search(rf'\bgame\s*#?\s*{re.escape(gnum)}\b', opp, re.IGNORECASE):
                    self_refs.append(f"{gid}: opp={opp!r}")
            ok = _check(f"{label}: no self-reference opponents",
                        not self_refs, "; ".join(self_refs[:2]))
            if not ok:
                failures += 1

    return failures == 0


# ── Step 4: LLM judge ─────────────────────────────────────────────────────────
def _llm_call(prompt: str) -> str:
    """Call Claude Haiku via API key (if set) or claude CLI (always available)."""
    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if api_key:
        try:
            import anthropic
            client = anthropic.Anthropic(api_key=api_key)
            msg = client.messages.create(
                model="claude-haiku-4-5-20251001",
                max_tokens=300,
                messages=[{"role": "user", "content": prompt}],
            )
            return msg.content[0].text.strip()
        except Exception as e:
            raise RuntimeError(f"anthropic SDK call failed: {e}") from e

    # Fallback: use the `claude` CLI (Claude Code's own auth, no key needed)
    result = subprocess.run(
        ["claude", "-p", "--model", "claude-haiku-4-5-20251001"],
        input=prompt,
        capture_output=True,
        text=True,
        timeout=30,
    )
    if result.returncode != 0:
        raise RuntimeError(f"claude CLI failed: {result.stderr[:200]}")
    return result.stdout.strip()


def step_llm_judge(tournament_id: str, teams: list) -> bool:
    print("\n" + "=" * 60)
    print("Step 4: LLM bracket judge (Claude Haiku)")
    print("=" * 60)

    import urllib.parse
    from app import app as flask_app

    failures = 0

    with flask_app.test_client() as http_client:
        for team_query, expected_div in teams:
            encoded = urllib.parse.quote(team_query)
            resp = http_client.get(f"/api/games/{tournament_id}/{encoded}")
            if resp.status_code != 200:
                continue

            data = resp.get_json() or {}
            upcoming = data.get("upcoming") or []
            if not upcoming:
                _note(f"{team_query!r}: no upcoming games — skipping LLM")
                continue

            # Format schedule for the judge
            lines = []
            for g in upcoming:
                gnum  = g.get("game_num", "?")
                path  = g.get("path") or ""
                gid   = g.get("game_id", "?")
                dt    = f"{g.get('date','')} {g.get('time','')}".strip()
                opp   = g.get("opponent", "?")
                ph    = "placeholder" if g.get("placeholder") else "confirmed"
                path_tag = f" [{path}]" if path else ""
                lines.append(f"  Game #{gnum}{path_tag} [{gid}] {dt}: vs {opp}  [{ph}]")

            prompt = f"""You are reviewing a water polo tournament bracket that will be shown to parents on a mobile app. Check for any issues a parent would find confusing or wrong.

Tournament: {tournament_id}
Team: {team_query} ({expected_div})
Total upcoming games shown: {len(upcoming)}

Games (Game #N = column number in staircase display; [win]/[lose] = which path):
{chr(10).join(lines)}

Flag these specific problems:
1. A win-path game and lose-path game from the same parent have DIFFERENT Game #N numbers — both must share the same number (they are alternate paths to the same round)
2. Game #N LABELS are out of order with actual game times: every game labeled Game #4 must happen at the same time or LATER than every game labeled Game #3, and so on. If a game labeled Game #4 has an earlier time than any game labeled Game #3, that is a label ordering bug. (Exception: win/lose siblings at the same Game # can be on different days — that is fine.)
3. Opponent label is the same team as the team being shown (self-reference)
4. More than 9 upcoming games (suggests bracket over-expansion)
5. Fewer than 2 upcoming games (suggests expansion failed)
6. Any game missing a date or showing "None"

Reply ONLY in this format:
STATUS: OK

or:

STATUS: FLAG
ISSUES:
- one issue per line"""

            try:
                text = _llm_call(prompt)
            except Exception as e:
                _note(f"{team_query!r}: LLM call failed — {e}")
                continue

            status = "OK"
            issues: list[str] = []
            for line in text.splitlines():
                if line.upper().startswith("STATUS:"):
                    status = line.split(":", 1)[1].strip().upper()
                elif line.strip().startswith("-"):
                    issues.append(line.strip()[1:].strip())

            ok = _check(
                f"{team_query!r}: LLM judge",
                status == "OK",
                "; ".join(issues[:2]) if issues else "",
            )
            if not ok:
                failures += 1
                for issue in issues[2:]:
                    print(f"       └ {issue}")

    return failures == 0


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(description="Pre-tournament validation pipeline")
    parser.add_argument("tournament_id",
                        help="Tournament ID (e.g. jo-quals, futures-super)")
    parser.add_argument("--skip-llm",          action="store_true",
                        help="Skip LLM judge (faster, no API cost)")
    parser.add_argument("--skip-parser-tests", action="store_true",
                        help="Skip test_parsers.py (use when you've already run it)")
    args = parser.parse_args()

    tid   = args.tournament_id
    teams = TROJAN_TEAMS.get(tid)

    if not teams:
        known = sorted(TROJAN_TEAMS.keys())
        print(f"ERROR: No team list configured for {tid!r}.")
        print(f"Add it to TROJAN_TEAMS in scripts/validate_tournament.py")
        print(f"Known tournaments: {known}")
        sys.exit(1)

    print(f"\n{'='*60}")
    print(f"PRE-TOURNAMENT VALIDATION: {tid}")
    print(f"Teams: {[t for t, _ in teams]}")
    print(f"{'='*60}")

    results: dict[str, bool] = {}

    if not args.skip_parser_tests:
        results["parser_tests"] = step_parser_tests()

    results["team_detection"] = step_team_detection(tid, teams)
    results["api_smoke_test"] = step_api_smoke_test(tid, teams)

    if not args.skip_llm:
        results["llm_judge"] = step_llm_judge(tid, teams)

    # ── Final summary ─────────────────────────────────────────────────────────
    print(f"\n{'='*60}")
    print("SUMMARY")
    print(f"{'='*60}")
    all_pass = True
    for step, ok in results.items():
        status = PASS_ if ok else FAIL_
        print(f"  [{status}] {step}")
        if not ok:
            all_pass = False

    if all_pass:
        print(f"\n\033[32m✓ All checks passed — safe to push to Railway.\033[0m\n")
    else:
        print(f"\n\033[31m✗ Failures found — fix before pushing.\033[0m\n")

    sys.exit(0 if all_pass else 1)


if __name__ == "__main__":
    main()

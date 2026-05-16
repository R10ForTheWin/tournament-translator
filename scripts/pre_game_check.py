#!/usr/bin/env python3
"""
Pre-game validation sweep for a tournament weekend.

Run the night before games to catch data issues before parents see them.
Every check here encodes a bug that was found the hard way on game day.

Usage (from project root):
    python3 scripts/pre_game_check.py futures-5
    python3 scripts/pre_game_check.py futures-5 --sheet "16u Boys"
    python3 scripts/pre_game_check.py futures-5 --download
    python3 scripts/pre_game_check.py futures-5 --download --sheet "16u Boys" --verbose

Flags:
    --download   Fetch a fresh copy from Google Sheets before checking.
                 Without this flag, uses the local cached file.
    --sheet X    Only check one sheet (e.g. "16u Boys"). Default: all sheets.
    --verbose    Print every team's results, not just problems.
"""
from __future__ import annotations
import argparse, os, sys, re, subprocess
from collections import defaultdict
from datetime import date, datetime

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)

from parsers.detect import load_and_parse
from app import (
    _expand_bracket_games, _build_wpl_game_tree,
    team_matches, strip_prefix, describe_slot,
    _tournament_meta, _SLOT_LIKE_RE,
    WPL_TOURNAMENTS, FUTURES_SHEETS_URL, EXCEL_DIR,
)

_WPL_KEYWORD = "futures wpl"  # substring match against filenames in EXCEL_DIR
_TOURNAMENT_KEYWORDS = {
    "kap7-intl":  "kap7 international",
    "kap7-cup":   "kap7 cup",
    "turbo-cup":  "turbo",
    "newport-invite": "newport",
}

def _find_local_excel(tournament_id: str):
    """Find the local Excel file for a tournament by keyword matching."""
    keyword = _WPL_KEYWORD if tournament_id in WPL_TOURNAMENTS \
              else _TOURNAMENT_KEYWORDS.get(tournament_id, "")
    if not keyword:
        return None
    for fname in os.listdir(EXCEL_DIR):
        if fname.endswith(".xlsx") and keyword.lower() in fname.lower():
            return os.path.join(EXCEL_DIR, fname)
    return None


# ── Thresholds ────────────────────────────────────────────────────────────────
# Tune these when a new championship format is encountered.
MIN_TOTAL_GAMES   = 2   # direct + extras must be at least this many
MIN_SUNDAY_GAMES  = 1   # must see at least 1 Sunday game/node somewhere
MAX_EXTRAS_RATIO  = 8   # extras > direct * ratio → likely contamination
CHAMP_WEEKEND_WINDOW = 1  # ±days to consider "this weekend"


# ── Colors ────────────────────────────────────────────────────────────────────
GREEN  = "\033[32m"
YELLOW = "\033[33m"
RED    = "\033[31m"
RESET  = "\033[0m"
BOLD   = "\033[1m"

def ok(msg):   return f"{GREEN}✅ {msg}{RESET}"
def warn(msg): return f"{YELLOW}⚠️  {msg}{RESET}"
def err(msg):  return f"{RED}❌ {msg}{RESET}"


# ── Helpers ───────────────────────────────────────────────────────────────────

_PTS_RE = re.compile(r'\s*-\s*\d+(\.\d+)?\s*PTS\.?\s*$', re.IGNORECASE)

def _all_teams_this_weekend(div_games: list, anchor: date) -> list[str]:
    """Return sorted unique team names that appear in this weekend's games.

    Only returns real team names — skips slot placeholders, composite slots,
    standings-table entries, and seed-number prefixes.
    """
    seen: set[str] = set()
    teams: list[str] = []
    for g in div_games:
        if not (g.get("date") and abs((g["date"] - anchor).days) <= CHAMP_WEEKEND_WINDOW):
            continue
        for slot in (g["white_team"], g["dark_team"]):
            name = strip_prefix(slot).strip().upper()
            # Strip "- N PTS." standings suffix
            name = _PTS_RE.sub("", name).strip()

            if not name or len(name) < 3:
                continue
            # Skip slots that still look like bracket references
            if re.search(r'\((?:WIN|LOS)\s+GM', name, re.IGNORECASE):
                continue
            if re.match(r'^(WIN|LOS|TBD|BYE)', name, re.IGNORECASE):
                continue
            # Skip bare seed numbers ("22", "13 - ")
            if re.match(r'^\d+\s*-?\s*$', name):
                continue
            # Skip pure pool positions ("A2", "H1")
            if re.match(r'^[A-Z]\d+$', name):
                continue
            # Skip if it's still a slot-like string after stripping
            if re.match(r'^\d', name):
                continue

            if name not in seen:
                seen.add(name)
                teams.append(name.title())
    return sorted(teams)


def _check_team(team: str, sheet: str, div_games: list, anchor: date) -> list[str]:
    """Run all checks for one team. Returns list of issue strings (empty = clean)."""
    issues = []

    # Direct games this weekend
    direct = sorted(
        [g for g in div_games
         if (team_matches(g["white_team"], team) or team_matches(g["dark_team"], team))
         and g.get("date") and abs((g["date"] - anchor).days) <= CHAMP_WEEKEND_WINDOW],
        key=lambda g: (g.get("date") or date.min, g.get("time") or datetime.min.time()),
    )
    if not direct:
        issues.append("no direct games found this weekend")
        return issues

    # Bracket expansion
    extras = _expand_bracket_games(team, direct, div_games)
    total  = len(direct) + len(extras)

    if total < MIN_TOTAL_GAMES:
        issues.append(f"only {total} total game(s) — expansion may be broken")

    # Cross-weekend contamination guard
    if len(direct) > 0 and len(extras) > len(direct) * MAX_EXTRAS_RATIO:
        issues.append(
            f"extras={len(extras)} >> direct={len(direct)} — "
            f"possible cross-weekend contamination"
        )

    # Sunday game presence (via extras or tree)
    sun_extras = [g for g in extras
                  if g.get("date") and g["date"].weekday() == 6]

    # Bracket tree
    tree = _build_wpl_game_tree(team, div_games, anchor_date=anchor)
    sun_tree = [n for n in tree if n.get("date") and n["date"].weekday() == 6]

    has_sunday = bool(sun_extras or sun_tree)
    if not has_sunday:
        issues.append(
            f"no Sunday games visible — expansion={len(extras)} extras, "
            f"tree={len(tree)} nodes — Sunday data may be missing from sheet"
        )

    if tree:
        # Unresolved slot strings in opponent labels
        bad_opps = []
        for n in tree:
            wt, dt = n["white_team"], n["dark_team"]
            try:
                opp_slot = dt if team.split()[-1].upper() in wt.upper() else wt
            except Exception:
                opp_slot = wt
            opp = describe_slot(opp_slot, div_games, ref_date=anchor)
            if _SLOT_LIKE_RE.match(opp):
                bad_opps.append(f"{n['game_id']}: {opp!r}")
        if bad_opps:
            issues.append(
                f"unresolved slot string(s) in opponent labels: "
                + ", ".join(bad_opps[:3])
            )

        # Placeholder discipline: inferred WIN-GM games must not be marked real
        inferred_real = []
        for n in tree:
            wt, dt = n["white_team"], n["dark_team"]
            is_inferred = (
                not team_matches(wt, team) and not team_matches(dt, team)
                and re.search(r'\bWIN\s+GM\s+#', wt + dt, re.IGNORECASE)
            )
            if is_inferred and not n.get("placeholder"):
                inferred_real.append(n["game_id"])
        if inferred_real:
            issues.append(
                f"inferred WIN-GM games incorrectly marked confirmed: "
                + str(inferred_real)
            )
    else:
        # Empty tree is only a problem for WPL teams (not round-robin-only formats)
        if any(re.search(r'^[A-Z]\d+-', s, re.IGNORECASE)
               for g in direct
               for s in (g["white_team"], g["dark_team"])
               if team_matches(s, team)):
            issues.append("bracket tree is empty for a pool-slot team")

    return issues


def _summary_line(team: str, direct: int, extras: int, tree: int,
                  sun: int, issues: list[str]) -> str:
    total = direct + extras
    stats = f"direct={direct} extras={extras} tree={tree} sun={sun}"
    if issues:
        return err(f"{team:<28s} {stats}")
    return ok(f"{team:<28s} {stats}")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("tournament", help="Tournament ID, e.g. futures-5")
    p.add_argument("--sheet",    help="Only check this sheet, e.g. '16u Boys'")
    p.add_argument("--download", action="store_true",
                   help="Download fresh spreadsheet before checking")
    p.add_argument("--verbose",  action="store_true",
                   help="Print all teams, not just problems")
    args = p.parse_args()

    meta = _tournament_meta(args.tournament)
    if not meta:
        print(err(f"Unknown tournament: {args.tournament!r}"))
        sys.exit(1)

    anchor = meta.get("date_start")
    if not anchor:
        print(err(f"No date_start configured for {args.tournament!r}"))
        sys.exit(1)

    print(f"\n{BOLD}{'='*64}{RESET}")
    print(f"{BOLD}Pre-game check: {meta.get('name', args.tournament)}{RESET}")
    print(f"Weekend: {meta.get('dates', str(anchor))}")

    # Find the local Excel file for this tournament
    excel_path = _find_local_excel(args.tournament)

    if args.download:
        if args.tournament in WPL_TOURNAMENTS:
            import shutil, urllib.request
            print(f"Downloading fresh spreadsheet from Google Sheets…")
            try:
                tmp, _ = urllib.request.urlretrieve(FUTURES_SHEETS_URL)
                if excel_path:
                    shutil.copy(tmp, excel_path)
                else:
                    # Save as a new file
                    excel_path = os.path.join(
                        EXCEL_DIR,
                        "2026 KAP7 Futures WPL - Southern California - Presented by BIWPA.xlsx"
                    )
                    shutil.copy(tmp, excel_path)
                print(f"  → Updated {os.path.basename(excel_path)}")
            except Exception as e:
                print(warn(f"  Download failed: {e}. Using cached file."))
        else:
            print(warn("--download only supported for WPL Futures tournaments"))

    if not excel_path:
        print(err(f"No local Excel file found for {args.tournament!r}"))
        print(f"  Try: python3 scripts/pre_game_check.py {args.tournament} --download")
        sys.exit(1)

    print(f"File: {os.path.basename(excel_path)}")
    print(f"{'='*64}{RESET}\n")

    games = load_and_parse(excel_path)

    # Determine sheets to check
    all_sheets = sorted({g["sheet"] for g in games})
    sheets_to_check = [args.sheet] if args.sheet else all_sheets

    grand_ok = grand_warn = grand_err = 0

    for sheet in sheets_to_check:
        div_games = [g for g in games if g["sheet"] == sheet]
        teams     = _all_teams_this_weekend(div_games, anchor)

        if not teams:
            if args.verbose:
                print(f"{YELLOW}{sheet}: no games found for {anchor}{RESET}")
            continue

        print(f"{BOLD}{sheet}{RESET}  ({len(teams)} teams this weekend)\n")

        team_issues: dict[str, list[str]] = {}
        for team in teams:
            direct = [g for g in div_games
                      if (team_matches(g["white_team"], team) or
                          team_matches(g["dark_team"], team))
                      and g.get("date") and abs((g["date"] - anchor).days) <= CHAMP_WEEKEND_WINDOW]
            extras = _expand_bracket_games(team, direct, div_games)
            tree   = _build_wpl_game_tree(team, div_games, anchor_date=anchor)
            sun    = len([n for n in tree if n.get("date") and n["date"].weekday() == 6])
            issues = _check_team(team, sheet, div_games, anchor)

            team_issues[team] = issues
            line = _summary_line(team, len(direct), len(extras), len(tree), sun, issues)

            if issues:
                grand_err += 1
                print(f"  {line}")
                for iss in issues:
                    print(f"      → {iss}")
            else:
                grand_ok += 1
                if args.verbose:
                    print(f"  {line}")

        print()

    # Summary
    total = grand_ok + grand_warn + grand_err
    print(f"{'='*64}")
    print(f"Results: {GREEN}{grand_ok} clean{RESET}  "
          f"{RED}{grand_err} with issues{RESET}  "
          f"(out of {total} teams checked)")

    if grand_err == 0:
        print(f"\n{GREEN}{BOLD}✅ All teams look good — safe to open for parents.{RESET}")
    else:
        print(f"\n{RED}{BOLD}❌ Fix the issues above before games start.{RESET}")

    print(f"{'='*64}\n")
    sys.exit(0 if grand_err == 0 else 1)


if __name__ == "__main__":
    main()

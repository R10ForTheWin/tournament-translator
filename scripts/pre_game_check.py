#!/usr/bin/env python3
"""
Pre-game validation sweep for a tournament weekend.

Run the night before games to catch data issues before parents see them.
Every check here encodes a bug that was found the hard way on game day.

For WPL Futures tournaments, a fresh copy is downloaded automatically.
For other tournaments, pass --download to fetch a fresh copy.

Usage (from project root):
    python3 scripts/pre_game_check.py futures-5
    python3 scripts/pre_game_check.py futures-5 --sheet "16u Boys"
    python3 scripts/pre_game_check.py futures-5 --sheet "16u Boys" --verbose
    python3 scripts/pre_game_check.py futures-5 --no-download   # skip download, use cache
    python3 scripts/pre_game_check.py kap7-cup --download       # non-WPL needs explicit flag

Flags:
    --no-download  Skip downloading fresh data (use local cached file).
                   For WPL tournaments, downloading is the default.
    --sheet X      Only check one sheet (e.g. "16u Boys"). Default: all sheets.
    --verbose      Print every team's results, not just problems.
"""
from __future__ import annotations
import argparse, os, sys, re, shutil, urllib.request
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

_WPL_KEYWORD = "futures wpl"
_TOURNAMENT_KEYWORDS = {
    "kap7-intl":      "kap7 international",
    "kap7-cup":       "kap7 cup",
    "turbo-cup":      "turbo",
    "newport-invite": "newport",
}

_WPL_LOCAL_FILENAME = (
    "2026 KAP7 Futures WPL - Southern California - Presented by BIWPA.xlsx"
)

def _find_local_excel(tournament_id: str) -> str | None:
    keyword = _WPL_KEYWORD if tournament_id in WPL_TOURNAMENTS \
              else _TOURNAMENT_KEYWORDS.get(tournament_id, "")
    if not keyword:
        return None
    for fname in os.listdir(EXCEL_DIR):
        if fname.endswith(".xlsx") and keyword.lower() in fname.lower():
            return os.path.join(EXCEL_DIR, fname)
    return None


def _download_wpl(excel_path: str | None) -> str | None:
    """Download fresh WPL sheet; save to excel_path (or default). Returns final path."""
    dest = excel_path or os.path.join(EXCEL_DIR, _WPL_LOCAL_FILENAME)
    print(f"  Downloading fresh spreadsheet from Google Sheets…", end=" ", flush=True)
    try:
        tmp, _ = urllib.request.urlretrieve(FUTURES_SHEETS_URL)
        shutil.copy(tmp, dest)
        print(f"✓  ({os.path.getsize(dest) // 1024} KB)")
        return dest
    except Exception as e:
        print(f"FAILED: {e}")
        return excel_path  # fall back to cached


# ── Thresholds ────────────────────────────────────────────────────────────────
MIN_TOTAL_GAMES      = 2   # direct + extras must be at least this many
MAX_EXTRAS_RATIO     = 8   # extras > direct × ratio → likely contamination
CHAMP_WEEKEND_WINDOW = 1   # ±days to consider "this weekend"


# ── Colors ────────────────────────────────────────────────────────────────────
GREEN  = "\033[32m"
YELLOW = "\033[33m"
RED    = "\033[31m"
RESET  = "\033[0m"
BOLD   = "\033[1m"

def _ok(msg):   return f"{GREEN}✅ {msg}{RESET}"
def _warn(msg): return f"{YELLOW}⚠️  {msg}{RESET}"
def _err(msg):  return f"{RED}❌ {msg}{RESET}"


# ── Team name extraction ──────────────────────────────────────────────────────

_PTS_RE = re.compile(r'\s*-\s*\d+(\.\d+)?\s*PTS\.?\s*$', re.IGNORECASE)

def _all_teams_this_weekend(div_games: list, anchor: date) -> list[str]:
    """Return sorted unique real team names playing this weekend.

    Filters out slot placeholders, composite slots, standings-table entries,
    and seed-number prefixes — only returns human-readable team names.
    """
    seen: set[str] = set()
    teams: list[str] = []
    for g in div_games:
        if not (g.get("date") and abs((g["date"] - anchor).days) <= CHAMP_WEEKEND_WINDOW):
            continue
        for slot in (g["white_team"], g["dark_team"]):
            name = strip_prefix(slot).strip().upper()
            name = _PTS_RE.sub("", name).strip()
            if not name or len(name) < 3:
                continue
            if re.search(r'\((?:WIN|LOS)\s+GM', name, re.IGNORECASE):
                continue
            if re.match(r'^(WIN|LOS|TBD|BYE)', name, re.IGNORECASE):
                continue
            if re.match(r'^\d+\s*-?\s*$', name):
                continue
            if re.match(r'^[A-Z]\d+$', name):
                continue
            if re.match(r'^\d', name):
                continue
            if name not in seen:
                seen.add(name)
                teams.append(name.title())
    return sorted(teams)


# ── Per-team checks ───────────────────────────────────────────────────────────

def _visible_sunday_count(direct: list, extras: list, tree: list) -> int:
    """Count unique Sunday games visible to a team across all sources."""
    sun_ids: set[str] = set()
    for g in direct + extras:
        if g.get("date") and g["date"].weekday() == 6:
            sun_ids.add(g["game_id"])
    for n in tree:
        if n.get("date") and n["date"].weekday() == 6:
            sun_ids.add(n["game_id"])
    return len(sun_ids)


def _check_team(team: str, div_games: list, anchor: date) -> tuple[list, list, list, int, list[str]]:
    """Run all checks. Returns (direct, extras, tree, sun_count, issues)."""
    issues = []

    direct = sorted(
        [g for g in div_games
         if (team_matches(g["white_team"], team) or team_matches(g["dark_team"], team))
         and g.get("date") and abs((g["date"] - anchor).days) <= CHAMP_WEEKEND_WINDOW],
        key=lambda g: (g.get("date") or date.min, g.get("time") or datetime.min.time()),
    )
    if not direct:
        return [], [], [], 0, ["no direct games found this weekend"]

    extras = _expand_bracket_games(team, direct, div_games)
    tree   = _build_wpl_game_tree(team, div_games, anchor_date=anchor)
    sun    = _visible_sunday_count(direct, extras, tree)

    total = len(direct) + len(extras)
    if total < MIN_TOTAL_GAMES:
        issues.append(f"only {total} total game(s) — expansion may be broken")

    if len(direct) > 0 and len(extras) > len(direct) * MAX_EXTRAS_RATIO:
        issues.append(
            f"extras={len(extras)} >> direct={len(direct)} — "
            "possible cross-weekend contamination"
        )

    if sun == 0:
        issues.append(
            f"no Sunday games visible (direct={len(direct)} extras={len(extras)} "
            f"tree={len(tree)}) — Sunday data may be missing from sheet"
        )

    if tree:
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
                "unresolved slot string(s) in opponent labels: "
                + ", ".join(bad_opps[:3])
            )

        inferred_real = [
            n["game_id"] for n in tree
            if not n.get("placeholder")
            and not team_matches(n["white_team"], team)
            and not team_matches(n["dark_team"], team)
            and re.search(r'\bWIN\s+GM\s+#', n["white_team"] + n["dark_team"], re.IGNORECASE)
        ]
        if inferred_real:
            issues.append(
                "inferred WIN-GM games incorrectly marked confirmed: "
                + str(inferred_real)
            )
    else:
        if any(re.search(r'^[A-Z]\d+-', s, re.IGNORECASE)
               for g in direct
               for s in (g["white_team"], g["dark_team"])
               if team_matches(s, team)):
            issues.append("bracket tree is empty for a pool-slot team")

    return direct, extras, tree, sun, issues


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    p = argparse.ArgumentParser(description=__doc__,
                                formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("tournament", help="Tournament ID, e.g. futures-5")
    p.add_argument("--sheet",       help="Only check this sheet, e.g. '16u Boys'")
    p.add_argument("--no-download", action="store_true",
                   help="Skip downloading fresh data; use local cached file.")
    p.add_argument("--verbose",     action="store_true",
                   help="Print all teams, not just problems")
    args = p.parse_args()

    meta = _tournament_meta(args.tournament)
    if not meta:
        print(_err(f"Unknown tournament: {args.tournament!r}"))
        sys.exit(1)

    anchor = meta.get("date_start")
    if not anchor:
        print(_err(f"No date_start configured for {args.tournament!r}"))
        sys.exit(1)

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n{BOLD}{'='*64}{RESET}")
    print(f"{BOLD}Pre-game check: {meta.get('name', args.tournament)}{RESET}")
    print(f"Weekend: {meta.get('dates', str(anchor))}   |   Run at: {now_str}")

    excel_path = _find_local_excel(args.tournament)

    should_download = (args.tournament in WPL_TOURNAMENTS) and not args.no_download
    if should_download:
        excel_path = _download_wpl(excel_path)
    elif not excel_path:
        # Non-WPL: need explicit --download
        print(_err(f"No local file for {args.tournament!r} — run with --download"))
        sys.exit(1)

    if not excel_path:
        print(_err(f"No Excel file available — download failed and no cache found"))
        sys.exit(1)

    print(f"File: {os.path.basename(excel_path)}")
    print(f"{'='*64}{RESET}\n")

    games = load_and_parse(excel_path)
    all_sheets = sorted({g["sheet"] for g in games})
    sheets_to_check = [args.sheet] if args.sheet else all_sheets

    grand_ok = grand_err = 0

    for sheet in sheets_to_check:
        div_games = [g for g in games if g["sheet"] == sheet]
        teams     = _all_teams_this_weekend(div_games, anchor)

        if not teams:
            if args.verbose:
                print(f"{YELLOW}{sheet}: no games found for {anchor}{RESET}\n")
            continue

        print(f"{BOLD}{sheet}{RESET}  ({len(teams)} teams)\n")

        for team in teams:
            direct, extras, tree, sun, issues = _check_team(team, div_games, anchor)
            stats = (f"direct={len(direct)} extras={len(extras)} "
                     f"tree={len(tree)} sun={sun}")
            line_body = f"{team:<28s} {stats}"

            if issues:
                grand_err += 1
                print(f"  {_err(line_body)}")
                for iss in issues:
                    print(f"      → {iss}")
            else:
                grand_ok += 1
                if args.verbose:
                    print(f"  {_ok(line_body)}")

        print()

    total = grand_ok + grand_err
    print(f"{'='*64}")
    print(f"Results: {GREEN}{grand_ok} clean{RESET}  {RED}{grand_err} with issues{RESET}"
          f"  (of {total} teams)")

    if grand_err == 0:
        print(f"\n{GREEN}{BOLD}✅ All teams look good — safe to open for parents.{RESET}")
    else:
        print(f"\n{RED}{BOLD}❌ Fix the issues above before games start.{RESET}")
    print(f"{'='*64}\n")
    sys.exit(0 if grand_err == 0 else 1)


if __name__ == "__main__":
    main()

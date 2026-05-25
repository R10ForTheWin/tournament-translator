#!/usr/bin/env python3
"""
Tournament Translator — Hourly bracket health monitor with LLM judge.

Two-layer checking:
  1. Deterministic: game count, bracket_confidence, bracket_warnings (fast, free)
  2. LLM judge: Claude Haiku reviews the full bracket display for semantic issues
     the deterministic layer can't catch (missing paths, over-expansion, etc.)

Notifications: POST to ntfy.sh if NTFY_TOPIC env var is set.
  Setup: install the ntfy app, subscribe to your topic, set NTFY_TOPIC on Railway.

Exit codes: 0=all green, 1=red issues, 2=yellow warnings

Usage:
    python3 scripts/tournament_monitor.py
    python3 scripts/tournament_monitor.py --api-url https://custom-url.railway.app
    python3 scripts/tournament_monitor.py --skip-llm   # deterministic only
"""
import sys, json, os, re, urllib.request, urllib.error, urllib.parse, argparse
from datetime import date
from collections import defaultdict


def _time_min(t):
    """Parse '9:00 AM' → minutes since midnight. Returns -1 on failure."""
    if not t:
        return -1
    m = re.match(r'(\d+):(\d+)\s*(AM|PM)', t.strip(), re.IGNORECASE)
    if not m:
        return -1
    h, mn, ampm = int(m.group(1)), int(m.group(2)), m.group(3).upper()
    if ampm == 'PM' and h != 12:
        h += 12
    if ampm == 'AM' and h == 12:
        h = 0
    return h * 60 + mn


def _day_order(d):
    """Return a sort key for a date string like 'Thursday, May 29'."""
    days = ['monday', 'tuesday', 'wednesday', 'thursday', 'friday', 'saturday', 'sunday']
    if not d:
        return 99
    first = d.split(',')[0].strip().lower()
    return days.index(first) if first in days else 99


def check_game_num_order(upcoming_games):
    """Return error string if any game_num N column starts earlier than game_num N-1."""
    by_num = defaultdict(list)
    for g in upcoming_games:
        by_num[g.get("game_num") or 0].append(g)
    prev_day, prev_min = -1, -1
    for n in sorted(by_num.keys()):
        dated = [g for g in by_num[n] if g.get("date") and g.get("time")]
        if not dated:
            continue
        dated.sort(key=lambda g: (_day_order(g["date"]), _time_min(g["time"])))
        earliest = dated[0]
        day = _day_order(earliest["date"])
        tmin = _time_min(earliest["time"])
        if day < prev_day or (day == prev_day and tmin < prev_min):
            return (f"Game #{n} starts {earliest['date']} {earliest['time']} "
                    f"but Game #{n-1} starts later — game_num labels out of order")
        prev_day, prev_min = day, tmin
    return None

RAILWAY_URL = "https://web-production-5a744.up.railway.app"

TROJAN_TEAMS = [
    "trojan cardinal",
    "trojan gold",
    "trojan silver",
]

TOURNAMENTS = [
    {"id": "jo-quals",        "name": "JO Qualifications",  "start": date(2026, 5, 29), "end": date(2026, 5, 31), "days": 3},
    {"id": "futures-super",   "name": "Futures Superfinal", "start": date(2026, 6, 26), "end": date(2026, 6, 28)},
    {"id": "junior-olympics", "name": "Junior Olympics",    "start": date(2026, 7, 23), "end": date(2026, 7, 26), "days": 3},
]

# Bracket game count thresholds — flag outside this range pre-tournament
MIN_GAMES = 2
MAX_GAMES = 9       # 2-day tournaments (WPL Futures, Kap7, Newport)
MAX_GAMES_3DAY = 18 # 3-day tournaments (JO Quals, Junior Olympics)


def active_tournaments():
    today = date.today()
    return [t for t in TOURNAMENTS if t["start"] <= today <= t["end"]]


def check_team(base_url, tournament_id, team):
    team_encoded = urllib.parse.quote(team)
    url = f"{base_url}/api/games/{tournament_id}/{team_encoded}"
    try:
        with urllib.request.urlopen(url, timeout=15) as resp:
            data = json.load(resp)
    except urllib.error.HTTPError as e:
        if e.code == 404:
            return None
        return {"error": f"HTTP {e.code} from {url}"}
    except Exception as e:
        return {"error": f"{type(e).__name__}: {e}"}

    played   = data.get("played")   or []
    upcoming = data.get("upcoming") or []
    bracket  = data.get("wpl_bracket") or []

    return {
        "team":               data.get("our_team_name", team),
        "division":           data.get("division", ""),
        "bracket_confidence": data.get("bracket_confidence", "unknown"),
        "display_mode":       data.get("display_mode", "unknown"),
        "bracket_warnings":   data.get("bracket_warnings") or [],
        "wpl_bracket_nodes":  len(bracket),
        "played":             len(played),
        "upcoming_count":     len(upcoming),
        "upcoming_games":     upcoming,  # full objects for LLM judge
    }


def llm_judge(team_name, tournament_name, division, upcoming_games, max_games=MAX_GAMES):
    """Ask Claude Haiku if this bracket display looks reasonable.

    Returns {"status": "OK"|"FLAG", "reason": str} or None if unavailable.
    """
    try:
        import anthropic
    except ImportError:
        return None

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not api_key:
        return None

    n = len(upcoming_games)

    # Fast pre-filter: obvious count anomalies don't need LLM
    if n < MIN_GAMES:
        return {"status": "FLAG", "reason": f"Only {n} upcoming game(s) — bracket expansion may have failed."}
    if n > max_games:
        return {"status": "FLAG", "reason": f"{n} upcoming games shown — likely over-expansion (expected ≤{max_games})."}

    lines = []
    for i, g in enumerate(upcoming_games, 1):
        ph    = "placeholder" if g.get("placeholder") else "named"
        path  = f" | path:{g['path']}" if g.get("path") else ""
        dt    = f"{g.get('date','')} {g.get('time','')}".strip()
        lines.append(
            f"{i}. [{g.get('game_id','')}] {dt} — {g.get('white_team','')} vs {g.get('dark_team','')} ({ph}{path})"
        )

    prompt = f"""You are reviewing a water polo bracket display about to be shown to parents on a mobile app.

Tournament: {tournament_name}
Division: {division}
Team: {team_name}
Total upcoming games: {n}

Games:
{chr(10).join(lines)}

A healthy bracket path has:
- 3–8 games (1 direct pool game + win/lose branches + bracket continuation)
- At least 1 named game where the team appears explicitly (not just W#/L# or finish-slot refs)
- A win path AND a lose path after pool play
- No duplicate bracket outcomes (same time/date slot appearing for different pool ranks)

Flag if something looks wrong that a parent would find confusing or misleading.

Reply ONLY in this format:
STATUS: OK

or:

STATUS: FLAG
REASON: one sentence"""

    try:
        client = anthropic.Anthropic(api_key=api_key)
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=80,
            messages=[{"role": "user", "content": prompt}],
        )
        text = msg.content[0].text.strip()
    except Exception as e:
        return {"status": "ERROR", "reason": str(e)}

    status, reason = "OK", ""
    for line in text.splitlines():
        if line.upper().startswith("STATUS:"):
            status = line.split(":", 1)[1].strip().upper()
        elif line.upper().startswith("REASON:"):
            reason = line.split(":", 1)[1].strip()
    return {"status": status, "reason": reason}


def send_notification(title, body):
    """Push notification via ntfy.sh. Requires NTFY_TOPIC env var."""
    topic = os.environ.get("NTFY_TOPIC")
    if not topic:
        return
    try:
        req = urllib.request.Request(
            f"https://ntfy.sh/{topic}",
            data=body.encode("utf-8"),
            headers={
                "Title":    title,
                "Priority": "high",
                "Tags":     "warning",
            },
        )
        urllib.request.urlopen(req, timeout=5)
    except Exception:
        pass  # notification failure is non-fatal


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--api-url",   default=RAILWAY_URL)
    parser.add_argument("--skip-llm",  action="store_true", help="Skip LLM judge, run deterministic checks only")
    args = parser.parse_args()
    base_url = args.api_url.rstrip("/")

    active = active_tournaments()
    if not active:
        print(f"No active tournament today ({date.today()}). Nothing to check.")
        sys.exit(0)

    red_issues    = []
    yellow_issues = []
    report_lines  = []
    notify_lines  = []

    for t in active:
        tid  = t["id"]
        name = t["name"]
        report_lines.append(f"\n=== {name} ({date.today()}) ===")

        for team in TROJAN_TEAMS:
            result = check_team(base_url, tid, team)
            if result is None:
                continue
            if "error" in result:
                msg = f"{team} [{tid}]: API error — {result['error']}"
                red_issues.append(msg)
                notify_lines.append(msg)
                report_lines.append(f"  ✗ {team}: {result['error']}")
                continue

            team_name = result["team"]
            conf      = result["bracket_confidence"]
            mode      = result["display_mode"]
            warns     = result["bracket_warnings"]
            nodes     = result["wpl_bracket_nodes"]
            n_up      = result["upcoming_count"]

            # ── Deterministic checks ────────────────────────────────────────
            max_games = MAX_GAMES_3DAY if t.get("days", 2) >= 3 else MAX_GAMES
            det_flags = list(warns)  # existing app-level warnings
            if n_up > max_games:
                det_flags.append(f"{n_up} upcoming games shown (expected ≤{max_games})")
            elif n_up < MIN_GAMES and conf != "green":
                det_flags.append(f"Only {n_up} upcoming game(s)")
            order_err = check_game_num_order(result["upcoming_games"])
            if order_err:
                det_flags.append(order_err)

            # ── LLM judge ───────────────────────────────────────────────────
            llm_result = None
            if not args.skip_llm:
                llm_result = llm_judge(
                    team_name, name,
                    result.get("division", ""),
                    result["upcoming_games"],
                    max_games=max_games,
                )

            # ── Build report line ───────────────────────────────────────────
            llm_tag = ""
            if llm_result:
                if llm_result["status"] == "FLAG":
                    llm_tag = f" | LLM: {llm_result['reason']}"
                elif llm_result["status"] == "ERROR":
                    llm_tag = f" | LLM-err: {llm_result['reason']}"

            icon = {"green": "✓", "yellow": "⚠", "red": "✗"}.get(conf, "?")
            report_lines.append(
                f"  {icon} {team_name}: conf={conf} mode={mode} "
                f"nodes={nodes} played={result['played']} upcoming={n_up}{llm_tag}"
            )
            for w in det_flags:
                report_lines.append(f"       → {w}")

            # ── Classify severity ───────────────────────────────────────────
            is_red    = conf == "red" or n_up > max_games
            is_yellow = (conf == "yellow" and det_flags) or (
                llm_result and llm_result["status"] == "FLAG"
            )

            summary = f"{team_name} [{name}]"
            all_issues = det_flags[:]
            if llm_result and llm_result["status"] == "FLAG":
                all_issues.append(f"LLM: {llm_result['reason']}")

            if all_issues:
                summary += ": " + "; ".join(all_issues[:2])

            if is_red:
                red_issues.append(summary)
                notify_lines.append(summary)
            elif is_yellow:
                yellow_issues.append(summary)
                notify_lines.append(summary)

    report = "\n".join(report_lines)
    print(report)

    if red_issues:
        print(f"\n{'='*50}")
        print(f"RED ISSUES ({len(red_issues)}):")
        for i in red_issues: print(f"  • {i}")
        send_notification(
            "Tournament Translator — RED",
            "\n".join(notify_lines[:5]),
        )
        sys.exit(1)
    elif yellow_issues:
        print(f"\n{'='*50}")
        print(f"YELLOW WARNINGS ({len(yellow_issues)}):")
        for i in yellow_issues: print(f"  • {i}")
        send_notification(
            "Tournament Translator — Warning",
            "\n".join(notify_lines[:5]),
        )
        sys.exit(2)
    else:
        print("\nAll checks passed ✓")
        sys.exit(0)


if __name__ == "__main__":
    main()

#!/usr/bin/env python3
"""
Tournament Translator — Hourly bracket health monitor.

Hits the live Railway API for each known Trojan team during active tournament
weekends. Exits 0 if all green, exits 1 if any RED/YELLOW issues found so the
calling agent knows to send an alert.

Usage:
    python3 scripts/tournament_monitor.py
    python3 scripts/tournament_monitor.py --api-url https://custom-url.railway.app
"""
import sys, json, urllib.request, urllib.error, argparse
from datetime import date

RAILWAY_URL = "https://web-production-5a744.up.railway.app"

TROJAN_TEAMS = [
    "trojan cardinal",
    "trojan gold",
    "trojan silver",
]

TOURNAMENTS = [
    {"id": "jo-quals",        "name": "JO Qualifications",   "start": date(2026, 5, 29), "end": date(2026, 5, 31)},
    {"id": "futures-super",   "name": "Futures Superfinal",  "start": date(2026, 6, 26), "end": date(2026, 6, 28)},
    {"id": "junior-olympics", "name": "Junior Olympics",     "start": date(2026, 7, 23), "end": date(2026, 7, 26)},
]


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
            return None   # team not in this tournament — skip
        return {"error": f"HTTP {e.code} from {url}"}
    except Exception as e:
        return {"error": f"{type(e).__name__}: {e}"}

    played   = data.get("played") or []
    upcoming = data.get("upcoming") or []
    bracket  = data.get("wpl_bracket") or []

    return {
        "team":               data.get("our_team_name", team),
        "bracket_confidence": data.get("bracket_confidence", "unknown"),
        "display_mode":       data.get("display_mode", "unknown"),
        "bracket_warnings":   data.get("bracket_warnings") or [],
        "wpl_bracket_nodes":  len(bracket),
        "played":             len(played),
        "upcoming":           len(upcoming),
    }


def main():
    import urllib.parse

    parser = argparse.ArgumentParser()
    parser.add_argument("--api-url", default=RAILWAY_URL)
    args = parser.parse_args()
    base_url = args.api_url.rstrip("/")

    active = active_tournaments()
    if not active:
        print(f"No active tournament today ({date.today()}). Nothing to check.")
        sys.exit(0)

    red_issues    = []
    yellow_issues = []
    report_lines  = []

    for t in active:
        tid  = t["id"]
        name = t["name"]
        report_lines.append(f"\n=== {name} ({date.today()}) ===")

        for team in TROJAN_TEAMS:
            result = check_team(base_url, tid, team)
            if result is None:
                continue
            if "error" in result:
                red_issues.append(f"{team} [{tid}]: API error — {result['error']}")
                report_lines.append(f"  ✗ {team}: {result['error']}")
                continue

            conf  = result["bracket_confidence"]
            mode  = result["display_mode"]
            warns = result["bracket_warnings"]
            nodes = result["wpl_bracket_nodes"]

            icon = {"green": "✓", "yellow": "⚠", "red": "✗"}.get(conf, "?")
            report_lines.append(
                f"  {icon} {result['team']}: confidence={conf} "
                f"mode={mode} nodes={nodes} "
                f"played={result['played']} upcoming={result['upcoming']}"
            )
            for w in warns:
                report_lines.append(f"       → {w}")

            summary = f"{result['team']} [{name}]: {'; '.join(warns[:2]) or mode}"
            if conf == "red":
                red_issues.append(summary)
            elif conf == "yellow" and warns:
                yellow_issues.append(summary)

    report = "\n".join(report_lines)
    print(report)

    if red_issues:
        print(f"\n{'='*50}")
        print(f"RED ISSUES ({len(red_issues)}):")
        for i in red_issues:
            print(f"  • {i}")
        sys.exit(1)
    elif yellow_issues:
        print(f"\n{'='*50}")
        print(f"YELLOW WARNINGS ({len(yellow_issues)}):")
        for i in yellow_issues:
            print(f"  • {i}")
        sys.exit(2)
    else:
        print(f"\nAll checks passed ✓")
        sys.exit(0)


if __name__ == "__main__":
    main()

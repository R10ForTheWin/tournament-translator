"""
Format A parser — Turbo OC Cup, Kap7 International, Newport Spring Invite.
Columns: DATE, TIME, LOCATION, GAME_ID, WHITE, W_SCORE, DARK, D_SCORE, COMMENTS, DIVISION
"""
import re
from datetime import datetime

SKIP_SHEETS = {
    "TEAM LISTING AND BRACKETS", "CHICLETS", "CHICLETS MASTER", "CHICLETS NEW",
    "MASTER BY DIVISION", "MASTER BY LOCATION", "MASTER BY TIME",
    "RAW ENTRIES", "MEDALS",
}

def parse(wb) -> list[dict]:
    games = []
    # game_id -> (white_team, sheet_name) of first occurrence. Dedup stays GLOBAL
    # across sheets (Format A workbooks often carry stale duplicate/draft division
    # sheets — e.g. "18U BOYS PLATINUM-22 TEAMS" alongside the real "-23 TEAMS" —
    # and rely on the first, real occurrence winning; cross-sheet collisions must
    # still be silently dropped, not renamed).
    seen_ids: dict[str, tuple[str, str]] = {}
    collision_count: dict[str, int] = {}  # base id -> # of extra occurrences so far
    for sheet_name in wb.sheetnames:
        if sheet_name.upper() in {s.upper() for s in SKIP_SHEETS}:
            continue
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=1, values_only=True):
            if len(row) < 8:
                continue
            date_val, time_val, location, game_id = row[0], row[1], row[2], row[3]
            white, w_score, dark, d_score = row[4], row[5], row[6], row[7]
            comments = row[8] if len(row) > 8 else None
            division = row[9] if len(row) > 9 else None

            if not isinstance(date_val, datetime):
                continue
            if not game_id or not isinstance(game_id, str):
                continue
            if not white or not dark:
                continue

            gid = game_id.strip()
            white_team = str(white).strip()

            if gid not in seen_ids:
                seen_ids[gid] = (white_team, sheet_name)
            elif seen_ids[gid] == (white_team, sheet_name):
                continue  # true duplicate (same id, same teams, same sheet) — skip
            elif seen_ids[gid][1] != sheet_name:
                continue  # collision from a different sheet (stale/duplicate division tab) — skip
            else:
                # Same ID, different teams, SAME sheet — organizer reused an ID by
                # mistake within one authoritative schedule. Rename to preserve
                # this game instead of silently dropping it.
                n = collision_count.get(gid, 0) + 1
                collision_count[gid] = n
                gid = f"{gid}-{chr(ord('B') + n - 1)}"
                seen_ids[gid] = (white_team, sheet_name)

            games.append({
                "date":       date_val.date(),
                "time":       time_val if hasattr(time_val, "hour") else None,
                "location":   str(location).strip() if location else "TBD",
                "game_id":    gid,
                "white_team": white_team,
                "white_score": _to_int(w_score),
                "dark_team":  str(dark).strip(),
                "dark_score": _to_int(d_score),
                "comments":   str(comments).strip() if comments else "",
                "division":   str(division).strip() if division else sheet_name,
                "sheet":      sheet_name,
                "played":     w_score is not None and d_score is not None,
            })
    return games


def _to_int(val):
    if val is None:
        return None
    try:
        return int(float(val))
    except (TypeError, ValueError):
        return None

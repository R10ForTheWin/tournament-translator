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
    seen_ids: set = set()
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
            if gid in seen_ids:
                continue
            seen_ids.add(gid)
            games.append({
                "date":       date_val.date(),
                "time":       time_val if hasattr(time_val, "hour") else None,
                "location":   str(location).strip() if location else "TBD",
                "game_id":    gid,
                "white_team": str(white).strip(),
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

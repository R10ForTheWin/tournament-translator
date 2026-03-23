"""
Format B parser — Kap7 Futures WPL / Southern California.
Columns: GAME_ID, TIME, WHITE, W_SCORE, DARK, D_SCORE, DIVISION, NOTES
Dates are embedded in section header rows, not in game rows.
"""
import re
from datetime import datetime, date, time as dtime

SKIP_SHEETS = {"InfoRules", "DivisionsStandings"}

# Matches section headers like "Weekend 4 - March 21, 2026" or "Day 1 - April 18"
_DATE_HEADER_RE = re.compile(
    r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
    r"Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
    r"\s+(\d{1,2})(?:,\s*(\d{4}))?",
    re.IGNORECASE,
)
_MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}

def parse(wb) -> list[dict]:
    games = []
    for sheet_name in wb.sheetnames:
        if sheet_name in SKIP_SHEETS:
            continue
        ws = wb[sheet_name]
        current_date = None

        for row in ws.iter_rows(min_row=1, values_only=True):
            if not any(c is not None for c in row):
                continue

            # Check for a date header row
            first = str(row[0]).strip() if row[0] else ""
            date_match = _DATE_HEADER_RE.search(first)
            if date_match and (row[3] is None or row[4] is None):
                current_date = _parse_date_from_match(date_match)
                continue

            # Game rows: col0=game_id (str), col1=time, col2=white, col3=w_score,
            #            col4=dark, col5=d_score, col6=division, col7=notes
            if len(row) < 5:
                continue

            game_id, time_val, white, w_score = row[0], row[1], row[2], row[3]
            dark = row[4] if len(row) > 4 else None
            d_score = row[5] if len(row) > 5 else None
            division = row[6] if len(row) > 6 else None
            notes = row[7] if len(row) > 7 else None

            if not game_id or not isinstance(game_id, str):
                continue
            if not white or not dark:
                continue
            # Skip rows that look like headers
            if str(white).strip().upper() in ("WHITE", "TEAM", "WHITE TEAM"):
                continue

            games.append({
                "date":        current_date,
                "time":        time_val if hasattr(time_val, "hour") else None,
                "location":    "TBD",
                "game_id":     str(game_id).strip(),
                "white_team":  _normalize_team(str(white).strip()),
                "white_score": _to_int(w_score),
                "dark_team":   _normalize_team(str(dark).strip()),
                "dark_score":  _to_int(d_score),
                "comments":    str(notes).strip() if notes else "",
                "division":    str(division).strip() if division else sheet_name,
                "sheet":       sheet_name,
                "played":      w_score is not None and d_score is not None,
            })
    return games


def _normalize_team(s: str) -> str:
    """Convert 'B1 - Trojan Gold' → 'B1-TROJAN GOLD' to match Format A style."""
    m = re.match(r"^([A-Z]\d+)\s*-\s*(.+)$", s, re.IGNORECASE)
    if m:
        return f"{m.group(1).upper()}-{m.group(2).strip().upper()}"
    # Placement like "1st in A - Team" → "1stA-TEAM"
    m2 = re.match(r"^(\d+(?:st|nd|rd|th))\s+in\s+([A-Z])\s*-\s*(.+)$", s, re.IGNORECASE)
    if m2:
        return f"{m2.group(1)}{m2.group(2).upper()}-{m2.group(3).strip().upper()}"
    return s.upper()


def _parse_date_from_match(m) -> date:
    month_str = m.group(0).split()[0][:3].lower()
    month = _MONTH_MAP.get(month_str, 1)
    day = int(m.group(1))
    year = int(m.group(2)) if m.group(2) else datetime.now().year
    try:
        return date(year, month, day)
    except ValueError:
        return None


def _to_int(val):
    if val is None:
        return None
    try:
        return int(float(val))
    except (TypeError, ValueError):
        return None

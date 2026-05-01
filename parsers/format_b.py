"""
Format B parser — Kap7 Futures WPL / Southern California.

Left-side columns (0–7):  GAME_ID, TIME, WHITE, W_SCORE, DARK, D_SCORE, DIVISION, NOTES
Right-side columns (9–16): same layout, used for Sunday games when Saturday and Sunday
                            are laid out side-by-side in the same rows.

Section header rows identify dates for the block below them.  The left half carries the
Saturday (or first-day) date; the right half carries the Sunday (or second-day) date.
"""
import re
from datetime import datetime, date, time as dtime, timedelta

from parsers.normalize import normalize_team_slot as _normalize_team

SKIP_SHEETS = {"InfoRules", "DivisionsStandings"}

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
# Section-header markers: SATURDAY / SUNDAY / DAY 1 / etc.
_DAY_HEADER_RE = re.compile(r'^(saturday|sunday|day\s*\d+)', re.IGNORECASE)


def parse(wb) -> list[dict]:
    games = []
    for sheet_name in wb.sheetnames:
        if sheet_name in SKIP_SHEETS:
            continue
        ws = wb[sheet_name]
        left_date  = None   # date for games in left columns (0–7)
        right_date = None   # date for games in right columns (9–16)

        for row in ws.iter_rows(min_row=1, values_only=True):
            if not any(c is not None for c in row):
                continue

            left_hdr  = str(row[0]).strip() if row[0] else ""
            right_hdr = str(row[9]).strip() if len(row) > 9 and row[9] else ""

            # Section header rows: contain day names (Saturday/Sunday) with date cells
            if _DAY_HEADER_RE.match(left_hdr):
                left_date  = _date_from_cell(row[1]) or _date_from_text(left_hdr) or left_date
                raw_right  = _date_from_cell(row[10] if len(row) > 10 else None)
                # If the right-side date is missing or implausible (wrong year / before left),
                # infer it as left_date + 1 day (Saturday → Sunday).
                if left_date and (raw_right is None or raw_right <= left_date):
                    right_date = left_date + timedelta(days=1)
                else:
                    right_date = raw_right or right_date
                continue

            # Also catch text-based date headers (legacy format)
            dm = _DATE_HEADER_RE.search(left_hdr)
            if dm and (len(row) < 4 or row[3] is None or row[4] is None):
                left_date = _parse_date_from_match(dm)
                continue

            # Parse left-side game (cols 0–7)
            g = _parse_game_cols(row, 0, left_date, sheet_name)
            if g:
                games.append(g)

            # Parse right-side game (cols 9–16)
            g2 = _parse_game_cols(row, 9, right_date, sheet_name)
            if g2:
                games.append(g2)

    return games


def _parse_game_cols(row, start: int, current_date, sheet_name: str):
    """Parse one game from `row` starting at column `start`."""
    if len(row) < start + 5:
        return None
    game_id  = row[start]
    time_val = row[start + 1]
    white    = row[start + 2]
    w_score  = row[start + 3]
    dark     = row[start + 4]
    d_score  = row[start + 5] if len(row) > start + 5 else None
    division = row[start + 6] if len(row) > start + 6 else None
    notes    = row[start + 7] if len(row) > start + 7 else None

    if not game_id or not isinstance(game_id, str):
        return None
    if not white or not dark:
        return None
    if str(white).strip().upper() in ("WHITE", "TEAM", "WHITE TEAM", "GAME #"):
        return None

    return {
        "date":        current_date,
        "time":        (time_val.time() if isinstance(time_val, datetime) else time_val)
                       if hasattr(time_val, "hour") else None,
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
    }



def _date_from_cell(val):
    """Extract a date from a cell value that may be a datetime object."""
    if val is None:
        return None
    if isinstance(val, datetime):
        # Reject obviously-wrong dates (e.g. Excel mis-parses "2/22" as Feb 22 of some year)
        # We trust dates only in 2025-2027 range.
        if 2025 <= val.year <= 2027:
            return val.date()
    if isinstance(val, date):
        if 2025 <= val.year <= 2027:
            return val
    return None


def _date_from_text(s: str):
    m = _DATE_HEADER_RE.search(s)
    return _parse_date_from_match(m) if m else None


def _parse_date_from_match(m):
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

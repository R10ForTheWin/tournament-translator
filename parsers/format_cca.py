"""
Format CCA parser — CCA JO Qualifications (Google Sheets CSV).

Fixed column layout (0-indexed):
  0: Date    1: Location    2: Time    3: Game#
  4: White Team    5: White Score    6: Dark Team    7: Dark Score    8: Comments

Slot formats handled:
  "D2 - Trojan Gold B"        → "D2-TROJAN GOLD B"   (single-letter pool slot)
  "BB1 - Trojan Cardinal A"   → "BB1-TROJAN CARDINAL A"  (bracket group + team)
  "Winner #10"                → "W#10"
  "Loser #58"                 → "L#58"
  "Winnner #17"               → "W#17"   (typo variant)
  "GGG1- Win #36- "           → "W#36"   (bracket group + win ref + trailing dash)
  "DDD1 - Lose #58 -"         → "L#58"   (bracket group + loss ref)
  "HH 1 - 2nd C"              → "2ndC-"  (bracket group + finish slot)
  "BBB1 - 2nd II"             → "2ndII-" (bracket group + multi-letter pool finish)
  "BB2 - 1st C"               → "1stC-"  (bracket group + finish ref)
  "AA2 - 1st D"               → "1stD-"  (bracket group + finish ref)
"""
import re
import csv
import io
from datetime import datetime

YEAR_HINT = 2026


def _parse_date(s: str):
    if not s:
        return None
    s = s.strip()
    # "26-May" or "29-May" (no year)
    m = re.match(r'^(\d{1,2})-([A-Za-z]{3,})$', s)
    if m:
        try:
            return datetime.strptime(f"{m.group(1)}-{m.group(2)}-{YEAR_HINT}", "%d-%b-%Y").date()
        except ValueError:
            pass
    # "5/24/26" or "5/24/2026"
    for fmt in ("%m/%d/%y", "%m/%d/%Y", "%Y-%m-%d"):
        try:
            d = datetime.strptime(s, fmt).date()
            if 2025 <= d.year <= 2027:
                return d
        except ValueError:
            pass
    return None


def _parse_time(s: str):
    if not s:
        return None
    s = s.strip().upper()
    for fmt in ("%I:%M %p", "%I:%M%p", "%H:%M"):
        try:
            return datetime.strptime(s, fmt).time()
        except ValueError:
            pass
    return None


def _to_int(s):
    if not s:
        return None
    try:
        return int(str(s).strip())
    except (ValueError, TypeError):
        return None


def _normalize_inner(s: str):
    """Parse a structured slot reference that has no bracket-group prefix.

    Returns normalized string, or None if not a recognized reference format
    (caller treats it as a plain team name).
    """
    s = s.strip()
    # Strip trailing " - " placeholder where team name will be filled in later
    s = re.sub(r'\s*[-–]\s*$', '', s).strip()
    if not s:
        return None

    # Win reference: "Winner #10", "Win #36", "Winnner #17" (triple-n typo)
    # Optionally followed by " - TEAM" for already-resolved slots
    m = re.match(r'^Win+e?r?\s*#?(\d+)(?:\s*[-–]\s*(.+))?$', s, re.IGNORECASE)
    if m:
        team = (m.group(2) or "").strip().upper()
        return f"W#{m.group(1)}-{team}" if team else f"W#{m.group(1)}"

    # Loss reference: "Loser #10", "Lose #58"
    m = re.match(r'^Los[se]*r?\s*#?(\d+)(?:\s*[-–]\s*(.+))?$', s, re.IGNORECASE)
    if m:
        team = (m.group(2) or "").strip().upper()
        return f"L#{m.group(1)}-{team}" if team else f"L#{m.group(1)}"

    # Finish slot: "1st C", "2nd D", "3rd GG", "4th II" (single or multi-letter pool)
    m = re.match(r'^(\d+)(st|nd|rd|th)\s+([A-Z]+)\s*$', s, re.IGNORECASE)
    if m:
        ordinal = m.group(1) + m.group(2).lower()
        pool = m.group(3).upper()
        return f"{ordinal}{pool}-"

    # Single-letter pool slot: "D2 - Trojan Gold B"
    m = re.match(r'^([A-Z]\d+)\s*[-–]\s*(.+)$', s, re.IGNORECASE)
    if m:
        pos = m.group(1).upper()
        team = m.group(2).strip().upper()
        return f"{pos}-{team}" if team else pos

    # Bare single-letter pool slot: "D2" (no team yet)
    m = re.match(r'^([A-Z]\d+)\s*$', s, re.IGNORECASE)
    if m:
        return m.group(1).upper()

    return None  # caller handles as plain team name


def _normalize_slot(raw: str) -> str:
    """Normalize a CCA JO Quals team slot to canonical form."""
    s = raw.strip()
    if not s or s.upper() in ("TBD", "BYE", "NONE"):
        return "TBD"

    # Bracket-group prefix requires 2+ consecutive uppercase letters (case-insensitive)
    # followed by optional space, one or more digits, then a dash.
    # Examples: "BB1 - Team", "GGG1- Win #36- ", "HH 1 - 2nd C", "AA2 - 1st D"
    bg = re.match(r'^([A-Z]{2,})\s*(\d+)\s*[-–]\s*(.+)', s, re.IGNORECASE)
    if bg:
        inner_raw = bg.group(3).strip()
        inner_norm = _normalize_inner(inner_raw)
        if inner_norm is not None:
            # Inner was a W#/L# ref or finish slot — strip bracket-group prefix
            return inner_norm
        # Plain team name inside bracket group — keep bracket position for context
        pos = (bg.group(1) + bg.group(2)).upper()
        team = inner_raw.upper()
        return f"{pos}-{team}" if team else pos

    # No multi-letter bracket prefix — try direct normalization
    return _normalize_inner(s) or s.upper()


def _compute_advancement(games: list) -> None:
    """Fill w_to/l_to by scanning W#N/L#N slot references across all games.

    For each game G that has a slot "W#N", set game N's w_to = G's game number.
    For each game G that has a slot "L#N", set game N's l_to = G's game number.
    This lets _build_njo_game_tree traverse the bracket via w_to/l_to links.
    """
    gnum_map: dict[str, dict] = {}
    for g in games:
        m = re.search(r'-(\d+)$', g["game_id"])
        if m:
            gnum_map[str(int(m.group(1)))] = g

    for g in games:
        this_m = re.search(r'-(\d+)$', g["game_id"])
        if not this_m:
            continue
        this_num = str(int(this_m.group(1)))

        for slot in (g["white_team"], g["dark_team"]):
            m = re.match(r'^([WL])#(\d+)', slot)
            if not m:
                continue
            wl = m.group(1).upper()
            ref_num = str(int(m.group(2)))
            src = gnum_map.get(ref_num)
            if src is None:
                continue
            if wl == "W" and src.get("w_to") is None:
                src["w_to"] = int(this_num)
            elif wl == "L" and src.get("l_to") is None:
                src["l_to"] = int(this_num)


def parse_csv(csv_text: str, division: str, id_prefix: str) -> list[dict]:
    """Parse a CCA JO Quals CSV tab into game dicts.

    Column layout (0-indexed, fixed across all venue blocks):
      0=Date, 1=Location, 2=Time, 3=Game#,
      4=White Team, 5=White Score, 6=Dark Team, 7=Dark Score, 8=Comments
    """
    reader = csv.reader(io.StringIO(csv_text))
    rows = list(reader)

    CI_DATE = 0; CI_LOC = 1; CI_TIME = 2; CI_GNUM = 3
    CI_WHITE = 4; CI_WS = 5; CI_DARK = 6; CI_DS = 7; CI_COMM = 8

    games: list[dict] = []
    seen_ids: set[str] = set()
    current_date = None
    current_loc = "TBD"

    for row in rows:
        row = list(row) + [""] * max(0, CI_COMM + 1 - len(row))

        def _get(ci, _row=row):
            v = _row[ci].strip() if ci < len(_row) else ""
            return v if v else None

        # Game number is required and must be a plain integer
        gnum_raw = _get(CI_GNUM)
        if not gnum_raw or not re.match(r'^\d+$', gnum_raw):
            continue

        game_id = f"{id_prefix}-{int(gnum_raw)}"
        if game_id in seen_ids:
            continue
        seen_ids.add(game_id)

        # Date carry-forward (game rows have their date in col 0)
        d = _parse_date(_get(CI_DATE))
        if d:
            current_date = d

        # Location carry-forward (skip "LOCATION" header text)
        loc_raw = _get(CI_LOC)
        if loc_raw and loc_raw.upper() != "LOCATION":
            current_loc = loc_raw

        white_raw = _get(CI_WHITE) or ""
        dark_raw  = _get(CI_DARK)  or ""
        if not white_raw and not dark_raw:
            continue

        white_s = _normalize_slot(white_raw) if white_raw else "TBD"
        dark_s  = _normalize_slot(dark_raw)  if dark_raw else "TBD"

        white_score = _to_int(_get(CI_WS))
        dark_score  = _to_int(_get(CI_DS))
        time_val    = _parse_time(_get(CI_TIME))
        comments    = _get(CI_COMM) or ""

        games.append({
            "date":        current_date,
            "time":        time_val,
            "location":    current_loc,
            "game_id":     game_id,
            "white_team":  white_s,
            "white_score": white_score,
            "dark_team":   dark_s,
            "dark_score":  dark_score,
            "comments":    comments,
            "division":    division,
            "sheet":       division,
            "played":      white_score is not None and dark_score is not None,
            "w_to":        None,
            "l_to":        None,
        })

    _compute_advancement(games)
    return games

"""
Format C parser — National Junior Olympics / JO Qualifications.

Each division has its own sheet. Within a sheet there are one or more
venue blocks, each preceded by a header row:
  Date | Time | Type | Location | Gm # | White | S | Dark | S | W to # | L to # | GMID

GMID (e.g. "16B-003", "18BQ-005") is the canonical game identifier.
Team slots use several formats:
  A2(8)-TEAM       pool slot with national seed
  8-TEAM           seed-number slot
  W11-TEAM         winner of game 11 (team filled in)
  L31-TEAM         loser of game 31 (team filled in)
  2ndA-TEAM        finish slot (2nd in pool A)
  1stA-            finish slot, team TBD
  W3-              winner ref, team TBD
  G1(L16)-TEAM     composite: pool G1 who is loser of game 16

Sheets to skip are those without the GMID column (master/reference sheets).
"""
import re
from datetime import datetime, date

from parsers.normalize import normalize_team_slot as _normalize_team

def _prettify_division(sheet_name: str) -> str:
    """Convert a raw NJO sheet name to a readable division label.

    '18U_F_Champ 47'       → '18U Girls Championship'
    '16U_M_CHAMP-41 teams' → '16U Boys Championship'
    '10U_C_Classic 24'     → '10U Co-Ed Classic'
    '18U_M_Invite 24'      → '18U Boys Invite'
    Returns the original name unchanged if pattern doesn't match.
    """
    s = sheet_name.strip()
    age_m = re.match(r'(\d+)U', s, re.IGNORECASE)
    if not age_m:
        return sheet_name
    age = age_m.group(1) + 'U'
    gen_m = re.search(r'_([MFC])_', s, re.IGNORECASE)
    gender = {'M': 'Boys', 'F': 'Girls', 'C': 'Co-Ed'}.get(
        gen_m.group(1).upper(), '') if gen_m else ''
    if   re.search(r'champ',   s, re.IGNORECASE): event = 'Championship'
    elif re.search(r'classic', s, re.IGNORECASE): event = 'Classic'
    elif re.search(r'invite',  s, re.IGNORECASE): event = 'Invite'
    else:                                          event = ''
    return ' '.join(p for p in [age, gender, event] if p)


SKIP_SHEETS = {
    "Reference",
    "MASTER BY DIVISION",
    "MASTER BY TIME",
    "MASTER BY LOCATION",
    "MASTER S2 CHICLETS",
}

_HEADER_COLS = {"gm #", "gmid"}   # must both be present in a header row


def _is_header_row(row) -> bool:
    vals = {str(c).strip().lower() for c in row if c is not None}
    return _HEADER_COLS.issubset(vals)


def _col_index(row, *names):
    """Return 0-based index of the first matching column name (case-insensitive)."""
    for name in names:
        for i, c in enumerate(row):
            if c is not None and str(c).strip().lower() == name.lower():
                return i
    return None


def _to_int(val):
    if val is None:
        return None
    s = str(val).strip()
    # Strip trailing non-numeric chars (e.g. "7.1" for OT wins → 7)
    m = re.match(r'^-?\d+', s)
    if m:
        return int(m.group())
    try:
        return int(float(val))
    except (TypeError, ValueError):
        return None


def _parse_date(val):
    if val is None:
        return None
    if isinstance(val, datetime):
        if 2024 <= val.year <= 2027:
            return val.date()
    if isinstance(val, date):
        if 2024 <= val.year <= 2027:
            return val
    return None


def _parse_time(val):
    if val is None:
        return None
    if isinstance(val, datetime):
        return val.time()
    if hasattr(val, 'hour'):
        return val
    return None


def _normalize_slot(raw: str) -> str:
    """Normalize a JO/NJO team slot to a canonical form.

    Examples:
      'A2(8)-SAN CLEMENTE BLACK'  → 'A2-SAN CLEMENTE BLACK'
      '8-DYNAMO'                  → '8-DYNAMO'
      'W11-DYNAMO'                → 'W#11-DYNAMO'
      'L31-TEAM'                  → 'L#31-TEAM'
      '2ndA-TEAM'                 → '2ndA-TEAM'
      'G1(L16)-TEAM'              → 'G1-TEAM'
      'W3-'                       → 'W#3'
      '1stA-'                     → '1stA-'
    """
    s = raw.strip()
    if not s or s.upper() in ("NONE", "TBD", "BYE"):
        return s.upper()

    # Win/loss game reference:  W11-TEAM  or  L31-TEAM
    m = re.match(r'^([WL])(\d+)-(.*)', s, re.IGNORECASE)
    if m:
        wl = m.group(1).upper()
        num = m.group(2)
        team = m.group(3).strip()
        suffix = f"-{team.upper()}" if team else ""
        return f"{wl}#{num}{suffix}"

    # Pool slot with national seed:  A2(8)-SAN CLEMENTE  or  G1(L16)-TEAM
    m = re.match(r'^([A-Z]\d+)\([^)]*\)-(.*)', s, re.IGNORECASE)
    if m:
        pos  = m.group(1).upper()
        team = m.group(2).strip().upper()
        return f"{pos}-{team}" if team else pos

    # NJO bracket advancement slots with underscore separators:
    #   2ND_A-PEGASUS    → 2ndA-PEGASUS   (finish slot, pool A)
    #   1ST_PT_M-CAPITAL → 1stM-CAPITAL   (bracket stage PT, pool M)
    #   3RD_AU_O-TEAM    → 3rdO-TEAM      (bracket stage AU, pool O)
    m = re.match(r'^(\d+(?:ST|ND|RD|TH))_(?:[A-Z]+_)?([A-Z])-(.+)', s, re.IGNORECASE)
    if m:
        ord_part  = m.group(1).capitalize()
        pool_lett = m.group(2).upper()
        team      = m.group(3).strip().upper()
        return f"{ord_part}{pool_lett}-{team}"

    # Stage-prefixed pool slots:  AU_M1-ELMHURST  pt_M2-CAPITAL  ni_C3-TEAM
    m = re.match(r'^[A-Za-z]+_([A-Z]\d+)-(.+)', s, re.IGNORECASE)
    if m:
        pos  = m.group(1).upper()
        team = m.group(2).strip().upper()
        return f"{pos}-{team}"

    # Standard pool slot (already handled by normalize_team_slot):  A2-TEAM
    # Finish slot:  2ndA-TEAM  1stA-  3rd_G-TEAM
    # Seed number:  8-DYNAMO  or  8-  (no team yet)
    # Fall through to normalize_team_slot for everything else
    return _normalize_team(s)


def parse(wb) -> list[dict]:
    games: list[dict] = []
    seen_gmids: set[str] = set()

    for sheet_name in wb.sheetnames:
        if sheet_name in SKIP_SHEETS:
            continue

        ws = wb[sheet_name]
        all_rows = list(ws.iter_rows(min_row=1, values_only=True))

        # Find all header rows — a sheet may have multiple venue blocks
        header_positions = []
        for i, row in enumerate(all_rows):
            if _is_header_row(row):
                header_positions.append(i)

        if not header_positions:
            continue  # no game schedule in this sheet

        for h_idx in header_positions:
            header_row = all_rows[h_idx]
            # Column indices
            ci_date  = _col_index(header_row, "date")
            ci_time  = _col_index(header_row, "time")
            ci_type  = _col_index(header_row, "type")
            ci_loc   = _col_index(header_row, "location")
            ci_gnum  = _col_index(header_row, "gm #")
            ci_white = _col_index(header_row, "white")
            ci_ws    = _col_index(header_row, "s")          # first S after white
            ci_dark  = _col_index(header_row, "dark")
            ci_ds    = None                                  # second S (after dark)
            ci_gmid  = _col_index(header_row, "gmid")

            # Second score column is the S right after the dark column
            if ci_dark is not None:
                for j in range(ci_dark + 1, len(header_row)):
                    if header_row[j] is not None and str(header_row[j]).strip().upper() == "S":
                        ci_ds = j
                        break

            if any(c is None for c in [ci_date, ci_gnum, ci_white, ci_dark, ci_gmid]):
                continue  # incomplete header, skip block

            # Determine end of this block (next header or end of sheet)
            next_h = header_positions[header_positions.index(h_idx) + 1] \
                     if h_idx != header_positions[-1] else len(all_rows)

            # Parse game rows within this block
            current_date = None
            current_loc  = "TBD"

            for row in all_rows[h_idx + 1:next_h]:
                if not any(c is not None for c in row):
                    continue

                # Date carry-forward
                d = _parse_date(row[ci_date]) if ci_date < len(row) else None
                if d:
                    current_date = d

                # Location carry-forward
                loc_val = row[ci_loc] if ci_loc is not None and ci_loc < len(row) else None
                if loc_val and str(loc_val).strip() and str(loc_val).strip() != "0":
                    current_loc = str(loc_val).strip()

                # GMID — required; skip non-game rows
                gmid_raw = row[ci_gmid] if ci_gmid < len(row) else None
                if not gmid_raw:
                    continue
                gmid = str(gmid_raw).strip()
                if not gmid or gmid.upper() in ("GMID", "NONE"):
                    continue
                # Must look like a real game ID (letters-digits-digits)
                if not re.search(r'\d', gmid):
                    continue
                if gmid in seen_gmids:
                    continue
                seen_gmids.add(gmid)

                # White / dark team
                white_raw = row[ci_white] if ci_white < len(row) else None
                dark_raw  = row[ci_dark]  if ci_dark  < len(row) else None
                if not white_raw or not dark_raw:
                    continue
                white_s = str(white_raw).strip()
                dark_s  = str(dark_raw).strip()
                if not white_s or not dark_s:
                    continue
                # Skip column-header lookalike rows
                if white_s.upper() in ("WHITE", "TEAM", "W"):
                    continue

                # Skip template rows where no team has been assigned yet
                # (e.g. "11-" = seed #11, TBD; these appear in blank schedule templates)
                if re.match(r'^\d+-?$', white_s) or re.match(r'^\d+-?$', dark_s):
                    continue

                white_score = _to_int(row[ci_ws] if ci_ws is not None and ci_ws < len(row) else None)
                dark_score  = _to_int(row[ci_ds] if ci_ds  is not None and ci_ds  < len(row) else None)

                time_val = _parse_time(row[ci_time] if ci_time is not None and ci_time < len(row) else None)
                game_type = str(row[ci_type]).strip() if ci_type is not None and ci_type < len(row) and row[ci_type] else ""

                games.append({
                    "date":        current_date,
                    "time":        time_val,
                    "location":    current_loc,
                    "game_id":     gmid,
                    "white_team":  _normalize_slot(white_s),
                    "white_score": white_score,
                    "dark_team":   _normalize_slot(dark_s),
                    "dark_score":  dark_score,
                    "comments":    game_type,
                    "division":    _prettify_division(sheet_name),
                    "sheet":       sheet_name,
                    "played":      white_score is not None and dark_score is not None,
                })

    return games

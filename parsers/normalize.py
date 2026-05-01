"""
Shared slot-normalization for water polo tournament parsers.

Converts raw team-name strings from Excel into the canonical slot format
that app.py's bracket-expansion logic expects:

  Pool slot:    'B1 - Trojan Gold'   -> 'B1-TROJAN GOLD'
  Finish slot:  '1st in A - Team'    -> '1stA-TEAM'
                '1st in E - '        -> '1stE-'   (TBD / empty team name OK)
  Win ref:      'Win Gm #330 - '     -> 'W#330'
  Loss ref:     'Los Gm #330 - '     -> 'L#330'
  Anything else -> uppercased as-is
"""

import re

_WIN_GM_RE = re.compile(r'^(win|w)\s*(?:gm|game)?\s*#?(\d+)', re.IGNORECASE)
_LOS_GM_RE = re.compile(r'^(los|l(?:os)?|lose?)\s*(?:gm|game)?\s*#?(\d+)', re.IGNORECASE)


def normalize_team_slot(s: str) -> str:
    """Return the canonical slot string for a raw team-name cell value."""
    # Win/loss game-reference slots
    wm = _WIN_GM_RE.match(s)
    if wm:
        return f"W#{wm.group(2)}"
    lm = _LOS_GM_RE.match(s)
    if lm:
        return f"L#{lm.group(2)}"

    # Pool slot: "B1 - Trojan Gold"
    m = re.match(r"^([A-Z]\d+)\s*-\s*(.*)$", s, re.IGNORECASE)
    if m:
        return f"{m.group(1).upper()}-{m.group(2).strip().upper()}"

    # Finish slot: "1st in A - Team" or "1st in E - " (empty team)
    m2 = re.match(r"^(\d+(?:st|nd|rd|th))\s+in\s+([A-Z])\s*-\s*(.*)$", s, re.IGNORECASE)
    if m2:
        return f"{m2.group(1)}{m2.group(2).upper()}-{m2.group(3).strip().upper()}"

    return s.upper()

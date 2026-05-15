"""
Deterministic validator for parsed game lists.

Runs between the parser and the bracket tree builder to catch bad output
before it reaches parents. Checks that are free, fast, and rules-based —
no LLM required for these invariants.

Returns a list of error strings. Empty list = valid.
"""
import re

_POOL_SLOT_RE   = re.compile(r'^[A-Z]\d+-', re.IGNORECASE)
_WL_SLOT_RE     = re.compile(r'^[WL]#',    re.IGNORECASE)
_FINISH_SLOT_RE = re.compile(r'^\d+(?:st|nd|rd|th)[A-Z]-', re.IGNORECASE)


def _valid_slot(slot: str) -> bool:
    """Return True if the value looks like a real team name or a valid bracket slot.

    Accepts:
      - Pool slots:   E3-TROJAN GOLD
      - W#/L# refs:   W#330, L#318
      - Finish slots: 1stE-, 4thF-
      - Plain names:  any string with at least one letter and length >= 2

    Rejects pure numbers, formulas, #ERROR, empty strings, and other garbage
    that an AI parser might hallucinate into a team column.
    """
    s = slot.strip()
    if not s or len(s) < 2:
        return False
    if _POOL_SLOT_RE.match(s) or _WL_SLOT_RE.match(s) or _FINISH_SLOT_RE.match(s):
        return True
    return bool(re.search(r'[A-Za-z]', s))


def validate_games(games: list) -> list:
    """Return a list of error strings. Empty list = all checks passed."""
    if not games:
        return ["No games parsed"]

    errors = []
    seen_ids: dict = {}

    for g in games:
        gid = str(g.get("game_id", "?"))

        # Duplicate game IDs — scoped to (game_id, sheet) so that WPL age-group
        # sheets which share a numbering sequence don't collide with each other.
        dup_key = (gid, g.get("sheet", ""))
        if dup_key in seen_ids:
            errors.append(f"Duplicate game_id '{gid}'")
        else:
            seen_ids[dup_key] = True

        # Both team slots present and structurally valid
        for field in ("white_team", "dark_team"):
            slot = str(g.get(field) or "").strip()
            if not slot:
                errors.append(f"Game {gid}: empty {field}")
            elif not _valid_slot(slot):
                errors.append(f"Game {gid}: suspicious {field} value {slot!r}")

        # Scores must be numeric or absent
        for field in ("white_score", "dark_score"):
            v = g.get(field)
            if v is not None and not isinstance(v, (int, float)):
                errors.append(f"Game {gid}: {field} is not numeric ({v!r})")

    return errors

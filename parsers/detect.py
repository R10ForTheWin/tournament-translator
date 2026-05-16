"""
Excel parser dispatcher for tournament schedules.

Parse priority (revised):
  1. AI-cached schema (disk or memory) — instant, primary path for any format
     seen before. No API call. Covers all novel formats after first encounter.
  2. Format A heuristic — fast-path for Kap7/Kahuna tournaments.
  3. Format B heuristic — fast-path for WPL Futures weekends.
  4. AI parser — makes one API call for truly new layouts, caches result.

Every parse attempt is gated through the deterministic validator. If a parser
returns games that fail validation, it is rejected and the next parser is tried.
This prevents silently wrong brackets from reaching parents.
"""
import io
import openpyxl
from parsers.validate import validate_games


def _load_wb(filepath_or_bytes):
    if isinstance(filepath_or_bytes, (bytes, bytearray)):
        return openpyxl.load_workbook(io.BytesIO(filepath_or_bytes), data_only=True)
    return openpyxl.load_workbook(filepath_or_bytes, data_only=True)


def _tag_and_validate(games: list, fmt: str) -> list:
    """Tag games with their format and run the validator.
    Returns tagged games on success, empty list on validation failure."""
    if not games:
        return []
    issues = validate_games(games)
    if issues:
        print(f"[validate] {fmt} parse rejected — {issues[:3]}")
        return []
    for g in games:
        g["format"] = fmt
    return games


def load_and_parse(filepath_or_bytes) -> list[dict]:
    wb = _load_wb(filepath_or_bytes)

    # 1. AI-cached schema — primary path for any previously seen layout.
    try:
        from parsers.format_ai import parse_cached
        games = _tag_and_validate(parse_cached(wb), "AI-cached")
        if games:
            return games
    except Exception as exc:
        print(f"[AI-cached] error: {exc}")

    # 2. Format A heuristic — Kap7/Kahuna (DATE col + LOCATION col).
    try:
        from parsers.format_a import parse as parse_a
        games = _tag_and_validate(parse_a(wb), "A")
        if games:
            return games
    except Exception as exc:
        print(f"[format_a] error: {exc}")

    # 3. Format B heuristic — WPL Futures (section-header dates, dual-column layout).
    try:
        from parsers.format_b import parse as parse_b
        games = _tag_and_validate(parse_b(wb), "B")
        if games:
            return games
    except Exception as exc:
        print(f"[format_b] error: {exc}")

    # 3b. Format C heuristic — NJO/JO Quals (GMID column, per-division sheets).
    try:
        from parsers.format_c import parse as parse_c
        games = _tag_and_validate(parse_c(wb), "C")
        if games:
            return games
    except Exception as exc:
        print(f"[format_c] error: {exc}")

    # 4. AI parser — one API call for unknown layouts, result cached to disk.
    try:
        from parsers.format_ai import parse as parse_ai
        games = _tag_and_validate(parse_ai(wb), "AI")
        if games:
            return games
    except Exception as exc:
        print(f"[AI parser] failed: {exc}")

    return []

"""
Auto-detect which parser format an Excel workbook uses.
Format A: has a DATE column (datetime objects) + LOCATION column
Format B: no DATE column, time objects only, Futures/WPL style
"""
import openpyxl
from datetime import datetime

SKIP = {"InfoRules", "DivisionsStandings", "TEAM LISTING AND BRACKETS",
        "CHICLETS", "CHICLETS MASTER", "CHICLETS NEW",
        "MASTER BY DIVISION", "MASTER BY LOCATION", "MASTER BY TIME",
        "RAW ENTRIES", "MEDALS"}


def detect_format(wb) -> str:
    """Return 'A' or 'B'."""
    for sheet_name in wb.sheetnames:
        if sheet_name.upper() in {s.upper() for s in SKIP}:
            continue
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=1, max_row=60, values_only=True):
            if len(row) < 4:
                continue
            if isinstance(row[0], datetime) and row[3] and isinstance(row[3], str):
                return "A"
    return "B"


def load_and_parse(filepath: str) -> list[dict]:
    wb = openpyxl.load_workbook(filepath, data_only=True)

    # Heuristic parsers first — fast, no API calls, covers all known formats.
    from parsers.format_a import parse as parse_a
    games = parse_a(wb)
    if games:
        for g in games:
            g["format"] = "A"
        return games

    from parsers.format_b import parse as parse_b
    games = parse_b(wb)
    if games:
        for g in games:
            g["format"] = "B"
        return games

    # AI parser as last resort — handles novel/unknown layouts via Claude Haiku.
    # Only runs when both heuristics return zero games.
    try:
        from parsers.format_ai import parse as parse_ai
        games = parse_ai(wb)
        if games:
            for g in games:
                g["format"] = "AI"
            return games
    except Exception as exc:
        print(f"[AI parser] failed: {exc}")

    return []

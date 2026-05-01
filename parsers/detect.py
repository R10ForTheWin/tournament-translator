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


def load_and_parse(filepath_or_bytes) -> list[dict]:
    """Load and parse an Excel workbook from a file path, bytes, or BytesIO object.

    Parse priority:
      1. Cached AI schema (disk or memory) — instant, no API call.
      2. Heuristic parsers (Format A, then Format B) — fast, no API call.
      3. AI parser (makes one API call, caches result for next time).
    """
    import io

    if isinstance(filepath_or_bytes, (bytes, bytearray)):
        wb = openpyxl.load_workbook(io.BytesIO(filepath_or_bytes), data_only=True)
    elif hasattr(filepath_or_bytes, "read"):
        wb = openpyxl.load_workbook(filepath_or_bytes, data_only=True)
    else:
        wb = openpyxl.load_workbook(filepath_or_bytes, data_only=True)

    # 1. Cached AI schema — instant (no API call).
    #    Returns [] if any sheet is unknown, so we fall through safely.
    try:
        from parsers.format_ai import parse_cached
        games = parse_cached(wb)
        if games:
            for g in games:
                g["format"] = "AI-cached"
            return games
    except Exception:
        pass

    # 2. Heuristic parsers — cover all previously known formats.
    from parsers.format_a import parse as parse_a
    games = parse_a(wb)
    if games:
        for g in games:
            g["format"] = "A"
        _queue_schema_cache(wb)
        return games

    from parsers.format_b import parse as parse_b
    games = parse_b(wb)
    if games:
        for g in games:
            g["format"] = "B"
        _queue_schema_cache(wb)
        return games

    # 3. AI parser — makes one API call for unknown layouts, caches result.
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


def _queue_schema_cache(wb):
    """After a heuristic parse succeeds, build the AI schema cache in the background
    if any sheet doesn't have one yet. This runs once per new file layout."""
    import threading
    try:
        from parsers.format_ai import _sample_rows, _hash_rows, _load_disk_schema, SKIP_SHEETS
        import openpyxl
        needs_cache = False
        for sheet_name in wb.sheetnames:
            if sheet_name.lower().strip() in SKIP_SHEETS:
                continue
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(max_row=60, values_only=True))
            sample = _sample_rows(rows)
            key = _hash_rows(sample)
            if not _load_disk_schema(key):
                needs_cache = True
                break
        if not needs_cache:
            return

        # Re-open the workbook bytes for the background thread (wb may not be thread-safe)
        import io, openpyxl as ox
        buf = io.BytesIO()
        # We can't re-serialize wb easily, so just flag that caching is needed.
        # The actual cache build happens via the /admin/cache-schema endpoint.
        print("[schema cache] New layout detected — call /admin/cache-schema to build AI cache.")
    except Exception:
        pass

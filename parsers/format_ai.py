"""
AI-assisted parser using Claude Haiku to handle non-standard Excel layouts.
Called when Format A and Format B both return zero games.

Claude analyzes a representative sample of rows from each sheet to identify
column structure, section header patterns, and date placement, then we apply
that schema to parse the full sheet programmatically.

Handles:
  - Standard per-row dates (Format A style)
  - Date-in-section-header rows (Format B style, text-based)
  - Date-in-section-header rows where the date is a datetime cell (Futures WPL)
  - Dual-column Saturday+Sunday layouts (Futures WPL Google Sheet)
"""

from __future__ import annotations

import json
import hashlib
import os
import re
import anthropic
from datetime import datetime, date as date_type, timedelta
from typing import Optional

from parsers.normalize import normalize_team_slot

# ---------------------------------------------------------------------------
# In-memory schema cache — keyed by hash of the sampled rows.
# Prevents redundant API calls when the same sheet is analyzed more than once.
# ---------------------------------------------------------------------------
_client: Optional[anthropic.Anthropic] = None
_schema_cache: dict = {}  # type: ignore[type-arg]

# Disk cache directory — survives server restarts.
# Lives at <project_root>/data/schema_cache/
_CACHE_DIR = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "data", "schema_cache",
)

SKIP_SHEETS = {
    "inforules", "divisionsstandings", "team listing and brackets",
    "chiclets", "chiclets master", "chiclets new",
    "master by division", "master by location", "master by time",
    "raw entries", "medals", "contact info", "emails",
}

# ---------------------------------------------------------------------------
# System prompt
# ---------------------------------------------------------------------------
_SYSTEM_PROMPT = """\
You are an expert at parsing water polo tournament schedule spreadsheets exported from Excel or Google Sheets.

Your job: analyze a sample of rows from one sheet and return a JSON schema that describes how to extract game records.

=== OUTPUT FORMAT ===
Return ONLY valid JSON (no markdown, no explanation):
{
  "skip": false,

  "date_in_row": true,
  "date_col": 0,

  "section_header_pattern": null,
  "section_date_col": null,
  "section_location_col": null,

  "time_col": 1,
  "location_col": 2,
  "game_id_col": 3,
  "white_col": 4,
  "white_score_col": 5,
  "dark_col": 6,
  "dark_score_col": 7,
  "division_col": null,

  "alt_game_id_col": null,
  "alt_time_col": null,
  "alt_white_col": null,
  "alt_white_score_col": null,
  "alt_dark_col": null,
  "alt_dark_score_col": null,
  "alt_division_col": null,
  "alt_date_is_next_day": false,

  "header_rows": [0],
  "default_division": "Sheet1"
}

=== FIELD RULES ===
skip                  — true if this sheet has no game rows (info/rules/standings/images).

Date placement — choose ONE of:
  date_in_row=true    — every game row has its own date cell. Set date_col to its 0-based index.
  date_in_row=false   — dates appear only in section-header rows between game blocks.
                        Set section_header_pattern (Python regex, re.IGNORECASE) to match those rows,
                        e.g. "^SATURDAY$|^SUNDAY$" or "Weekend \\d+".

section_date_col      — if date_in_row=false AND the date is stored as a cell value (datetime or
                        date object) in the section-header row, set this to its column index.
                        If the date must be parsed from the text of col 0, leave null.
section_location_col  — column in the section-header row that holds the venue/pool name, or null.

time_col .. dark_score_col — 0-based column indices for the PRIMARY (Saturday / only) game block.
division_col          — 0-based index for division label, or null.

alt_* columns         — some sheets place a SECOND game block (Sunday games) in the same rows,
                        starting at a higher column offset. Set alt_game_id_col etc. for that block.
                        Leave all null if there is only one game block per row.
alt_date_is_next_day  — true if alt-block games happen on the day after the section-header date
                        (typical for Saturday+Sunday side-by-side layouts).

header_rows           — list of 0-based row indices that are column-label rows to skip (e.g. [9]).
default_division      — fallback label when division_col is null (usually the sheet name).

=== DOMAIN KNOWLEDGE ===
• Divisions: Platinum, Gold, Silver, Bronze, or age codes (B18, G14, "1.0", "1--10").
• Team names often include seeds: "A1 - Newport Beach", "3 - La Jolla United",
  "C2 (Win Gm #336) - AETOS", "1st in A - San Diego Dons". All are valid team names.
• Bracket slots: "W#3" / "Win Gm #3" = winner of game 3. "L#7" / "Los Gm #7" = loser of game 7.
• Scores of 0 are valid; null/empty = game not yet played.
• Game IDs may look like "18UB 336", "G1", "101", "10Cag01".
• Section-header rows often look like: SATURDAY | datetime | Host: | City | Location: | Venue
  or: "Weekend 4 - April 11-12, 2026" as a single merged cell.
• Some sheets have BOTH Saturday games (cols 0–7) and Sunday games (cols 9–16) in the same rows.
  The Sunday date is usually the day after the Saturday date.
• Skip-worthy sheet names: InfoRules, DivisionsStandings, CHICLETS, MASTER BY DIVISION, etc.
• The row format in the sample uses \" | \" as a separator; empty cells show as empty string.
"""


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _get_client() -> "anthropic.Anthropic":
    global _client
    if _client is None:
        _client = anthropic.Anthropic()
    return _client


def _rows_to_text(rows: list) -> str:
    lines = []
    for i, row in enumerate(rows):
        cells = [str(c) if c is not None else "" for c in row]
        lines.append(f"R{i}: {' | '.join(cells)}")
    return "\n".join(lines)


def _hash_rows(rows: list) -> str:
    return hashlib.sha256(repr(rows).encode()).hexdigest()[:20]


def _hash_schema_key(sheet_name: str, all_rows: list) -> str:
    """Stable cache key based on sheet name + first non-empty row (structural signature only).

    Hashing row data caused cache misses whenever scores were updated or teams
    were added — even though the column layout was identical. The first non-empty
    row is the structural fingerprint; data rows are irrelevant to the schema.
    """
    header_row = None
    for row in all_rows[:15]:
        cells = tuple(str(c).strip() if c is not None else "" for c in row)
        if any(cells):
            header_row = cells
            break
    key_data = repr((sheet_name.strip(), header_row))
    return hashlib.sha256(key_data.encode()).hexdigest()[:20]


_MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
}
_DATE_RE = re.compile(
    r"(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
    r"Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
    r"\s+(\d{1,2})(?:,?\s*(\d{4}))?",
    re.IGNORECASE,
)


def _extract_date_from_text(s: str) -> Optional[date_type]:
    m = _DATE_RE.search(s)
    if not m:
        return None
    month = _MONTH_MAP.get(m.group(1)[:3].lower())
    if not month:
        return None
    day = int(m.group(2))
    year = int(m.group(3)) if m.group(3) else datetime.now().year
    try:
        return date_type(year, month, day)
    except ValueError:
        return None


def _coerce_date(val) -> Optional[date_type]:
    """Convert a cell value to a date, however it's stored."""
    if isinstance(val, datetime):
        return val.date()
    if isinstance(val, date_type):
        return val
    if isinstance(val, str) and val.strip():
        return _extract_date_from_text(val)
    return None


def _safe_get(row: tuple, idx: Optional[int]):
    if idx is None or not (0 <= idx < len(row)):
        return None
    return row[idx]


def _to_int(val) -> Optional[int]:
    if val is None:
        return None
    try:
        return int(float(val))
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------------------
# Smart row sampling — send first 5 rows + rows starting where game data begins
# ---------------------------------------------------------------------------

_GAME_SECTION_STARTS = re.compile(
    r"^(saturday|sunday|game\s*#?|date|weekend\s+\d)", re.IGNORECASE
)


def _sample_rows(all_rows: list, sample_size: int = 45) -> list:
    """Return a representative sample: preamble + the first game section."""
    # Find where the first game section or header row appears
    start = 0
    for i, row in enumerate(all_rows):
        first = str(row[0]).strip() if row[0] else ""
        if _GAME_SECTION_STARTS.match(first):
            start = max(0, i - 1)
            break

    # Always include the very first few rows for context
    head = all_rows[:min(5, start)]
    body = all_rows[start: start + sample_size]

    # Deduplicate while preserving order
    seen = set()
    result = []
    for row in head + body:
        key = id(row)
        if key not in seen:
            seen.add(key)
            result.append(row)
    return result[:sample_size]


# ---------------------------------------------------------------------------
# Disk-cache helpers
# ---------------------------------------------------------------------------

def _disk_cache_path(cache_key: str) -> str:
    return os.path.join(_CACHE_DIR, f"{cache_key}.json")


def _load_disk_schema(cache_key: str):
    """Return cached schema dict from disk, or None if not present / unreadable."""
    path = _disk_cache_path(cache_key)
    if os.path.exists(path):
        try:
            with open(path) as f:
                return json.load(f)
        except Exception:
            pass
    return None


def _save_disk_schema(cache_key: str, schema: dict) -> None:
    os.makedirs(_CACHE_DIR, exist_ok=True)
    path = _disk_cache_path(cache_key)
    with open(path, "w") as f:
        json.dump(schema, f, indent=2)


# ---------------------------------------------------------------------------
# Schema detection via Claude
# ---------------------------------------------------------------------------

def _get_schema(sheet_name: str, all_rows: list) -> dict:
    sample = _sample_rows(all_rows)
    cache_key = _hash_schema_key(sheet_name, all_rows)

    # 1. Memory cache (fastest)
    if cache_key in _schema_cache:
        return _schema_cache[cache_key]

    # 2. Disk cache (survives restarts — no API call)
    schema = _load_disk_schema(cache_key)
    if schema is not None:
        _schema_cache[cache_key] = schema
        return schema

    # 3. API call — only for truly new/unseen layouts
    rows_text = _rows_to_text(sample)
    client = _get_client()

    response = client.messages.create(
        model="claude-haiku-4-5-20251001",
        max_tokens=1024,
        system=[{
            "type": "text",
            "text": _SYSTEM_PROMPT,
            "cache_control": {"type": "ephemeral"},
        }],
        messages=[{
            "role": "user",
            "content": (
                f"Sheet name: {sheet_name}\n\n"
                f"Sample rows (R0 = first sampled row, indices relative to sample):\n"
                f"{rows_text}\n\n"
                "Return ONLY valid JSON."
            ),
        }],
    )

    text = response.content[0].text.strip()
    if "```" in text:
        for chunk in text.split("```"):
            chunk = chunk.strip()
            if chunk.startswith("json"):
                chunk = chunk[4:].strip()
            if chunk.startswith("{"):
                text = chunk
                break

    schema = json.loads(text)
    schema.setdefault("default_division", sheet_name)

    # Persist to disk so next restart skips the API call
    _save_disk_schema(cache_key, schema)
    _schema_cache[cache_key] = schema
    return schema


# ---------------------------------------------------------------------------
# Row-level parsing
# ---------------------------------------------------------------------------

_HEADER_CELLS = {"white", "team", "white team", "home", "game", "game id",
                 "game #", "time", "dark", "dark team"}

_TIME_RE = re.compile(r"^\d{1,2}:\d{2}")


def _is_valid_time(val) -> bool:
    """Return True if val looks like a game time (time object or HH:MM string)."""
    if val is None:
        return False
    if hasattr(val, "hour"):
        return True
    if isinstance(val, str):
        return bool(_TIME_RE.match(val.strip()))
    return False


def _extract_games_from_row(
    row: tuple,
    current_date: Optional[date_type],
    current_location: str,
    schema: dict,
    sheet_name: str,
    col_prefix: str = "",
) -> list[dict]:
    """Extract one or two games from a row using primary or alt column mapping."""
    gid_col = schema.get(f"{col_prefix}game_id_col")
    time_col = schema.get(f"{col_prefix}time_col")
    white_col = schema.get(f"{col_prefix}white_col")
    ws_col = schema.get(f"{col_prefix}white_score_col")
    dark_col = schema.get(f"{col_prefix}dark_col")
    ds_col = schema.get(f"{col_prefix}dark_score_col")
    div_col = schema.get(f"{col_prefix}division_col")
    loc_col = schema.get("location_col") if not col_prefix else None

    game_id = _safe_get(row, gid_col)
    white = _safe_get(row, white_col)
    dark = _safe_get(row, dark_col)

    if not game_id or not white or not dark:
        return []
    if str(white).strip().lower() in _HEADER_CELLS:
        return []

    # If a time column is defined, require a valid time value — this filters out
    # standings tables and description rows that happen to have content in game columns.
    time_raw = _safe_get(row, time_col)
    if time_col is not None and not _is_valid_time(time_raw):
        return []

    w_score = _safe_get(row, ws_col)
    d_score = _safe_get(row, ds_col)
    division_val = _safe_get(row, div_col)
    location_val = _safe_get(row, loc_col) if loc_col is not None else None

    return [{
        "date":        current_date,
        "time":        time_raw if hasattr(time_raw, "hour") else None,
        "location":    str(location_val).strip() if location_val else current_location,
        "game_id":     str(game_id).strip(),
        "white_team":  normalize_team_slot(str(white).strip()),
        "white_score": _to_int(w_score),
        "dark_team":   normalize_team_slot(str(dark).strip()),
        "dark_score":  _to_int(d_score),
        "comments":    "",
        "division":    str(division_val).strip() if division_val else schema.get("default_division", sheet_name),
        "sheet":       sheet_name,
        "played":      w_score is not None and d_score is not None,
    }]


def _parse_sheet(all_rows: list, schema: dict, sheet_name: str) -> list[dict]:
    games = []
    current_date: Optional[date_type] = None
    current_location = "TBD"

    header_rows = set(schema.get("header_rows") or [])
    date_in_row = schema.get("date_in_row", True)
    date_col = schema.get("date_col")
    section_header_pattern = schema.get("section_header_pattern")
    section_date_col = schema.get("section_date_col")
    section_location_col = schema.get("section_location_col")
    has_alt = schema.get("alt_game_id_col") is not None
    alt_date_is_next_day = schema.get("alt_date_is_next_day", False)

    hdr_re = re.compile(section_header_pattern, re.IGNORECASE) if section_header_pattern else None

    for row_idx, row in enumerate(all_rows):
        if row_idx in header_rows:
            continue
        if not any(c is not None for c in row):
            continue

        first_cell = str(row[0]).strip() if row[0] is not None else ""

        # ---- Section-header rows ----
        if not date_in_row and hdr_re and first_cell and hdr_re.search(first_cell):
            # Date from a specific column (e.g. col 1 holds a datetime)
            if section_date_col is not None:
                d = _coerce_date(_safe_get(row, section_date_col))
                if d:
                    current_date = d
            else:
                # Parse date from the text of col 0
                d = _extract_date_from_text(first_cell)
                if d:
                    current_date = d

            # Location from section header
            if section_location_col is not None:
                loc_raw = _safe_get(row, section_location_col)
                if loc_raw:
                    current_location = str(loc_raw).strip()
            continue

        # ---- Per-row date ----
        if date_in_row and date_col is not None:
            d = _coerce_date(_safe_get(row, date_col))
            if d:
                current_date = d

        # ---- Primary game block ----
        games.extend(_extract_games_from_row(
            row, current_date, current_location, schema, sheet_name, col_prefix=""
        ))

        # ---- Alt (Sunday) game block ----
        if has_alt:
            alt_date = (current_date + timedelta(days=1)) if (current_date and alt_date_is_next_day) else current_date
            games.extend(_extract_games_from_row(
                row, alt_date, current_location, schema, sheet_name, col_prefix="alt_"
            ))

    return games


# ---------------------------------------------------------------------------
# Public entry points
# ---------------------------------------------------------------------------

def parse_cached(wb) -> list[dict]:
    """Parse using only cached schemas (memory or disk) — no API calls.

    Returns [] if any non-skipped sheet lacks a cached schema, signalling
    detect.py to fall back to the heuristic parsers instead.
    """
    games = []
    for sheet_name in wb.sheetnames:
        if sheet_name.lower().strip() in SKIP_SHEETS:
            continue
        ws = wb[sheet_name]
        all_rows = list(ws.iter_rows(values_only=True))
        if not all_rows:
            continue

        cache_key = _hash_schema_key(sheet_name, all_rows)

        # Try memory cache first, then disk cache
        schema = _schema_cache.get(cache_key) or _load_disk_schema(cache_key)
        if schema is None:
            # Unknown layout — signal caller to use heuristic parsers
            return []

        # Warm the memory cache for future calls in the same process
        if cache_key not in _schema_cache:
            _schema_cache[cache_key] = schema

        if schema.get("skip"):
            continue

        try:
            sheet_games = _parse_sheet(all_rows, schema, sheet_name)
            games.extend(sheet_games)
        except Exception as exc:
            print(f"[AI parser] parse_cached failed for '{sheet_name}': {exc}")

    return games


def parse(wb) -> list[dict]:
    games = []
    for sheet_name in wb.sheetnames:
        if sheet_name.lower().strip() in SKIP_SHEETS:
            continue

        ws = wb[sheet_name]
        all_rows = list(ws.iter_rows(values_only=True))
        if not all_rows:
            continue

        try:
            schema = _get_schema(sheet_name, all_rows)
        except Exception as exc:
            print(f"[AI parser] schema detection failed for '{sheet_name}': {exc}")
            continue

        if schema.get("skip"):
            continue

        try:
            sheet_games = _parse_sheet(all_rows, schema, sheet_name)
            games.extend(sheet_games)
        except Exception as exc:
            print(f"[AI parser] parse failed for '{sheet_name}': {exc}")

    return games

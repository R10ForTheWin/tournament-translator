"""
Tournament Translator — Flask app
"""
from __future__ import annotations
import os, re, json, glob, io, time, base64, random, threading
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from functools import lru_cache
import requests
from flask import Flask, render_template, jsonify, request, abort

from parsers.detect import load_and_parse as _load_and_parse

_parse_cache: dict[str, tuple[float, list]] = {}

def load_and_parse(filepath) -> list[dict]:
    """Cached wrapper: re-parses only when the file changes on disk."""
    if filepath == _JO_QUALS_SENTINEL:
        return _fetch_jo_quals_games()
    key = str(filepath)
    try:
        mtime = os.path.getmtime(key)
    except OSError:
        return _load_and_parse(filepath)
    cached = _parse_cache.get(key)
    if cached and cached[0] == mtime:
        return list(cached[1])
    games = _load_and_parse(filepath)
    _parse_cache[key] = (mtime, games)
    return list(games)

app = Flask(__name__)

EXCEL_DIR    = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tournaments Excels")
RESULTS_DIR  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "results")
COMMENTS_DIR  = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "comments")
FEEDBACK_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "feedback.json")
ADMIN_PW    = os.environ.get("ADMIN_PASSWORD", "trojan")  # override via Railway env var

# ── Live URL sources ────────────────────────────────────────────────────────
# All tournament URLs baked in code — updated here after each tournament is scheduled.
# Phone UI (user_urls.json) overrides these within a session.
FUTURES_SHEETS_ID  = "1AkX3vwOU9CIc3cymacG2F-uXz-_Gi_A8yR40dEbDpMQ"
FUTURES_SHEETS_URL = f"https://docs.google.com/spreadsheets/d/{FUTURES_SHEETS_ID}/export?format=xlsx"
WPL_TOURNAMENTS    = {"futures-2", "futures-3", "futures-4", "futures-5", "futures-super"}

JO_QUALS_SHEETS_ID   = "1IAZNJosmCYMjgDw_YZ792hEP3EAxFX9cJnJqzBgoIJU"
JO_QUALS_GID_18U     = "1123307693"
JO_QUALS_GID_16U     = "2019076064"
JO_QUALS_TOURNAMENTS = {"jo-quals"}
_JO_QUALS_SENTINEL   = "__jo_quals__"

TOURNAMENT_URLS = {
    "kap7-intl":      "",  # update before Jan 2027 tournament
    "kap7-cup":       "https://onedrive.live.com/:x:/g/personal/6f253ef3afcfe1c8/IQDxdebmKQFASaux2nx7kWcvASg2jqJRHO5Kj9EwKW4D82o?rtime=GrzV592c3kg&redeem=aHR0cHM6Ly8xZHJ2Lm1zL3gvYy82ZjI1M2VmM2FmY2ZlMWM4L0lRRHhkZWJtS1FGQVNhdXgybng3a1djdkFTZzJqcUpSSE81S2o5RXdLVzREODJvP2U9WEdNa1FB",
    "turbo-cup":      "https://1drv.ms/x/c/6f253ef3afcfe1c8/IQB7PJXtfzNsT74lTYhWpOeXASFcmpB96L1OpYL_E6HBMM0?e=UsRrMb",
    "newport-invite": "https://onedrive.live.com/download?resid=6F253EF3AFCFE1C8!66694&authkey=!AO8pyWY0qwL2sYE",
    "jo-quals":       "",  # update when schedule is posted
    "junior-olympics":"",  # update when schedule is posted
}

PRESET_URL_TOURNAMENTS = WPL_TOURNAMENTS | {k for k, v in TOURNAMENT_URLS.items() if v}
USER_URLS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "user_urls.json")

_URL_CACHE: dict   = {}   # {url: (fetched_at, bytes)}
URL_CACHE_TTL      = 300  # re-fetch at most every 5 minutes

_LIVE_SCORES: dict = {}   # {(tournament_id, game_id): {our_score, opp_score, quarter, updated_at}}
def _load_feedback() -> list:
    try:
        with open(FEEDBACK_FILE) as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return []

def _save_feedback(entries: list) -> None:
    with open(FEEDBACK_FILE, "w") as f:
        json.dump(entries, f)

_FEEDBACK: list = _load_feedback()

_MONTH_MAP = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
              "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
_RE_YEAR   = re.compile(r"(\d{4})")
_RE_MONTH  = re.compile(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)", re.I)

os.makedirs(RESULTS_DIR,  exist_ok=True)
os.makedirs(COMMENTS_DIR, exist_ok=True)
os.makedirs(EXCEL_DIR,    exist_ok=True)

# ── Tournament registry ────────────────────────────────────────────────────────

KNOWN_TOURNAMENTS = [
    {"id": "kap7-intl",       "name": "Kap7 International",     "dates": "Jan 31–Feb 1, 2026"},
    {"id": "futures-2",       "name": "Futures Weekend 2",      "dates": "Feb 21–22, 2026",
     "date_start": date(2026, 2, 21), "date_end": date(2026, 2, 22)},
    {"id": "turbo-cup",       "name": "Turbo OC Cup",           "dates": "Mar 7–8, 2026"},
    {"id": "newport-invite",  "name": "Newport Spring Invite",  "dates": "Mar 14–15, 2026"},
    {"id": "futures-3",       "name": "Futures Weekend 3",      "dates": "Mar 21–22, 2026",
     "date_start": date(2026, 3, 21), "date_end": date(2026, 3, 22)},
    {"id": "kap7-cup",        "name": "Kap7 Cup",               "dates": "Apr 18–19, 2026",
     "date_start": date(2026, 4, 18), "date_end": date(2026, 4, 19)},
    {"id": "futures-4",       "name": "Futures Weekend 4",      "dates": "May 2–3, 2026",
     "date_start": date(2026, 5, 2),  "date_end": date(2026, 5, 3)},
    {"id": "futures-5",       "name": "Futures Weekend 5",      "dates": "May 16–17, 2026",
     "date_start": date(2026, 5, 16), "date_end": date(2026, 5, 17)},
    {"id": "jo-quals",        "name": "JO Qualifications",      "dates": "May 29–31, 2026",
     "date_start": date(2026, 5, 29), "date_end": date(2026, 5, 31)},
    {"id": "futures-super",   "name": "Futures Superfinal",     "dates": "Jun 26–28, 2026",
     "date_start": date(2026, 6, 26), "date_end": date(2026, 6, 28)},
    {"id": "junior-olympics", "name": "Junior Olympics",        "dates": "Jul 23–26, 2026",
     "date_start": date(2026, 7, 23), "date_end": date(2026, 7, 26)},
]

# Map tournament id → excel filename (partial match, case-insensitive)
FILE_MAP = {
    "turbo-cup":      "TURBO OC CUP",
    "kap7-intl":      "KAP7 INTERNATIONAL",
    "newport-invite": "NEWPORT",
    "futures-2":      "KAP7 Futures WPL",
    "futures-3":      "KAP7 Futures WPL",
    "futures-4":      "KAP7 Futures WPL",
    "futures-5":      "KAP7 Futures WPL",
    "futures-super":  "KAP7 Futures WPL",
}


def _load_user_urls() -> dict:
    env_urls = os.environ.get("USER_URLS_JSON", "")
    try:
        base = json.loads(env_urls) if env_urls else {}
    except Exception:
        base = {}
    try:
        if os.path.exists(USER_URLS_FILE):
            with open(USER_URLS_FILE) as f:
                base.update(json.load(f))
    except Exception:
        pass
    return base

def _save_user_url(tournament_id: str, url: str):
    urls = _load_user_urls()
    urls[tournament_id] = url
    with open(USER_URLS_FILE, "w") as f:
        json.dump(urls, f, indent=2)


def _fetch_url(url: str, *, onedrive=False) -> bytes | None:
    """Fetch Excel bytes from a URL with 5-min cache. Returns None on failure."""
    now = time.time()
    cached = _URL_CACHE.get(url)
    if cached and now - cached[0] < URL_CACHE_TTL:
        return cached[1]
    try:
        if onedrive:
            token = base64.urlsafe_b64encode(url.encode()).rstrip(b"=").decode()
            fetch_url = f"https://api.onedrive.com/v1.0/shares/u!{token}/root/content"
        elif "onedrive.live.com" in url or "1drv.ms" in url:
            sep = "&" if "?" in url else "?"
            fetch_url = url + sep + "download=1"
        else:
            fetch_url = url
        resp = requests.get(fetch_url, allow_redirects=True, timeout=30)
        resp.raise_for_status()
        _URL_CACHE[url] = (now, resp.content)
        return resp.content
    except Exception as exc:
        app.logger.warning("Fetch failed (%s): %s", url, exc)
        return cached[1] if cached else None


def find_excel(tournament_id: str):
    """Return an Excel file path, BytesIO from a live URL, or None.
    User-pasted URL (stored in user_urls.json) takes priority over all presets."""
    if tournament_id in JO_QUALS_TOURNAMENTS:
        return _JO_QUALS_SENTINEL
    user_url = _load_user_urls().get(tournament_id)
    if user_url:
        data = _fetch_url(user_url)
        if data:
            return io.BytesIO(data)
    # WPL tournaments always fetch live from Google Sheets — never use a
    # local file, which would be a stale snapshot from a past weekend.
    if tournament_id in WPL_TOURNAMENTS:
        data = _fetch_url(FUTURES_SHEETS_URL)
        return io.BytesIO(data) if data else None
    keyword = FILE_MAP.get(tournament_id, "")
    if keyword:
        for f in os.listdir(EXCEL_DIR):
            if f.endswith(".xlsx") and keyword.lower() in f.lower():
                return os.path.join(EXCEL_DIR, f)
    preset = TOURNAMENT_URLS.get(tournament_id, "")
    if preset:
        data = _fetch_url(preset)
        return io.BytesIO(data) if data else None
    return None


def _cache_age(tournament_id: str) -> int | None:
    """Seconds since the Excel was last fetched, or None if not yet cached."""
    url = TOURNAMENT_URLS.get(tournament_id) or _load_user_urls().get(tournament_id)
    if not url:
        url = FUTURES_SHEETS_URL if tournament_id in WPL_TOURNAMENTS else None
    if not url:
        return None
    cached = _URL_CACHE.get(url)
    return int(time.time() - cached[0]) if cached else None


def _fetch_jo_quals_games() -> list[dict]:
    """Fetch live game data for JO Qualifications from the CCA Google Sheets."""
    from parsers.format_cca import parse_csv
    base = (f"https://docs.google.com/spreadsheets/d/{JO_QUALS_SHEETS_ID}"
            f"/gviz/tq?tqx=out:csv&gid=")
    tabs = [
        (JO_QUALS_GID_18U, "18U Boys", "18U"),
        (JO_QUALS_GID_16U, "16U Boys", "16U"),
    ]
    all_games: list[dict] = []
    for gid, division, prefix in tabs:
        data = _fetch_url(base + gid)
        if not data:
            app.logger.warning("JO Quals: failed to fetch %s (gid=%s)", division, gid)
            continue
        try:
            csv_text = data.decode("utf-8")
        except UnicodeDecodeError:
            csv_text = data.decode("latin-1")
        games = parse_csv(csv_text, division, prefix)
        for g in games:
            g["format"] = "CCA"
        all_games.extend(games)
    return all_games


def _tournament_meta(tournament_id: str):
    """Return the KNOWN_TOURNAMENTS entry for this id, or None."""
    return next((t for t in KNOWN_TOURNAMENTS if t["id"] == tournament_id), None)


def _filter_by_dates(games, tournament_id: str):
    """If the tournament has a date range, filter games to that range.
    Games with date=None are kept only when no date range is set."""
    meta = _tournament_meta(tournament_id)
    if not meta or "date_start" not in meta:
        return games
    d0, d1 = meta["date_start"], meta["date_end"]
    return [g for g in games if g.get("date") and d0 <= g["date"] <= d1]


def all_excels():
    """All .xlsx files in the Excel directory."""
    return [f for f in os.listdir(EXCEL_DIR) if f.endswith(".xlsx")]


# ── Head-to-head normalization ──────────────────────────────────────────────────

# Tier mapping for Trojan (known colors → A/B/C)
_TROJAN_TIER = {
    "cardinal": "A", "red": "A", "platinum": "A",
    "gold": "B", "blue": "B",
    "silver": "C", "white": "C", "bronze": "C",
}
# Colors to strip from end of non-Trojan names for club normalization
_STRIP_COLOR = re.compile(
    r"\s+(gold|silver|cardinal|red|blue|white|platinum|bronze|"
    r"black|green|orange|purple|maroon|gray|grey)\s*$",
    re.IGNORECASE,
)
_EXPLICIT_LETTER = re.compile(r"^(.*?)\s+\(?([ABC])\)?\s*$", re.IGNORECASE)
_EXPLICIT_NUMBER = re.compile(r"^(.*?)\s+([123])\s*$")


def normalize_opp(raw: str):
    """Return (club_key, tier) for head-to-head matching.
    club_key is uppercase, stripped of tier indicators.
    tier is 'A', 'B', 'C', or None."""
    name = strip_prefix(raw).strip()

    # Trojan — use known color→tier mapping
    if "TROJAN" in name.upper():
        for color, tier in _TROJAN_TIER.items():
            if re.search(rf"\b{re.escape(color)}\b", name, re.IGNORECASE):
                return ("TROJAN", tier)
        return ("TROJAN", "A")   # bare TROJAN = top team

    # Explicit letter suffix: "Newport Beach A", "San Clemente C"
    m = _EXPLICIT_LETTER.match(name)
    if m:
        return (m.group(1).strip().upper(), m.group(2).upper())

    # Number suffix: "Newport 1", "Newport 2"
    m = _EXPLICIT_NUMBER.match(name)
    if m:
        return (m.group(1).strip().upper(), "ABC"[int(m.group(2)) - 1])

    # Strip trailing color word to normalize club name; tier unknown
    club = _STRIP_COLOR.sub("", name).strip()
    return (club.upper(), None)


def clubs_match(a: str, b: str) -> bool:
    """Fuzzy club match — handles minor spelling differences like
    'NEWPORT' vs 'NEWPORT BEACH'."""
    a, b = a.upper().strip(), b.upper().strip()
    return a == b or a in b or b in a


def h2h_key(raw: str) -> str:
    club, tier = normalize_opp(raw)
    return f"{club}|{tier or ''}"


# ── Team friendly-name helpers ─────────────────────────────────────────────────

_TIER_ORDER = {"cardinal": 0, "gold": 1, "platinum": 1, "silver": 2, "bronze": 3}
_AGE_WORDS   = re.compile(r"(\d{1,2})U", re.IGNORECASE)
_GENDER_BOYS = re.compile(r"\bBOYS?\b", re.IGNORECASE)
_GENDER_GIRLS = re.compile(r"\bGIRLS?\b", re.IGNORECASE)
_GENDER_COED  = re.compile(r"\bCOED\b", re.IGNORECASE)

def _parse_sheet(sheet: str):
    """Return (gender, age_str) from an Excel sheet name, e.g. '16U BOYS GOLD-11 TEAMS' → ('Boys', '16U')."""
    s = sheet.upper()
    age_m = _AGE_WORDS.search(s)
    age = f"{age_m.group(1)}U" if age_m else None
    if _GENDER_GIRLS.search(s):   gender = "Girls"
    elif _GENDER_BOYS.search(s):  gender = "Boys"
    elif _GENDER_COED.search(s):  gender = "Coed"
    else:                          gender = None
    return gender, age

def _parse_tier(team_name: str):
    """Return display tier from team name: Gold, Cardinal, Silver, etc."""
    t = team_name.upper()
    for keyword, label in [("GOLD","Gold"),("CARDINAL","Cardinal"),("PLATINUM","Platinum"),
                            ("SILVER","Silver"),("BRONZE","Bronze"),("NORTH","North"),("SOUTH","South")]:
        if keyword in t:
            return label
    return None

def friendly_team_name(team_name: str, sheet: str):
    """'TROJAN GOLD' + '16U BOYS PLATINUM-11 TEAMS' → 'Trojan Gold · Boys 16U'.
    Also strips embedded age/gender words that Kahuna includes in the team name itself,
    e.g. 'TROJAN 16 BOYS 19U CARDINAL' → 'Trojan 16 Cardinal · Boys 19U'."""
    gender, age = _parse_sheet(sheet)
    # Strip embedded age/gender words from the team name (Kahuna includes them)
    name = team_name.strip()
    if gender:
        name = _GENDER_BOYS.sub("", name) if gender == "Boys" else name
        name = _GENDER_GIRLS.sub("", name) if gender == "Girls" else name
        name = _GENDER_COED.sub("", name) if gender == "Coed" else name
    if age:
        name = re.sub(rf"\b{re.escape(age)}\b", "", name, flags=re.IGNORECASE)
    name = re.sub(r"\s{2,}", " ", name).strip().title()
    age_gender = " ".join(p for p in [gender, age] if p)
    return f"{name} · {age_gender}" if age_gender else name or None

def _team_sort_key(team: dict):
    """Sort: Boys before Girls, age desc (18U→10U), tier Cardinal→Gold→Silver."""
    friendly = team.get("friendly") or ""
    gender_ord = 0 if "Boys" in friendly else (1 if "Girls" in friendly else 2)
    age_m = _AGE_WORDS.search(friendly)
    age_ord = -(int(age_m.group(1))) if age_m else 0   # negate for descending
    tier = _parse_tier(team["name"]) or ""
    tier_ord = _TIER_ORDER.get(tier.lower(), 9)
    return (gender_ord, age_ord, tier_ord)


# ── Game helpers ───────────────────────────────────────────────────────────────

_PREFIX_RE = re.compile(
    r"^(?:\d+(?:st|nd|rd|th)[A-Z]+-|[A-Z]+\d+\s*\([^)]+\)\s*-\s*|[WL]#[^-\s]+-?|[A-Z]+\d+\s*-\s*|\d+\s*-\s*)(.*)",
    re.IGNORECASE,
)

def strip_prefix(s: str) -> str:
    m = _PREFIX_RE.match(s.strip())
    return m.group(1).strip() if m else s.strip()

def is_trojan(team_slot: str) -> bool:
    return "TROJAN" in strip_prefix(team_slot).upper()

def team_matches(slot: str, name: str) -> bool:
    return name.upper() in strip_prefix(slot).upper()

def _title(s: str) -> str:
    return s.title() if s else s

def _pool_teams_for_group(group: str, division_games: list) -> list[str]:
    """All team names seeded in a pool group, ordered by seed number."""
    teams = {}
    for g in division_games:
        for slot in (g["white_team"], g["dark_team"]):
            m = _POOL_SLOT_RE.match(slot.strip())
            if m and m.group(1).upper() == group.upper():
                seed = int(m.group(2))
                name = m.group(3).strip()
                if name:
                    teams[seed] = name
    return [teams[k] for k in sorted(teams)]

# ── Place Predictor ────────────────────────────────────────────────────────────

def _ordinal(n: int) -> str:
    return {1: "1st", 2: "2nd", 3: "3rd"}.get(n, f"{n}th")

def _find_team_pool_group(team: str, division_games: list):
    for g in division_games:
        for slot in (g["white_team"], g["dark_team"]):
            m = _POOL_SLOT_RE.match(slot.strip())
            if m and team_matches(m.group(3), team):
                return m.group(1).upper()
    return None

def _standings_for_group(group: str, division_games: list, extra_outcomes: dict = None) -> list:
    """Pool standings with optional simulated outcomes for unplayed games.
    Handles both direct pool-slot games and W#/L# bracket games within the pool."""
    extra_outcomes = extra_outcomes or {}
    team_stats: dict = {}

    for name in _pool_teams_for_group(group, division_games):
        team_stats[name] = {"team": name, "wins": 0, "losses": 0, "gf": 0, "ga": 0}

    # Pass 1: direct pool-slot games (B1-X vs B2-Y format)
    direct_ids: set = set()
    game_results: dict = {}  # game_id -> (winner_key, loser_key) for W#/L# resolution

    for g in division_games:
        wm = _POOL_SLOT_RE.match(g["white_team"].strip())
        dm = _POOL_SLOT_RE.match(g["dark_team"].strip())
        if not wm or not dm:
            continue
        if wm.group(1).upper() != group.upper() or dm.group(1).upper() != group.upper():
            continue
        direct_ids.add(g["game_id"])
        wt, dt = wm.group(3).strip(), dm.group(3).strip()
        ws, ds = g.get("white_score"), g.get("dark_score")
        gid = g["game_id"]

        if g.get("played") and ws is not None and ds is not None:
            white_wins, wgf, dgf = ws > ds, ws, ds
        elif gid in extra_outcomes:
            white_wins, wgf, dgf = extra_outcomes[gid], 0, 0
        else:
            continue

        wkey = next((k for k in team_stats if team_matches(k, wt)), None)
        dkey = next((k for k in team_stats if team_matches(k, dt)), None)
        if not wkey or not dkey:
            continue

        if white_wins:
            team_stats[wkey]["wins"] += 1; team_stats[dkey]["losses"] += 1
        else:
            team_stats[dkey]["wins"] += 1; team_stats[wkey]["losses"] += 1
        team_stats[wkey]["gf"] += wgf; team_stats[wkey]["ga"] += dgf
        team_stats[dkey]["gf"] += dgf; team_stats[dkey]["ga"] += wgf
        game_results[gid] = (wkey if white_wins else dkey, dkey if white_wins else wkey)

    # Pass 2: W#/L# bracket games within the pool (e.g. L#2 vs L#4 tiebreaker games)
    def _resolve_wl(slot: str) -> str | None:
        wm = _WL_SLOT_RE.match(slot.strip())
        if not wm:
            return None
        ref = re.search(r'(\d+)$', wm.group(1))
        if not ref:
            return None
        ref_num = str(int(ref.group(1)))
        ref_gid = next((gid for gid in direct_ids if _game_num(gid) == ref_num), None)
        if not ref_gid or ref_gid not in game_results:
            return None
        winner, loser = game_results[ref_gid]
        return winner if slot.strip()[0].upper() == 'W' else loser

    for g in division_games:
        if g["game_id"] in direct_ids:
            continue
        wt = _resolve_wl(g["white_team"])
        dt = _resolve_wl(g["dark_team"])
        if not wt or not dt:
            continue
        ws, ds = g.get("white_score"), g.get("dark_score")
        gid = g["game_id"]

        if g.get("played") and ws is not None and ds is not None:
            white_wins, wgf, dgf = ws > ds, ws, ds
        elif gid in extra_outcomes:
            white_wins, wgf, dgf = extra_outcomes[gid], 0, 0
        else:
            continue

        if white_wins:
            team_stats[wt]["wins"] += 1; team_stats[dt]["losses"] += 1
        else:
            team_stats[dt]["wins"] += 1; team_stats[wt]["losses"] += 1
        team_stats[wt]["gf"] += wgf; team_stats[wt]["ga"] += dgf
        team_stats[dt]["gf"] += dgf; team_stats[dt]["ga"] += wgf

    return sorted(team_stats.values(),
                  key=lambda s: (-s["wins"], -(s["gf"] - s["ga"]), -s["gf"]))


def _pool_bracket_games(group: str, division_games: list) -> list:
    """Return W#/L# games within the pool: both slots reference direct pool games."""
    direct_ids = {
        g["game_id"] for g in division_games
        if (m := _POOL_SLOT_RE.match(g["white_team"].strip())) and
           (_POOL_SLOT_RE.match(g["dark_team"].strip())) and
           m.group(1).upper() == group.upper()
    }

    def _refs_direct(slot: str) -> bool:
        wm = _WL_SLOT_RE.match(slot.strip())
        if not wm:
            return False
        ref = re.search(r'(\d+)$', wm.group(1))
        if not ref:
            return False
        ref_num = str(int(ref.group(1)))
        return any(_game_num(gid) == ref_num for gid in direct_ids)

    return [g for g in division_games
            if g["game_id"] not in direct_ids
            and _refs_direct(g["white_team"]) and _refs_direct(g["dark_team"])]


def _resolve_slot_for_sim(slot: str, group_standings: dict, game_results: dict) -> str | None:
    """Resolve a bracket slot to a team name for Monte Carlo simulation.
    group_standings: {group: [team_name, ...]} ordered 1st..last
    game_results: {game_id: (winner_name, loser_name)}
    """
    slot = slot.strip()

    pm = _POOL_SLOT_RE.match(slot)
    if pm:
        return pm.group(3).strip()

    # Composite+W#/L# like "E1(4thB)L#7": W#/L# takes priority over pool label
    name = strip_prefix(slot)
    if name != slot and name and re.match(r'^[WL]#', name, re.IGNORECASE):
        wm = _WL_SLOT_RE.match(name)
        if wm:
            ref = re.search(r'(\d+)$', wm.group(1))
            if ref:
                ref_num = str(int(ref.group(1)))
                ref_gid = next((gid for gid in game_results if _game_num(gid) == ref_num), None)
                if ref_gid:
                    winner, loser = game_results[ref_gid]
                    return winner if name[0].upper() == 'W' else loser

    wm = _WL_SLOT_RE.match(slot)
    if wm:
        ref = re.search(r'(\d+)$', wm.group(1))
        if ref:
            ref_num = str(int(ref.group(1)))
            ref_gid = next((gid for gid in game_results if _game_num(gid) == ref_num), None)
            if ref_gid:
                winner, loser = game_results[ref_gid]
                return winner if slot[0].upper() == 'W' else loser

    # Composite: K4(2ndB) or K4(2ndB)-
    cm = re.search(r'\((\d+)(?:st|nd|rd|th)([A-Z])\)', slot, re.IGNORECASE)
    if cm:
        rank = int(cm.group(1))
        teams = group_standings.get(cm.group(2).upper(), [])
        return teams[rank - 1] if len(teams) >= rank else None

    # Simple finish slot: 1stA-
    fm = _FINISH_SLOT_RE.match(slot)
    if fm:
        group = fm.group(1).upper()
        rank_m = re.search(r'^(\d+)', slot)
        rank = int(rank_m.group(1)) if rank_m else 1
        teams = group_standings.get(group, [])
        return teams[rank - 1] if len(teams) >= rank else None

    return None


def _tournament_finish_probs(team: str, division_games: list, n_trials: int = 500) -> tuple:
    """Monte Carlo simulation of final tournament placement.
    Returns (our_probs, all_team_probs) where each is {placement_int: probability_0_to_1}.
    """
    all_groups = sorted({
        m.group(1).upper()
        for g in division_games
        for slot in (g["white_team"], g["dark_team"])
        if (m := _POOL_SLOT_RE.match(slot.strip()))
    })
    if not all_groups:
        return {}

    all_pool_teams: dict = {}
    for group in all_groups:
        for t in _pool_teams_for_group(group, division_games):
            all_pool_teams[t] = group
    if not all_pool_teams:
        return {}

    our_team = next((t for t in all_pool_teams if team_matches(t, team)), None)
    if not our_team:
        return {}, {}

    # Pool phase games (direct pool games + intra-pool bracket games)
    def _is_pool_slot(g):
        return bool(_POOL_SLOT_RE.match(g["white_team"].strip()) and
                    _POOL_SLOT_RE.match(g["dark_team"].strip()))

    pool_direct = [g for g in division_games if _is_pool_slot(g)]
    intra_bracket: list = []
    for grp in all_groups:
        intra_bracket.extend(_pool_bracket_games(grp, division_games))
    pool_phase_ids = {g["game_id"] for g in pool_direct + intra_bracket}

    # Sunday placement bracket games (composite slots)
    bracket_games = sorted(
        (g for g in division_games if g["game_id"] not in pool_phase_ids),
        key=lambda g: (g.get("date") or date.min, g.get("time") or datetime.min.time()),
    )

    unplayed_pool_phase = [g for g in division_games
                           if g["game_id"] in pool_phase_ids and not g.get("played")]

    placement_counts: dict = {}
    all_placement_counts: dict = {t: {} for t in all_pool_teams}

    for _ in range(n_trials):
        pool_outcomes = {g["game_id"]: random.random() < 0.5 for g in unplayed_pool_phase}

        group_standings: dict = {}
        for grp in all_groups:
            st = _standings_for_group(grp, division_games, pool_outcomes)
            group_standings[grp] = [s["team"] for s in st]

        game_results: dict = {}  # game_id -> (winner_name, loser_name)
        # [bracket_wins, last_game_date_ordinal, last_game_minute]
        team_rec: dict = {t: [0, 0, 0] for t in all_pool_teams}

        for g in bracket_games:
            wt = _resolve_slot_for_sim(g["white_team"], group_standings, game_results)
            dt = _resolve_slot_for_sim(g["dark_team"], group_standings, game_results)
            if not wt or not dt:
                continue

            if g.get("played") and g.get("white_score") is not None:
                white_won = g["white_score"] > g["dark_score"]
            else:
                white_won = random.random() < 0.5

            winner, loser = (wt, dt) if white_won else (dt, wt)
            game_results[g["game_id"]] = (winner, loser)

            gdate = g.get("date") or date.min
            gtime = g.get("time") or datetime.min.time()
            date_ord = gdate.toordinal() if gdate != date.min else 0
            time_min = gtime.hour * 60 + gtime.minute if gdate != date.min else 0

            for name, won in ((winner, True), (loser, False)):
                key = next((t for t in all_pool_teams if team_matches(t, name)), None)
                if key:
                    if won:
                        team_rec[key][0] += 1
                    team_rec[key][1] = max(team_rec[key][1], date_ord)
                    team_rec[key][2] = max(team_rec[key][2], time_min)

        def _rank_key(t):
            wins, date_ord, time_min = team_rec[t]
            grp = all_pool_teams[t]
            pool_rank = next((i for i, pt in enumerate(group_standings.get(grp, []))
                              if team_matches(pt, t)), 99)
            return (-wins, -date_ord, -time_min, pool_rank)

        sorted_teams = sorted(all_pool_teams.keys(), key=_rank_key)
        for i, t in enumerate(sorted_teams):
            p = i + 1
            all_placement_counts[t][p] = all_placement_counts[t].get(p, 0) + 1
        placement = next((i + 1 for i, t in enumerate(sorted_teams) if t == our_team), None)
        if placement is not None:
            placement_counts[placement] = placement_counts.get(placement, 0) + 1

    total = sum(placement_counts.values())
    our_probs = {p: c / total for p, c in placement_counts.items()} if total else {}
    all_probs = {
        t: {p: c / total for p, c in counts.items()}
        for t, counts in all_placement_counts.items()
    } if total else {}
    return our_probs, all_probs


@app.route("/api/place-predictor/<tournament_id>/<path:team>")
def api_place_predictor(tournament_id, team):
    excel = find_excel(tournament_id)
    if not excel:
        abort(404)

    sheet = request.args.get("sheet")
    games = _filter_by_dates(load_and_parse(excel), tournament_id)
    division_games = [g for g in games if sheet is None or g["sheet"] == sheet]

    group = _find_team_pool_group(team, division_games)
    if not group:
        return jsonify({"pool": None, "finish_probs": []})

    standings    = _standings_for_group(group, division_games)
    current_rank = next((i + 1 for i, s in enumerate(standings) if team_matches(s["team"], team)), None)
    team_stats   = next((s for s in standings if team_matches(s["team"], team)), {})
    finish_probs, all_team_probs = _tournament_finish_probs(team, division_games)

    all_groups = {
        m.group(1).upper()
        for g in division_games
        for slot in (g["white_team"], g["dark_team"])
        if (m := _POOL_SLOT_RE.match(slot.strip()))
    }
    total_div_teams = sum(len(_pool_teams_for_group(g, division_games)) for g in all_groups)

    def _modal_placement(probs):
        return max(probs.items(), key=lambda x: x[1])[0] if probs else 999

    predicted_standings = sorted(
        [{"team": t,
          "predicted_rank": _modal_placement(all_team_probs.get(t, {})),
          "pct": round(max(all_team_probs.get(t, {}).values(), default=0) * 100),
          "is_mine": team_matches(t, team)}
         for t in all_team_probs],
        key=lambda x: x["predicted_rank"],
    )

    return jsonify({
        "pool": {
            "group":        group,
            "current_rank": current_rank,
            "total_teams":  len(standings),
            "wins":         team_stats.get("wins", 0),
            "losses":       team_stats.get("losses", 0),
        },
        "finish_probs": [
            {"rank": r, "label": _ordinal(r), "pct": round(p * 100)}
            for r, p in sorted(finish_probs.items())
        ],
        "total_div_teams": total_div_teams,
        "predicted_standings": predicted_standings,
    })


def _game_by_num(num: str, division_games: list):
    for g in division_games:
        if _game_num(g["game_id"]) == num:
            return g
    return None

_PTS_SUFFIX_RE = re.compile(r'\s*-\s*\d+\s*PTS\.?\s*$', re.IGNORECASE)

def describe_slot(slot: str, division_games: list = None, ref_date=None) -> str:
    """Return a human-readable opponent label.
    With division_games, resolves bracket slots to actual team names using standings.
    ref_date: when provided, scopes standings lookups to ±2 days to prevent stale
    results from previous weekends (WPL pool letters repeat across weekends)."""
    slot = slot.strip()
    name = strip_prefix(slot)
    if name != slot and name:
        # Strip WPL championship seeding suffix e.g. "OLYMPUS - 9 PTS." → "OLYMPUS"
        name = _PTS_SUFFIX_RE.sub('', name).strip()
        # If strip_prefix left us with a W#/L# reference (e.g. "E1(4thB)L#7"), resolve it.
        if re.match(r'^[WL]#', name, re.IGNORECASE):
            return describe_slot(name, division_games, ref_date)
        # CCA extended: strip left bracket-group prefix, then recurse if still contains
        # a resolvable reference like "4TH B - L55" or "3RD G (WINNER 46)"
        if (re.search(r'\bL(\d+)\s*$', name, re.IGNORECASE)
                or re.search(r'\bWinner\s+\d+', name, re.IGNORECASE)):
            return describe_slot(name, division_games, ref_date)
        return name

    if division_games:
        # Scope to current weekend when ref_date is provided — pool letters repeat
        # across WPL weekends; without scoping, standings bleed in from prior weeks.
        dg = ([g for g in division_games
               if not g.get("date") or abs((g["date"] - ref_date).days) <= 2]
              if ref_date else division_games)

        # W#N / L#N → resolve to actual winner/loser if game has been played
        # Also handles WPL championship format: "WIN GM #N" / "LOS GM #N"
        wm = re.match(r'^([WL])#([^-\s]+)', slot, re.IGNORECASE)
        win_gm = re.search(r'\b(WIN|LOS)\s+GM\s+#(\d+)', slot, re.IGNORECASE) if not wm else None
        if wm or win_gm:
            if wm:
                want_winner = wm.group(1).upper() == "W"
                ref = re.search(r'(\d+)$', wm.group(2))
                ref_num = str(int(ref.group(1))) if ref else None
            else:
                want_winner = win_gm.group(1).upper() == "WIN"
                ref_num = str(int(win_gm.group(2)))
            if ref_num:
                ref_game = _game_by_num(ref_num, dg)
                if ref_game:
                    t1 = describe_slot(ref_game["white_team"], dg)
                    t2 = describe_slot(ref_game["dark_team"], dg)
                    if ref_game.get("played") and ref_game.get("white_score") is not None:
                        white_won = ref_game["white_score"] > ref_game["dark_score"]
                        resolved = (t1 if white_won else t2) if want_winner else (t2 if white_won else t1)
                        return _title(resolved)
                    elif t1 and t2:
                        wl = "Winner" if want_winner else "Loser"
                        return f"{wl} of {_title(t1)} or {_title(t2)}"

        # Pool-finish slot: 1stB- (simple) or K4(1stB) (composite)
        gm = _COMPOSITE_SLOT_RE.search(slot) or _FINISH_SLOT_RE.match(slot)
        if gm:
            group = gm.group(1).upper()
            rank_m = re.search(r'(\d+)', slot)
            rank = int(rank_m.group(1)) if rank_m else None
            ordinal_m = re.search(r'(\d+(?:st|nd|rd|th))', slot, re.IGNORECASE)
            ordinal = ordinal_m.group(1) if ordinal_m else "?"
            if rank:
                standings = _standings_for_group(group, dg)
                if standings and len(standings) >= rank:
                    any_played = any(s["wins"] > 0 or s["losses"] > 0 for s in standings)
                    if any_played:
                        return _title(standings[rank - 1]["team"])
                    else:
                        team_str = " or ".join(_title(s["team"]) for s in standings)
                        return f"{ordinal} in Pool {group} ({team_str})"

    # Fallbacks without division_games
    m = re.match(r"^([WL])#([^-\s]+)", slot, re.IGNORECASE)
    if m:
        wl = "Winner" if m.group(1).upper() == "W" else "Loser"
        return f"{wl} of game #{m.group(2).strip('#')}"
    m = re.match(r"^(\d+(?:st|nd|rd|th))([A-Z])-?\s*$", slot, re.IGNORECASE)
    if m:
        return f"{m.group(1)} in Pool {m.group(2)}"
    m = re.match(r"^[A-Z]\d+\(([^)]+)\)-?\s*$", slot, re.IGNORECASE)
    if m:
        return m.group(1)
    # CCA extended: "Winner N" or "(Winner N)" anywhere in slot
    wm = re.search(r'\bWinner\s+(\d+)', slot, re.IGNORECASE)
    if wm:
        if division_games:
            ref_game = _game_by_num(str(int(wm.group(1))), division_games)
            if ref_game:
                t1 = describe_slot(ref_game["white_team"], division_games)
                t2 = describe_slot(ref_game["dark_team"], division_games)
                if ref_game.get("played") and ref_game.get("white_score") is not None:
                    white_won = ref_game["white_score"] > ref_game["dark_score"]
                    return _title(t1 if white_won else t2)
                elif t1 and t2:
                    return f"Winner of {_title(t1)} or {_title(t2)}"
        return f"Winner of game #{wm.group(1)}"
    # CCA extended: "LN" at end of slot — e.g. "BB1-4TH G - L46"
    lm = re.search(r'\bL(\d+)\s*$', slot, re.IGNORECASE)
    if lm:
        if division_games:
            ref_game = _game_by_num(str(int(lm.group(1))), division_games)
            if ref_game:
                t1 = describe_slot(ref_game["white_team"], division_games)
                t2 = describe_slot(ref_game["dark_team"], division_games)
                if ref_game.get("played") and ref_game.get("white_score") is not None:
                    white_won = ref_game["white_score"] > ref_game["dark_score"]
                    return _title(t2 if white_won else t1)
                elif t1 and t2:
                    return f"Loser of {_title(t1)} or {_title(t2)}"
        return f"Loser of game #{lm.group(1)}"
    return slot

def _game_num(game_id: str):
    # Allows alphabetic-only suffix after digits (e.g. "18UB 402-B" from ID-collision dedup)
    m = re.search(r"(\d+)[A-Za-z-]*$", game_id)
    return str(int(m.group(1))) if m else None


def _build_division_rounds(division_games: list) -> dict[str, int]:
    """Assign round numbers to all games in a division via two-pass topological sort.

    The standard tournament structure has a gap between W#/L# bracket games and
    composite-slot placement games (Sunday AM) — no explicit W#/L# link bridges
    them. A single Kahn's pass puts composite games at round 1 (wrong). Fix:

    Pass 1  — Kahn's on non-composite games with no composite predecessors.
    Middle  — Composite-slot games assigned after Pass 1 max, sorted by time.
    Pass 2  — Remaining non-composite games (Sunday PM W#/L# referencing Sunday
              AM composite games) assigned after their now-known predecessors.

    Returns: game_id -> round (1-indexed).
    """
    by_num: dict[str, str] = {}
    for g in division_games:
        n = _game_num(g["game_id"])
        if n:
            by_num[n] = g["game_id"]

    def _is_composite(g) -> bool:
        for slot in (g["white_team"], g["dark_team"]):
            s = slot.strip()
            if _FINISH_SLOT_RE.match(s) or _COMPOSITE_SLOT_RE.search(s):
                return True
        return False

    # Build full W#/L# predecessor/successor maps
    preds: dict[str, set] = {g["game_id"]: set() for g in division_games}
    succs: dict[str, list] = {g["game_id"]: []   for g in division_games}
    for g in division_games:
        for slot in (g["white_team"], g["dark_team"]):
            s = slot.strip()
            wm = _WL_SLOT_RE.match(s)
            if wm:
                ref = re.search(r"(\d+)$", wm.group(1))
                if ref:
                    pid = by_num.get(str(int(ref.group(1))))
                    if pid and pid != g["game_id"]:
                        preds[g["game_id"]].add(pid)
                        succs[pid].append(g["game_id"])

    comp_ids    = {g["game_id"] for g in division_games if _is_composite(g)}
    non_comp    = {g["game_id"] for g in division_games} - comp_ids
    # Pass 1: non-composite games whose predecessors are all non-composite
    pass1       = {gid for gid in non_comp if not any(p in comp_ids for p in preds[gid])}

    rounds: dict[str, int] = {}
    in_deg = {gid: sum(1 for p in preds[gid] if p in pass1) for gid in pass1}
    queue  = [gid for gid, d in in_deg.items() if d == 0]
    while queue:
        gid = queue.pop(0)
        pred_r = [rounds[p] for p in preds[gid] if p in rounds]
        rounds[gid] = (max(pred_r) + 1) if pred_r else 1
        for succ in succs[gid]:
            if succ in pass1 and succ not in rounds:
                in_deg[succ] -= 1
                if in_deg[succ] == 0:
                    queue.append(succ)

    # Middle: composite-slot games sorted by (date, time); same time → same round
    max_p1 = max(rounds.values(), default=0)
    comp_sorted = sorted(
        [g for g in division_games if g["game_id"] in comp_ids],
        key=lambda g: (g.get("date") or date.min, g.get("time") or datetime.min.time()),
    )
    slot_r: dict[tuple, int] = {}
    off = 0
    for g in comp_sorted:
        key = (g.get("date"), g.get("time"))
        if key not in slot_r:
            slot_r[key] = max_p1 + 1 + off
            off += 1
        rounds[g["game_id"]] = slot_r[key]

    # Pass 2: non-composite games with composite predecessors (Sunday PM W#/L#)
    pass2 = non_comp - pass1
    changed = True
    while changed:
        changed = False
        for gid in list(pass2):
            if gid not in rounds and all(p in rounds for p in preds[gid]):
                pred_r = [rounds[p] for p in preds[gid] if p in rounds]
                rounds[gid] = (max(pred_r) + 1) if pred_r else max_p1 + 1
                changed = True

    return rounds

_POOL_SLOT_RE    = re.compile(r'^([A-Z])(\d+)-(.+)', re.IGNORECASE)
_WL_SLOT_RE      = re.compile(r'^[WL]#([^-\s]+)', re.IGNORECASE)   # dash optional (bare W#2 before scores)
_FINISH_SLOT_RE  = re.compile(r'^\d+(?:st|nd|rd|th)(?:\s+in\s+)?([A-Z])\s*-', re.IGNORECASE)
# Composite bracket slots like K4(1stG)- or K4(1stG) — group letter is inside parens
_COMPOSITE_SLOT_RE = re.compile(r'\(\d+(?:st|nd|rd|th)([A-Z])\)', re.IGNORECASE)

def _team_won(team: str, game: dict):
    """True if team won, False if lost, None if not yet played."""
    ws, ds = game.get("white_score"), game.get("dark_score")
    if ws is None or ds is None:
        return None
    yours = ws if team_matches(game["white_team"], team) else ds
    opp   = ds if team_matches(game["white_team"], team) else ws
    if yours > opp: return True
    if yours < opp: return False
    return None

def _expand_bracket_games(team: str, direct_games: list, division_games: list) -> list:
    """Return all bracket games the team can potentially reach, across all days.

    Each returned game gets a 'placeholder' bool:
      False = confirmed (result already determined this path)
      True  = possible but not yet confirmed

    Progressive narrowing: once a game is played, the losing path is dropped
    so only the correct branch remains.
    """
    if not direct_games:
        return []
    seen_ids = {g["game_id"] for g in direct_games}

    # Pool groups this team is seeded in, keyed by group letter → list of dates.
    # Tracking dates lets us scope finish-slot expansion to the same weekend,
    # preventing contamination when pool letters repeat across WPL weekends.
    groups: dict[str, list] = {}
    for g in direct_games:
        for slot in (g["white_team"], g["dark_team"]):
            m = _POOL_SLOT_RE.match(slot.strip())
            if m and team_matches(slot, team):
                letter = m.group(1).upper()
                d = g.get("date")
                grp_dates = groups.setdefault(letter, [])
                if d not in grp_dates:
                    grp_dates.append(d)

    # For seed-number prelim games (e.g. "13 - TROJAN CARDINAL"), _POOL_SLOT_RE
    # won't match and groups stays empty.  Scan downstream division games for
    # WIN/LOS GM #N references to our direct games and extract the pool letter
    # from those slots so finish-slot expansion can work.
    if not groups:
        direct_nums = {_game_num(g["game_id"]) for g in direct_games
                       if _game_num(g["game_id"])}
        for dg in division_games:
            for slot in (dg["white_team"], dg["dark_team"]):
                s = slot.strip()
                wgm = re.search(r'\b(WIN|LOS)\s+GM\s+#(\d+)', s, re.IGNORECASE)
                if wgm and str(int(wgm.group(2))) in direct_nums:
                    pm = re.match(r'^([A-Z])', s, re.IGNORECASE)
                    if pm:
                        letter = pm.group(1).upper()
                        d = dg.get("date")
                        grp_dates = groups.setdefault(letter, [])
                        if d not in grp_dates:
                            grp_dates.append(d)

    def _same_weekend(cand_date, grp_letter: str) -> bool:
        """True if cand_date is within 2 days of any date the team was in grp_letter."""
        if cand_date is None:
            return True
        grp_dates = [d for d in groups.get(grp_letter, []) if d is not None]
        return not grp_dates or any(abs((cand_date - d).days) <= 2 for d in grp_dates)

    # Pool standings per group, scoped to the same weekend(s) the team played in.
    # Cross-weekend contamination (same letter, different teams) is excluded by date.
    grp_pool_games: dict[str, list] = {}
    for g in division_games:
        wm = _POOL_SLOT_RE.match(g["white_team"].strip())
        dm = _POOL_SLOT_RE.match(g["dark_team"].strip())
        if wm and dm:
            grp = wm.group(1).upper()
            if grp in groups and _same_weekend(g.get("date"), grp):
                grp_pool_games.setdefault(grp, []).append(g)
    team_pool_ranks: dict[str, int] = {}
    for grp in groups:
        pool_games = grp_pool_games.get(grp, [])
        if pool_games and all(
            g.get("white_score") is not None and g.get("dark_score") is not None
            for g in pool_games
        ):
            standings = _standings_for_group(grp, pool_games)
            for i, s in enumerate(standings):
                if team_matches(s["team"], team):
                    team_pool_ranks[grp] = i + 1
                    break
    # Override with ranks from explicitly named finish-slot games (e.g. "2ndE-TROJAN GOLD").
    # These are more reliable than standings when pool letters repeat across WPL weekends.
    # direct_games is chronological, so later iterations win for the same group.
    for g in direct_games:
        for slot in (g["white_team"], g["dark_team"]):
            s = slot.strip()
            fm = _FINISH_SLOT_RE.match(s)
            if fm and team_matches(s, team):
                grp = fm.group(1).upper()
                rank_m = re.search(r'(\d+)', fm.group(0))
                if rank_m:
                    team_pool_ranks[grp] = int(rank_m.group(1))

    # game_num -> (game_dict, is_placeholder, ph_depth)
    # ph_depth: 0=pool slot, 1=composite slot, 2=W#/L# slot (stop expanding)
    def _direct_depth(g: dict) -> int:
        for slot in (g["white_team"], g["dark_team"]):
            if not team_matches(slot, team):
                continue
            s = slot.strip()
            if _WL_SLOT_RE.match(s):
                return 2
            if _FINISH_SLOT_RE.match(s) or _COMPOSITE_SLOT_RE.search(s):
                return 1
        return 0

    reachable: dict[str, tuple] = {}
    for g in direct_games:
        n = _game_num(g["game_id"])
        if n:
            d = _direct_depth(g)
            reachable[n] = (g, d > 0, d)

    extras = []
    # Only deduplicate true composite slots (e.g. "K4(1stG)") — simple finish slots
    # like "1stE-" can appear in multiple distinct games (WPL has 2 Sunday games per finish).
    composite_added: set = set()
    for g in direct_games:
        for slot in (g["white_team"], g["dark_team"]):
            s = slot.strip()
            cm = _COMPOSITE_SLOT_RE.search(s)
            if cm and cm.group(1).upper() in groups:
                composite_added.add(cm.group(1).upper())
    changed = True
    while changed:
        changed = False
        for g in division_games:
            if g["game_id"] in seen_ids:
                continue
            add_placeholder = None
            add_ph_depth = 1
            add_pool_rank = None
            add_grp = None

            for slot in (g["white_team"], g["dark_team"]):
                s = slot.strip()

                # Pool-finish bracket (1stA-, K4(1stG), etc.)
                fm = _FINISH_SLOT_RE.match(s) or _COMPOSITE_SLOT_RE.search(s)
                if fm and fm.group(1).upper() in groups:
                    grp = fm.group(1).upper()
                    if not _same_weekend(g.get("date"), grp):
                        continue  # different weekend — pool letter reused, skip
                    rank_m = re.search(r'(\d+)', fm.group(0))
                    rank = int(rank_m.group(1)) if rank_m else None
                    if rank and grp in team_pool_ranks and team_pool_ranks[grp] != rank:
                        break
                    # For true composite slots, show only one per group to avoid clutter.
                    # Simple finish slots (1stE-, 2ndE-, …) are allowed multiple times.
                    is_composite = bool(_COMPOSITE_SLOT_RE.search(s))
                    if is_composite and grp not in team_pool_ranks and grp in composite_added:
                        break
                    add_pool_rank = rank
                    add_grp = grp
                    add_placeholder = True
                    add_ph_depth = 2  # successors of placement games expand one more level
                    break

                # W#/L# bracket (standard W#N / L#N, plus CCA extended formats)
                wm = _WL_SLOT_RE.match(s)
                # CCA extended: "LN" at end of slot (no #) — e.g. "BB1-4TH G - L46"
                lm_ext = re.search(r'\bL(\d+)\s*$', s, re.IGNORECASE) if not wm else None
                # CCA extended: "Winner N" anywhere — e.g. "3RD G (WINNER 46)"
                wm_ext = re.search(r'\bWinner\s+(\d+)', s, re.IGNORECASE) if not wm else None
                if wm or lm_ext or wm_ext:
                    if wm:
                        ref = re.search(r'(\d+)$', wm.group(1))
                        if not ref:
                            continue
                        ref_num = str(int(ref.group(1)))
                        is_win_slot = s[0].upper() == 'W'
                    elif lm_ext:
                        ref_num = str(int(lm_ext.group(1)))
                        is_win_slot = False  # "L46" = loser
                    else:
                        ref_num = str(int(wm_ext.group(1)))
                        is_win_slot = True   # "Winner 46" = winner
                    if ref_num not in reachable:
                        continue
                    src_game, src_ph, src_depth = reachable[ref_num]
                    if src_depth >= 4:
                        continue  # stop expanding beyond 4 levels of uncertainty
                    won = _team_won(team, src_game)

                    # Drop paths made impossible by a known result
                    if won is True  and not is_win_slot: continue
                    if won is False and     is_win_slot: continue

                    ph = src_ph or (won is None)
                    add_placeholder = ph if add_placeholder is None else (add_placeholder and ph)
                    add_ph_depth = (src_depth + 1) if ph else 0
                    break

                # WIN GM #N / LOS GM #N — WPL championship prelim-to-pool slot
                # e.g. "E2 (WIN GM #399) -" or "F1 (LOS GM #399) -"
                wgm = re.search(r'\b(WIN|LOS)\s+GM\s+#(\d+)', s, re.IGNORECASE)
                if wgm:
                    ref_num = str(int(wgm.group(2)))
                    if ref_num not in reachable:
                        continue
                    src_game, src_ph, src_depth = reachable[ref_num]
                    if src_depth >= 2:
                        continue
                    want_win = wgm.group(1).upper() == "WIN"
                    won = _team_won(team, src_game)
                    if won is True  and not want_win: continue
                    if won is False and     want_win: continue
                    ph = src_ph or (won is None)
                    add_placeholder = ph if add_placeholder is None else (add_placeholder and ph)
                    add_ph_depth = (src_depth + 1) if ph else 0
                    break

            if add_placeholder is not None:
                if _game_num(g["game_id"]) is None:
                    continue  # skip non-game rows (e.g. embedded standings entries)
                # Per-day cap: max 4 depth-2+ games per calendar day so Saturday can't
                # crowd out Sunday. Only depth-2+ games count toward the cap so that
                # depth-1 direct W/L successors (which are exempt) don't consume a slot
                # and block win/lose siblings later in the same BFS pass.
                # Absolute cap of 20 as a safety net.
                _gdate = g.get("date", "")
                if _gdate and add_ph_depth >= 2:
                    _capped_day = sum(1 for e in extras
                                      if e.get("date") == _gdate and e.get("_ph_depth", 0) >= 2)
                    if _capped_day >= 4:
                        continue
                if len(extras) >= 20:
                    continue
                g_copy = dict(g)
                g_copy["placeholder"] = add_placeholder
                g_copy["_ph_depth"] = add_ph_depth
                if add_pool_rank:
                    g_copy["pool_rank"] = add_pool_rank
                if add_grp:
                    g_copy["pool_rank_group"] = add_grp
                    composite_added.add(add_grp)
                extras.append(g_copy)
                seen_ids.add(g["game_id"])
                n = _game_num(g["game_id"])
                # Don't add collision-renamed games (e.g. "18UB 389-B") to reachable
                # under their base number — their number collides with another real game
                # and would cause false expansions via WIN/LOS GM # lookup.
                is_renamed = bool(re.search(r'-[A-Z]+$', g["game_id"]))
                if n and not is_renamed:
                    reachable[n] = (g_copy, add_placeholder, add_ph_depth)
                changed = True

    return extras


def _build_wpl_game_tree(team: str, division_games: list, anchor_date=None) -> list:
    """Build a WPL weekend game tree for a team.

    Handles two formats automatically:
      - Bracket crossover (e.g. Trojan Gold): 1 named game → W#/L# Saturday bracket → finish-slot Sunday
      - Round-robin + placement (e.g. Trojan Cardinal): N named games → finish-slot placement game

    Each returned node is a game dict plus:
      placeholder    – bool
      src_game_id    – game_id of predecessor (None for root)
      src_path       – "win" | "lose" | None (None for round-robin chain)
      win_next_ids   – game_ids if team wins (may equal lose_next_ids in round-robin)
      lose_next_ids  – game_ids if team loses
      sunday_pair_id – paired Sunday game_id (for 1st/4th place doubles)
      tree_format    – "bracket" | "roundrobin"
    """
    # Sort division games chronologically once
    sorted_games = sorted(
        division_games,
        key=lambda g: (g.get("date") or date.min, g.get("time") or datetime.min.time()),
    )

    def _make_node(g, src_game_id, src_path, is_placeholder, fmt):
        node = dict(g)
        node["placeholder"]    = is_placeholder
        node["src_game_id"]    = src_game_id
        node["src_path"]       = src_path
        node["win_next_ids"]   = []
        node["lose_next_ids"]  = []
        node["sunday_pair_id"] = None
        node["tree_format"]    = fmt
        return node

    # ── 1. Find ALL explicitly-named pool-slot games for this team ──────────
    explicit_games = []
    for g in sorted_games:
        for slot in (g["white_team"], g["dark_team"]):
            m = _POOL_SLOT_RE.match(slot.strip())
            if m and team_matches(m.group(3), team):
                if g["game_id"] not in {x["game_id"] for x in explicit_games}:
                    explicit_games.append(g)
                break

    # Note: do NOT return early when explicit_games is empty. Teams whose Weekend 5
    # format uses only seed-number slots (e.g. "34 - BACK BAY") have never appeared in
    # pool-slot format, so the outer search finds nothing. The seed-number detection
    # block below handles them when anchor_date is provided.

    # Narrow to the most recent weekend cluster (WPL has 5 weekends on one sheet).
    # anchor_date: the latest date any game directly involves this team (including
    # non-pool-slot formats like "13 - TROJAN CARDINAL"). If the team's most recent
    # games use a format _POOL_SLOT_RE can't see, recent_explicit will be empty and
    # we fall back to seed-number format detection below.
    ref_date = anchor_date or (max(g["date"] for g in explicit_games if g.get("date"))
                               if any(g.get("date") for g in explicit_games) else None)
    if ref_date:
        recent_explicit = [g for g in explicit_games
                           if g.get("date") and abs((g["date"] - ref_date).days) <= 3]
    else:
        recent_explicit = explicit_games[-1:]

    # Track games added via WIN GM # inference (not direct pool-slot match).
    # These are conditional on winning the prelim and must be marked placeholder.
    inferred_game_ids: set[str] = set()

    if not recent_explicit:
        # Current weekend uses seed-number format (e.g. "13 - TROJAN CARDINAL"),
        # not pool-slot format.  Find the prelim game(s) directly, then follow
        # WIN GM #N forward to discover the pool-phase games for this weekend.
        if not anchor_date:
            return []
        def _prelim_slot_match(slot: str) -> bool:
            """Exact match after stripping slot prefix — prevents 'South Coast' from
            absorbing 'South Coast B' prelim games via substring team_matches."""
            return strip_prefix(slot).strip().upper() == team.upper()

        def _is_pool_pos_slot(s: str) -> bool:
            # Matches pool-position slots like "E2-TEAM", "E2 (WIN GM #399)-TEAM"
            # AND finish-slot positions like "2ndE-TEAM", "1stH-TEAM" — both must be
            # skipped during prelim detection to prevent them being treated as seed games.
            return bool(re.match(r'^[A-Z]\d+[-\s(]', s.strip(), re.IGNORECASE)
                        or re.match(r'^\d+(?:st|nd|rd|th)[A-Z]-', s.strip(), re.IGNORECASE))

        prelim_nums: set[str] = set()
        for g in sorted_games:
            if not (g.get("date") and abs((g["date"] - anchor_date).days) <= 3):
                continue
            wt, dt = g["white_team"], g["dark_team"]
            if not (_prelim_slot_match(wt) or _prelim_slot_match(dt)):
                continue
            # Skip pool-position slots the organiser filled in mid-tournament
            # (e.g. "F3 (LOS GM #401) - ROSE BOWL"). Adding them to recent_explicit
            # causes the downstream WIN/LOS GM # search to run against the wrong pool,
            # and their collision-renamed game IDs (e.g. "18UB 389-B" → "389") would
            # contaminate prelim_nums with a different team's prelim number.
            matching = wt if _prelim_slot_match(wt) else dt
            if _is_pool_pos_slot(matching):
                continue
            gid = g["game_id"]
            if gid not in {x["game_id"] for x in explicit_games}:
                explicit_games.append(g)
                recent_explicit.append(g)
                n = _game_num(gid)
                if n:
                    prelim_nums.add(n)
        # Check prelim result to pick the correct downstream path.
        # Pre-game (not yet played) → show WIN path as the planned scenario.
        # After a loss → follow LOS GM # consolation path instead.
        _prelim_result = (
            _team_won(team, recent_explicit[0]) if recent_explicit else None
        )
        gm_pattern = (
            r'\bLOS\s+GM\s+#(\d+)' if _prelim_result is False
            else r'\bWIN\s+GM\s+#(\d+)'
        )

        if prelim_nums:
            for g in sorted_games:
                if g["game_id"] in {x["game_id"] for x in explicit_games}:
                    continue
                if not (g.get("date") and abs((g["date"] - anchor_date).days) <= 3):
                    continue
                for slot in (g["white_team"], g["dark_team"]):
                    wgm = re.search(gm_pattern, slot.strip(), re.IGNORECASE)
                    if wgm and str(int(wgm.group(1))) in prelim_nums:
                        explicit_games.append(g)
                        recent_explicit.append(g)
                        inferred_game_ids.add(g["game_id"])
                        break
            # If WIN GM # found nothing and result is unknown, fall back to LOS GM #.
            # Some pools reference winners by relative position (L#1, L#2) rather than
            # WIN GM #N, making the win path untraceable from the prelim.  The lose
            # path uses LOS GM #N and IS traceable — show it as the planned scenario.
            if not inferred_game_ids and _prelim_result is None:
                for g in sorted_games:
                    if g["game_id"] in {x["game_id"] for x in explicit_games}:
                        continue
                    if not (g.get("date") and abs((g["date"] - anchor_date).days) <= 3):
                        continue
                    for slot in (g["white_team"], g["dark_team"]):
                        lgm = re.search(r'\bLOS\s+GM\s+#(\d+)', slot.strip(), re.IGNORECASE)
                        if lgm and str(int(lgm.group(1))) in prelim_nums:
                            explicit_games.append(g)
                            recent_explicit.append(g)
                            inferred_game_ids.add(g["game_id"])
                            break

            # Slot-position follow-up: if LOS GM # pool games were found (directly or
            # via fallback), scan for additional games by pool-position prefix (e.g. "C4").
            # This catches Sunday games whose organizer filled in stale game-number
            # references from a prior weekend instead of the correct current-weekend numbers.
            if inferred_game_ids:
                los_slot_pos: str | None = None
                existing_ids = {x["game_id"] for x in explicit_games}
                for g in explicit_games:
                    if g["game_id"] not in inferred_game_ids:
                        continue
                    for slot in (g["white_team"], g["dark_team"]):
                        lgm = re.search(r'\bLOS\s+GM\s+#(\d+)', slot.strip(), re.IGNORECASE)
                        if lgm and str(int(lgm.group(1))) in prelim_nums:
                            pm = re.match(r'^([A-Z]\d+)\s*[-\(]', slot.strip(), re.IGNORECASE)
                            if pm:
                                los_slot_pos = pm.group(1).upper()
                                break
                    if los_slot_pos:
                        break
                if los_slot_pos:
                    for g in sorted_games:
                        if g["game_id"] in existing_ids:
                            continue
                        if not (g.get("date") and abs((g["date"] - anchor_date).days) <= 3):
                            continue
                        for slot in (g["white_team"], g["dark_team"]):
                            pm = re.match(r'^([A-Z]\d+)\s*[-\(]', slot.strip(), re.IGNORECASE)
                            if pm and pm.group(1).upper() == los_slot_pos:
                                explicit_games.append(g)
                                recent_explicit.append(g)
                                inferred_game_ids.add(g["game_id"])
                                break

        if not recent_explicit:
            return []

    root = recent_explicit[0]
    root_num = _game_num(root["game_id"])
    if not root_num:
        return []

    # Pool group letter (e.g. "E" from "E3-TROJAN GOLD")
    pool_group = None
    for slot in (root["white_team"], root["dark_team"]):
        m = _POOL_SLOT_RE.match(slot.strip())
        if m and team_matches(m.group(3), team):
            pool_group = m.group(1).upper()
            break
    # Fallback: extract pool letter from WIN or LOS GM # game in recent_explicit.
    # Handles both the win path (E pool) and the lose/consolation path (F pool).
    if pool_group is None and root_num:
        for g in recent_explicit:
            if g["game_id"] == root["game_id"]:
                continue
            for slot in (g["white_team"], g["dark_team"]):
                s = slot.strip()
                wgm = re.search(r'\b(?:WIN|LOS)\s+GM\s+#(\d+)', s, re.IGNORECASE)
                if wgm and str(int(wgm.group(1))) == root_num:
                    pm = re.match(r'^([A-Z])', s, re.IGNORECASE)
                    if pm:
                        pool_group = pm.group(1).upper()
                        break
            if pool_group:
                break

    # ── 2. Check for W#/L# bracket games off the root ──────────────────────
    win_sat_game = lose_sat_game = None
    for g in sorted_games:
        if g["game_id"] in {x["game_id"] for x in explicit_games}:
            continue
        # Skip games with no date — these are bogus rows parsed from the seed table
        # (rows 24-28) which produce W#N slots that falsely trigger FORMAT A.
        if not g.get("date"):
            continue
        for slot in (g["white_team"], g["dark_team"]):
            s = slot.strip()
            wm = _WL_SLOT_RE.match(s)
            if not wm:
                continue
            ref = re.search(r'(\d+)$', wm.group(1))
            if not ref or str(int(ref.group(1))) != root_num:
                continue
            if s[0].upper() == 'W':
                win_sat_game = g
            else:
                lose_sat_game = g

    # ── 3. Finish-slot games for this pool group ─────────────────────────────
    explicit_ids = {x["game_id"] for x in explicit_games}
    bracket_ids  = {g["game_id"] for g in [win_sat_game, lose_sat_game] if g}
    root_date    = root.get("date")
    sunday_by_rank = {}   # rank(int) → [game, ...]
    if pool_group:
        for g in sorted_games:
            if g["game_id"] in explicit_ids | bracket_ids:
                continue
            # Only look within the same weekend — pool letters repeat across WPL weekends
            if root_date and g.get("date") and abs((g["date"] - root_date).days) > 3:
                continue
            for slot in (g["white_team"], g["dark_team"]):
                s = slot.strip()
                fm = _FINISH_SLOT_RE.match(s)
                if fm and fm.group(1).upper() == pool_group:
                    if _game_num(g["game_id"]) is None:
                        break  # skip embedded standings rows
                    rank_m = re.search(r'^(\d+)', s)
                    if rank_m:
                        sunday_by_rank.setdefault(int(rank_m.group(1)), []).append(g)
                    break

    out  = []
    seen = set()

    # ═══════════════════════════════════════════════════════════════════════
    # FORMAT A — Bracket crossover (win/lose Saturday split)
    # e.g. Trojan Gold: Game 330 → W#330 or L#330 → Sunday finish slots
    # ═══════════════════════════════════════════════════════════════════════
    if win_sat_game or lose_sat_game:
        finish_map = {}
        if win_sat_game:  finish_map[win_sat_game["game_id"]] = {"win": 1, "lose": 2}
        if lose_sat_game: finish_map[lose_sat_game["game_id"]] = {"win": 3, "lose": 4}

        root_result = _team_won(team, root)

        root_node = _make_node(root, None, None, False, "bracket")
        out.append(root_node); seen.add(root["game_id"])

        sat2_nodes = {}
        for sat2_game, path in [(win_sat_game, "win"), (lose_sat_game, "lose")]:
            if sat2_game is None or sat2_game["game_id"] in seen:
                continue
            if root_result is True:   is_ph = (path == "lose")
            elif root_result is False: is_ph = (path == "win")
            else:                      is_ph = True

            node = _make_node(sat2_game, root["game_id"], path, is_ph, "bracket")
            out.append(node); seen.add(sat2_game["game_id"])
            sat2_nodes[sat2_game["game_id"]] = node
            if path == "win": root_node["win_next_ids"].append(sat2_game["game_id"])
            else:             root_node["lose_next_ids"].append(sat2_game["game_id"])

        for sat2_game, _ in [(win_sat_game, "win"), (lose_sat_game, "lose")]:
            if sat2_game is None:
                continue
            sat2_node   = sat2_nodes.get(sat2_game["game_id"])
            sat2_result = _team_won(team, sat2_game)
            for outcome, rank in finish_map.get(sat2_game["game_id"], {}).items():
                sat2_ph = sat2_node["placeholder"]
                if sat2_result is True:   is_sun_ph = sat2_ph or (outcome == "lose")
                elif sat2_result is False: is_sun_ph = sat2_ph or (outcome == "win")
                else:                      is_sun_ph = True
                sun_nodes = []
                for sg in sunday_by_rank.get(rank, []):
                    if sg["game_id"] in seen: continue
                    sn = _make_node(sg, sat2_game["game_id"], outcome, is_sun_ph, "bracket")
                    out.append(sn); seen.add(sg["game_id"]); sun_nodes.append(sn)
                sun_ids = [sn["game_id"] for sn in sun_nodes]
                if outcome == "win": sat2_node["win_next_ids"].extend(sun_ids)
                else:                sat2_node["lose_next_ids"].extend(sun_ids)
                if len(sun_nodes) == 2:
                    sun_nodes[0]["sunday_pair_id"] = sun_nodes[1]["game_id"]
                    sun_nodes[1]["sunday_pair_id"] = sun_nodes[0]["game_id"]

    # ═══════════════════════════════════════════════════════════════════════
    # FORMAT B — Round-robin + placement (multiple named games, no W#/L# split)
    # e.g. Trojan Cardinal: Game 307 → Game 310 → Game 351 → placement TBD
    # ═══════════════════════════════════════════════════════════════════════
    else:
        # Inferred pool games (win or lose path) are placeholder only while the
        # prelim hasn't been played.  Once it's decided — either way — the path
        # is confirmed and those games are no longer hypothetical.
        prelim_decided = _team_won(team, recent_explicit[0]) if inferred_game_ids else None
        prev_node = None
        for g in recent_explicit:
            if g["game_id"] in seen: continue
            is_ph = (g["game_id"] in inferred_game_ids) and (prelim_decided is None)
            node = _make_node(g, prev_node["game_id"] if prev_node else None,
                              None, is_ph, "roundrobin")
            if prev_node:
                # Same next game regardless of result (round-robin)
                prev_node["win_next_ids"].append(g["game_id"])
                prev_node["lose_next_ids"].append(g["game_id"])
            out.append(node); seen.add(g["game_id"])
            prev_node = node

        # After all round-robin games, one finish-slot placement game
        if prev_node and pool_group:
            # Highest-priority: organiser has already written the team's name into the
            # specific placement slot (e.g. "1stH-TROJAN GOLD").  Read it directly —
            # no score inference needed, works even before all scores are entered.
            named_placement = None
            for rank_num, games_list in sunday_by_rank.items():
                for pg in games_list:
                    for slot in (pg["white_team"], pg["dark_team"]):
                        if team_matches(slot, team):
                            named_placement = (pg, rank_num)
                            break
                    if named_placement:
                        break
                if named_placement:
                    break

            if named_placement:
                placement_games = [named_placement]
            else:
                # Fall back to score-based rank inference (official spreadsheet scores only).
                all_played = all(g.get("white_score") is not None for g in recent_explicit)
                if all_played:
                    standings = _standings_for_group(pool_group, division_games)
                    rank = next((i + 1 for i, s in enumerate(standings)
                                 if team_matches(s["team"], team)), None)
                    # Fallback: standings can't identify team when opponents use composite
                    # slots (e.g. "H2 (WIN GM #410) - SOUTH COAST").  Use W/L count instead.
                    if rank is None and recent_explicit:
                        wins  = sum(1 for g in recent_explicit
                                    if g.get("white_score") is not None
                                    and (( team_matches(g["white_team"], team) and g["white_score"] > g["dark_score"])
                                      or (not team_matches(g["white_team"], team) and g["dark_score"] > g["white_score"])))
                        worst = max(sunday_by_rank.keys()) if sunday_by_rank else len(recent_explicit)
                        if   wins == len(recent_explicit): rank = 1
                        elif wins == 0:                    rank = worst
                        else:                              rank = 2
                    placement_games = [(pg, rank) for pg in sunday_by_rank.get(rank, [])] if rank else []
                else:
                    # Neither the slots nor scores tell us the rank yet — show all options.
                    placement_games = []
                    for rank_num, games_list in sunday_by_rank.items():
                        for pg in games_list:
                            placement_games.append((pg, rank_num))
                    placement_games.sort(
                        key=lambda x: (x[0].get("date") or date.min, x[0].get("time") or datetime.min.time()),
                    )
                    placement_games = placement_games[:4]

            for pg_item in placement_games:
                pg, pg_rank = pg_item if isinstance(pg_item, tuple) else (pg_item, None)
                if pg["game_id"] in seen: continue
                if _game_num(pg["game_id"]) is None: continue  # skip embedded standings rows
                pn = _make_node(pg, prev_node["game_id"], None, True, "roundrobin")
                pn["placement_rank"] = pg_rank
                prev_node["win_next_ids"].append(pg["game_id"])
                prev_node["lose_next_ids"].append(pg["game_id"])
                out.append(pn); seen.add(pg["game_id"])

    # Invariant: a played game is never a placeholder.
    # Scores in the spreadsheet are ground truth — clear placeholder regardless of
    # how the node was constructed. This catches any future build-time assumptions
    # that go stale once the organiser enters results.
    for node in out:
        if node.get("played"):
            node["placeholder"] = False

    return out


def _build_njo_game_tree(team: str, division_games: list, anchor_date=None) -> list:
    """Build a bracket tree for NJO / JO-Quals format games.

    Unlike WPL (which infers bracket paths from slot-reference strings), NJO
    stores explicit advancement: each game has w_to (game# winner goes to) and
    l_to (game# loser goes to).  We follow those links forward from the team's
    earliest game to produce a win-path tree.

    Each node returned has the same shape as WPL tree nodes so the frontend
    staircase renderer works unchanged.
    """
    if not division_games:
        return []

    # Build a map from game-number string → game, scoped to the same sheet
    # (GMIDs like "16B-003" encode a prefix + 3-digit number).
    def _gnum_str(gid: str):
        m = re.search(r'-(\d+)$', gid)
        return str(int(m.group(1))) if m else None

    gnum_map: dict[str, dict] = {}
    for g in division_games:
        n = _gnum_str(g["game_id"])
        if n and n not in gnum_map:
            gnum_map[n] = g

    # Find all games where this team appears
    my_games = sorted(
        [g for g in division_games if team_matches(g["white_team"], team)
         or team_matches(g["dark_team"], team)],
        key=lambda g: (g.get("date") or date.min, g.get("time") or datetime.min.time()),
    )
    if not my_games:
        return []

    # Restrict to games around the anchor date (current tournament weekend)
    if anchor_date:
        my_games = [g for g in my_games
                    if g.get("date") and abs((g["date"] - anchor_date).days) <= 3]
    if not my_games:
        return []

    def _make_njo_node(g, src_id, src_path, is_ph):
        node = dict(g)
        node["placeholder"]    = is_ph
        node["src_game_id"]    = src_id
        node["src_path"]       = src_path
        node["win_next_ids"]   = []
        node["lose_next_ids"]  = []
        node["sunday_pair_id"] = None
        node["tree_format"]    = "bracket"
        return node

    out: list[dict] = []
    seen: set[str] = set()

    def _follow(game, src_id, src_path, is_ph, depth=0):
        if depth > 8 or game["game_id"] in seen:
            return None
        seen.add(game["game_id"])
        node = _make_njo_node(game, src_id, src_path, is_ph)
        out.append(node)

        played = game.get("played", False)
        won = _team_won(team, game) if played else None  # True/False/None

        # Follow win path
        w_num = str(game.get("w_to")) if game.get("w_to") is not None else None
        if w_num and w_num in gnum_map:
            next_g = gnum_map[w_num]
            if next_g["game_id"] not in seen:
                # Only follow if team appears or result unknown
                involved = (team_matches(next_g["white_team"], team)
                            or team_matches(next_g["dark_team"], team))
                next_ph = is_ph or (won is False)  # placeholder if we lost
                if involved or won is not False:
                    child = _follow(next_g, game["game_id"], "win", next_ph, depth + 1)
                    if child:
                        node["win_next_ids"].append(next_g["game_id"])

        # Follow lose path
        l_num = str(game.get("l_to")) if game.get("l_to") is not None else None
        if l_num and l_num in gnum_map:
            next_g = gnum_map[l_num]
            if next_g["game_id"] not in seen:
                involved = (team_matches(next_g["white_team"], team)
                            or team_matches(next_g["dark_team"], team))
                next_ph = is_ph or (won is True)  # placeholder if we won
                if involved or won is not True:
                    child = _follow(next_g, game["game_id"], "lose", next_ph, depth + 1)
                    if child:
                        node["lose_next_ids"].append(next_g["game_id"])

        return node

    root = my_games[0]
    _follow(root, None, None, False)
    return out


def find_next_games(game, division_games):
    num = _game_num(game["game_id"])
    if not num:
        return None, None
    winner_next = loser_next = None
    for g in division_games:
        if g["game_id"] == game["game_id"]:
            continue
        for slot in (g["white_team"], g["dark_team"]):
            # Standard W#N / L#N format
            pm = re.match(r"^([WL])#([^-\s]+)", slot)
            if pm:
                ref = re.search(r"(\d+)$", pm.group(2))
                ref_num = str(int(ref.group(1))) if ref else pm.group(2)
                if ref_num == num:
                    if pm.group(1).upper() == "W": winner_next = g
                    else:                           loser_next  = g
                continue
            # CCA extended: "LN" (no #) at end of slot — e.g. "BB1-4TH G - L46"
            lm = re.search(r'\bL(\d+)\s*$', slot, re.IGNORECASE)
            if lm and str(int(lm.group(1))) == num:
                loser_next = g
                continue
            # CCA extended: "Winner N" / "(Winner N)" — e.g. "3RD G (WINNER 46)"
            wm = re.search(r'\bWinner\s+(\d+)', slot, re.IGNORECASE)
            if wm and str(int(wm.group(1))) == num:
                winner_next = g
    return winner_next, loser_next


def _ordinal(n: int) -> str:
    suffix = 'th' if 11 <= n % 100 <= 13 else {1:'st', 2:'nd', 3:'rd'}.get(n % 10, 'th')
    return f"{n}{suffix}"

def _futures_weekend_num(tournament_id: str) -> int | None:
    """Return the weekend number for a Futures tournament ID (futures-2 → 2), or None."""
    m = re.match(r'^futures-(\d+)$', tournament_id)
    return int(m.group(1)) if m else None


def _parse_standing_row(row, start_col: int):
    """Parse one side of a DivisionsStandings team row.
    Returns (team_str, reg_wins, sw, sl, reg_losses, points) or None."""
    try:
        team_str = str(row[start_col]).strip()
        if not team_str:
            return None
        def _num(v):
            try: return float(v) if v is not None else 0.0
            except (TypeError, ValueError): return 0.0
        rw  = _num(row[start_col + 1] if len(row) > start_col + 1 else None)
        sw  = _num(row[start_col + 2] if len(row) > start_col + 2 else None)
        sl  = _num(row[start_col + 3] if len(row) > start_col + 3 else None)
        rl  = _num(row[start_col + 4] if len(row) > start_col + 4 else None)
        pts = _num(row[start_col + 5] if len(row) > start_col + 5 else None)
        return (team_str, rw, sw, sl, rl, pts)
    except Exception:
        return None


def _read_futures_pool_standing(excel_src, weekend_num: int, team: str, sheet_name: str) -> dict | None:
    """Read a team's pool standing from the DivisionsStandings tab for a specific Futures weekend.

    Returns {pool, rank, total, reg_wins, shootout_wins, shootout_losses, reg_losses, points}
    or None if not found.
    """
    try:
        if hasattr(excel_src, 'seek'):
            excel_src.seek(0)
        import openpyxl
        wb = openpyxl.load_workbook(excel_src, data_only=True, read_only=True)
        if 'DivisionsStandings' not in wb.sheetnames:
            return None
        ws = wb['DivisionsStandings']
        rows = list(ws.iter_rows(max_row=600, values_only=True))
    except Exception:
        return None

    gender, age = _parse_sheet(sheet_name)
    if not age or not gender:
        return None
    age_key    = age.upper()     # e.g. "16U"
    gender_key = gender.upper()  # e.g. "BOYS"
    weekend_pat = rf'Weekend\s*{weekend_num}\b'
    team_upper  = team.upper().strip()

    # Find the row-range for our section (age + gender + weekend)
    section_start = section_end = None
    for i, row in enumerate(rows):
        if not row or not isinstance(row[0], str):
            continue
        cell0 = row[0].strip()
        is_section_hdr = (
            re.search(rf'\b{re.escape(age_key)}\b', cell0, re.I) and
            re.search(rf'\b{re.escape(gender_key)}\b', cell0, re.I) and
            re.search(weekend_pat, cell0, re.I)
        )
        if is_section_hdr:
            section_start = i
        elif section_start is not None and section_end is None:
            # Stop at the next different top-level section header
            if re.search(r'\d+U\s*(BOYS|GIRLS)', cell0, re.I) and re.search(r'Weekend', cell0, re.I):
                section_end = i
                break

    if section_start is None:
        return None
    if section_end is None:
        section_end = len(rows)

    # Collect all team entries from both left (cols 0-5) and right (cols 7-12) sides
    all_entries = []
    skip_patterns = [
        lambda c: re.search(r'\d+U\s*(BOYS|GIRLS)', c, re.I),        # section header
        lambda c: re.match(r'\d+u?\s*(boys|girls)\s*-\s*D\d+', c, re.I),  # div header
        lambda c: re.search(r'Regulation|Shootout', c, re.I),          # col header
    ]

    for row in rows[section_start:section_end]:
        if not row:
            continue
        # Left side (col 0)
        if isinstance(row[0], str):
            c = row[0].strip()
            if c and not any(p(c) for p in skip_patterns):
                e = _parse_standing_row(row, 0)
                if e:
                    all_entries.append(e)
        # Right side (col 7)
        if len(row) > 7 and isinstance(row[7], str):
            c = row[7].strip()
            if c and not any(p(c) for p in skip_patterns):
                e = _parse_standing_row(row, 7)
                if e:
                    all_entries.append(e)

    # Find my team entry and its pool letter prefix
    my_entry = my_pool = None
    for entry in all_entries:
        team_str = entry[0]
        clean = strip_prefix(team_str).upper()
        if team_upper in clean or clean in team_upper:
            my_entry = entry
            pm = re.match(r'^([A-Z])\d+', team_str.strip())
            my_pool = pm.group(1) if pm else None
            break

    if my_entry is None or my_pool is None:
        return None

    # Collect all teams in the same pool (same letter prefix) and rank by points
    pool_entries = [e for e in all_entries if re.match(rf'^{re.escape(my_pool)}\d+', e[0].strip())]
    pool_entries.sort(key=lambda x: -x[5])
    rank = next((i + 1 for i, e in enumerate(pool_entries) if e[0] == my_entry[0]), 1)

    _, rw, sw, sl, rl, pts = my_entry
    return {
        'pool':            my_pool,
        'rank':            rank,
        'total':           len(pool_entries),
        'reg_wins':        int(rw),
        'shootout_wins':   int(sw),
        'shootout_losses': int(sl),
        'reg_losses':      int(rl),
        'points':          int(pts),
    }


def _read_futures_team_div_map(excel_src) -> dict:
    """Return {(AGE_UPPER, GENDER_UPPER, CLEAN_NAME_UPPER): div_num} from DivisionsStandings.

    Keyed by age+gender+name so that e.g. 'Trojan Gold' in 16U Boys and 18U Boys
    are tracked separately. Most-recent weekend wins per key.
    """
    try:
        if hasattr(excel_src, 'seek'):
            excel_src.seek(0)
        import openpyxl
        wb = openpyxl.load_workbook(excel_src, data_only=True, read_only=True)
        if 'DivisionsStandings' not in wb.sheetnames:
            return {}
        ws = wb['DivisionsStandings']
        rows = list(ws.iter_rows(max_row=800, values_only=True))
    except Exception:
        return {}

    team_div: dict = {}   # (age, gender, clean) → (weekend_num, div_num)
    current_weekend = None
    current_div = None
    current_age = None
    current_gender = None

    for row in rows:
        if not row:
            continue
        cell0 = row[0] if len(row) > 0 else None
        if isinstance(cell0, str):
            c0 = cell0.strip()
            if not c0:
                continue
            wm = re.search(r'Weekend\s*(\d+)', c0, re.I)
            am = _AGE_WORDS.search(c0)
            if wm and am and re.search(r'BOYS|GIRLS|COED', c0, re.I):
                current_weekend = int(wm.group(1))
                current_age = f"{am.group(1)}U".upper()
                current_gender = ("GIRLS" if re.search(r'GIRLS', c0, re.I)
                                  else "COED" if re.search(r'COED', c0, re.I)
                                  else "BOYS")
                current_div = None
                continue
            dm = re.match(r'\d+u?\s*(boys|girls|coed)\s*-\s*D(\d+)', c0, re.I)
            if dm:
                current_div = int(dm.group(2))
                continue
        if current_weekend is None or current_div is None or current_age is None:
            continue
        for start_col in (0, 7):
            if len(row) <= start_col:
                continue
            v = row[start_col]
            if not isinstance(v, str):
                continue
            c = v.strip()
            if not c:
                continue
            if (re.search(r'\d+U\s*(BOYS|GIRLS)', c, re.I) or
                    re.match(r'\d+u?\s*(boys|girls)\s*-\s*D\d+', c, re.I) or
                    re.search(r'Regulation|Shootout', c, re.I)):
                continue
            clean = strip_prefix(c).upper().strip()
            if not clean:
                continue
            key = (current_age, current_gender, clean)
            prev = team_div.get(key)
            if prev is None or current_weekend >= prev[0]:
                team_div[key] = (current_weekend, current_div)

    return {k: v[1] for k, v in team_div.items()}


def _read_futures_cumulative_standings(excel_src, team: str, sheet_name: str,
                                        max_weekend: int | None = None):
    """Aggregate season standings for all teams in the same age/gender group across all Futures weekends.

    Teams are grouped by age/gender only (not division) so promotion/relegation between
    weekends doesn't drop any weekend's data. Returns (label, standings_list) where each
    entry has: {name, points, reg_wins, shootout_wins, shootout_losses, reg_losses, rank, total, is_mine}.
    max_weekend: if set, only include weekends ≤ this number (used to compute previous-weekend rank).
    Returns (None, None) on failure.
    """
    try:
        if hasattr(excel_src, 'seek'):
            excel_src.seek(0)
        import openpyxl
        wb = openpyxl.load_workbook(excel_src, data_only=True, read_only=True)
        if 'DivisionsStandings' not in wb.sheetnames:
            return None, None
        ws = wb['DivisionsStandings']
        rows = list(ws.iter_rows(max_row=800, values_only=True))
    except Exception:
        return None, None

    gender, age = _parse_sheet(sheet_name)
    if not age or not gender:
        return None, None
    age_key    = age.upper()
    gender_key = gender.upper()
    team_upper = team.upper().strip()

    def _is_our_section(c):
        return (re.search(rf'\b{re.escape(age_key)}\b', c, re.I) and
                re.search(rf'\b{re.escape(gender_key)}\b', c, re.I) and
                re.search(r'Weekend\s*\d+', c, re.I))

    def _is_any_section(c):
        return bool(re.search(r'\d+U\s*(BOYS|GIRLS)', c, re.I) and re.search(r'Weekend', c, re.I))

    def _div_num(c):
        m = re.match(r'\d+u?\s*(boys|girls)\s*-\s*D(\d+)', c, re.I)
        return int(m.group(2)) if m else None

    def _is_div_hdr(c):
        return _div_num(c) is not None

    def _skip(c):
        return _is_any_section(c) or _is_div_hdr(c) or bool(re.search(r'Regulation|Shootout', c, re.I))

    # Accumulate totals per team across all weekends (no division filter).
    # Also track each team's division from their most recent weekend.
    totals = {}        # clean_name → {name, points, reg_wins, …}
    team_div = {}      # clean_name → (weekend_num, div_num) — most recent weekend's div
    in_section = False
    current_weekend = None
    current_div = None

    for row in rows:
        if not row:
            continue
        cell0 = row[0] if len(row) > 0 else None
        if isinstance(cell0, str):
            c0 = cell0.strip()
            if c0:
                if _is_our_section(c0):
                    m = re.search(r'Weekend\s*(\d+)', c0, re.I)
                    current_weekend = int(m.group(1)) if m else None
                    in_section = (max_weekend is None or (current_weekend is not None and current_weekend <= max_weekend))
                    current_div = None
                    continue
                elif _is_any_section(c0):
                    in_section = False
                    current_weekend = None
                    current_div = None
                    continue
                elif in_section:
                    dn = _div_num(c0)
                    if dn is not None:
                        current_div = dn
                        continue
        if not in_section:
            continue

        for start_col in (0, 7):
            if len(row) <= start_col:
                continue
            v = row[start_col]
            if not isinstance(v, str):
                continue
            c = v.strip()
            if not c or _skip(c):
                continue
            entry = _parse_standing_row(row, start_col)
            if not entry:
                continue
            team_str, rw, sw, sl, rl, pts = entry
            clean = strip_prefix(team_str).upper().strip()
            if not clean:
                continue
            if clean not in totals:
                totals[clean] = dict(name=strip_prefix(team_str), points=0,
                                     reg_wins=0, shootout_wins=0, shootout_losses=0, reg_losses=0)
            t = totals[clean]
            t['points']          += int(pts)
            t['reg_wins']        += int(rw)
            t['shootout_wins']   += int(sw)
            t['shootout_losses'] += int(sl)
            t['reg_losses']      += int(rl)
            # Keep track of the most recent division this team appeared in
            if current_div is not None and current_weekend is not None:
                prev = team_div.get(clean)
                if prev is None or current_weekend >= prev[0]:
                    team_div[clean] = (current_weekend, current_div)

    if not totals:
        return None, None

    # Locate our team — try exact substring, then word-level match
    team_words = [w for w in team_upper.split() if len(w) > 2]
    my_clean = next((c for c in totals
                     if team_upper in c or c in team_upper
                     or (team_words and all(w in c for w in team_words))), None)

    label = f"{age} {gender.title()} · Season"

    standings = []
    for clean, t in totals.items():
        div_num = team_div.get(clean, (None, None))[1]
        row_out = dict(t, is_mine=(clean == my_clean),
                       division=f"D{div_num}" if div_num else None,
                       div_num=div_num if div_num is not None else 99)
        standings.append(row_out)

    standings.sort(key=lambda x: (x['div_num'], -x['points'], -x['reg_wins']))

    # Assign rank and total within each division separately
    from collections import Counter
    div_counts = Counter(s['div_num'] for s in standings)
    div_rank: dict = {}
    for s in standings:
        dn = s['div_num']
        div_rank[dn] = div_rank.get(dn, 0) + 1
        s['rank'] = div_rank[dn]
        s['total'] = div_counts[dn]

    return label, standings


def _tournament_records(division_games: list) -> dict:
    """W/L/T record keyed by uppercase team name for all played games in the division."""
    records: dict = {}
    for g in division_games:
        if not g.get("played") or g.get("white_score") is None:
            continue
        wt = strip_prefix(g["white_team"]).upper()
        dt = strip_prefix(g["dark_team"]).upper()
        if not wt or not dt:
            continue
        ws, ds = g["white_score"], g["dark_score"]
        for t in (wt, dt):
            records.setdefault(t, {"wins": 0, "losses": 0, "ties": 0})
        if ws > ds:
            records[wt]["wins"] += 1; records[dt]["losses"] += 1
        elif ds > ws:
            records[dt]["wins"] += 1; records[wt]["losses"] += 1
        else:
            records[wt]["ties"] += 1; records[dt]["ties"] += 1
    return records


def _infer_placement(played_out: list) -> str | None:
    """Parse the last played game's comment for an ordinal (e.g. '3rd', '13th').
    Win → that place; Loss → that place + 1."""
    if not played_out:
        return None
    last = played_out[-1]
    comment = (last.get('comments') or '').lower()
    result  = last.get('result')   # 'win' | 'loss' | 'tie'
    m = re.search(r'\b(\d+)(?:st|nd|rd|th)\b', comment)
    if not m:
        return None
    place = int(m.group(1))
    if result == 'win':  return _ordinal(place)
    if result == 'loss': return _ordinal(place + 1)
    return None


def _bracket_letter_from_comment(comment: str) -> str | None:
    """Extract bracket letter from 'I bracket I2,I3' → 'I'."""
    m = re.match(r'^([A-Z])\s+bracket\b', (comment or '').strip(), re.IGNORECASE)
    return m.group(1).upper() if m else None


def _bracket_standings(bracket_letter: str, all_div_games: list) -> dict:
    """Return {team_name: wins} for all teams in this bracket (played games only)."""
    standings: dict = {}
    for g in all_div_games:
        c = (g.get('comments') or '').strip()
        if not re.match(rf'^{re.escape(bracket_letter)}\s+bracket\b', c, re.IGNORECASE):
            continue
        if not g.get('played'):
            continue
        wt = strip_prefix(g['white_team']).upper()
        dt = strip_prefix(g['dark_team']).upper()
        standings.setdefault(wt, 0)
        standings.setdefault(dt, 0)
        ws, ds = g.get('white_score'), g.get('dark_score')
        if ws is not None and ds is not None:
            if ws > ds:   standings[wt] += 1
            elif ds > ws: standings[dt] += 1
    return standings


def _read_bracket_range(excel_src, sheet_name: str, bracket_letter: str) -> tuple | None:
    """Read bracket placement range from Excel header row.
    Matches cells like 'I (7th-12th)' → returns (7, 12)."""
    try:
        if hasattr(excel_src, 'seek'):
            excel_src.seek(0)
        import openpyxl
        wb = openpyxl.load_workbook(excel_src, data_only=True, read_only=True)
        if sheet_name not in wb.sheetnames:
            return None
        ws = wb[sheet_name]
        for row in ws.iter_rows(max_row=30, values_only=True):
            for cell in row:
                if not isinstance(cell, str):
                    continue
                m = re.match(
                    rf'^{re.escape(bracket_letter)}\s*\((\d+)(?:st|nd|rd|th)[-–](\d+)(?:st|nd|rd|th)\)',
                    cell.strip(), re.IGNORECASE,
                )
                if m:
                    return int(m.group(1)), int(m.group(2))
    except Exception:
        pass
    return None


def _estimate_placement(played_out: list, all_div_games: list,
                        excel_src, sheet_name: str, team: str) -> str | None:
    """Fallback: estimate placement from bracket standings when no explicit ordinal exists.
    Returns e.g. '8th (estimated)' or None."""
    if not played_out:
        return None
    last = played_out[-1]
    bracket = _bracket_letter_from_comment(last.get('comments') or '')
    if not bracket:
        return None

    standings = _bracket_standings(bracket, all_div_games)
    if not standings:
        return None

    # Rank teams by wins (descending); find this team's position
    my_key = team.upper().strip()
    ranked = sorted(standings.items(), key=lambda x: -x[1])
    rank = next(
        (i + 1 for i, (t, _) in enumerate(ranked) if my_key in t or t in my_key),
        None,
    )
    if rank is None:
        return None

    # Map bracket rank → overall place using the range from the Excel header
    bracket_range = _read_bracket_range(excel_src, sheet_name, bracket)
    place = (bracket_range[0] + (rank - 1)) if bracket_range else rank
    return f"{_ordinal(place)} (estimated)"

def _fmt_time(t) -> str:
    return t.strftime("%I:%M %p").lstrip("0") if t else "TBD"

def _fmt_date(d) -> str:
    return d.strftime("%A, %b %-d") if d else "TBD"

def _fmt_score(game) -> str:
    ws, ds = game["white_score"], game["dark_score"]
    if ws is None or ds is None: return None
    return f"{ws} – {ds}"

def _result_str(game, team) -> str:
    ws, ds = game["white_score"], game["dark_score"]
    if ws is None or ds is None: return None
    your_white = team_matches(game["white_team"], team)
    yours = ws if your_white else ds
    opp   = ds if your_white else ws
    if yours > opp:  return "win"
    if yours < opp:  return "loss"
    return "tie"

def _team_opp_slot(g: dict, team: str, dg: list, my_game_ids: set = None) -> str:
    """Return the opponent's slot for game g.

    Handles three cases where team_matches() alone fails:
      1. Direct name match (normal games) — delegates to team_matches
      2. Finish-slot placeholder (1stG-, 2ndH-) — uses pool_rank_group stored by
         _expand_bracket_games to identify which slot is the team's
      3. W#/L# placeholder (L#23, W#38) — checks if the referenced game is in
         the team's known game list (my_game_ids) to identify the team's slot
    """
    white, dark = g["white_team"], g["dark_team"]

    # Direct name match
    if team_matches(white, team): return dark
    if team_matches(dark, team):  return white

    # Finish-slot: pool_rank_group tells us which pool letter is ours
    pr  = g.get("pool_rank")
    grp = g.get("pool_rank_group")
    if pr and grp:
        for slot, other in ((white, dark), (dark, white)):
            fm = _FINISH_SLOT_RE.match(slot.strip())
            if fm and fm.group(1).upper() == grp:
                rank_m = re.search(r'(\d+)', slot)
                if rank_m and int(rank_m.group(1)) == pr:
                    return other

    # W#/L# reference: team is in the slot that points at one of their known games
    if my_game_ids:
        for slot, other in ((white, dark), (dark, white)):
            # Standard W#N / L#N
            wm = re.match(r'^[WL]#(\d+)', slot, re.IGNORECASE)
            if wm:
                ref_num = str(int(wm.group(1)))
                ref_game = _game_by_num(ref_num, dg)
                if ref_game and ref_game["game_id"] in my_game_ids:
                    return other
            # CCA extended: "LN" at end of slot — e.g. "BB1-4TH G - L46"
            lm = re.search(r'\bL(\d+)\s*$', slot, re.IGNORECASE)
            if lm:
                ref_num = str(int(lm.group(1)))
                ref_game = _game_by_num(ref_num, dg)
                if ref_game and ref_game["game_id"] in my_game_ids:
                    return other
            # CCA extended: "Winner N" — e.g. "3RD G (WINNER 46)"
            wm2 = re.search(r'\bWinner\s+(\d+)', slot, re.IGNORECASE)
            if wm2:
                ref_num = str(int(wm2.group(1)))
                ref_game = _game_by_num(ref_num, dg)
                if ref_game and ref_game["game_id"] in my_game_ids:
                    return other

    # None of the strategies resolved the slot — log so we know about new formats.
    import logging as _logging
    _logging.warning(
        "_team_opp_slot: unresolved slot for team=%r game=%s white=%r dark=%r pr=%r grp=%r",
        team, g.get("game_id"), white, dark, g.get("pool_rank"), g.get("pool_rank_group"),
    )
    return white  # default: team is dark, opponent is white


def _next_summary(g, team, dg=None, ref_date=None, my_game_ids=None):
    opp = _team_opp_slot(g, team, dg or [], my_game_ids)
    return {
        "opponent": describe_slot(opp, dg, ref_date=ref_date),
        "date":     _fmt_date(g["date"]),
        "time":     _fmt_time(g["time"]),
        "location": g["location"],
    }


# ── Results store ──────────────────────────────────────────────────────────────

def results_path(tournament_id: str) -> str:
    return os.path.join(RESULTS_DIR, f"{tournament_id}.json")

def load_results(tournament_id: str) -> dict:
    path = results_path(tournament_id)
    if os.path.exists(path):
        with open(path) as f:
            return json.load(f)
    return {}

def save_results(tournament_id: str, data: dict):
    with open(results_path(tournament_id), "w") as f:
        json.dump(data, f, indent=2, default=str)


# ── Routes ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/tournaments")
def api_tournaments():
    today = datetime.now(ZoneInfo('America/Los_Angeles')).date()
    out = []
    for t in KNOWN_TOURNAMENTS:
        excel = find_excel(t["id"])
        # has_file: Excel exists; has_excel: Excel has games in this tournament's date range
        has_file  = excel is not None
        has_excel = False
        if excel:
            try:
                games = _filter_by_dates(load_and_parse(excel), t["id"])
                has_excel = len(games) > 0
            except Exception:
                has_file = False
        year_match  = _RE_YEAR.search(t["dates"])
        month_match = _RE_MONTH.search(t["dates"])
        yr = int(year_match.group(1)) if year_match else today.year
        mo = _MONTH_MAP.get(month_match.group(1).lower(), 1) if month_match else 1
        t_date = date(yr, mo, 1)
        start = t.get("date_start") or t_date
        end   = t.get("date_end")   or start
        is_past = end < today
        days_until = (start - today).days if not is_past else None
        out.append({**t, "has_excel": has_excel, "has_file": has_file,
                    "has_preset_url": t["id"] in PRESET_URL_TOURNAMENTS,
                    "past": is_past, "days_until": days_until,
                    "_sort_date": t_date})
    # Sort: upcoming first (chronological), past last (reverse chronological)
    out.sort(key=lambda x: (x["past"], x["_sort_date"] if not x["past"] else -x["_sort_date"].toordinal()))
    for x in out: del x["_sort_date"]
    return jsonify(out)


@app.route("/api/trojan-teams/<tournament_id>")
def api_trojan_teams(tournament_id):
    excel = find_excel(tournament_id)
    if not excel:
        return jsonify([])
    games = _filter_by_dates(load_and_parse(excel), tournament_id)
    # Count games per (name, sheet) so we can pick the best sheet when duplicates exist
    counts = {}
    for g in games:
        for slot in (g["white_team"], g["dark_team"]):
            name = strip_prefix(slot)
            if "TROJAN" in name.upper():
                key = (name, g["sheet"])
                counts[key] = counts.get(key, 0) + 1

    # Build candidate list
    candidates = []
    for (name, sheet), count in counts.items():
        candidates.append({"name": name, "sheet": sheet,
                           "friendly": friendly_team_name(name, sheet), "_count": count})

    # Deduplicate by friendly name — same team may appear in multiple bracket splits
    # Keep the sheet with the most games for that team
    deduped = {}
    for t in candidates:
        key = t["friendly"] or f"{t['name']}|{t['sheet']}"
        if key not in deduped or t["_count"] > deduped[key]["_count"]:
            deduped[key] = t

    div_map = {}
    if tournament_id in WPL_TOURNAMENTS and excel:
        if hasattr(excel, 'seek'):
            excel.seek(0)
        div_map = _read_futures_team_div_map(excel)

    def _team_fields(t):
        gender, age = _parse_sheet(t["sheet"])
        tier = _parse_tier(t["sheet"])
        age_gender = " ".join(p for p in [gender, age] if p) or None
        clean = strip_prefix(t["name"]).upper().strip()
        key = (age.upper() if age else "", (gender.upper() if gender else ""), clean)
        wpl_div = div_map.get(key)
        return {"name": t["name"], "sheet": t["sheet"], "friendly": t["friendly"],
                "age_gender": age_gender, "tier": tier, "wpl_div": wpl_div}

    teams = [_team_fields(t) for t in deduped.values()]
    teams.sort(key=_team_sort_key)
    return jsonify(teams)


def _run_bracket_llm_check(team: str, nodes: list, warnings: list) -> None:
    """Background: ask Haiku to explain bracket issues in plain English. Logs only."""
    try:
        import anthropic
        client = anthropic.Anthropic()
        summary = json.dumps([{
            "game_id":      n["game_id"],
            "date":         n.get("date"),
            "opponent":     n.get("opponent"),
            "placeholder":  n.get("placeholder"),
            "win_next_ids": n.get("win_next_ids"),
            "lose_next_ids":n.get("lose_next_ids"),
        } for n in nodes], indent=2)
        resp = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=200,
            messages=[{"role": "user", "content":
                f'Water polo bracket for "{team}" has issues: {warnings}\n\nNodes:\n{summary}\n\n'
                f'In 1-2 sentences: what is wrong and what should the bracket look like?'
            }]
        )
        print(f"[bracket-judge] {team}: {resp.content[0].text}")
    except Exception as exc:
        print(f"[bracket-judge] failed: {exc}")


_SLOT_LIKE_RE = re.compile(
    r'^(?:\d+(?:st|nd|rd|th)(?:\s+in\s+)?[A-Z]-|[A-Z]\d+[-\(]|[WL]#|WIN\s+GM|LOS\s+GM)',
    re.IGNORECASE,
)


def _derive_expected_bracket(team: str, div_games: list, anchor_date) -> dict:
    """Read the schedule and derive exactly what the bracket tree should contain.

    This is the ground truth: we trace the game graph from the schedule itself,
    independent of _build_wpl_game_tree, and return what the tree SHOULD look like.

    Returns a dict:
      prelim_ids      – game_ids where team appears by name this weekend
      win_pool_ids    – game_ids reachable via WIN GM # references from prelims
      placement_ids   – finish-slot placement game_ids for the team's pool
      expected_depth  – number of unique sequential rounds (prelim + pool + 1 placement slot)
      pool_letter     – pool letter (e.g. 'E'), or None if not found
    """
    if not anchor_date:
        return {}

    def _within(g) -> bool:
        return bool(g.get("date") and abs((g["date"] - anchor_date).days) <= 3)

    def _exact_match(slot: str, t: str) -> bool:
        """Exact team name match after stripping slot prefix.
        Prevents 'South Coast' from absorbing 'South Coast B' games via substring."""
        return strip_prefix(slot).strip().upper() == t.upper()

    # Step 1: direct games — team appears by name this weekend.
    # Use exact matching (not substring) so 'South Coast' doesn't absorb 'South Coast B'.
    direct = [
        g for g in div_games
        if _within(g) and (_exact_match(g["white_team"], team)
                           or _exact_match(g["dark_team"], team))
    ]
    if not direct:
        return {}

    seen_ids = {g["game_id"] for g in direct}

    # Step 2: detect format — pool-seed (E1-IMPERIAL) vs prelim (13 - TROJAN CARDINAL)
    pool_letter = None
    for g in direct:
        for slot in (g["white_team"], g["dark_team"]):
            m = _POOL_SLOT_RE.match(slot.strip())
            if m and team_matches(slot, team):
                pool_letter = m.group(1).upper()
                break
        if pool_letter:
            break

    if pool_letter:
        # Pool-seed team: direct games ARE their pool games; no prelim round.
        # Find additional pool games for this pool that the team plays (WIN GM #
        # from their direct game numbers, or pool-letter second-day games).
        direct_nums = {_game_num(g["game_id"]) for g in direct if _game_num(g["game_id"])}
        extra_pool = []
        for g in div_games:
            if g["game_id"] in seen_ids or not _within(g):
                continue
            for slot in (g["white_team"], g["dark_team"]):
                pm = _POOL_SLOT_RE.match(slot.strip())
                if pm and pm.group(1).upper() == pool_letter and team_matches(slot, team):
                    extra_pool.append(g)
                    seen_ids.add(g["game_id"])
                    break
        pool_games = direct + extra_pool
        prelim_games: list = []
        win_pool = pool_games
    else:
        # Prelim-path team: direct games are prelims; find WIN GM # pool games.
        # Mid-tournament the organiser fills team names into pool-position slots like
        # "B3 (WIN GM #392) - LA JOLLA UNITED".  Those games appear in `direct` via
        # _exact_match but are NOT prelims.  Filter them out so prelim_nums contains
        # only the actual prelim game number, and remove pool-position games from
        # seen_ids so the WIN GM # search below can still find them.
        def _is_pool_position_game(g: dict) -> bool:
            for s in (g["white_team"], g["dark_team"]):
                if re.match(r'^[A-Z]\d+[-\s(]', s.strip(), re.IGNORECASE):
                    return True
            return False
        prelim_games = [g for g in direct if not _is_pool_position_game(g)] or direct
        pool_pos_ids = {g["game_id"] for g in direct if _is_pool_position_game(g)}
        seen_ids -= pool_pos_ids   # allow WIN GM # search to find these games
        prelim_nums  = {_game_num(g["game_id"]) for g in prelim_games
                        if _game_num(g["game_id"])}
        win_pool = []
        for g in div_games:
            if g["game_id"] in seen_ids or not _within(g):
                continue
            for slot in (g["white_team"], g["dark_team"]):
                wgm = re.search(r'\bWIN\s+GM\s+#(\d+)', slot, re.IGNORECASE)
                if wgm and str(int(wgm.group(1))) in prelim_nums:
                    win_pool.append(g)
                    seen_ids.add(g["game_id"])
                    break
        # Extract pool letter from win-pool slots
        for g in win_pool:
            for slot in (g["white_team"], g["dark_team"]):
                wgm = re.search(r'\bWIN\s+GM\s+#(\d+)', slot, re.IGNORECASE)
                if wgm and str(int(wgm.group(1))) in prelim_nums:
                    pm = re.match(r'^([A-Z])', slot.strip(), re.IGNORECASE)
                    if pm:
                        pool_letter = pm.group(1).upper()
                        break
            if pool_letter:
                break

    # Step 3: finish-slot placement games for this pool
    placement = []
    if pool_letter:
        for g in div_games:
            if g["game_id"] in seen_ids or not _within(g):
                continue
            for slot in (g["white_team"], g["dark_team"]):
                fm = _FINISH_SLOT_RE.match(slot.strip())
                if fm and fm.group(1).upper() == pool_letter:
                    if _game_num(g["game_id"]):
                        placement.append(g)
                        seen_ids.add(g["game_id"])
                    break

    # ── Step 4 (prelim-path teams only): trace the lose / consolation path ───
    # Use a fresh seen set so win-path games don't block lose-path discovery.
    lose_pool: list = []
    lose_pool_letter: str | None = None
    lose_placement: list = []

    if prelim_games:
        lose_seen = {g["game_id"] for g in prelim_games}
        for g in div_games:
            if g["game_id"] in lose_seen or not _within(g):
                continue
            for slot in (g["white_team"], g["dark_team"]):
                lgm = re.search(r'\bLOS\s+GM\s+#(\d+)', slot, re.IGNORECASE)
                if lgm and str(int(lgm.group(1))) in prelim_nums:
                    lose_pool.append(g)
                    lose_seen.add(g["game_id"])
                    if not lose_pool_letter:
                        pm = re.match(r'^([A-Z])', slot.strip(), re.IGNORECASE)
                        if pm:
                            lose_pool_letter = pm.group(1).upper()
                    break

        # After the LOS GM # search, also scan by pool-position slot (e.g. "C4")
        # to catch Sunday games whose organizer filled in stale game numbers from a
        # prior weekend instead of the correct current-weekend numbers.
        if lose_pool:
            lose_slot_pos: str | None = None
            for g in lose_pool:
                for slot in (g["white_team"], g["dark_team"]):
                    lgm = re.search(r'\bLOS\s+GM\s+#(\d+)', slot, re.IGNORECASE)
                    if lgm and str(int(lgm.group(1))) in prelim_nums:
                        pm = re.match(r'^([A-Z]\d+)\s*[-\(]', slot.strip(), re.IGNORECASE)
                        if pm:
                            lose_slot_pos = pm.group(1).upper()
                            break
                if lose_slot_pos:
                    break
            if lose_slot_pos:
                for g in div_games:
                    if g["game_id"] in lose_seen or not _within(g):
                        continue
                    for slot in (g["white_team"], g["dark_team"]):
                        pm = re.match(r'^([A-Z]\d+)\s*[-\(]', slot.strip(), re.IGNORECASE)
                        if pm and pm.group(1).upper() == lose_slot_pos:
                            lose_pool.append(g)
                            lose_seen.add(g["game_id"])
                            break

        if lose_pool_letter:
            for g in div_games:
                if g["game_id"] in lose_seen or not _within(g):
                    continue
                for slot in (g["white_team"], g["dark_team"]):
                    fm = _FINISH_SLOT_RE.match(slot.strip())
                    if fm and fm.group(1).upper() == lose_pool_letter:
                        if _game_num(g["game_id"]):
                            lose_placement.append(g)
                            lose_seen.add(g["game_id"])
                        break

    # ── Compute expected depths ────────────────────────────────────────────
    prelim_slots      = len({(g["date"], g["time"]) for g in prelim_games})
    pool_slots        = len({(g["date"], g["time"]) for g in win_pool})
    placement_slot    = 1 if placement else 0
    expected_depth    = prelim_slots + pool_slots + placement_slot

    lose_pool_slots      = len({(g["date"], g["time"]) for g in lose_pool})
    lose_placement_slot  = 1 if lose_placement else 0
    lose_expected_depth  = prelim_slots + lose_pool_slots + lose_placement_slot

    return {
        "prelim_ids":          [g["game_id"] for g in prelim_games],
        "win_pool_ids":        [g["game_id"] for g in win_pool],
        "placement_ids":       [g["game_id"] for g in placement],
        "expected_depth":      expected_depth,
        "pool_letter":         pool_letter,
        "lose_pool_ids":       [g["game_id"] for g in lose_pool],
        "lose_placement_ids":  [g["game_id"] for g in lose_placement],
        "lose_pool_letter":    lose_pool_letter,
        "lose_expected_depth": lose_expected_depth,
        "is_pool_seed":        bool(pool_letter and not prelim_games),
    }


def _check_bracket_structure(team: str, tree: list,
                              div_games: list = None, anchor_date=None) -> list[str]:
    """Deterministic structural assertions on the WPL bracket tree. Never calls an LLM.

    Pass div_games + anchor_date to also validate against the schedule ground truth —
    this catches mismatches like an empty tree when the schedule expects 4 rounds,
    or a tree with too few/many depths relative to what the schedule specifies.

    Structural checks (always run):
      - Empty tree  → wpl_bracket null → frontend falls back to raw chronological
        game_nums, producing non-sequential GAME labels like GAME 2 → GAME 6 → GAME 7.
      - Single node → prelim only; WIN GM # expansion failed.
      - Multiple roots, dangling next_ids, duplicate game_ids, node count > 10.

    Ground-truth checks (when div_games + anchor_date provided):
      - Expected bracket depth (from schedule) vs actual tree depth.
      - Saturday and Sunday both present.
      - All win-path schedule games accounted for in tree.
    """
    issues = []

    # ── Structural checks ──────────────────────────────────────────────────────
    if not tree:
        issues.append(
            "tree is empty — wpl_bracket will be null; frontend falls back to "
            "chronological game_nums including both win-path and lose-path games, "
            "producing non-sequential GAME labels (e.g. GAME 2 → GAME 6 → GAME 7)"
        )
        if div_games and anchor_date:
            expected = _derive_expected_bracket(team, div_games, anchor_date)
            if expected.get("expected_depth"):
                d = expected["expected_depth"]
                issues[-1] += (
                    f"; schedule expects {d} round(s): "
                    f"prelim={expected['prelim_ids']}, "
                    f"pool={expected['win_pool_ids']}, "
                    f"placement={len(expected['placement_ids'])} option(s)"
                )
        return issues

    if len(tree) == 1:
        issues.append(
            f"tree has only 1 node ({tree[0]['game_id']}) — prelim only; "
            "WIN GM # pool-phase expansion failed; ensure div_games_for_tree "
            "uses _all_games (unfiltered), not the date-filtered subset"
        )

    roots = [n for n in tree if not n.get("src_game_id")]
    if len(roots) != 1:
        issues.append(
            f"expected exactly 1 root node, found {len(roots)}: "
            f"{[r['game_id'] for r in roots]}"
        )

    tree_ids = {n["game_id"] for n in tree}
    dangling = {
        nid
        for n in tree
        for nid in (n.get("win_next_ids") or []) + (n.get("lose_next_ids") or [])
        if nid not in tree_ids
    }
    if dangling:
        issues.append(f"dangling next_ids (referenced but not in tree): {sorted(dangling)}")

    seen: set = set()
    dupes: set = set()
    for n in tree:
        gid = n["game_id"]
        if gid in seen:
            dupes.add(gid)
        seen.add(gid)
    if dupes:
        issues.append(f"duplicate game_ids in tree: {sorted(dupes)}")

    if len(tree) > 10:
        issues.append(
            f"tree has {len(tree)} nodes — expected ≤10 for any WPL weekend format; "
            "possible wrong anchor_date or runaway expansion"
        )

    # ── Ground-truth checks ────────────────────────────────────────────────────
    if div_games and anchor_date:
        expected = _derive_expected_bracket(team, div_games, anchor_date)
        if expected:
            # Determine which path the tree should reflect.
            # Prelim not played → show win path (planned).  Won → win path confirmed.
            # Lost → lose/consolation path.  Pool-seed teams have no prelim.
            prelim_result: bool | None = None
            if expected.get("prelim_ids") and tree:
                root = next((n for n in tree if not n.get("src_game_id")), None)
                if root:
                    prelim_result = _team_won(team, root)

            on_lose_path  = (prelim_result is False)
            # If the WIN path is untraceable (pool uses relative L#N positions rather
            # than WIN GM #N), fall back to the lose path as ground truth.  This avoids
            # a spurious depth mismatch when the tree builder correctly populates the
            # lose-path games (the only traceable pool phase for such teams).
            # Win path untraceable: pool uses relative L#N positions, not WIN GM #N.
            # Only fall back to lose path as ground truth when:
            #   - exactly 1 prelim game (multiple = organiser filled in pool names,
            #     not a real prelim set; those games were already removed from prelim_ids)
            #   - team did NOT win their prelim (a winner's untraceable L pool is a
            #     known limitation, not a lose-path problem)
            win_path_missing = (not expected.get("win_pool_ids")
                                and not expected.get("is_pool_seed")
                                and expected.get("lose_pool_ids")
                                and len(expected.get("prelim_ids", [])) == 1
                                and prelim_result is not True)
            if win_path_missing:
                on_lose_path = True
            exp_pool_ids  = (expected["lose_pool_ids"] if on_lose_path
                             else expected["win_pool_ids"])
            exp_plac_ids  = (expected["lose_placement_ids"] if on_lose_path
                             else expected["placement_ids"])
            exp_depth     = (expected["lose_expected_depth"] if on_lose_path
                             else expected["expected_depth"])

            # Unique sequential rounds: unique (date, time) slots in tree,
            # with placement alternatives (different times, mutually exclusive)
            # counted as just 1 round.
            tree_slots: dict = {}
            for n in tree:
                key = (n.get("date"), n.get("time"))
                if key[0]:
                    tree_slots[key] = tree_slots.get(key, 0) + 1
            actual_unique_slot_count = len(tree_slots)
            placement_in_tree = len(exp_plac_ids)
            adjusted_depth = actual_unique_slot_count - max(0, placement_in_tree - 1)

            if adjusted_depth != exp_depth:
                path_label = "lose" if on_lose_path else "win"
                issues.append(
                    f"schedule expects {exp_depth} sequential round(s) on {path_label} path "
                    f"(prelim={len(expected['prelim_ids'])}, "
                    f"pool={len(exp_pool_ids)}, "
                    f"placement=1 of {len(exp_plac_ids)}), "
                    f"but tree has {adjusted_depth} unique rounds "
                    f"({actual_unique_slot_count} raw slots, "
                    f"{placement_in_tree} placement options in tree)"
                )

            # All schedule games for the active path should appear in the tree
            tree_ids_all = {n["game_id"] for n in tree}
            for gid in expected["prelim_ids"] + exp_pool_ids:
                if gid not in tree_ids_all:
                    path_label = "lose" if on_lose_path else "win"
                    issues.append(
                        f"schedule game {gid!r} is on the {path_label} path "
                        f"but missing from tree"
                    )

            # Lose path must exist in the schedule for prelim-path teams that use
            # WIN GM # format.  Teams using W#/L# FORMAT A bracket crossover have
            # their consolation in L# slots (not LOS GM #) — don't warn for those.
            uses_win_gm_format = bool(expected.get("win_pool_ids"))
            if (expected.get("prelim_ids") and not expected.get("is_pool_seed")
                    and uses_win_gm_format):
                if not expected["lose_pool_ids"]:
                    issues.append(
                        f"lose path (LOS GM # consolation games) not found in schedule "
                        f"for prelim game(s) {expected['prelim_ids']} — "
                        f"if team loses the prelim, the app will show only 1 game"
                    )
                elif expected["lose_expected_depth"] != expected["expected_depth"]:
                    issues.append(
                        f"unbalanced paths: win depth={expected['expected_depth']} "
                        f"vs lose depth={expected['lose_expected_depth']} — "
                        f"one path has fewer games than the other"
                    )

            # Saturday AND Sunday should both be present
            sat = [n for n in tree if n.get("date") and n["date"].weekday() == 5]
            sun = [n for n in tree if n.get("date") and n["date"].weekday() == 6]
            if not sat:
                issues.append("no Saturday games in tree — Saturday data missing or wrong anchor")
            if not sun:
                issues.append("no Sunday games in tree — Sunday data missing from spreadsheet")

    return issues


def _bracket_has_cycle(nodes: list) -> bool:
    """DFS cycle detection on bracket graph. Returns True if a cycle exists."""
    node_map = {n["game_id"]: n for n in nodes}
    visited: set = set()
    rec_stack: set = set()

    def dfs(gid: str) -> bool:
        visited.add(gid)
        rec_stack.add(gid)
        node = node_map.get(gid)
        if node:
            for nid in (node.get("win_next_ids") or []) + (node.get("lose_next_ids") or []):
                if nid not in visited:
                    if dfs(nid):
                        return True
                elif nid in rec_stack:
                    return True
        rec_stack.discard(gid)
        return False

    for n in nodes:
        if n["game_id"] not in visited:
            if dfs(n["game_id"]):
                return True
    return False


def _validate_wpl_bracket(team: str, nodes: list, upcoming: list = None,
                           serialized_nodes: list = None) -> tuple[str, list]:
    """Deterministic bracket integrity checks.

    Returns (confidence, warnings) where confidence is 'green' | 'yellow' | 'red'.
    RED  → bracket is not trustworthy; caller must force display_mode='flat_schedule'.
    YELLOW → bracket is usable but has warnings visible to admin/debug users.
    GREEN  → bracket is structurally sane and safe to render.

    Fires a background LLM diagnosis when issues are found.
    """
    red: list   = []
    yellow: list = []

    # ── Empty tree ─────────────────────────────────────────────────────────────
    if not nodes:
        red.append("bracket is empty — wpl_bracket could not be built")
        _fire_llm_check(team, nodes, red)
        return "red", red

    all_ids    = {n["game_id"] for n in nodes}
    seen_ids: set = set()
    node_map   = {n["game_id"]: n for n in nodes}
    non_placeholder = [n for n in nodes if not n.get("placeholder")]

    # ── Duplicate game IDs ─────────────────────────────────────────────────────
    for n in nodes:
        gid = n["game_id"]
        if gid in seen_ids:
            red.append(f"duplicate game_id {gid!r} in bracket")
        seen_ids.add(gid)

    # ── Dangling next_ids ──────────────────────────────────────────────────────
    for n in nodes:
        gid = n["game_id"]
        for ref_id in (n.get("win_next_ids") or []):
            if ref_id not in all_ids:
                red.append(f"game {gid}: win_next_id {ref_id!r} missing from tree")
        for ref_id in (n.get("lose_next_ids") or []):
            if ref_id not in all_ids:
                red.append(f"game {gid}: lose_next_id {ref_id!r} missing from tree")

    # ── Cycle detection ────────────────────────────────────────────────────────
    if _bracket_has_cycle(nodes):
        red.append("bracket graph contains a cycle — impossible bracket structure")

    # ── Node count ceiling ────────────────────────────────────────────────────
    # FORMAT A max = 7 (root + 2 Saturday + 4 Sunday).
    # FORMAT B max ≈ 5–6.  Anything above 7 = runaway expansion.
    if len(nodes) > 7:
        red.append(
            f"bracket has {len(nodes)} nodes — expected ≤7; "
            "likely a mis-identified prelim or finish-slot pulling in unrelated games"
        )

    # ── Played game still placeholder (invariant violation) ───────────────────
    for n in nodes:
        if n.get("played") and n.get("placeholder"):
            red.append(
                f"game {n['game_id']}: played=True but placeholder=True — "
                "score exists in spreadsheet but node will render as upcoming"
            )

    # ── Played game missing score ──────────────────────────────────────────────
    for n in nodes:
        if n.get("played") and not n.get("placeholder"):
            if n.get("white_score") is None or n.get("dark_score") is None:
                red.append(
                    f"game {n['game_id']}: played=True but score is None — "
                    "game marked played with no score data"
                )

    # ── All nodes are placeholders ─────────────────────────────────────────────
    if not non_placeholder:
        red.append("all bracket nodes are placeholders — no real games confirmed")

    # ── Self-referential opponent ──────────────────────────────────────────────
    for n in nodes:
        if not n.get("placeholder") and team_matches(n.get("opponent", ""), team):
            red.append(f"game {n['game_id']}: opponent resolves to own team")

    # ── Foreign game detection ─────────────────────────────────────────────────
    # Each non-root node must name the team directly OR reference a tree game
    # via WIN GM #N / LOS GM #N.
    tree_game_nums = {_game_num(n["game_id"]) for n in nodes} - {None}
    for n in nodes:
        if not n.get("src_game_id"):
            continue  # root exempt
        gid = n["game_id"]
        direct = (team_matches(n.get("white_team", ""), team)
                  or team_matches(n.get("dark_team", ""), team))
        if not direct:
            ref_found = any(
                (wgm := re.search(r'\b(?:WIN|LOS)\s+GM\s+#(\d+)', slot, re.IGNORECASE))
                and str(int(wgm.group(1))) in tree_game_nums
                for slot in (n.get("white_team", ""), n.get("dark_team", ""))
            )
            if not ref_found:
                red.append(
                    f"game {gid}: foreign game — neither slot names {team!r} "
                    f"nor references a tree game via WIN/LOS GM # "
                    f"(slots: {n.get('white_team')!r} / {n.get('dark_team')!r})"
                )

    # ── Chronological monotonicity ─────────────────────────────────────────────
    # Walking win_next_ids from root, each node's datetime must be ≥ parent's.
    root_node = next((n for n in nodes if not n.get("src_game_id")), None)
    if root_node:
        visited2: set = set()
        stack = [(root_node, None)]
        while stack:
            cur, parent_dt = stack.pop()
            cid = cur["game_id"]
            if cid in visited2:
                continue
            visited2.add(cid)
            cur_dt = (datetime.combine(cur["date"], cur["time"])
                      if cur.get("date") and cur.get("time") else None)
            if parent_dt and cur_dt and cur_dt < parent_dt:
                red.append(
                    f"game {cid}: datetime {cur_dt} precedes parent {parent_dt} "
                    "— backwards time jump, likely a foreign game"
                )
            for nid in (cur.get("win_next_ids") or []):
                if nid in node_map and nid not in visited2:
                    stack.append((node_map[nid], cur_dt or parent_dt))

    # ── Missing Sunday node ────────────────────────────────────────────────────
    if upcoming:
        from datetime import date as _date
        upcoming_dates = {g.get("date") for g in upcoming if g.get("date")}
        node_dates    = {n.get("date") for n in nodes   if n.get("date")}
        all_sat = upcoming_dates and all(
            isinstance(d, _date) and d.weekday() == 5 for d in upcoming_dates
        )
        has_sun_node = any(
            isinstance(d, _date) and d.weekday() == 6 for d in node_dates
        )
        if all_sat and len(upcoming_dates) >= 1 and not has_sun_node:
            yellow.append(
                f"all {len(upcoming)} upcoming game(s) are on Saturday with no Sunday "
                "node in bracket — Sunday placement data may be missing"
            )

    # ── Unresolved opponent slot strings (yellow — data quality only) ──────────
    if serialized_nodes:
        for sn in serialized_nodes:
            opp = sn.get("opponent", "")
            if opp and _SLOT_LIKE_RE.match(opp):
                yellow.append(
                    f"game {sn['game_id']}: opponent {opp!r} is an unresolved slot string"
                )

    # ── Classify confidence ────────────────────────────────────────────────────
    all_warnings = red + yellow
    if red:
        confidence = "red"
    elif yellow:
        confidence = "yellow"
    else:
        confidence = "green"

    if all_warnings:
        print(f"[bracket-validate] {team} → {confidence}: {all_warnings}")
        threading.Thread(
            target=_run_bracket_llm_check,
            args=(team, nodes, all_warnings),
            daemon=True,
        ).start()

    return confidence, all_warnings


def _fire_llm_check(team, nodes, warnings):
    """Convenience wrapper for the background LLM diagnosis thread."""
    if warnings:
        threading.Thread(
            target=_run_bracket_llm_check,
            args=(team, nodes, warnings),
            daemon=True,
        ).start()


def _compute_tree_layout(nodes: list, team: str) -> dict:
    """Assign column (BFS depth), path_condition, and eliminated flag to each node.

    column:         pure BFS depth from root (1-indexed) — no secondary/neutral adjustments
    path_condition: "win" | "lose" | None (neutral/always shown) from parent edge
    eliminated:     True if a played result means the team went the other way

    Returns {game_id: {column, path_condition, eliminated}}.
    """
    if not nodes:
        return {}

    node_map = {n["game_id"]: n for n in nodes}
    root = next((n for n in nodes if not n.get("src_game_id")), None)
    if not root:
        return {}

    # BFS: column = depth from root (1-indexed)
    from collections import deque as _dq
    column_map: dict[str, int] = {}
    q = _dq([(root["game_id"], 1)])
    visited: set = set()
    while q:
        gid, col = q.popleft()
        if gid in visited:
            continue
        visited.add(gid)
        column_map[gid] = col
        n = node_map.get(gid)
        if not n:
            continue
        for nid in (n.get("win_next_ids") or []) + (n.get("lose_next_ids") or []):
            if nid not in visited:
                q.append((nid, col + 1))

    # path_condition from parent edge
    path_cond: dict = {}
    for n in nodes:
        win_ids  = set(n.get("win_next_ids")  or [])
        lose_ids = set(n.get("lose_next_ids") or [])
        for nid in win_ids | lose_ids:
            if nid in win_ids and nid in lose_ids:
                path_cond[nid] = None   # neutral — always shown
            elif nid in win_ids:
                path_cond[nid] = "win"
            else:
                path_cond[nid] = "lose"

    # eliminated: propagate from played results
    eliminated: set = set()

    def _mark_elim(gid: str) -> None:
        if gid in eliminated:
            return
        eliminated.add(gid)
        nd = node_map.get(gid)
        if nd:
            for nid in (nd.get("win_next_ids") or []) + (nd.get("lose_next_ids") or []):
                _mark_elim(nid)

    for n in nodes:
        if not n.get("played"):
            continue
        won = _team_won(team, n)
        win_ids  = set(n.get("win_next_ids")  or [])
        lose_ids = set(n.get("lose_next_ids") or [])
        neutral  = win_ids & lose_ids
        if won is True:
            for nid in lose_ids - neutral:
                _mark_elim(nid)
        elif won is False:
            for nid in win_ids - neutral:
                _mark_elim(nid)

    return {
        gid: {
            "column":         col,
            "path_condition": path_cond.get(gid),
            "eliminated":     gid in eliminated,
        }
        for gid, col in column_map.items()
    }


def _build_canonical_bracket(team: str, our_team_name: str, wpl_bracket: list,
                              confidence: str, warnings: list, display_mode: str,
                              sheet_name: str) -> dict | None:
    """Build the canonical bracket object — single source of truth for rendering.

    guaranteed_games: nodes that are confirmed in the team's bracket path (not placeholder)
    possible_games:   placeholder nodes (uncertain path, shown dimmed)
    edges:            explicit graph edges with win/loss/always conditions
    """
    if not wpl_bracket:
        return None

    guaranteed = [n for n in wpl_bracket if not n.get("placeholder")]
    possible   = [n for n in wpl_bracket if n.get("placeholder")]

    edges = []
    for n in wpl_bracket:
        win_ids  = set(n.get("win_next_ids")  or [])
        lose_ids = set(n.get("lose_next_ids") or [])
        for nid in win_ids | lose_ids:
            in_win  = nid in win_ids
            in_lose = nid in lose_ids
            condition = "always" if (in_win and in_lose) else ("win" if in_win else "loss")
            edges.append({"from": n["game_id"], "to": nid, "condition": condition})

    return {
        "team":               our_team_name,
        "division":           sheet_name,
        "bracket_confidence": confidence,
        "display_mode":       display_mode,
        "bracket_warnings":   warnings or [],
        "guaranteed_games":   guaranteed,
        "possible_games":     possible,
        "edges":              edges,
    }


def _last_meeting(team: str, opponent: str, all_games: list,
                  before_date=None, sheet: str = None) -> dict | None:
    """Return the most recent played game between team and opponent.

    Searches all_games for a played game where both teams appear.
    sheet: if provided, restricts search to that division sheet so a 16u Boys
    lookup doesn't surface 12u Boys results for the same team name.
    before_date excludes games on or after that date so the current game
    isn't counted as its own last meeting.
    """
    if not opponent or _SLOT_LIKE_RE.match(opponent):
        return None  # opponent is TBD / unresolved slot — nothing to look up
    best = None
    for g in all_games:
        if not g.get("played"):
            continue
        if sheet and g.get("sheet") != sheet:
            continue
        if before_date and g.get("date") and g["date"] >= before_date:
            continue
        wt = strip_prefix(g["white_team"]).strip()
        dt = strip_prefix(g["dark_team"]).strip()
        if not (team_matches(wt, team) or team_matches(dt, team)):
            continue
        if not (team_matches(wt, opponent) or team_matches(dt, opponent)):
            continue
        if best is None or (g.get("date") and (best.get("date") is None
                                                or g["date"] > best["date"])):
            best = g
    if best is None:
        return None
    color = "WHITE" if team_matches(strip_prefix(best["white_team"]).strip(), team) else "DARK"
    ws = best.get("white_score") or 0
    ds = best.get("dark_score")  or 0
    return {
        "date":      _fmt_date(best["date"]),
        "our_score": ws if color == "WHITE" else ds,
        "opp_score": ds if color == "WHITE" else ws,
    }


@app.route("/api/games/<tournament_id>/<path:team>")
def api_games(tournament_id, team):
    excel = find_excel(tournament_id)
    if not excel:
        abort(404)

    _all_games = load_and_parse(excel)
    games      = _filter_by_dates(_all_games, tournament_id)

    # Weekend reference date for scoping describe_slot standings lookups.
    # Prevents stale pool standings from previous weekends (pool letters repeat in WPL).
    _meta = _tournament_meta(tournament_id)
    _weekend_ref_date = _meta.get("date_start") if _meta else None

    sheet    = request.args.get("sheet")
    my_games = [g for g in games
                if (sheet is None or g["sheet"] == sheet)
                and (team_matches(g["white_team"], team) or team_matches(g["dark_team"], team))]

    # ── Phase 2: team identity scoping ────────────────────────────────────────
    # Team names (e.g. "Trojan Gold") are reused across age groups and divisions.
    # If no sheet was specified and the team appears in multiple sheets, lock to
    # the one with the most direct game matches — that's the relevant division.
    # This prevents 12u Boys games bleeding into a 16u Boys search.
    if sheet is None and my_games:
        from collections import Counter
        sheet_counts = Counter(g["sheet"] for g in my_games)
        if len(sheet_counts) > 1:
            primary_sheet = sheet_counts.most_common(1)[0][0]
            print(f"[team-scope] {team!r}: found in {sorted(sheet_counts)} — "
                  f"scoping to {primary_sheet!r} ({sheet_counts[primary_sheet]} games)")
            my_games = [g for g in my_games if g["sheet"] == primary_sheet]

    div_map = {}
    for g in games:
        div_map.setdefault(g["sheet"], []).append(g)

    # Expand to include bracket games (pool-finish and W#/L# slots) the team can reach.
    # Scoped to the primary sheet — never expand across age groups.
    primary_sheets = {g["sheet"] for g in my_games}
    for sheet_key in primary_sheets:
        direct = [g for g in my_games if g["sheet"] == sheet_key]
        my_games.extend(_expand_bracket_games(team, direct, div_map.get(sheet_key, [])))

    my_games.sort(key=lambda g: (g["date"] or date.min, g["time"] or datetime.min.time()))

    # NJO/CCA: deduplicate by game_id only — round-robin sections (GG/HH/II in 18U,
    # BB/AA in 16U) legitimately produce multiple games sharing the same pool_rank
    # metadata; keeping only one per (pool_rank_group, pool_rank) pair would hide them.
    if tournament_id in {"jo-quals", "junior-olympics"}:
        seen_gids: set = set()
        deduped = []
        for g in my_games:
            if g["game_id"] not in seen_gids:
                seen_gids.add(g["game_id"])
                deduped.append(g)
        my_games = deduped

    # Pre-compute which games are the win/lose bracket path of another game in the list
    my_game_ids = {g["game_id"] for g in my_games}
    bracket_path = {}  # game_id -> "win" | "lose"
    for g in my_games:
        dg_pre = div_map.get(g["sheet"], [])
        wn, ln = find_next_games(g, dg_pre)
        if wn and wn["game_id"] in my_game_ids:
            bracket_path[wn["game_id"]] = "win"
        if ln and ln["game_id"] in my_game_ids:
            bracket_path[ln["game_id"]] = "lose"

    played_out   = []
    upcoming_out = []

    # Build adjacency map: game_id -> unique successor game_ids in this team's game list
    _adj: dict[str, list] = {}
    for g in my_games:
        dg_pre = div_map.get(g["sheet"], [])
        wn, ln = find_next_games(g, dg_pre)
        succs: list[str] = []
        for nxt in (wn, ln):
            if nxt and nxt["game_id"] in my_game_ids and nxt["game_id"] not in succs:
                succs.append(nxt["game_id"])
        if succs:
            _adj[g["game_id"]] = succs

    # ── Game numbering: BFS depth from root game ────────────────────────────
    # Win-path and lose-path games from the same parent are at the same depth
    # and share the same game_num ("Game 2"), regardless of their time/date.
    # Root = chronologically first game not reachable as a successor.
    # Games not reached by BFS (pool placement orphans) get the fallback.
    _game_num_map: dict[str, int] = {}
    if _adj:
        from collections import deque as _deque
        _all_succs = {nid for succs in _adj.values() for nid in succs}
        # Use only the earliest non-successor as root (not all of them).
        # Other non-successors (pool placement games) fall to the fallback below.
        _non_succs = sorted(
            [g for g in my_games if g["game_id"] not in _all_succs],
            key=lambda g: (g.get("date") or date.min, g.get("time") or datetime.min.time()),
        )
        if _non_succs:
            _game_num_map[_non_succs[0]["game_id"]] = 1
            _bfs_q: _deque = _deque([(_non_succs[0]["game_id"], 1)])
            while _bfs_q:
                _gid, _d = _bfs_q.popleft()
                for _nid in _adj.get(_gid, []):
                    if _nid not in _game_num_map:
                        _game_num_map[_nid] = _d + 1
                        _bfs_q.append((_nid, _d + 1))
    # Chronological fallback for games not reached by BFS (or no tree at all)
    _slot_to_num: dict[tuple, int] = {}
    _counter = max(_game_num_map.values(), default=0)
    for g in sorted(
        [g for g in my_games if g["game_id"] not in _game_num_map],
        key=lambda g: (g.get("date") or date.min, g.get("time") or datetime.min.time()),
    ):
        is_placement = any(_FINISH_SLOT_RE.match(s.strip())
                           for s in (g["white_team"], g["dark_team"]))
        _key = (g.get("date"), "__placement__") if is_placement else (g.get("date"), g.get("time"))
        if _key not in _slot_to_num:
            _counter += 1
            _slot_to_num[_key] = _counter
        _game_num_map[g["game_id"]] = _slot_to_num[_key]

    # Re-number so game_num labels are strictly sequential and chronological.
    # BFS assigns depth-based numbers; fallback appends at depth+N. Either can
    # produce a higher-numbered group that starts earlier than a lower-numbered
    # one (e.g. placement alternatives at 11 AM getting game_num=4 while the
    # bracket follow-ons at 12 PM keep game_num=3). Fix: find each group's
    # earliest game time, sort groups by that time, then re-label 1, 2, 3, ...
    _gn_first: dict[int, tuple] = {}
    for _g in my_games:
        _gn = _game_num_map.get(_g["game_id"])
        if _gn is None:
            continue
        _d, _t = _g.get("date"), _g.get("time")
        if _d is None or _t is None:
            continue
        _key2 = (_d, _t)
        if _gn not in _gn_first or _key2 < _gn_first[_gn]:
            _gn_first[_gn] = _key2
    if _gn_first:
        _sorted_gn = sorted(_gn_first.keys(), key=lambda n: (_gn_first[n], n))
        _gn_remap  = {old: (new + 1) for new, old in enumerate(_sorted_gn)}
        _game_num_map = {gid: _gn_remap.get(gn, gn)
                         for gid, gn in _game_num_map.items()}

    # Merge adjacent win-only / lose-only placeholder columns into one.
    # Happens when placement games fall out of BFS into the time-keyed fallback,
    # giving their win and lose successors different game_nums even though they're
    # siblings (both reachable from the same upstream game, just on different paths).
    # Use bracket_path (already computed above) to get the win/lose assignment
    # for each game, since path isn't set on the raw game objects yet.
    def _raw_path(g):
        if g.get("pool_rank"):
            return f"pool_{g['pool_rank']}"
        return bracket_path.get(g["game_id"])

    _by_gn: dict[int, list] = {}
    for g in my_games:
        _gn = _game_num_map.get(g["game_id"])
        if _gn:
            _by_gn.setdefault(_gn, []).append(g)
    for _n in sorted(_by_gn.keys()):
        if _n + 1 not in _by_gn:
            continue
        _col_a, _col_b = _by_gn[_n], _by_gn[_n + 1]
        _paths_a = {_raw_path(g) for g in _col_a}
        _paths_b = {_raw_path(g) for g in _col_b}
        _all_ph_a = all(g.get("placeholder") for g in _col_a)
        _all_ph_b = all(g.get("placeholder") for g in _col_b)
        if (_all_ph_a and _all_ph_b
                and _paths_a <= {"win", "lose"} and _paths_b <= {"win", "lose"}
                and _paths_a | _paths_b == {"win", "lose"}):
            for g in _col_b:
                _game_num_map[g["game_id"]] = _n

    # Close any gaps left by the sibling-merge (e.g. 1,2,3,4,6,7 → 1,2,3,4,5,6)
    _used_gns = sorted(set(_game_num_map.values()))
    _gap_remap = {old: (new + 1) for new, old in enumerate(_used_gns)}
    _game_num_map = {gid: _gap_remap[gn] for gid, gn in _game_num_map.items()}

    # Merge multiple pure-single-day columns that share the same calendar day.
    # e.g. two Saturday game_nums (pool placements + W/L branches) collapse into
    # one "Game 3" column, and all Sunday games collapse into one column.
    # A column is "pure single-day" only if ALL its games are on the same date.
    # Mixed columns (e.g. Fri win + Sat lose in game_num=2) are left alone.
    _by_gn2: dict[int, list] = {}
    for _g2 in my_games:
        _gn2 = _game_num_map.get(_g2["game_id"])
        if _gn2:
            _by_gn2.setdefault(_gn2, []).append(_g2)
    _gn_unique_days: dict[int, set] = {
        _n2: {_g2.get("date") for _g2 in _gs2 if _g2.get("date")}
        for _n2, _gs2 in _by_gn2.items()
    }
    _day_pure_gns: dict = {}
    for _n2, _days2 in _gn_unique_days.items():
        if len(_days2) == 1:
            _d2 = next(iter(_days2))
            _day_pure_gns.setdefault(_d2, []).append(_n2)
    for _d2, _gns2 in _day_pure_gns.items():
        if len(_gns2) <= 1:
            continue
        _target_gn = min(_gns2)
        for _n2 in _gns2:
            if _n2 == _target_gn:
                continue
            for _g2 in _by_gn2.get(_n2, []):
                _game_num_map[_g2["game_id"]] = _target_gn
    # Re-close gaps after day-merge
    _used_gns2 = sorted(set(_game_num_map.values()))
    _gap_remap2 = {old: (new + 1) for new, old in enumerate(_used_gns2)}
    _game_num_map = {gid: _gap_remap2[gn] for gid, gn in _game_num_map.items()}

    show_records = tournament_id not in WPL_TOURNAMENTS
    sheet_records: dict = {}
    if show_records:
        for sk, sg in div_map.items():
            sheet_records[sk] = _tournament_records(sg)

    now_la = datetime.now(ZoneInfo('America/Los_Angeles')).replace(tzinfo=None)

    for g in my_games:
        game_num = _game_num_map.get(g["game_id"], 1)
        dg     = div_map.get(g["sheet"], [])
        gid    = g["game_id"]
        opp_sl = _team_opp_slot(g, team, dg, my_game_ids)
        color  = "WHITE" if opp_sl == g["dark_team"] else "DARK"

        our_rec  = None
        opp_rec  = None
        if show_records:
            trec    = sheet_records.get(g["sheet"], {})
            our_key = strip_prefix(g["white_team"] if color == "WHITE" else g["dark_team"]).upper()
            opp_key = strip_prefix(opp_sl).upper()
            our_rec = trec.get(our_key)
            opp_rec = trec.get(opp_key)

        opponent_label = describe_slot(opp_sl, dg, ref_date=_weekend_ref_date)
        # Skip games where the opponent resolves to our own team (W#N self-play artifact)
        if team_matches(opponent_label, team):
            continue

        is_placement_alt = any(_FINISH_SLOT_RE.match(s.strip())
                               for s in (g["white_team"], g["dark_team"]))
        base = {
            "game_id":      gid,
            "date":         _fmt_date(g["date"]),
            "time":         _fmt_time(g["time"]),
            "location":     g["location"],
            "opponent":     opponent_label,
            "your_color":   color,
            "game_num":     game_num,
            "is_alternative": is_placement_alt,
            "path":        (f"pool_{g.get('pool_rank')}" if g.get('pool_rank') else None) or bracket_path.get(gid),
            "placeholder": g.get("placeholder", False),
            "our_record":  our_rec,
            "opp_record":  opp_rec,
        }

        winner_next, loser_next = find_next_games(g, dg)

        if g["played"]:
            ws = g.get("white_score") or 0
            ds = g.get("dark_score")  or 0
            base["score"]     = _fmt_score(g)
            base["our_score"] = ws if color == "WHITE" else ds
            base["opp_score"] = ds if color == "WHITE" else ws
            result = _result_str(g, team)
            base["result"] = result
            next_game = winner_next if result == "win" else loser_next if result == "loss" else None
            if next_game:
                base["next"] = _next_summary(next_game, team, dg, ref_date=_weekend_ref_date, my_game_ids=my_game_ids)
            played_out.append(base)
        else:
            # Detect if this game is currently in progress (window: -15 min to +2 hr from start)
            is_current = False
            if g.get("date") and g.get("time"):
                game_dt = datetime.combine(g["date"], g["time"])
                elapsed_s = (now_la - game_dt).total_seconds()
                is_current = -900 <= elapsed_s <= 7200
            base["is_current"] = is_current
            live = _LIVE_SCORES.get((tournament_id, gid))
            if live:
                base["live_score"] = live
                base["is_current"] = True
            scenarios = {}
            # Suppress scenarios on placement-alternative games (pool-finish slots like
            # "1stG-", "2ndH-") — their follow-on games are deeper placement rounds that
            # parents don't need to preview, and the verbose opponent labels create clutter.
            if not is_placement_alt:
                # Suppress a scenario if that game is already shown as its own card
                if winner_next and winner_next["game_id"] not in my_game_ids:
                    scenarios["win"]  = _next_summary(winner_next, team, dg, ref_date=_weekend_ref_date, my_game_ids=my_game_ids)
                if loser_next and loser_next["game_id"] not in my_game_ids:
                    scenarios["lose"] = _next_summary(loser_next,  team, dg, ref_date=_weekend_ref_date, my_game_ids=my_game_ids)
            base["scenarios"] = scenarios if scenarios else None
            base["last_meeting"] = _last_meeting(team, opponent_label, _all_games,
                                                  before_date=g.get("date"),
                                                  sheet=g.get("sheet"))
            upcoming_out.append(base)

    placement = _infer_placement(played_out)
    if placement is None:
        last_played_raw = next((g for g in reversed(my_games) if g.get('played')), None)
        if last_played_raw:
            if hasattr(excel, 'seek'):
                excel.seek(0)
            placement = _estimate_placement(
                played_out,
                div_map.get(last_played_raw['sheet'], []),
                excel,
                last_played_raw['sheet'],
                team,
            )

    # For Futures weekends: surface pool standing + cumulative season standings
    pool_standing        = None
    cumulative_standings = None
    cumulative_division  = None
    if tournament_id in WPL_TOURNAMENTS and my_games:
        sheet = my_games[0]['sheet']
        weekend_num = _futures_weekend_num(tournament_id)
        if weekend_num is not None:
            if hasattr(excel, 'seek'):
                excel.seek(0)
            pool_standing = _read_futures_pool_standing(excel, weekend_num, team, sheet)
        if hasattr(excel, 'seek'):
            excel.seek(0)
        cumulative_division, cumulative_standings = _read_futures_cumulative_standings(excel, team, sheet)

        # Compute rank movement vs previous weekend (show ↑/↓ on scoreboard)
        if cumulative_standings and weekend_num and weekend_num >= 3:
            if hasattr(excel, 'seek'):
                excel.seek(0)
            _, prev_standings = _read_futures_cumulative_standings(
                excel, team, sheet, max_weekend=weekend_num - 1)
            if prev_standings:
                prev_rank_by_name = {}
                for s in prev_standings:
                    key = s['name'].upper().strip()
                    if s.get('division') == cumulative_standings[0].get('division') if cumulative_standings else True:
                        prev_rank_by_name[key] = s['rank']
                # Build per-division prev rank maps
                prev_rank_by_div: dict = {}
                for s in prev_standings:
                    div = s.get('division') or 'D?'
                    prev_rank_by_div.setdefault(div, {})[s['name'].upper().strip()] = s['rank']
                for s in cumulative_standings:
                    div = s.get('division') or 'D?'
                    key = s['name'].upper().strip()
                    prev_r = prev_rank_by_div.get(div, {}).get(key)
                    if prev_r is not None:
                        s['rank_movement'] = prev_r - s['rank']  # positive = moved up
                    else:
                        s['rank_movement'] = None  # new team this weekend

    # WPL crossover game tree
    wpl_bracket = None
    _njo_tournaments = {"jo-quals", "junior-olympics"}
    if tournament_id in WPL_TOURNAMENTS and my_games:
        tree_sheet = my_games[0]['sheet']
        div_games_for_tree = [g for g in _all_games if g['sheet'] == tree_sheet]
        latest_team_date = max((g["date"] for g in my_games if g.get("date")), default=None)
        tree = _build_wpl_game_tree(team, div_games_for_tree, anchor_date=latest_team_date)
    elif tournament_id in _njo_tournaments and my_games:
        tree_sheet = my_games[0]['sheet']
        div_games_for_tree = [g for g in _all_games if g['sheet'] == tree_sheet]
        latest_team_date = max((g["date"] for g in my_games if g.get("date")), default=None)
        tree = _build_njo_game_tree(team, div_games_for_tree, anchor_date=latest_team_date)
        # NJO uses w_to/l_to integer links, not WPL-style WIN GM # slots — skip
        # ground-truth depth checks (_derive_expected_bracket is WPL-specific).
        struct_issues = _check_bracket_structure(team, tree)
        if struct_issues:
            for _si in struct_issues:
                print(f"[bracket-struct] {team!r} | {tournament_id}: {_si}", flush=True)

    if tree:
        # Serialize tree nodes: format dates/times, add opponent label
        def _serialize_tree_node(node, dg):
            gid = node["game_id"]
            # Try raw slot match first; fall back to resolving W#/L# refs
            if team_matches(node["white_team"], team):
                opp_sl, color = node["dark_team"],  "WHITE"
            elif team_matches(node["dark_team"], team):
                opp_sl, color = node["white_team"], "DARK"
            else:
                t_w = describe_slot(node["white_team"], dg, ref_date=latest_team_date)
                t_d = describe_slot(node["dark_team"],  dg, ref_date=latest_team_date)
                if team_matches(t_w, team):
                    opp_sl, color = node["dark_team"],  "WHITE"
                elif team_matches(t_d, team):
                    opp_sl, color = node["white_team"], "DARK"
                else:
                    opp_sl, color = node["white_team"], "DARK"
            opp_name = describe_slot(opp_sl, dg, ref_date=latest_team_date)
            # Guard: if resolved opponent still equals our own team, flip slots
            if team_matches(opp_name, team):
                other_sl = node["white_team"] if opp_sl == node["dark_team"] else node["dark_team"]
                opp_name = describe_slot(other_sl, dg, ref_date=latest_team_date)
            is_current = False
            if node.get("date") and node.get("time"):
                game_dt = datetime.combine(node["date"], node["time"])
                elapsed_s = (now_la - game_dt).total_seconds()
                is_current = -900 <= elapsed_s <= 7200
            live = _LIVE_SCORES.get((tournament_id, gid))
            if live and not node.get("placeholder"):
                is_current = True
            d = {
                "game_id":        gid,
                "date":           _fmt_date(node["date"]),
                "time":           _fmt_time(node["time"]),
                "location":       node["location"],
                "opponent":       opp_name,
                "your_color":     color,
                "placeholder":    node["placeholder"],
                "played":         bool(node.get("played", False)),
                "placement_rank": node.get("placement_rank"),
                "src_game_id":    node["src_game_id"],
                "src_path":       node["src_path"],
                "win_next_ids":   node["win_next_ids"],
                "lose_next_ids":  node["lose_next_ids"],
                "sunday_pair_id": node["sunday_pair_id"],
                "is_current":     is_current,
            }
            if node.get("played") and not node.get("placeholder"):
                ws = node.get("white_score") or 0
                ds = node.get("dark_score")  or 0
                d["score"]     = _fmt_score(node)
                d["our_score"] = ws if color == "WHITE" else ds
                d["opp_score"] = ds if color == "WHITE" else ws
                d["result"]    = _result_str(node, team)
            if not node.get("played"):
                d["last_meeting"] = _last_meeting(
                    team, opp_name, dg, before_date=node.get("date"))
            if live:
                d["live_score"] = live
            return d
        wpl_bracket = [_serialize_tree_node(n, div_games_for_tree) for n in tree]

        # Annotate serialized nodes with tree layout (column, path_condition, eliminated).
        layout = _compute_tree_layout(tree, team)
        for sn in wpl_bracket:
            info = layout.get(sn["game_id"], {})
            sn["column"]         = info.get("column", 1)
            sn["path_condition"] = info.get("path_condition")
            sn["eliminated"]     = info.get("eliminated", False)
            # Aliases kept for validate_tournament.py backwards compatibility
            sn["game_num"]       = sn["column"]
            sn["path"]           = sn["path_condition"]
            sn["is_alternative"] = False

    # ── Bracket confidence + display mode ────────────────────────────────────
    # WPL tournaments always attempt a bracket. Non-WPL/NJO: no bracket expected.
    is_bracket_tournament = (tournament_id in WPL_TOURNAMENTS
                             or tournament_id in {"jo-quals", "junior-olympics"})

    if wpl_bracket:
        bracket_confidence, bracket_warnings = _validate_wpl_bracket(
            team, tree, upcoming=upcoming_out, serialized_nodes=wpl_bracket)
    elif is_bracket_tournament and my_games:
        # Bracket expected but missing — hard RED
        bracket_confidence = "red"
        bracket_warnings   = ["bracket could not be built — wpl_bracket is null; "
                               "check tree builder logs for this team"]
    else:
        bracket_confidence = "green"
        bracket_warnings   = []

    # RED forces flat schedule; GREEN/YELLOW allow bracket
    display_mode = "bracket" if bracket_confidence in ("green", "yellow") and wpl_bracket else "flat_schedule"

    our_team_name = team.title()
    if my_games:
        sname = my_games[0].get("sheet", "")
        fn = friendly_team_name(team, sname)
        if fn:
            our_team_name = fn.split("·")[0].strip()

    # Build canonical bracket object (Phase 4 — single source of truth for rendering)
    canonical_bracket = _build_canonical_bracket(
        team, our_team_name, wpl_bracket,
        bracket_confidence, bracket_warnings, display_mode,
        my_games[0].get("sheet", "") if my_games else "",
    ) if wpl_bracket else None

    parse_format = games[0].get("format") if games else None
    return jsonify({
        "team":                 team,
        "our_team_name":        our_team_name,
        "played":               played_out,
        "upcoming":             upcoming_out,
        "placement":            placement,
        "pool_standing":        pool_standing,
        "cumulative_standings": cumulative_standings,
        "cumulative_division":  cumulative_division,
        "wpl_bracket":          wpl_bracket,
        "canonical_bracket":    canonical_bracket,
        "bracket_confidence":   bracket_confidence,
        "display_mode":         display_mode,
        "bracket_warnings":     bracket_warnings or None,
        "cache_age_s":          _cache_age(tournament_id),
        "cache_ttl_s":          URL_CACHE_TTL,
        "parse_format":         parse_format,
    })


@app.route("/api/live-score/<tournament_id>/<game_id>", methods=["GET"])
def api_live_score_get(tournament_id, game_id):
    return jsonify(_LIVE_SCORES.get((tournament_id, game_id)))


@app.route("/api/live-score/<tournament_id>/<game_id>", methods=["POST"])
def api_live_score_post(tournament_id, game_id):
    body = request.get_json(force=True)
    _LIVE_SCORES[(tournament_id, game_id)] = {
        "our_score": max(0, int(body.get("our_score", 0))),
        "opp_score": max(0, int(body.get("opp_score", 0))),
        "quarter":   max(1, min(5, int(body.get("quarter", 1)))),
        "updated_at": time.time(),
    }
    return jsonify({"ok": True})


def _comments_path(tournament_id: str) -> str:
    return os.path.join(COMMENTS_DIR, f"{tournament_id}.json")

def _comments_expired(tournament_id: str) -> bool:
    """True if tournament ended more than 24 h ago."""
    meta = _tournament_meta(tournament_id)
    if not meta or not meta.get("date_end"):
        return False
    cutoff = datetime.combine(meta["date_end"], datetime.min.time()) + timedelta(hours=24)
    return datetime.utcnow() > cutoff

def _load_comments(tournament_id: str) -> dict:
    if _comments_expired(tournament_id):
        return {}
    path = _comments_path(tournament_id)
    if os.path.exists(path):
        try:
            with open(path) as f:
                return json.load(f)
        except Exception:
            return {}
    return {}

def _save_comments(tournament_id: str, data: dict):
    with open(_comments_path(tournament_id), "w") as f:
        json.dump(data, f)

@app.route("/api/comments/<tournament_id>/<game_id>", methods=["GET"])
def api_comments_get(tournament_id, game_id):
    all_comments = _load_comments(tournament_id)
    return jsonify(all_comments.get(game_id, []))

@app.route("/api/comments/<tournament_id>/<game_id>", methods=["POST"])
def api_comments_post(tournament_id, game_id):
    if _comments_expired(tournament_id):
        abort(410)  # Gone — tournament window closed
    data = request.get_json(silent=True) or {}
    text = str(data.get("text", "")).strip()[:280]
    if not text:
        abort(400)
    all_comments = _load_comments(tournament_id)
    all_comments.setdefault(game_id, []).append({"text": text, "ts": int(time.time())})
    _save_comments(tournament_id, all_comments)
    return jsonify({"ok": True, "count": len(all_comments[game_id])})

@app.route("/api/feedback", methods=["POST"])
def api_feedback_post():
    data = request.get_json(silent=True) or {}
    text = str(data.get("text", "")).strip()[:1000]
    if not text:
        abort(400)
    entry = {"text": text, "ts": int(time.time())}
    _FEEDBACK.append(entry)
    _save_feedback(_FEEDBACK)
    print(f"[feedback] {entry}", flush=True)
    return jsonify({"ok": True})

@app.route("/api/feedback", methods=["GET"])
def api_feedback_get():
    return jsonify(_FEEDBACK)


@app.route("/api/result/<tournament_id>/<game_id>", methods=["POST"])
def api_lock_result(tournament_id, game_id):
    """Lock a yes/no result for a game. First submission wins."""
    data = request.get_json(force=True)
    trojan_won = data.get("trojan_won")
    team       = data.get("team", "TROJAN")
    if trojan_won is None:
        abort(400)

    results = load_results(tournament_id)
    if game_id in results:
        return jsonify({"ok": True, "already_set": True, **results[game_id]})

    entry = {"trojan_won": trojan_won, "team": team, "locked_at": datetime.now().isoformat()}
    results[game_id] = entry
    save_results(tournament_id, results)
    return jsonify({"ok": True, **entry})


@app.route("/api/result/<tournament_id>/<game_id>/unlock", methods=["POST"])
def api_unlock_result(tournament_id, game_id):
    """Unlock a result (password required)."""
    data = request.get_json(force=True)
    if data.get("password") != ADMIN_PW:
        abort(403)

    results = load_results(tournament_id)
    results.pop(game_id, None)
    save_results(tournament_id, results)
    return jsonify({"ok": True})


@app.route("/api/upload", methods=["POST"])
def api_upload():
    """Upload a new Excel file (password required)."""
    pw = request.form.get("password") or (request.get_json(force=True, silent=True) or {}).get("password")
    if pw != ADMIN_PW:
        abort(403)

    f = request.files.get("file")
    if not f or not f.filename.endswith(".xlsx"):
        abort(400)

    safe_name = re.sub(r"[^\w\s\-.]", "", f.filename).strip()
    dest = os.path.join(EXCEL_DIR, safe_name)
    f.save(dest)
    return jsonify({"ok": True, "filename": safe_name})


@app.route("/api/tournaments/<tournament_id>/url", methods=["POST"])
def api_set_url(tournament_id):
    """Store a user-provided URL for any tournament (password required)."""
    if not any(t["id"] == tournament_id for t in KNOWN_TOURNAMENTS):
        abort(404)
    data = request.get_json(force=True)
    if data.get("password") != ADMIN_PW:
        abort(403)
    url = (data.get("url") or "").strip()
    if not url:
        abort(400)
    content = _fetch_url(url)
    if not content:
        return jsonify({"ok": False, "error": "Could not fetch that URL. Make sure it's publicly accessible."}), 400
    try:
        load_and_parse(io.BytesIO(content))
    except Exception:
        return jsonify({"ok": False, "error": "URL was fetched but doesn't appear to be a valid Excel schedule."}), 400
    _save_user_url(tournament_id, url)
    return jsonify({"ok": True})


@app.route("/api/h2h/<path:team>")
def api_h2h(team):
    """Return head-to-head record for a Trojan team vs every opponent,
    aggregated across all available Excel files."""
    my_club, my_tier = normalize_opp(team)

    records = {}   # h2h_key → {club, tier, wins, losses, ties, games}

    for fname in all_excels():
        fpath = os.path.join(EXCEL_DIR, fname)
        try:
            file_games = load_and_parse(fpath)
        except Exception:
            continue

        # Find a friendly tournament name for this file
        t_name = fname
        for t in KNOWN_TOURNAMENTS:
            if find_excel(t["id"]) == fpath:
                t_name = t["name"]
                break

        for g in file_games:
            if not g.get("played"):
                continue

            w_club, w_tier = normalize_opp(g["white_team"])
            d_club, d_tier = normalize_opp(g["dark_team"])

            if w_club == "TROJAN" and w_tier == my_tier:
                opp_raw = g["dark_team"]
            elif d_club == "TROJAN" and d_tier == my_tier:
                opp_raw = g["white_team"]
            else:
                continue

            opp_club, opp_tier = normalize_opp(opp_raw)
            if opp_club == "TROJAN":
                continue   # skip Trojan-vs-Trojan

            key = h2h_key(opp_raw)
            result = _result_str(g, team)
            rec = records.setdefault(key, {
                "club": opp_club.title(),
                "tier": opp_tier,
                "wins": 0, "losses": 0, "ties": 0,
                "games": [],
            })
            if result == "win":    rec["wins"]   += 1
            elif result == "loss": rec["losses"] += 1
            else:                  rec["ties"]   += 1

            rec["games"].append({
                "tournament": t_name,
                "date":       _fmt_date(g["date"]),
                "score":      _fmt_score(g),
                "result":     result,
            })

    for rec in records.values():
        rec["games"].sort(key=lambda x: x["date"] or "")

    return jsonify(records)


@app.route("/api/raw-slots/<tournament_id>/<sheet_name>")
def api_raw_slots(tournament_id, sheet_name):
    excel = find_excel(tournament_id)
    if not excel:
        abort(404)
    games = [g for g in load_and_parse(excel) if g["sheet"] == sheet_name]
    return jsonify([{"game_id": g["game_id"], "date": str(g["date"]), "white": g["white_team"], "dark": g["dark_team"], "ws": g.get("white_score"), "ds": g.get("dark_score")} for g in games])


@app.route("/api/status")
def api_status():
    """Diagnostic endpoint: Excel fetch health, cache state, game counts."""
    import datetime as dt
    now = time.time()
    results = []
    for t in KNOWN_TOURNAMENTS:
        tid = t["id"]
        url = TOURNAMENT_URLS.get(tid) or _load_user_urls().get(tid)
        cached = _URL_CACHE.get(url) if url else None
        fetched_at = dt.datetime.utcfromtimestamp(cached[0]).strftime("%H:%M:%S UTC") if cached else None
        age_s = int(now - cached[0]) if cached else None
        stale = age_s is not None and age_s > URL_CACHE_TTL

        game_count = None
        sheet_count = None
        error = None
        try:
            excel = find_excel(tid)
            if excel:
                games = load_and_parse(excel)
                game_count = len(games)
                sheet_count = len({g["sheet"] for g in games})
            else:
                error = "no excel"
        except Exception as e:
            error = str(e)[:120]

        results.append({
            "id":          tid,
            "name":        t["name"],
            "has_url":     bool(url),
            "cached":      cached is not None,
            "fetched_at":  fetched_at,
            "cache_age_s": age_s,
            "stale":       stale,
            "games":       game_count,
            "sheets":      sheet_count,
            "error":       error,
        })
    return jsonify({"server_time": dt.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"),
                    "deploy_id": "1def61f",  # bump on each deploy to confirm Railway picked up latest code
                    "tournaments": results})


@app.route("/api/client-error", methods=["POST"])
def api_client_error():
    """Receives JS errors from the browser and prints them to Railway logs."""
    try:
        data = request.get_json(force=True, silent=True) or {}
        print(f"[CLIENT ERROR] {data.get('message','')} | "
              f"{data.get('source','')}:{data.get('line','')} | "
              f"url={data.get('url','')} | "
              f"ua={data.get('ua','')[:80]} | "
              f"stack={str(data.get('stack',''))[:300]}", flush=True)
    except Exception:
        pass
    return '', 204


@app.route("/api/debug/standings/<tournament_id>")
def api_debug_standings(tournament_id):
    """Diagnostic: show raw DivisionsStandings tab content so we can see why standings fail."""
    excel = find_excel(tournament_id)
    if not excel:
        return jsonify({"error": "no excel"}), 404
    try:
        import openpyxl, io as _io
        if hasattr(excel, 'seek'):
            excel.seek(0)
        wb = openpyxl.load_workbook(excel, data_only=True, read_only=True)
        if 'DivisionsStandings' not in wb.sheetnames:
            return jsonify({"error": "DivisionsStandings tab not found",
                            "sheets": wb.sheetnames})
        ws = wb['DivisionsStandings']
        # Show all non-blank rows up to 800 (full range the standings parser reads)
        preview = []
        for i, row in enumerate(ws.iter_rows(max_row=800, values_only=True)):
            if any(c is not None for c in row):
                preview.append({"row": i + 1, "cells": [str(c)[:60] if c is not None else None for c in row[:9]]})
        return jsonify({"sheets": wb.sheetnames, "preview": preview})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/admin/cache-schema/<tournament_id>", methods=["POST"])
def api_cache_schema(tournament_id):
    """One-time endpoint: analyze the tournament Excel with Claude and cache the
    layout schema to disk. After this runs, all subsequent parses are instant
    (no further API calls). Safe to call multiple times — no-ops if already cached."""
    excel = find_excel(tournament_id)
    if not excel:
        return jsonify({"error": "no excel found"}), 404
    try:
        if isinstance(excel, (bytes, bytearray)):
            data = excel
        elif hasattr(excel, "read"):
            data = excel.read()
        elif isinstance(excel, str) and excel.startswith("http"):
            data = _fetch_url(excel)
        else:
            data = open(excel, "rb").read()
        import io, openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        from parsers.format_ai import parse as parse_ai, _sample_rows, _hash_rows, _load_disk_schema, SKIP_SHEETS
        results = {}
        for sheet_name in wb.sheetnames:
            if sheet_name.lower().strip() in SKIP_SHEETS:
                results[sheet_name] = "skipped"
                continue
            ws = wb[sheet_name]
            rows = list(ws.iter_rows(max_row=60, values_only=True))
            sample = _sample_rows(rows)
            key = _hash_rows(sample)
            if _load_disk_schema(key):
                results[sheet_name] = f"already cached ({key})"
            else:
                results[sheet_name] = f"analyzing..."
        # Run full AI parse to build all schemas
        wb.close()
        wb = openpyxl.load_workbook(io.BytesIO(data), data_only=True)
        games = parse_ai(wb)
        return jsonify({"tournament": tournament_id, "games_found": len(games), "sheets": results})
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ── Pre-game sweep ────────────────────────────────────────────────────────────

_swept_tournament_ids: set[str] = set()   # prevent re-running in same process

def _run_pre_game_sweep(tournament_id: str) -> dict:
    """Validate all teams for a tournament weekend. Logs to stdout.

    Returns dict with 'ok' count, 'issues' list (for API endpoint).
    Called automatically 12-25 hours before each tournament, and on demand
    via /api/health/pre-game-check/<tournament_id>.
    """
    meta = _tournament_meta(tournament_id)
    if not meta or not meta.get("date_start"):
        return {"error": f"no date_start for {tournament_id}"}

    anchor = meta["date_start"]
    print(f"[pre-game-check] starting sweep: {tournament_id} ({anchor})")

    excel = find_excel(tournament_id)
    if not excel:
        msg = f"[pre-game-check] {tournament_id}: no excel file — skipping"
        print(msg)
        return {"error": msg}

    try:
        all_games = load_and_parse(excel)
    except Exception as e:
        msg = f"[pre-game-check] {tournament_id}: parse error — {e}"
        print(msg)
        return {"error": msg}

    _swept_tournament_ids.add(tournament_id)

    n_ok = 0
    all_issues: list[dict] = []

    by_sheet: dict[str, list] = {}
    for g in all_games:
        by_sheet.setdefault(g["sheet"], []).append(g)

    _pts_re = re.compile(r'\s*-\s*\d+(\.\d+)?\s*PTS\.?\s*$', re.IGNORECASE)

    for sheet, sheet_games in by_sheet.items():
        seen: set[str] = set()
        teams: list[str] = []
        for g in sheet_games:
            if not (g.get("date") and abs((g["date"] - anchor).days) <= 1):
                continue
            for slot in (g["white_team"], g["dark_team"]):
                name = strip_prefix(slot).strip().upper()
                name = _pts_re.sub("", name).strip()
                if (not name or len(name) < 3
                        or re.search(r'\bGM\s*#', name, re.IGNORECASE)
                        or re.match(r'^(WIN|LOS|TBD|\d)', name, re.IGNORECASE)
                        or re.match(r'^[A-Z]\d+$', name)):
                    continue
                if name not in seen:
                    seen.add(name)
                    teams.append(name.title())

        for team in teams:
            direct = [g for g in sheet_games
                      if (team_matches(g["white_team"], team) or
                          team_matches(g["dark_team"], team))
                      and g.get("date") and abs((g["date"] - anchor).days) <= 1]
            if not direct:
                continue

            extras = _expand_bracket_games(team, direct, sheet_games)
            tree   = _build_wpl_game_tree(team, sheet_games, anchor_date=anchor)

            # Count all unique Sunday game IDs visible to this team
            sun_ids: set[str] = set()
            for item in direct + extras:
                if item.get("date") and item["date"].weekday() == 6:
                    sun_ids.add(item["game_id"])
            for node in tree:
                if node.get("date") and node["date"].weekday() == 6:
                    sun_ids.add(node["game_id"])

            issues: list[str] = []
            total = len(direct) + len(extras)
            if total < 2:
                issues.append(f"only {total} total game(s)")
            if not sun_ids:
                issues.append(
                    f"no Sunday games visible "
                    f"(direct={len(direct)} extras={len(extras)} tree={len(tree)})"
                )
            if len(extras) > len(direct) * 8:
                issues.append(
                    f"extras={len(extras)} >> direct={len(direct)} — "
                    "possible cross-weekend contamination"
                )
            # Check for unresolved slot strings in tree nodes
            for n in tree:
                wt, dt = n["white_team"], n["dark_team"]
                try:
                    opp_slot = dt if team.split()[-1].upper() in wt.upper() else wt
                except Exception:
                    opp_slot = wt
                opp = describe_slot(opp_slot, sheet_games, ref_date=anchor)
                if _SLOT_LIKE_RE.match(opp):
                    issues.append(f"unresolved slot in tree: {n['game_id']} opp={opp!r}")
                    break

            if issues:
                for iss in issues:
                    print(f"[pre-game-check] ⚠  {sheet}/{team}: {iss}")
                all_issues.append({"sheet": sheet, "team": team, "issues": issues})
            else:
                n_ok += 1

    if all_issues:
        print(f"[pre-game-check] {tournament_id}: "
              f"{len(all_issues)} team(s) with issues, {n_ok} clean")
    else:
        print(f"[pre-game-check] {tournament_id}: all {n_ok} teams OK ✓")

    return {"tournament": tournament_id, "ok": n_ok, "issues": all_issues}


def _pre_game_monitor():
    """Background thread: runs a pre-game sweep 12-25 hours before tournament start.

    Checks once per hour. When a tournament is 12-25 hours away and hasn't
    been swept yet this process, fires _run_pre_game_sweep in a daemon thread.
    Logs prominently to Railway stdout so issues surface before game day.
    """
    import time as _time
    _time.sleep(60)  # let app fully start before first check
    while True:
        try:
            now = datetime.now(ZoneInfo('America/Los_Angeles')).replace(tzinfo=None)
            for t in KNOWN_TOURNAMENTS:
                tid   = t.get("id", "")
                start = t.get("date_start")
                if not start or tid in _swept_tournament_ids:
                    continue
                hours = (datetime.combine(start, datetime.min.time()) - now).total_seconds() / 3600
                if 12 <= hours <= 25:
                    print(f"[pre-game-check] {hours:.0f}h until {t.get('name', tid)} "
                          f"— launching pre-game sweep")
                    threading.Thread(
                        target=_run_pre_game_sweep, args=(tid,), daemon=True
                    ).start()
        except Exception as e:
            print(f"[pre-game-monitor] error: {e}")
        _time.sleep(3600)  # re-check every hour


@app.route("/api/health/pre-game-check/<tournament_id>")
def api_pre_game_check(tournament_id: str):
    """Manual trigger for pre-game validation sweep. Returns JSON results.

    Hit this any time to get a health report for all teams in a tournament:
        curl https://<railway-url>/api/health/pre-game-check/futures-5
    """
    # Remove from swept set so this always re-runs on demand
    _swept_tournament_ids.discard(tournament_id)
    result = _run_pre_game_sweep(tournament_id)
    if "error" in result:
        return jsonify(result), 404
    return jsonify(result)


# Start background monitor
threading.Thread(target=_pre_game_monitor, daemon=True).start()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port, debug=False)

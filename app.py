"""
Tournament Translator — Flask app
"""
from __future__ import annotations
import os, re, json, glob, io, time, base64, random
from datetime import datetime, date
from zoneinfo import ZoneInfo
from functools import lru_cache
import requests
from flask import Flask, render_template, jsonify, request, abort

from parsers.detect import load_and_parse

app = Flask(__name__)

EXCEL_DIR   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tournaments Excels")
RESULTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "results")
ADMIN_PW    = os.environ.get("ADMIN_PASSWORD", "trojan")  # override via Railway env var

# ── Live URL sources ────────────────────────────────────────────────────────
# All tournament URLs baked in code — updated here after each tournament is scheduled.
# Phone UI (user_urls.json) overrides these within a session.
FUTURES_SHEETS_ID  = "1AkX3vwOU9CIc3cymacG2F-uXz-_Gi_A8yR40dEbDpMQ"
FUTURES_SHEETS_URL = f"https://docs.google.com/spreadsheets/d/{FUTURES_SHEETS_ID}/export?format=xlsx"
WPL_TOURNAMENTS    = {"futures-2", "futures-3", "futures-4", "futures-5", "futures-super"}

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

_MONTH_MAP = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
              "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
_RE_YEAR   = re.compile(r"(\d{4})")
_RE_MONTH  = re.compile(r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)", re.I)

os.makedirs(RESULTS_DIR, exist_ok=True)
os.makedirs(EXCEL_DIR, exist_ok=True)

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
    user_url = _load_user_urls().get(tournament_id)
    if user_url:
        data = _fetch_url(user_url)
        if data:
            return io.BytesIO(data)
    keyword = FILE_MAP.get(tournament_id, "")
    if keyword:
        for f in os.listdir(EXCEL_DIR):
            if f.endswith(".xlsx") and keyword.lower() in f.lower():
                return os.path.join(EXCEL_DIR, f)
    if tournament_id in WPL_TOURNAMENTS:
        data = _fetch_url(FUTURES_SHEETS_URL)
        return io.BytesIO(data) if data else None
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
    r"^(?:\d+(?:st|nd|rd|th)[A-Z]-|[A-Z]\d+\([^)]+\)-?|[WL]#[^-\s]+-?|[A-Z]\d+-|\d+-)(.*)",
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

def describe_slot(slot: str, division_games: list = None) -> str:
    """Return a human-readable opponent label.
    With division_games, resolves bracket slots to actual team names using standings."""
    slot = slot.strip()
    name = strip_prefix(slot)
    if name != slot and name:
        # If strip_prefix left us with a W#/L# reference (e.g. "E1(4thB)L#7"), resolve it.
        if re.match(r'^[WL]#', name, re.IGNORECASE):
            return describe_slot(name, division_games)
        return name

    if division_games:
        # W#N / L#N → resolve to actual winner/loser if game has been played
        wm = re.match(r'^([WL])#([^-\s]+)', slot, re.IGNORECASE)
        if wm:
            want_winner = wm.group(1).upper() == "W"
            ref = re.search(r'(\d+)$', wm.group(2))
            if ref:
                ref_game = _game_by_num(str(int(ref.group(1))), division_games)
                if ref_game:
                    t1 = describe_slot(ref_game["white_team"], division_games)
                    t2 = describe_slot(ref_game["dark_team"], division_games)
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
                standings = _standings_for_group(group, division_games)
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
    return slot

def _game_num(game_id: str):
    m = re.search(r"(\d+)$", game_id)
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
_FINISH_SLOT_RE  = re.compile(r'^\d+(?:st|nd|rd|th)([A-Z])-', re.IGNORECASE)
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

    # Pool groups this team is seeded in (e.g. 'G' from 'G2-TROJAN SILVER')
    groups = set()
    for g in direct_games:
        for slot in (g["white_team"], g["dark_team"]):
            m = _POOL_SLOT_RE.match(slot.strip())
            if m and team_matches(slot, team):
                groups.add(m.group(1).upper())

    # Determine team's actual pool finish per group if pool play is complete.
    # Used to prune composite-slot Sunday games for ranks we didn't finish.
    grp_pool_games: dict[str, list] = {}
    for g in division_games:
        wm = _POOL_SLOT_RE.match(g["white_team"].strip())
        dm = _POOL_SLOT_RE.match(g["dark_team"].strip())
        if wm and dm:
            grp_pool_games.setdefault(wm.group(1).upper(), []).append(g)
    team_pool_ranks: dict[str, int] = {}
    for grp in groups:
        pool_games = grp_pool_games.get(grp, [])
        if pool_games and all(
            g.get("white_score") is not None and g.get("dark_score") is not None
            for g in pool_games
        ):
            standings = _standings_for_group(grp, division_games)
            for i, s in enumerate(standings):
                if team_matches(s["team"], team):
                    team_pool_ranks[grp] = i + 1
                    break

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
    # Pre-populate from direct composite games so we don't add duplicates
    composite_added: set = set()
    for g in direct_games:
        for slot in (g["white_team"], g["dark_team"]):
            s = slot.strip()
            fm = _FINISH_SLOT_RE.match(s) or _COMPOSITE_SLOT_RE.search(s)
            if fm and fm.group(1).upper() in groups:
                composite_added.add(fm.group(1).upper())
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
                    rank_m = re.search(r'(\d+)', fm.group(0))
                    rank = int(rank_m.group(1)) if rank_m else None
                    if rank and grp in team_pool_ranks and team_pool_ranks[grp] != rank:
                        break
                    # When pool play is incomplete, show only one composite placeholder per group
                    if grp not in team_pool_ranks and grp in composite_added:
                        break
                    add_pool_rank = rank
                    add_grp = grp
                    add_placeholder = True
                    add_ph_depth = 2  # composite games don't expand further
                    break

                # W#/L# bracket
                wm = _WL_SLOT_RE.match(s)
                if wm:
                    ref = re.search(r'(\d+)$', wm.group(1))
                    if not ref:
                        continue
                    ref_num = str(int(ref.group(1)))
                    if ref_num not in reachable:
                        continue
                    src_game, src_ph, src_depth = reachable[ref_num]
                    if src_depth >= 2:
                        continue  # stop expanding beyond 2 levels of uncertainty
                    is_win_slot = s[0].upper() == 'W'
                    won = _team_won(team, src_game)

                    # Drop paths made impossible by a known result
                    if won is True  and not is_win_slot: continue
                    if won is False and     is_win_slot: continue

                    ph = src_ph or (won is None)
                    add_placeholder = ph if add_placeholder is None else (add_placeholder and ph)
                    add_ph_depth = (src_depth + 1) if ph else 0
                    break

            if add_placeholder is not None:
                g_copy = dict(g)
                g_copy["placeholder"] = add_placeholder
                if add_pool_rank:
                    g_copy["pool_rank"] = add_pool_rank
                if add_grp:
                    composite_added.add(add_grp)
                extras.append(g_copy)
                seen_ids.add(g["game_id"])
                n = _game_num(g["game_id"])
                if n:
                    reachable[n] = (g_copy, add_placeholder, add_ph_depth)
                changed = True

    return extras


def find_next_games(game, division_games):
    num = _game_num(game["game_id"])
    if not num:
        return None, None
    winner_next = loser_next = None
    for g in division_games:
        if g["game_id"] == game["game_id"]:
            continue
        for slot in (g["white_team"], g["dark_team"]):
            pm = re.match(r"^([WL])#([^-\s]+)", slot)
            if pm:
                ref = re.search(r"(\d+)$", pm.group(2))
                ref_num = str(int(ref.group(1))) if ref else pm.group(2)
                if ref_num == num:
                    if pm.group(1).upper() == "W": winner_next = g
                    else:                           loser_next  = g
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


def _read_futures_cumulative_standings(excel_src, team: str, sheet_name: str):
    """Aggregate season standings for all teams in the same age/gender group across all Futures weekends.

    Teams are grouped by age/gender only (not division) so promotion/relegation between
    weekends doesn't drop any weekend's data. Returns (label, standings_list) where each
    entry has: {name, points, reg_wins, shootout_wins, shootout_losses, reg_losses, rank, total, is_mine}.
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
                    in_section = True
                    m = re.search(r'Weekend\s*(\d+)', c0, re.I)
                    current_weekend = int(m.group(1)) if m else None
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

    # Locate our team
    my_clean = next((c for c in totals if team_upper in c or c in team_upper), None)
    if my_clean is None:
        return None, None

    label = f"{age} {gender.title()} · Season"

    standings = []
    for clean, t in totals.items():
        div_num = team_div.get(clean, (None, None))[1]
        row_out = dict(t, is_mine=(clean == my_clean),
                       division=f"D{div_num}" if div_num else None)
        standings.append(row_out)

    standings.sort(key=lambda x: (-x['points'], -x['reg_wins']))
    total = len(standings)
    for i, s in enumerate(standings):
        s['rank'] = i + 1
        s['total'] = total

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

def _next_summary(g, team, dg=None):
    opp = g["dark_team"] if team_matches(g["white_team"], team) else g["white_team"]
    return {
        "opponent": describe_slot(opp, dg),
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
        is_past = start < today
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

    def _team_fields(t):
        gender, age = _parse_sheet(t["sheet"])
        tier = _parse_tier(t["sheet"]) or _parse_tier(t["name"])
        age_gender = " ".join(p for p in [gender, age] if p) or None
        return {"name": t["name"], "sheet": t["sheet"], "friendly": t["friendly"],
                "age_gender": age_gender, "tier": tier}

    teams = [_team_fields(t) for t in deduped.values()]
    teams.sort(key=_team_sort_key)
    return jsonify(teams)


@app.route("/api/games/<tournament_id>/<path:team>")
def api_games(tournament_id, team):
    excel = find_excel(tournament_id)
    if not excel:
        abort(404)

    games    = _filter_by_dates(load_and_parse(excel), tournament_id)

    sheet    = request.args.get("sheet")
    my_games = [g for g in games
                if (sheet is None or g["sheet"] == sheet)
                and (team_matches(g["white_team"], team) or team_matches(g["dark_team"], team))]

    div_map = {}
    for g in games:
        div_map.setdefault(g["sheet"], []).append(g)

    # Expand to include bracket games (pool-finish and W#/L# slots) the team can reach
    for sheet_key in {g["sheet"] for g in my_games}:
        direct = [g for g in my_games if g["sheet"] == sheet_key]
        my_games.extend(_expand_bracket_games(team, direct, div_map.get(sheet_key, [])))

    my_games.sort(key=lambda g: (g["date"] or date.min, g["time"] or datetime.min.time()))

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

    # ── Game numbering: chronological time-slot order ───────────────────────
    # Each unique (date, time) gets the next sequential game number.
    # Games at the same time slot (alternate bracket paths) share a number.
    my_games_ranked = sorted(
        my_games,
        key=lambda g: (g.get("date") or date.min, g.get("time") or datetime.min.time()),
    )
    _game_num_map: dict[str, int] = {}
    slot_to_num: dict[tuple, int] = {}
    counter = 0
    for g in my_games_ranked:
        key = (g.get("date"), g.get("time"))
        if key not in slot_to_num:
            counter += 1
            slot_to_num[key] = counter
        _game_num_map[g["game_id"]] = slot_to_num[key]

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
        opp_sl = g["dark_team"] if team_matches(g["white_team"], team) else g["white_team"]
        color  = "WHITE" if team_matches(g["white_team"], team) else "DARK"

        our_rec  = None
        opp_rec  = None
        if show_records:
            trec    = sheet_records.get(g["sheet"], {})
            our_key = strip_prefix(g["white_team"] if color == "WHITE" else g["dark_team"]).upper()
            opp_key = strip_prefix(opp_sl).upper()
            our_rec = trec.get(our_key)
            opp_rec = trec.get(opp_key)

        opponent_label = describe_slot(opp_sl, dg)
        # Skip games where the opponent resolves to our own team (W#N self-play artifact)
        if team_matches(opponent_label, team):
            continue

        base = {
            "game_id":    gid,
            "date":       _fmt_date(g["date"]),
            "time":       _fmt_time(g["time"]),
            "location":   g["location"],
            "opponent":   opponent_label,
            "your_color": color,
            "game_num":    game_num,
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
                base["next"] = _next_summary(next_game, team, dg)
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
            # Suppress a scenario if that game is already shown as its own card
            if winner_next and winner_next["game_id"] not in my_game_ids:
                scenarios["win"]  = _next_summary(winner_next, team, dg)
            if loser_next and loser_next["game_id"] not in my_game_ids:
                scenarios["lose"] = _next_summary(loser_next,  team, dg)
            base["scenarios"] = scenarios if scenarios else None
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

    our_team_name = team.title()
    if my_games:
        sname = my_games[0].get("sheet", "")
        fn = friendly_team_name(team, sname)
        if fn:
            our_team_name = fn.split("·")[0].strip()

    return jsonify({
        "team":                 team,
        "our_team_name":        our_team_name,
        "played":               played_out,
        "upcoming":             upcoming_out,
        "placement":            placement,
        "pool_standing":        pool_standing,
        "cumulative_standings": cumulative_standings,
        "cumulative_division":  cumulative_division,
        "cache_age_s":          _cache_age(tournament_id),
        "cache_ttl_s":          URL_CACHE_TTL,
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
    return jsonify({"server_time": dt.datetime.utcnow().strftime("%Y-%m-%d %H:%M:%S UTC"), "tournaments": results})


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port, debug=False)

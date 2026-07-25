"""
Tournament Translator — Flask app
"""
from __future__ import annotations
import os, re, json, glob, io, time, base64, random, threading, math, concurrent.futures, heapq
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from functools import lru_cache
import requests
from flask import Flask, render_template, jsonify, request, abort

from parsers.detect import load_and_parse as _load_and_parse

_NJO_TOURNAMENTS = {"jo-quals", "junior-olympics"}

_parse_cache: dict[str, tuple[float, list]] = {}
_jo_quals_game_sig: int = -1  # tracks last-seen game count to detect data changes

def load_and_parse(filepath) -> list[dict]:
    """Cached wrapper: re-parses only when the file changes on disk."""
    global _jo_quals_game_sig
    if filepath == _JO_QUALS_SENTINEL:
        games = _fetch_jo_quals_games()
        sig = len(games)
        if sig != _jo_quals_game_sig:
            _jo_quals_game_sig = sig
            # Fire smoke test in background whenever the game count changes.
            # Safe: smoke test uses the URL cache, won't trigger another fetch cycle.
            threading.Thread(
                target=_run_trojan_smoke_test, args=("jo-quals",), daemon=True
            ).start()
        return games
    # Live-fetched data (BytesIO from a URL) has no filesystem mtime to key
    # on. find_excel tags it with a cache key tied to the URL + fetch time,
    # so repeated requests within the 5-min URL cache TTL reuse the parse
    # instead of re-running the full parser on every request.
    live_key = getattr(filepath, "_tt_cache_key", None)
    if live_key is not None:
        cached = _parse_cache.get(live_key)
        if cached and cached[0] == live_key:
            return list(cached[1])
        filepath.seek(0)
        games = _load_and_parse(filepath)
        _parse_cache[live_key] = (live_key, games)
        return list(games)
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
_DATA_DIR    = os.environ.get("DATA_DIR", os.path.join(os.path.dirname(os.path.abspath(__file__)), "data"))
RESULTS_DIR  = os.path.join(_DATA_DIR, "results")
COMMENTS_DIR = os.path.join(_DATA_DIR, "comments")
FEEDBACK_FILE = os.path.join(_DATA_DIR, "feedback.json")
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

QUIKSILVER_SHEETS_ID  = "18yVkTqV4amoIyESXsB1RzZSs1TI0_EKQ"
QUIKSILVER_GID_16U     = "472292980"

# 2026 Junior Olympics public schedule (Season 2, boys divisions). Single
# tournament per doc -- no gid-scoping needed (confirmed no cross-tournament
# contamination, unlike Quiksilver Cup's shared sheet).
JO_SHEETS_ID  = "1ycEOkayVwo_h37vL98PTXbzEnBpRU_-3S9l6NeiwCc4"
JO_SHEETS_URL = f"https://docs.google.com/spreadsheets/d/{JO_SHEETS_ID}/export?format=xlsx"

TOURNAMENT_URLS = {
    "kap7-intl":      "",  # update before Jan 2027 tournament
    "kap7-cup":       "https://onedrive.live.com/:x:/g/personal/6f253ef3afcfe1c8/IQDxdebmKQFASaux2nx7kWcvASg2jqJRHO5Kj9EwKW4D82o?rtime=GrzV592c3kg&redeem=aHR0cHM6Ly8xZHJ2Lm1zL3gvYy82ZjI1M2VmM2FmY2ZlMWM4L0lRRHhkZWJtS1FGQVNhdXgybng3a1djdkFTZzJqcUpSSE81S2o5RXdLVzREODJvP2U9WEdNa1FB",
    "turbo-cup":      "https://1drv.ms/x/c/6f253ef3afcfe1c8/IQB7PJXtfzNsT74lTYhWpOeXASFcmpB96L1OpYL_E6HBMM0?e=UsRrMb",
    "newport-invite": "https://onedrive.live.com/download?resid=6F253EF3AFCFE1C8!66694&authkey=!AO8pyWY0qwL2sYE",
    "jo-quals":       "",  # update when schedule is posted
    "junior-olympics": JO_SHEETS_URL,
    "quiksilver-cup":  (f"https://docs.google.com/spreadsheets/d/{QUIKSILVER_SHEETS_ID}"
                         f"/export?format=xlsx&gid={QUIKSILVER_GID_16U}"),
}

PRESET_URL_TOURNAMENTS = WPL_TOURNAMENTS | {k for k, v in TOURNAMENT_URLS.items() if v}
USER_URLS_FILE = os.path.join(_DATA_DIR, "user_urls.json")

_URL_CACHE: dict   = {}   # {url: (fetched_at, bytes)}
URL_CACHE_TTL      = 300  # re-fetch at most every 5 minutes

# A hung external host (seen in production: Google Sheets export occasionally
# stalls past even a 30s requests-level read timeout) must never be able to
# block a request past this ceiling. requests.get's own timeout parameter has
# not reliably enforced this in practice, so the fetch runs on a worker thread
# with a hard wall-clock deadline; if it blows past FETCH_HARD_TIMEOUT we give
# up and fall back to cached data (the leaked thread just finishes on its own
# later and its result is discarded). 12s was too tight in practice -- the JO
# sheet alone has legitimately taken 30-35s under normal (non-hung) load, which
# an earlier version of this cap treated as a failure on every cold start.
_FETCH_EXECUTOR = concurrent.futures.ThreadPoolExecutor(max_workers=8, thread_name_prefix="url-fetch")
FETCH_HARD_TIMEOUT = 50

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

# Admin email notify (feedback, smoke-test/pre-game alerts): sent via
# Resend's HTTPS API. Railway blocks outbound SMTP and, separately,
# ntfy.sh's host specifically (both confirmed live via /api/health/net-diag
# -- ENETUNREACH on both, while api.resend.com and every other tested host
# were reachable), so plain SMTP and ntfy were dead ends. Set RESEND_API_KEY
# as a Railway variable to enable; silently does nothing if unset.
ADMIN_NOTIFY_EMAIL = "djnurre@gmail.com"
RESEND_API_KEY      = os.environ.get("RESEND_API_KEY", "")

def _send_admin_email(subject: str, text: str) -> None:
    if not RESEND_API_KEY:
        return
    def _send():
        try:
            requests.post(
                "https://api.resend.com/emails",
                headers={"Authorization": f"Bearer {RESEND_API_KEY}"},
                json={
                    "from": "Tournament Translator <onboarding@resend.dev>",
                    "to": [ADMIN_NOTIFY_EMAIL],
                    "subject": subject,
                    "text": text,
                },
                timeout=10,
            )
        except Exception as exc:
            app.logger.warning("Admin email notify failed (%s): %s", subject, exc)
    threading.Thread(target=_send, daemon=True).start()

def _notify_feedback_email(text: str) -> None:
    _send_admin_email("Tournament Translator feedback", text)

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
    {"id": "quiksilver-cup",  "name": "Quiksilver Cup",         "dates": "Jul 10–12, 2026",
     "date_start": date(2026, 7, 10), "date_end": date(2026, 7, 12)},
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


_FETCH_INFLIGHT: set = set()  # urls currently being refreshed in the background
_FETCH_INFLIGHT_LOCK = threading.Lock()

def _do_fetch(url: str, onedrive: bool) -> bytes:
    if onedrive:
        token = base64.urlsafe_b64encode(url.encode()).rstrip(b"=").decode()
        fetch_url = f"https://api.onedrive.com/v1.0/shares/u!{token}/root/content"
    elif "onedrive.live.com" in url or "1drv.ms" in url:
        sep = "&" if "?" in url else "?"
        fetch_url = url + sep + "download=1"
    else:
        fetch_url = url

    def _do_get():
        resp = requests.get(fetch_url, allow_redirects=True, timeout=45)
        resp.raise_for_status()
        return resp.content
    content = _FETCH_EXECUTOR.submit(_do_get).result(timeout=FETCH_HARD_TIMEOUT)
    # A 2xx status is not proof of a real spreadsheet -- Google/OneDrive can
    # return an HTTP 200 interstitial/rate-limit HTML page instead of the
    # actual file under heavy load, which resp.raise_for_status() has no way
    # to catch (it only looks at the status code). Every real .xlsx is a ZIP
    # archive, always starting with the "PK" magic bytes; an HTML error page
    # never does. Skipping the cache write in that case means _fetch_url's
    # existing stale-while-revalidate behavior keeps serving the last GOOD
    # cached copy, and a later retry gets another chance -- instead of this
    # one bad response silently overwriting good data with something that
    # parses to zero games. Confirmed live 2026-07-25: Junior Olympics sat
    # stuck at 0 games for many minutes this way during a run of Google
    # rate-limit responses, self-inflicted by an unusually high deploy
    # frequency that same evening.
    if not content.startswith(b"PK"):
        raise ValueError(f"fetched content is not a valid .xlsx (got {content[:80]!r})")
    _URL_CACHE[url] = (time.time(), content)
    return content

def _refresh_url_background(url: str, onedrive: bool):
    """Kick off a background re-fetch, deduped so only one runs per URL at a
    time. Callers keep serving the stale-but-still-cached copy in the
    meantime instead of blocking on the network.

    The check-and-add into _FETCH_INFLIGHT must be one atomic step under a
    lock -- with 8 gthread request-handling threads, several near-simultaneous
    requests for the same stale URL could each see "not in the set yet" and
    all fire their own fetch before any of them got to the add. Confirmed
    live 2026-07-25: a burst of 4 near-simultaneous fetches to the same
    junior-olympics URL landed within 16 seconds during a live tournament
    evening (many parents' phones hitting a stale cache at once), which was
    actively adding to the request volume Google was already 429 rate-
    limiting us for and prolonging the outage."""
    with _FETCH_INFLIGHT_LOCK:
        if url in _FETCH_INFLIGHT:
            return
        _FETCH_INFLIGHT.add(url)
    def _run():
        try:
            _do_fetch(url, onedrive)
        except Exception as exc:
            app.logger.warning("Background refresh failed (%s): %s", url, exc)
        finally:
            _FETCH_INFLIGHT.discard(url)
    threading.Thread(target=_run, daemon=True).start()

def _fetch_url(url: str, *, onedrive=False) -> bytes | None:
    """Fetch Excel bytes from a URL with a 5-min cache.

    Stale-while-revalidate: once data has been fetched at least once, an
    expired cache entry is still served immediately while a background
    thread refreshes it, so a visitor only ever blocks on the live network
    on the very first fetch (e.g. right after a deploy). Returns None on
    failure with nothing cached yet."""
    now = time.time()
    cached = _URL_CACHE.get(url)
    if cached and now - cached[0] < URL_CACHE_TTL:
        return cached[1]
    if cached:
        _refresh_url_background(url, onedrive)
        return cached[1]
    try:
        return _do_fetch(url, onedrive)
    except Exception as exc:
        app.logger.warning("Fetch failed (%s): %s", url, exc)
        return None


def _bytesio_for_url(url: str, data: bytes) -> io.BytesIO:
    """Wrap live-fetched bytes, tagging them with a cache key tied to the
    URL cache's fetch timestamp so load_and_parse can reuse a prior parse
    instead of re-parsing on every request."""
    bio = io.BytesIO(data)
    fetched_at = _URL_CACHE.get(url, (None,))[0]
    bio._tt_cache_key = f"{url}@{fetched_at}"
    return bio


def find_excel(tournament_id: str):
    """Return an Excel file path, BytesIO from a live URL, or None.
    User-pasted URL (stored in user_urls.json) takes priority over all presets."""
    if tournament_id in JO_QUALS_TOURNAMENTS:
        return _JO_QUALS_SENTINEL
    user_url = _load_user_urls().get(tournament_id)
    if user_url:
        data = _fetch_url(user_url)
        if data:
            return _bytesio_for_url(user_url, data)
    # WPL tournaments always fetch live from Google Sheets — never use a
    # local file, which would be a stale snapshot from a past weekend.
    if tournament_id in WPL_TOURNAMENTS:
        data = _fetch_url(FUTURES_SHEETS_URL)
        return _bytesio_for_url(FUTURES_SHEETS_URL, data) if data else None
    keyword = FILE_MAP.get(tournament_id, "")
    if keyword:
        for f in os.listdir(EXCEL_DIR):
            if f.endswith(".xlsx") and keyword.lower() in f.lower():
                return os.path.join(EXCEL_DIR, f)
    preset = TOURNAMENT_URLS.get(tournament_id, "")
    if preset:
        data = _fetch_url(preset)
        return _bytesio_for_url(preset, data) if data else None
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

_ALL_HISTORICAL_GAMES_CACHE: list | None = None

def _all_historical_games() -> list:
    """All played games across every LOCAL Excel file only, cached in memory.

    Do not call this directly for any opponent-history feature (last_meeting,
    head-to-head, strength ratings, etc.) -- live-fetched tournaments (JO
    Quals, Junior Olympics, Quiksilver, WPL) never land in EXCEL_DIR as a
    local file, so this alone silently misses a tournament's own live games.
    That exact gap has shipped as a real bug three separate times (matchup
    history's sheet-scoping, the bracket tree's last_meeting, then the same
    bracket-tree bug again for JO). Use _h2h_source() instead, which combines
    this with whatever live game list the caller already has in hand.
    """
    global _ALL_HISTORICAL_GAMES_CACHE
    if _ALL_HISTORICAL_GAMES_CACHE is not None:
        return _ALL_HISTORICAL_GAMES_CACHE
    combined = []
    for fname in all_excels():
        fpath = os.path.join(EXCEL_DIR, fname)
        try:
            combined.extend(load_and_parse(fpath))
        except Exception:
            pass
    _ALL_HISTORICAL_GAMES_CACHE = combined
    return combined


def _h2h_source(*sources: list) -> list:
    """Combine any number of game lists (typically a tournament's own
    live-fetched games plus _all_historical_games()) into one deduped list,
    for any feature that needs a team's full opponent history: last_meeting,
    head-to-head tallies, strength ratings, etc.

    This is the ONE sanctioned way to build that combined list -- see the
    warning on _all_historical_games(). Earlier occurrences of this exact
    gap were each an independent inline combine-and-dedup block that someone
    had to remember to write correctly; routing every caller through this
    function means there is only one place left to get it right.

    Dedup keeps the FIRST occurrence of each game_id, so argument order
    encodes precedence -- pass the source you want to win on conflict first
    (e.g. live data before local archives, so a live tournament's own
    up-to-date copy of a game wins over a possibly-stale archived one).
    """
    seen: set = set()
    combined: list = []
    for source in sources:
        for g in source:
            gid = g.get("game_id")
            if gid in seen:
                continue
            seen.add(gid)
            combined.append(g)
    return combined


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
    # [A-Z]+(?:_[A-Z]+)? in the ordinal branch also covers compound
    # tier_pool codes seen in NJO's "Public Sched" export ("3RD AU_P-",
    # "2ND BZ_R-" -- tier "AU"/"BZ" + pool "P"/"R"), which the plain
    # [A-Z]+ alternative can't match past the underscore, leaving the whole
    # slot un-stripped and surfacing as a phantom team name. The optional
    # (?:\([^)]+\)\s*)? right before the dash covers the same tier_pool
    # code combined with a parenthetical win/loss reference, e.g.
    # "2ND BZ_S(L97)-TROJAN GOLD" or "1ST AU_S(W97)-TPCM SHARKS" -- found
    # live in the real 2026 Junior Olympics sheet, unresolved as of
    # 2026-07-14 but due to resolve to a real team name (including one of
    # our own) as pool play concludes.
    #
    # The bracket-position branch (composite pool-slot + parenthetical game
    # reference, e.g. "K1(2ndB)-TEAM") needed the identical (?:_[A-Z]+)?
    # extension: found live 2026-07-23 as pool play actually started
    # resolving these on the real sheet -- "NI_B1(W18)-TROJAN GOLD" and
    # "PT_P3(W44)-TROJAN CARDINAL" (compound bracket-position code, letters
    # + underscore + letter + digit, not just letters + digit) went
    # un-stripped the same way and surfaced as two literal phantom "team"
    # buttons on the home screen team list, right next to the real Trojan
    # Gold/Cardinal buttons.
    r"^(?:\d+(?:st|nd|rd|th)\s*(?:in\s+)?[A-Z]+(?:_[A-Z]+)?\s*(?:\([^)]+\)\s*)?-\s*"
    r"|[A-Z]+(?:_[A-Z]+)?\d+\s*\([^)]+\)\s*-\s*|[WL]\s*#\s*\d+\s*-?\s*|[A-Z]+\d+\s*-\s*|\d+\s*-\s*)(.*)",
    re.IGNORECASE,
)

def strip_prefix(s: str) -> str:
    m = _PREFIX_RE.match(s.strip())
    return m.group(1).strip() if m else s.strip()

def is_trojan(team_slot: str) -> bool:
    return "TROJAN" in strip_prefix(team_slot).upper()

def team_matches(slot: str, name: str) -> bool:
    stripped = strip_prefix(slot).upper()
    if name.upper() in stripped:
        return True
    # Organizers sometimes hand-type just the club name ("Trojan") into a
    # newly-resolved slot, dropping the team-color qualifier ("Cardinal" /
    # "Gold" / etc.) that actually distinguishes which specific team it is
    # (seen live: "K1(2ndB)- TROJAN" instead of "... TROJAN CARDINAL").
    # Without this, that game silently never matches the real team at all —
    # it's this app's whole job to track Trojan-affiliated teams, so treating
    # a bare "TROJAN" as a match for any Trojan-team search is a much safer
    # failure mode than dropping the team's own game from its own schedule.
    if stripped == "TROJAN" and "TROJAN" in name.upper():
        return True
    return False

# 'a' intentionally excluded: single-letter A/B/C are pool group names, not articles.
_LOWER_WORDS = frozenset({'in', 'of', 'or', 'and', 'the', 'an', 'at', 'by', 'for', 'to', 'vs'})

def _capitalize_word(w: str) -> str:
    """Capitalize the first letter in a word, skipping leading punctuation ('(trojan' → '(Trojan')."""
    for i, ch in enumerate(w):
        if ch.isalpha():
            return w[:i] + ch.upper() + w[i + 1:]
    return w

def _title(s: str) -> str:
    """Title-case that doesn't capitalize after digits ('1st' not '1St'),
    keeps common prepositions lowercase mid-string, and handles leading punctuation."""
    if not s:
        return s
    words = s.split()
    out = []
    for i, w in enumerate(words):
        if w and w[0].isdigit():                           # "1st" → "1st"
            out.append(w.lower())
        elif i > 0 and w.strip('().,').lower() in _LOWER_WORDS:
            out.append(w.lower())
        else:
            out.append(_capitalize_word(w))
    return ' '.join(out)


def _tournament_is_past(tournament_id: str) -> bool:
    """Mirrors the is_past logic in api_tournaments -- used to suppress
    smoke-test/pre-game-check alerts for a tournament that has already
    concluded. Nobody wants an email about a division from a past event."""
    t = next((x for x in KNOWN_TOURNAMENTS if x["id"] == tournament_id), None)
    if not t:
        return False
    today = datetime.now(ZoneInfo('America/Los_Angeles')).date()
    year_match  = _RE_YEAR.search(t["dates"])
    month_match = _RE_MONTH.search(t["dates"])
    yr = int(year_match.group(1)) if year_match else today.year
    mo = _MONTH_MAP.get(month_match.group(1).lower(), 1) if month_match else 1
    t_date = date(yr, mo, 1)
    start = t.get("date_start") or t_date
    end   = t.get("date_end")   or start
    return end < today

def _send_ntfy(tournament_id: str, title: str, body: str) -> None:
    """Smoke-test / pre-game-check alert email. Kept the old name (many call
    sites) but this now goes through Resend, same as feedback notify -- the
    original ntfy.sh version required NTFY_TOPIC, which was never actually
    set on Railway, so this path had never sent a real alert. Suppressed
    entirely for past tournaments -- nobody wants alerts about old data."""
    if _tournament_is_past(tournament_id):
        return
    _send_admin_email(title, body)

def _pool_teams_for_group(group: str, division_games: list) -> list[str]:
    """All team names seeded in a pool group, ordered by seed number."""
    teams = {}
    for g in division_games:
        for slot in (g["white_team"], g["dark_team"]):
            s = slot.strip()
            m = _POOL_SLOT_RE.match(s) or _COMPOSITE_POOL_SEED_RE.match(s)
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

    # Pass 1: direct pool-slot games (B1-X vs B2-Y format, or the WPL championship
    # crossover-pool seed variant "B1 (WIN GM #N) - X")
    direct_ids: set = set()
    game_results: dict = {}  # game_id -> (winner_key, loser_key) for W#/L# resolution

    for g in division_games:
        wt_raw, dt_raw = g["white_team"].strip(), g["dark_team"].strip()
        wm = _POOL_SLOT_RE.match(wt_raw) or _COMPOSITE_POOL_SEED_RE.match(wt_raw)
        dm = _POOL_SLOT_RE.match(dt_raw) or _COMPOSITE_POOL_SEED_RE.match(dt_raw)
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

    all_pool_teams: dict = {}
    _seed_num_by_team: dict = {}
    if all_groups:
        for group in all_groups:
            for t in _pool_teams_for_group(group, division_games):
                all_pool_teams[t] = group
    else:
        # No lettered pools anywhere in this division -- some tournament
        # formats (Kap7 Intl's single-division seeded brackets, 2026 Junior
        # Olympics Champ/Classic divisions) skip round-robin pools entirely
        # and seed straight into a single-elimination bracket. Root-round
        # slots name the seed directly ("12-TROJAN CARDINAL"); strip_prefix's
        # bare "\d+-" alternative already extracts the team name from these
        # -- the same path that already makes the Schedule/Bracket view work
        # correctly for these divisions (see _PREFIX_RE). Build the roster
        # from those roots instead of a pool letter, keeping the actual seed
        # number too so the pre-results strength prior below can still
        # reflect real seeding instead of guessing everyone's even.
        for g in division_games:
            for slot in (g["white_team"], g["dark_team"]):
                s = slot.strip()
                if re.match(r'^[WL]#', s, re.IGNORECASE):
                    continue  # a reference to another game, not a root entrant
                sm = re.match(r'^(\d+)\s*-\s*(.+)$', s)
                if not sm:
                    continue
                name = sm.group(2).strip()
                if _SLOT_LIKE_RE.match(name):
                    continue
                all_pool_teams.setdefault(name, None)  # no pool letter
                _seed_num_by_team.setdefault(name, int(sm.group(1)))

    if not all_pool_teams:
        return {}, {}

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

    # Opponent-adjusted strength via iterative Bradley-Terry MLE (MM algorithm).
    # Each game contributes a fractional "win" based on goal margin and recency.
    # Winning by more against a strong opponent converges to a higher strength score.
    # P(A beats B) = s_A / (s_A + s_B).
    _MARGIN_CAP   = 8
    _HALF_LIFE_WK = 3.0
    _PRIOR_WT     = 1.0   # regularization: half-win pseudo-game vs ghost team at strength 1
    _BT_ITERS     = 50

    def _team_key(name: str):
        return next((t for t in all_pool_teams if team_matches(t, name)), None)

    def _slot_name(slot: str) -> str:
        m = _POOL_SLOT_RE.match(slot.strip())
        if m:
            return m.group(3).strip()
        # Seeded-bracket root slots ("12-TROJAN CARDINAL") and other
        # non-pool prefixed forms: same strip_prefix + slot-like guard
        # already used as the fallback in the bracket-resolution loop below,
        # so a played root-round game in a no-pool-letter division still
        # feeds the strength model instead of being silently dropped.
        stripped = strip_prefix(slot.strip())
        return stripped if stripped and not _SLOT_LIKE_RE.match(stripped) else slot.strip()

    def _recency_wt(gdate) -> float:
        if not gdate:
            return 0.5
        weeks = max(0.0, (date.today() - gdate).days / 7.0)
        return math.exp(-weeks * math.log(2) / _HALF_LIFE_WK)

    # Infer pre-tournament seed so seed_perf can replace the uniform 0.5 prior,
    # letting the model respect real tournament seeding before any games are
    # played (the entire prediction, right when a bracket first posts).
    if _seed_num_by_team:
        # Seeded-bracket division: the sheet already states each team's real
        # overall seed directly, no inference needed.
        _n_seeded = max(_seed_num_by_team.values())
        _seed_perf: dict = {t: (_n_seeded + 1 - s) / _n_seeded
                            for t, s in _seed_num_by_team.items()}
    else:
        # Pool-letter division: infer seed from pool slot assignment (standard
        # snake draft). Pool A gets odd rounds (A1=seed1, A2=seed16 for 8
        # pools, A3=seed17, …), pool B gets the next, etc.
        _n_pools = len(all_groups)
        _pool_seed_rank: dict = {}
        # _pi = position within all_groups (already sorted), not ord(letter)
        # - ord('A'). A compound tier_pool group like "BZ_M" (see
        # _POOL_SLOT_RE) is more than one character, so ord() on it crashed
        # outright. Position-in-sorted-order is identical to the old
        # ord()-based index for every plain single-letter group (A=0, B=1,
        # ...), and is a reasonable, crash-free fallback for a compound one
        # -- true snake-draft seed order cannot be reliably inferred from a
        # compound code's alphabetic value either way, so this does not
        # regress accuracy, only availability. Found live 2026-07-25.
        for _grp in all_groups:
            _pi = all_groups.index(_grp)
            for _sp, _t in enumerate(_pool_teams_for_group(_grp, division_games), 1):
                _ri = _sp - 1
                if _ri % 2 == 0:
                    _pool_seed_rank[_t] = _ri * _n_pools + _pi + 1
                else:
                    _pool_seed_rank[_t] = _ri * _n_pools + (_n_pools - 1 - _pi) + 1

        _n_seeded = max(_pool_seed_rank.values()) if _pool_seed_rank else 1
        _seed_perf: dict = {t: (_n_seeded + 1 - r) / _n_seeded
                            for t, r in _pool_seed_rank.items()}

    # Collect game records: (white_key, dark_key, white_frac_win, recency_weight)
    _bt_records: list = []

    for _g in _h2h_source(_all_historical_games(), division_games):
        if not _g.get("played") or _g.get("white_score") is None:
            continue
        _wk = _team_key(_slot_name(_g["white_team"]))
        _dk = _team_key(_slot_name(_g["dark_team"]))
        if not _wk or not _dk:
            continue
        _ws, _ds = _g["white_score"], _g["dark_score"]
        _margin  = min(abs(_ws - _ds), _MARGIN_CAP)
        _mf      = math.log(1 + _margin) / math.log(1 + _MARGIN_CAP)
        _rw      = _recency_wt(_g.get("date"))
        _w_perf  = (0.5 + 0.5 * _mf) if _ws > _ds else (0.5 - 0.5 * _mf)
        _bt_records.append((_wk, _dk, _w_perf, _rw))

    # Index by team so each MM iteration is O(games) not O(teams * games)
    _team_games: dict = {t: [] for t in all_pool_teams}
    for _ki, _kj, _p, _rw in _bt_records:
        _team_games[_ki].append((_kj, _p,       _rw))
        _team_games[_kj].append((_ki, 1.0 - _p, _rw))

    # MM update: s_i = W_i / D_i
    #   W_i = Σ (recency * frac_win)        — weighted fractional wins
    #   D_i = Σ recency / (s_i + s_opp)    — expected wins under current model
    # Regularisation prior: one pseudo-game vs ghost (strength 1) where the
    # win credit = seed_perf (seed-1 team ≈ 1.0, seed-N team ≈ 1/N).
    # Before any real games this recovers the seeding order; as games are
    # played the real results gradually dominate.
    _s: dict = {t: 1.0 for t in all_pool_teams}

    for _ in range(_BT_ITERS):
        _s_new: dict = {}
        for _t in all_pool_teams:
            _W = _PRIOR_WT * _seed_perf.get(_t, 0.5)
            _D = _PRIOR_WT / (_s[_t] + 1.0)
            for _opp, _p, _rw in _team_games[_t]:
                _W += _rw * _p
                _D += _rw / (_s[_t] + _s[_opp])
            _s_new[_t] = _W / _D if _D > 0 else _s[_t]
        # Normalise to mean=1 each iteration for numerical stability
        _mean = sum(_s_new.values()) / len(_s_new) if _s_new else 1.0
        _s = {t: v / _mean for t, v in _s_new.items()}

    def _win_prob(key_a: str, key_b: str) -> float:
        sa, sb = _s.get(key_a, 1.0), _s.get(key_b, 1.0)
        return sa / (sa + sb)

    placement_counts: dict = {}
    all_placement_counts: dict = {t: {} for t in all_pool_teams}

    for _ in range(n_trials):
        pool_outcomes = {}
        for _pg in unplayed_pool_phase:
            _wm = _POOL_SLOT_RE.match(_pg["white_team"].strip())
            _dm = _POOL_SLOT_RE.match(_pg["dark_team"].strip())
            _wk = _team_key(_wm.group(3).strip()) if _wm else None
            _dk = _team_key(_dm.group(3).strip()) if _dm else None
            pool_outcomes[_pg["game_id"]] = (
                random.random() < (_win_prob(_wk, _dk) if _wk and _dk else 0.5)
            )

        group_standings: dict = {}
        for grp in all_groups:
            st = _standings_for_group(grp, division_games, pool_outcomes)
            group_standings[grp] = [s["team"] for s in st]

        game_results: dict = {}  # game_id -> (winner_name, loser_name)

        # Seed game_results with pool phase outcomes so W#N bracket slots can
        # look up the winner/loser of a pool game (CCA format references pool
        # games by number; they never appear in the bracket loop itself).
        for _pg in division_games:
            if _pg["game_id"] not in pool_phase_ids:
                continue
            _wm = _POOL_SLOT_RE.match(_pg["white_team"].strip())
            _dm = _POOL_SLOT_RE.match(_pg["dark_team"].strip())
            if not _wm or not _dm:
                continue
            _wt_n, _dt_n = _wm.group(3).strip(), _dm.group(3).strip()
            if _pg.get("played") and _pg.get("white_score") is not None:
                _pw = _pg["white_score"] > _pg["dark_score"]
            elif _pg["game_id"] in pool_outcomes:
                _pw = pool_outcomes[_pg["game_id"]]
            else:
                continue
            game_results[_pg["game_id"]] = (_wt_n, _dt_n) if _pw else (_dt_n, _wt_n)

        # [bracket_wins, last_game_date_ordinal, last_game_minute]
        team_rec: dict = {t: [0, 0, 0] for t in all_pool_teams}
        team_final_rank: dict = {}

        for g in bracket_games:
            wt = _resolve_slot_for_sim(g["white_team"], group_standings, game_results)
            dt = _resolve_slot_for_sim(g["dark_team"], group_standings, game_results)
            # CCA slots embed the resolved team name after a dash
            # ("W#12 - TROJAN CARDINAL (A)", "1ST C - TEAM") — strip_prefix extracts it.
            if not wt:
                _sn = strip_prefix(g["white_team"].strip())
                if _sn and not _SLOT_LIKE_RE.match(_sn):
                    wt = _team_key(_sn)
            if not dt:
                _sn = strip_prefix(g["dark_team"].strip())
                if _sn and not _SLOT_LIKE_RE.match(_sn):
                    dt = _team_key(_sn)
            if not wt or not dt:
                continue

            if g.get("played") and g.get("white_score") is not None:
                white_won = g["white_score"] > g["dark_score"]
            else:
                _wk, _dk = _team_key(wt), _team_key(dt)
                white_won = random.random() < (_win_prob(_wk, _dk) if _wk and _dk else 0.5)

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

        # Placement games state their exact final rank directly in "comments"
        # (e.g. "11th" -> winner finishes 11th, loser 12th). Raw bracket
        # win-count is not a valid placement proxy for these formats: every
        # team gets the same number of Saturday/Sunday games regardless of
        # which tier bracket they're in, so a team stuck in the "17th place"
        # bracket can rack up the same win count as one in the "1st place"
        # bracket. Prefer the stated rank whenever a game resolves both real
        # team names and a bare-ordinal comment; fall back to the old
        # win-count heuristic only when it doesn't (earlier rounds, formats
        # without this convention, or a slot that never got a real name).
        for g in bracket_games:
            rm = _PLACEMENT_COMMENT_RE.match(g.get("comments") or "")
            if not rm:
                continue
            gid = g["game_id"]
            if gid not in game_results:
                continue
            winner, loser = game_results[gid]
            rank = int(rm.group(1))
            wk, lk = _team_key(winner), _team_key(loser)
            if wk:
                team_final_rank[wk] = rank
            if lk:
                team_final_rank[lk] = rank + 1

        def _fallback_key(t):
            wins, date_ord, time_min = team_rec[t]
            grp = all_pool_teams[t]
            pool_rank = next((i for i, pt in enumerate(group_standings.get(grp, []))
                              if team_matches(pt, t)), 99)
            return (-wins, -date_ord, -time_min, pool_rank)

        # Assign the stated rank directly rather than sorting-by-key and using
        # list position -- a placement game that fails to resolve both real
        # team names (e.g. a spelling mismatch between rounds) leaves a gap in
        # the known ranks, and position-based numbering would silently shift
        # every subsequent team's placement by the size of that gap. Teams
        # without a known rank fill in whichever numbers are left over, still
        # ordered by the old win-count heuristic among themselves.
        _unknown = sorted(
            (t for t in all_pool_teams if t not in team_final_rank),
            key=_fallback_key,
        )
        _available = sorted(set(range(1, len(all_pool_teams) + 1)) - set(team_final_rank.values()))
        rank_of = dict(team_final_rank)
        for t, r in zip(_unknown, _available):
            rank_of[t] = r
        # Any leftover unknown teams past len(_available) (rank collisions in
        # the data) fall back to the end, past every assigned rank.
        for t in _unknown[len(_available):]:
            rank_of[t] = len(all_pool_teams) + 1

        sorted_teams = sorted(all_pool_teams.keys(), key=lambda t: rank_of[t])
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

    if not sheet:
        # Auto-detect: find whichever sheet this team appears in as a pool slot
        for _g in games:
            for _slot in (_g["white_team"], _g["dark_team"]):
                _m = _POOL_SLOT_RE.match(_slot.strip())
                if _m and team_matches(_m.group(3).strip(), team):
                    sheet = _g.get("sheet")
                    break
            if sheet:
                break

    division_games = [g for g in games if sheet is None or g["sheet"] == sheet]

    group = _find_team_pool_group(team, division_games)
    finish_probs, all_team_probs = _tournament_finish_probs(team, division_games)
    # group is None both for a genuinely-unposted schedule AND for a division
    # that has no round-robin pool stage at all (a straight seeded
    # single-elimination bracket -- see _tournament_finish_probs). Only the
    # first case has nothing to show; the second still gets bracket-position
    # predictions, just no pool-standings card (no group letter to show).
    if not group and not finish_probs:
        return jsonify({"pool": None, "finish_probs": []})

    standings    = _standings_for_group(group, division_games) if group else []
    current_rank = next((i + 1 for i, s in enumerate(standings) if team_matches(s["team"], team)), None)
    team_stats   = next((s for s in standings if team_matches(s["team"], team)), {})

    all_groups = {
        m.group(1).upper()
        for g in division_games
        for slot in (g["white_team"], g["dark_team"])
        if (m := _POOL_SLOT_RE.match(slot.strip()))
    }
    total_div_teams = (sum(len(_pool_teams_for_group(g, division_games)) for g in all_groups)
                        if all_groups else len(all_team_probs))

    def _modal_placement(probs):
        return max(probs.items(), key=lambda x: x[1])[0] if probs else 999

    # Build win/loss record for every team across ALL played games (pool + bracket).
    _all_records: dict = {}
    for _g in division_games:
        if not _g.get("played") or _g.get("white_score") is None:
            continue
        _wm = _POOL_SLOT_RE.match(_g["white_team"].strip())
        _dm = _POOL_SLOT_RE.match(_g["dark_team"].strip())
        _wname = _wm.group(3).strip() if _wm else strip_prefix(_g["white_team"].strip())
        _dname = _dm.group(3).strip() if _dm else strip_prefix(_g["dark_team"].strip())
        if not _wname or not _dname:
            continue
        if _SLOT_LIKE_RE.match(_wname) or _SLOT_LIKE_RE.match(_dname):
            continue
        _wk = next((t for t in all_team_probs if team_matches(t, _wname)), None)
        _dk = next((t for t in all_team_probs if team_matches(t, _dname)), None)
        if not _wk or not _dk:
            continue
        _all_records.setdefault(_wk, [0, 0])
        _all_records.setdefault(_dk, [0, 0])
        if _g["white_score"] > _g["dark_score"]:
            _all_records[_wk][0] += 1
            _all_records[_dk][1] += 1
        else:
            _all_records[_dk][0] += 1
            _all_records[_wk][1] += 1

    predicted_standings = sorted(
        [{"team": t,
          "predicted_rank": _modal_placement(all_team_probs.get(t, {})),
          "pct": round(max(all_team_probs.get(t, {}).values(), default=0) * 100),
          "is_mine": team_matches(t, team),
          "wins":   _all_records.get(t, (0, 0))[0],
          "losses": _all_records.get(t, (0, 0))[1]}
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
        } if group else None,
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
                or re.search(r'\bWinner\s+G?\d+', name, re.IGNORECASE)
                or re.search(r'\bLoser\s+G\d+', name, re.IGNORECASE)):
            return describe_slot(name, division_games, ref_date)
        # Double-prefixed: "BB2-1ST C - ORWP" → after first strip: "1ST C - ORWP" (still slot-like)
        if _SLOT_LIKE_RE.match(name):
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
        # Composite bracket-position/tier-pool code whose ENTIRE remaining
        # content is a game-reference parenthetical with no team name filled
        # in yet, e.g. "AG_T2(W59)" -- the organizer already seeded next
        # round's slot with a self-referential composite ("whoever becomes
        # AG_T2, i.e. whoever wins game 59") before that game has been
        # played, so there is genuinely no resolved name to show yet. Same
        # shape describe_slot already handles for plain "W#59"/"L#59" text,
        # just wrapped in a composite code instead -- found live 2026-07-23
        # while tracing why Friday's hypothetical bracket did not render for
        # Trojan Gold 16U even though the organizer had already published
        # Friday's real game rows (_resolve_slot_code_games's own fix, same
        # investigation, is what makes these rows reachable in the tree at
        # all; this is what makes their opponent text readable once reached).
        code_gm = (re.match(r'^[A-Z]+(?:_[A-Z]+)?\d*\s*\(([WL])(\d+)\)\s*$', slot, re.IGNORECASE)
                   if not (wm or win_gm) else None)
        if wm or win_gm or code_gm:
            if wm:
                want_winner = wm.group(1).upper() == "W"
                ref = re.search(r'(\d+)$', wm.group(2))
                ref_num = str(int(ref.group(1))) if ref else None
            elif win_gm:
                want_winner = win_gm.group(1).upper() == "WIN"
                ref_num = str(int(win_gm.group(2)))
            else:
                want_winner = code_gm.group(1).upper() == "W"
                ref_num = str(int(code_gm.group(2)))
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
            # Rank must come from the digit immediately before st/nd/rd/th, not
            # just the first digit in the slot — "K1(2ndB)-" has an unrelated
            # bracket-position digit ("K1") before the real ordinal ("2nd").
            ordinal_m = re.search(r'(\d+)(?:st|nd|rd|th)', slot, re.IGNORECASE)
            rank = int(ordinal_m.group(1)) if ordinal_m else None
            ordinal = ordinal_m.group(0) if ordinal_m else "?"
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
    # Bare numeric game_id with no prefix (e.g. "111.0" instead of
    # "16BX-111") -- confirmed live 2026-07-25, the JO GMID column
    # occasionally comes through this way. The regex below is written for
    # "PREFIX-NNN[-suffix]" shaped ids and treats "." as an arbitrary
    # non-trailing character it cannot cross, so on "111.0" it matched only
    # the trailing "0" -- silently breaking find_next_games' W#N/L#N lookup
    # for the whole division (every reference resolved to nonexistent game
    # "0"), stranding real, already-determined games as TBD. Checked first,
    # ahead of the regex, since a real "PREFIX-NNN" id is never also valid
    # float() input.
    try:
        f = float(game_id)
        if f == int(f):
            return str(int(f))
    except (TypeError, ValueError):
        pass
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

# [A-Z]+(?:_[A-Z])? (not a bare [A-Z]) so this also matches a compound
# tier_pool code, e.g. "BZ_M3(L51)-TROJAN GOLD" (tier "BZ" + pool "M"),
# not just a plain single-letter pool like "M3-TROJAN GOLD". The optional
# trailing (?:\([^)]*\))? covers the same code with its parenthetical
# game-reference still attached. Both forms are a strict superset of the
# old pattern -- every previously-matching input matches identically,
# unchanged. Found live 2026-07-25: this shape went completely
# unrecognized as a pool-membership slot, so the app never learned which
# pool group Trojan Gold 16U was actually in, and silently never found
# the real 3rd (placement) game that a completed 3-team round-robin pool
# produces once both round-robin games are played -- a parent needed that
# game before it happened, not after.
_POOL_SLOT_RE    = re.compile(r'^([A-Z]+(?:_[A-Z])?)(\d+)(?:\([^)]*\))?-(.+)', re.IGNORECASE)
# WPL championship-weekend crossover pool seed: "G2 (WIN GM #409) - SAN CLEMENTE" —
# the seed's team comes from a previous round's result, not a fixed name. Same
# capture-group layout as _POOL_SLOT_RE (letter, seed num, name) so callers can
# use either interchangeably once matched.
_COMPOSITE_POOL_SEED_RE = re.compile(
    r'^([A-Z])(\d+)\s*\([^)]*GM\s*#\d+[^)]*\)\s*-\s*(.+)$', re.IGNORECASE)
_WL_SLOT_RE      = re.compile(r'^[WL]#([^-\s]+)', re.IGNORECASE)   # dash optional (bare W#2 before scores)
# Accepts "2ndJ-" (no space), "2nd in J-" (worded), and "2ND J-" (bare space,
# no "in" -- seen live in JO's 10U/12U sheets, e.g. "2ND J- vs 2ND G-";
# previously unmatched, silently dropping that team's crossover game). Also
# accepts no trailing hyphen at all -- confirmed live 2026-07-22 in the JO
# 18U Invite sheet ("3rd A", "2nd A" with no team name and no dash yet, pool
# A not finished): previously fell through to raw-text display and tripped
# the bracket-confidence validator's unresolved-slot check (false yellow).
# Safe because the letter must be the entire rest of the string in that case
# -- a real team name would have more characters after it.
#
# [A-Z]+(?:_[A-Z])? (not a bare [A-Z]) for the same reason as _POOL_SLOT_RE
# above: a compound tier_pool code needs its own finish-slot to match too,
# e.g. "3RD BZ_M-" (the placement game for whoever finishes 3rd in pool
# "BZ_M"), not just "3rd M-". Strict superset of the old pattern -- every
# previously-matching single-letter input still matches identically.
_FINISH_SLOT_RE  = re.compile(r'^\d+(?:st|nd|rd|th)(?:\s+in\s+|\s+)?([A-Z]+(?:_[A-Z])?)(?:\s*-|\s*$)', re.IGNORECASE)
# Composite bracket slots like K4(1stG)- or K4(1stG) — group letter is inside parens
_COMPOSITE_SLOT_RE = re.compile(r'\(\d+(?:st|nd|rd|th)([A-Z])\)', re.IGNORECASE)
# Placement games often state their exact final rank as a bare ordinal in the
# comments column (e.g. "11th" -- winner finishes 11th, loser 12th). Strict
# full-string match so it never fires on freeform notes that merely mention
# an ordinal in passing.
_PLACEMENT_COMMENT_RE = re.compile(r'^\s*(\d+)(?:st|nd|rd|th)\s*$', re.IGNORECASE)
# WPL championship prelim-to-pool slot, e.g. "E2 (WIN GM #399)" / "F1 (LOS GM #399)"
_GM_WINLOS_RE = re.compile(r'\b(WIN|LOS)\s+GM\s+#(\d+)', re.IGNORECASE)

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

def _resolved_belongs_to_other_team(slot: str, team: str) -> bool:
    """True if a W#/L# (or WIN/LOS GM #N) slot has already been explicitly
    resolved to a real team name -- the organizer typed it in directly, e.g.
    "W#149-WEST SUBURBAN" -- and that name is not `team`. Such a slot
    definitively belongs to whoever it names; it is never a genuine branch
    for anyone else, regardless of whether the upstream game's result is
    otherwise ambiguous (see _team_won's tie case)."""
    resolved = strip_prefix(slot)
    return bool(resolved and resolved != slot
                and not _SLOT_LIKE_RE.match(resolved)
                and not team_matches(resolved, team))

def _resolve_slot_code_games(code: str, division_games: list) -> list:
    """Every game whose white_team/dark_team slot matches a bracket
    advancement code (e.g. "ni_D3" matches a slot starting with "NI_D3",
    case-insensitive) -- used when a W to#/L to# column holds a text slot
    code instead of a plain game number (some multi-stage formats seed the
    winner/loser directly into a later group-stage slot rather than a single
    numbered game). A code can legitimately match more than one game -- e.g.
    a 3-team round-robin sub-bracket plays every team against every other,
    so "NI_D3" appears in two separate games (vs NI_D1 and vs NI_D2) -- so
    this returns every match, not just the first.

    Also matches a code immediately followed by a parenthetical (e.g.
    "AG_T1(W51)"), not just a hyphenated team name -- the organizer's sheet
    can seed a future round's slot with only a self-referential composite
    reference and no resolved opponent at all yet ("whoever becomes AG_T1,
    which is whoever wins game 51"), with nothing between the code and the
    parenthesis. Missing this meant the app fell back to a generic, empty
    "TBD" placeholder for these games even when the real sheet already had
    the actual game rows entered -- found live 2026-07-23 while checking
    why Friday's hypothetical bracket wasn't showing for Trojan Gold 16U,
    even though the organizer had already published Friday's structure."""
    code_norm = code.strip().upper()
    if not code_norm:
        return []
    matches = []
    for g in division_games:
        for slot in (g["white_team"], g["dark_team"]):
            s = slot.strip().upper()
            if s == code_norm or s.startswith(code_norm + "-") or s.startswith(code_norm + "("):
                matches.append(g)
                break
    return matches

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
                    # This is a bracket app — before our rank in this group is known,
                    # every candidate slot (1st/2nd/3rd) is a genuine branch, not a
                    # guess to suppress. Add each one as its own alternative, tagged
                    # with its own pool_rank so the UI can label them distinctly
                    # ("If 1st in Pool" / "If 2nd in Pool" / "If 3rd in Pool"), the
                    # same way win/lose branches work for elimination brackets.
                    add_pool_rank = rank
                    add_grp = grp
                    add_placeholder = True
                    add_ph_depth = 2  # successors of placement games expand one more level
                    break

                # W#/L# bracket (standard W#N / L#N, plus CCA extended formats)
                wm = _WL_SLOT_RE.match(s)
                # CCA extended: "LN" at end of slot (no #) — e.g. "BB1-4TH G - L46"
                lm_ext = re.search(r'\bL(\d+)\s*$', s, re.IGNORECASE) if not wm else None
                # CCA extended: "Winner N" or "Winner GN" — e.g. "3RD G (WINNER 46)", "AAA4 (WINNER G33)"
                wm_ext = re.search(r'\bWinner\s+G?(\d+)', s, re.IGNORECASE) if not wm else None
                # CCA extended: "Loser GN" — e.g. "BBB2 (LOSER G31)"
                lm_gext = re.search(r'\bLoser\s+G(\d+)', s, re.IGNORECASE) if not wm and not lm_ext else None
                if wm or lm_ext or wm_ext or lm_gext:
                    if wm:
                        ref = re.search(r'(\d+)$', wm.group(1))
                        if not ref:
                            continue
                        ref_num = str(int(ref.group(1)))
                        is_win_slot = s[0].upper() == 'W'
                    elif lm_ext:
                        ref_num = str(int(lm_ext.group(1)))
                        is_win_slot = False  # "L46" = loser
                    elif lm_gext:
                        ref_num = str(int(lm_gext.group(1)))
                        is_win_slot = False  # "Loser G31" = loser
                    else:
                        ref_num = str(int(wm_ext.group(1)))
                        is_win_slot = True   # "Winner 46" / "Winner G33" = winner

                    # If this slot has already been explicitly resolved to a
                    # real, different team's name (organizer typed it in
                    # directly, e.g. "W#149-WEST SUBURBAN"), it definitively
                    # belongs to that team, not ours -- skip outright,
                    # regardless of whether the upstream game's win/loss is
                    # otherwise ambiguous (_team_won returns None for a tie,
                    # same as "not yet played", which previously let BOTH
                    # downstream branches through as candidates for us). Found
                    # live 2026-07-19 via a round-by-round replay of real 2025
                    # Junior Olympics results: a tied upstream game let a
                    # different team's real, already-played 9th-place game
                    # get silently attributed to our own played history as a
                    # loss against a team we never played.
                    if _resolved_belongs_to_other_team(s, team):
                        continue

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
                    if _resolved_belongs_to_other_team(s, team):
                        continue
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
                # Per-day cap: max 8 depth-2+ games per calendar day so Saturday can't
                # crowd out Sunday. Only depth-2+ games count toward the cap so that
                # depth-1 direct W/L successors (which are exempt) don't consume a slot
                # and block win/lose siblings later in the same BFS pass.
                # 8 (not 4) because a 3-way pool-rank branch (1st/2nd/3rd, each its own
                # 2-game round robin) is up to 6 legitimate alternatives on its own —
                # this is a bracket app, showing every real branch is the point.
                # Absolute cap of 20 as a safety net.
                _gdate = g.get("date", "")
                if _gdate and add_ph_depth >= 2:
                    _capped_day = sum(1 for e in extras
                                      if e.get("date") == _gdate and e.get("_ph_depth", 0) >= 2)
                    if _capped_day >= 8:
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


def _literal_team_count(sheet_games: list) -> int:
    """Count distinct teams in a sheet, from the STABLE Round-1 pool seeding
    only (LETTER+DIGIT-TEAMNAME, e.g. "B2-Trojan Cardinal") — never from
    composite/resolved slots elsewhere in the sheet.

    Round-1 seeding is fixed once at the start of the tournament and never
    edited again. Composite Round-2/3 slots, by contrast, get hand-typed
    resolved names added live during the tournament, in whatever format the
    organizer happens to use — e.g. "K1(2ndB)- TROJAN" instead of the full
    "TROJAN CARDINAL". Counting those as distinct teams inflates the total
    (confirmed live: 21 instead of 18), which breaks the whole-number ratio
    _expected_games_per_team(_by_day) needs and silently disables the TBD
    stub feature for the entire division, not just the affected team.
    """
    teams: set = set()
    for g in sheet_games:
        for slot in (g["white_team"], g["dark_team"]):
            m = _POOL_SLOT_RE.match(slot.strip())
            if m:
                teams.add(m.group(3).strip().upper())
    return len(teams)


def _expected_games_per_team(sheet_games: list) -> int | None:
    """If every team in this sheet's division plays the same total number of
    games, return that fixed count. True for round-robin + crossover formats
    (Quiksilver Cup, CCA pool play) where the bracket size is fixed in
    advance. False for elimination formats (WPL, Kap7) where winning or
    losing changes how many games a team plays, so there is no single
    "expected total" — callers must get None back and skip any
    placeholder-count logic rather than guess.

    Heuristic: (distinct game_ids * 2 team-slots) / (distinct literal team
    names) must divide evenly. Composite/finish-slot placeholders are
    excluded from the team count since they aren't real names yet.
    """
    if not sheet_games:
        return None
    n_teams = _literal_team_count(sheet_games)
    if not n_teams:
        return None
    game_ids = {g["game_id"] for g in sheet_games}
    ratio = (len(game_ids) * 2) / n_teams
    return int(ratio) if ratio == int(ratio) else None


def _expected_games_per_team_by_day(sheet_games: list) -> dict | None:
    """Same fixed-count logic as _expected_games_per_team, but broken down by
    calendar date. Every date's rows already carry a real date/time/location
    in this format even before an opponent is resolvable, so a "games
    remaining" placeholder can say WHICH DAY it falls on without needing to
    know who or where. Returns {date: expected_games_that_day} or None if any
    date's breakdown doesn't divide evenly (bail rather than guess)."""
    if not sheet_games:
        return None
    n_teams = _literal_team_count(sheet_games)
    if not n_teams:
        return None
    by_date: dict = {}
    for g in sheet_games:
        d = g.get("date")
        if d:
            by_date.setdefault(d, set()).add(g["game_id"])
    result: dict = {}
    for d, ids in by_date.items():
        ratio = (len(ids) * 2) / n_teams
        if ratio != int(ratio):
            return None
        result[d] = int(ratio)
    return result


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
        if m:
            return str(int(m.group(1)))
        # Bare numeric GMID with no prefix (e.g. "111.0" instead of
        # "16BX-111") -- confirmed live 2026-07-25, the JO GMID column
        # occasionally comes through this way, and the hyphen-only regex
        # above then matches nothing for the entire sheet, leaving
        # gnum_map empty and silently breaking every w_to/l_to advancement
        # in the division (fell back to a weaker text-slot matcher that
        # does not cover pool/placement branches, stranding Trojan Gold's
        # real Saturday/Sunday games as TBD stubs even though the sheet
        # already had them filled in). Same fallback shape as
        # _norm_game_num below, applied to the game_id side too.
        try:
            f = float(gid)
        except (TypeError, ValueError):
            return None
        return str(int(f)) if f == int(f) else None

    def _norm_game_num(v):
        """Normalize a w_to/l_to cell value to the same clean-integer string
        _gnum_str uses, e.g. 131.0 or "131.0" -> "131" -- openpyxl can read
        one particular numeric cell as a float even when every other
        w_to/l_to cell in the same column reads as a clean int, depending
        on that cell's own number format. A raw str(v) then never matches
        gnum_map's clean-integer keys, and (since the result still looks
        like a string) silently falls through to the non-numeric slot-code
        fallback below instead, which finds nothing either -- confirmed
        live 2026-07-25: this dead-ended Trojan Gold 16U's real placement
        game (16BX-111, itself only reachable via a separate fix earlier
        the same day) right where it should have continued into the next
        two real games, instead of a synthetic "further TBD" stub. Returns
        the value unchanged (as a string) for a genuinely non-numeric code
        (e.g. "ag_T1"), so that fallback still works correctly for it."""
        if v is None:
            return None
        s = str(v)
        try:
            f = float(s)
        except (TypeError, ValueError):
            return s
        return str(int(f)) if f == int(f) else s

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
        node["pool_next"]      = {}  # {game_id: rank} -- pool-finish placement branches
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

        # Advancement is normally tracked via the w_to/l_to columns. Some
        # seasons' sheets leave those columns blank and encode advancement
        # the older way instead -- W#N/L#N text inside the next game's own
        # slot (the same convention WPL/CCA formats use). Resolve w_to/l_to
        # first; for whichever side is missing, fall back to the same
        # slot-reference lookup find_next_games uses elsewhere, so the tree
        # isn't silently truncated to a single node when the columns are
        # empty. Confirmed live: the 2026 JO sheet leaves w_to/l_to blank
        # for every game and relies entirely on the text-slot convention.
        w_num = _norm_game_num(game.get("w_to"))
        l_num = _norm_game_num(game.get("l_to"))
        win_next_g  = gnum_map.get(w_num) if w_num else None
        lose_next_g = gnum_map.get(l_num) if l_num else None
        if win_next_g is None or lose_next_g is None:
            fb_win, fb_lose = find_next_games(game, division_games)
            if win_next_g is None:
                win_next_g = fb_win
            if lose_next_g is None:
                lose_next_g = fb_lose

        # Follow win path
        if win_next_g and win_next_g["game_id"] not in seen:
            # Only follow if team appears or result unknown
            involved = (team_matches(win_next_g["white_team"], team)
                        or team_matches(win_next_g["dark_team"], team))
            # If not involved by name, and this slot has already been
            # explicitly resolved to a different, real team (not just
            # unresolved), it definitively belongs to that team -- never
            # follow into it, regardless of whether `won` is otherwise
            # ambiguous. `won is None` covers BOTH "not yet played" (where
            # speculative branching into a genuinely unresolved slot is the
            # intended hypothetical preview) AND "tied in regulation" (where
            # it previously let an already-named different team's real,
            # played game get misattributed into our own tree). Found live
            # 2026-07-19, same root cause as the twin fix in
            # _expand_bracket_games -- see _resolved_belongs_to_other_team.
            belongs_to_other = not involved and any(
                _resolved_belongs_to_other_team(s, team)
                for s in (win_next_g["white_team"], win_next_g["dark_team"])
            )
            next_ph = is_ph or (won is False)  # placeholder if we lost
            if not belongs_to_other and (involved or won is not False):
                child = _follow(win_next_g, game["game_id"], "win", next_ph, depth + 1)
                if child:
                    node["win_next_ids"].append(win_next_g["game_id"])

        # Follow lose path
        if lose_next_g and lose_next_g["game_id"] not in seen:
            involved = (team_matches(lose_next_g["white_team"], team)
                        or team_matches(lose_next_g["dark_team"], team))
            belongs_to_other = not involved and any(
                _resolved_belongs_to_other_team(s, team)
                for s in (lose_next_g["white_team"], lose_next_g["dark_team"])
            )
            next_ph = is_ph or (won is True)  # placeholder if we won
            if not belongs_to_other and (involved or won is not True):
                child = _follow(lose_next_g, game["game_id"], "lose", next_ph, depth + 1)
                if child:
                    node["lose_next_ids"].append(lose_next_g["game_id"])

        # Slot-code advancement (w_to/l_to holding a bracket seed code like
        # "ni_D3" instead of a plain game number): only reached when neither
        # the w_to/l_to game-number lookup nor the W#/L#-text fallback above
        # found anything. Unlike those, a slot code can resolve to MULTIPLE
        # games at once (a round-robin sub-bracket plays every pairing), so
        # every match is attached as its own reachable branch rather than
        # picking just one. Found live 2026-07-19: Trojan Gold 18U's Invite
        # bracket dead-ended after its Day 1 cross game because "w to #" /
        # "l to #" held slot codes ("ni_D3", "cu_C1"), which the parser used
        # to silently discard as unparseable instead of preserving them.
        if win_next_g is None and isinstance(game.get("w_to"), str) and won is not False:
            for g2 in _resolve_slot_code_games(game["w_to"], division_games):
                if g2["game_id"] in seen:
                    continue
                involved2 = (team_matches(g2["white_team"], team)
                             or team_matches(g2["dark_team"], team))
                if not involved2 and any(_resolved_belongs_to_other_team(s, team)
                                          for s in (g2["white_team"], g2["dark_team"])):
                    continue
                child = _follow(g2, game["game_id"], "win", is_ph or (won is False), depth + 1)
                if child:
                    node["win_next_ids"].append(g2["game_id"])
        if lose_next_g is None and isinstance(game.get("l_to"), str) and won is not True:
            for g2 in _resolve_slot_code_games(game["l_to"], division_games):
                if g2["game_id"] in seen:
                    continue
                involved2 = (team_matches(g2["white_team"], team)
                             or team_matches(g2["dark_team"], team))
                if not involved2 and any(_resolved_belongs_to_other_team(s, team)
                                          for s in (g2["white_team"], g2["dark_team"])):
                    continue
                child = _follow(g2, game["game_id"], "lose", is_ph or (won is True), depth + 1)
                if child:
                    node["lose_next_ids"].append(g2["game_id"])

        return node

    root = my_games[0]
    _follow(root, None, None, False)

    # Some teams have multiple disconnected game segments within the same
    # tournament weekend -- e.g. a placement/consolation phase that starts a
    # fresh w_to/l_to numbering sequence unconnected to the pool-phase
    # bracket, or later round-robin games with no advancement links at all.
    # Following only the very first chain silently drops these later
    # segments entirely (confirmed live: a real team lost 6 of 9 actual
    # games from the tree this way). Start a fresh traversal for every one
    # of the team's known games not already reached by the first chain.
    for g in my_games:
        if g["game_id"] not in seen:
            _follow(g, None, None, False)

    # Pool-finish placement games (e.g. "2ndB-") are a separate branch type
    # from win/lose bracket games -- which one is real depends on the team's
    # pool STANDING, not a single game's result, so they can't be reached via
    # w_to/l_to or a W#/L# reference at all. Without this, the tree silently
    # omits the post-pool placement round even though it's a real upcoming
    # game (confirmed live: TROJAN GOLD's 18U Invite bracket tree stopped at
    # 2 nodes despite a 3rd, real placement game being reachable). Attach
    # every rank candidate as a pool_next branch of the team's LAST pool-
    # phase game, the same relationship _expand_bracket_games already
    # captures for the schedule list.
    _team_pool_group = None
    for g in my_games:
        for slot in (g["white_team"], g["dark_team"]):
            if _POOL_SLOT_RE.match(slot.strip()) and team_matches(slot, team):
                _team_pool_group = _POOL_SLOT_RE.match(slot.strip()).group(1).upper()
                break
        if _team_pool_group:
            break
    if _team_pool_group:
        pool_phase_games = sorted(
            [g for g in my_games
             for slot in (g["white_team"], g["dark_team"])
             if _POOL_SLOT_RE.match(slot.strip()) and team_matches(slot, team)
             and g["game_id"] in seen],
            key=lambda g: (g.get("date") or date.min, g.get("time") or datetime.min.time()),
        )
        if pool_phase_games:
            anchor = pool_phase_games[-1]
            anchor_node = next((n for n in out if n["game_id"] == anchor["game_id"]), None)
            if anchor_node:
                covered_ranks: set = set()
                a_date = None
                for g2 in division_games:
                    if g2["game_id"] in seen:
                        continue
                    for slot in (g2["white_team"], g2["dark_team"]):
                        s2 = slot.strip()
                        fm = _FINISH_SLOT_RE.match(s2) or _COMPOSITE_SLOT_RE.search(s2)
                        if fm and fm.group(1).upper() == _team_pool_group:
                            # Same guard as the win/lose follow above: a
                            # finish-slot candidate already explicitly
                            # resolved to a different real team belongs to
                            # that team, not ours, regardless of which rank
                            # it represents.
                            if _resolved_belongs_to_other_team(s2, team):
                                break
                            rank_m = re.search(r'(\d+)', fm.group(0))
                            if not rank_m:
                                break
                            rank = int(rank_m.group(1))
                            has_score = (g2.get("played") and g2.get("white_score") is not None
                                         and g2.get("dark_score") is not None)
                            child = _follow(g2, anchor["game_id"], f"pool_{rank}",
                                             not has_score, depth=1)
                            if child:
                                anchor_node["pool_next"][g2["game_id"]] = rank
                                covered_ranks.add(rank)
                                a_date = a_date or g2.get("date")
                            break

                # TBD stub for any pool-finish rank that's structurally
                # possible (or, once the pool is fully decided, the team's
                # own real rank) but has no discoverable game/slot in the
                # sheet at all -- e.g. a 3-team pool where the organizer's
                # bracket only publishes a cross-game for 2nd place, and
                # 1st/3rd advance directly via a seeding rule that lives
                # nowhere in the spreadsheet. Rather than silently show one
                # branch and drop the other two, show an honest "TBD" card
                # for each so the connector line has somewhere real to end.
                # Found live 2026-07-19/20: Trojan Gold 18U's Pool B is
                # exactly this shape.
                pool_size = len(_pool_teams_for_group(_team_pool_group, division_games))
                if pool_size > len(covered_ranks):
                    standings = _standings_for_group(_team_pool_group, division_games)
                    pool_decided = bool(standings) and all(
                        s["wins"] + s["losses"] >= pool_size - 1 for s in standings
                    )
                    if pool_decided:
                        our_rank = next((i + 1 for i, s in enumerate(standings)
                                          if team_matches(s["team"], team)), None)
                        missing_ranks = ({our_rank} - covered_ranks) if our_rank else set()
                    else:
                        missing_ranks = set(range(1, pool_size + 1)) - covered_ranks
                    for rank in sorted(missing_ranks):
                        stub_id = f"__tbd_pool_{anchor['game_id']}_{rank}"
                        anchor_node["pool_next"][stub_id] = rank
                        out.append({
                            "game_id":        stub_id,
                            "date":           a_date or anchor.get("date"),
                            "time":           None,
                            "location":       None,
                            "white_team":     "TBD",
                            "dark_team":      "TBD",
                            "played":         False,
                            "placeholder":    True,
                            "tbd_stub":       True,
                            "src_game_id":    anchor["game_id"],
                            "src_path":       f"pool_{rank}",
                            "win_next_ids":   [],
                            "lose_next_ids":  [],
                            "pool_next":      {},
                            "sunday_pair_id": None,
                            "tree_format":    "bracket",
                        })

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
            # Standard W#N / L#N format (with or without spaces: W#12, W #12, W # 12)
            pm = re.match(r"^([WL])\s*#\s*([^-\s]+)", slot)
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
            # CCA extended: "Winner N" / "Winner GN" — e.g. "3RD G (WINNER 46)", "AAA4 (WINNER G33)"
            wm = re.search(r'\bWinner\s+G?(\d+)', slot, re.IGNORECASE)
            if wm and str(int(wm.group(1))) == num:
                winner_next = g
                continue
            # CCA extended: "Loser GN" — e.g. "BBB2 (LOSER G31)"
            lm_g = re.search(r'\bLoser\s+G(\d+)', slot, re.IGNORECASE)
            if lm_g and str(int(lm_g.group(1))) == num:
                loser_next = g
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

    # Organizers sometimes hand-type just the bare club name ("TROJAN") into
    # a newly-resolved bracket slot instead of the full team name ("TROJAN
    # CARDINAL") used elsewhere in the sheet -- silently splitting one team's
    # record across two dict keys (seen live: Quiksilver Cup's Saturday/Sunday
    # K-bracket slots use bare "TROJAN" while Friday's pool slots spell out
    # "TROJAN CARDINAL"). Fold the bare key into the one specific team it
    # unambiguously refers to.
    for bare in [k for k in records if " " not in k]:
        specific = [k for k in records if k != bare and k.startswith(bare + " ")]
        if len(specific) == 1:
            tgt = specific[0]
            for field in ("wins", "losses", "ties"):
                records[tgt][field] += records[bare][field]
            # Keep the bare key pointing at the same (now-merged) record so a
            # lookup via either spelling returns the full combined total.
            records[bare] = records[tgt]

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
    your_dark  = team_matches(game["dark_team"], team)
    if not your_white and not your_dark:
        # Direct name match failed — team may be identified only via a
        # pool-rank slot (e.g. "H1(1stB)-"), not a literal name in this game.
        # Without this, the team is silently assumed to be on the dark side,
        # which inverts win/loss whenever they're actually white.
        pr, grp = game.get("pool_rank"), game.get("pool_rank_group")
        if pr and grp:
            for slot, is_white in ((game["white_team"], True), (game["dark_team"], False)):
                s = slot.strip()
                fm = _COMPOSITE_SLOT_RE.search(s) or _FINISH_SLOT_RE.match(s)
                if fm and fm.group(1).upper() == grp:
                    rank_m = re.search(r'(\d+)(?:st|nd|rd|th)', s, re.IGNORECASE)
                    if rank_m and int(rank_m.group(1)) == pr:
                        your_white = is_white
                        break
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

    # Finish-slot: pool_rank_group tells us which pool letter is ours.
    # Handles both simple ("1stB-") and composite ("H1(1stB)-") slot formats —
    # composite slots don't match _FINISH_SLOT_RE (anchored to a leading digit),
    # so _COMPOSITE_SLOT_RE must be tried too or these games are never matched.
    pr  = g.get("pool_rank")
    grp = g.get("pool_rank_group")
    if pr and grp:
        for slot, other in ((white, dark), (dark, white)):
            s = slot.strip()
            fm = _COMPOSITE_SLOT_RE.search(s) or _FINISH_SLOT_RE.match(s)
            if fm and fm.group(1).upper() == grp:
                # Rank is the digit immediately before st/nd/rd/th, not just the
                # first digit in the slot — a composite slot's bracket-position
                # number (e.g. the "1" in "K1(2ndB)-") is a different number.
                rank_m = re.search(r'(\d+)(?:st|nd|rd|th)', s, re.IGNORECASE)
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


_PAST_TOURNAMENT_STATUS_CACHE: dict[str, tuple[bool, bool]] = {}  # {id: (has_file, has_excel)}

@app.route("/api/tournaments")
def api_tournaments():
    today = datetime.now(ZoneInfo('America/Los_Angeles')).date()
    out = []
    for t in KNOWN_TOURNAMENTS:
        year_match  = _RE_YEAR.search(t["dates"])
        month_match = _RE_MONTH.search(t["dates"])
        yr = int(year_match.group(1)) if year_match else today.year
        mo = _MONTH_MAP.get(month_match.group(1).lower(), 1) if month_match else 1
        t_date = date(yr, mo, 1)
        start = t.get("date_start") or t_date
        end   = t.get("date_end")   or start
        is_past = end < today
        days_until = (start - today).days if not is_past else None

        # A tournament that already ended has a fixed, final schedule — its
        # live URL (if any) will never produce new data and, for old events,
        # is often a since-expired OneDrive share link. Re-fetching it every
        # 5 minutes forever just adds latency (or a hang) to every visitor's
        # home-screen load for no benefit, so resolve it once per process
        # lifetime and reuse that answer instead of hitting the network again.
        if is_past and t["id"] in _PAST_TOURNAMENT_STATUS_CACHE:
            has_file, has_excel = _PAST_TOURNAMENT_STATUS_CACHE[t["id"]]
        else:
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
            if is_past:
                _PAST_TOURNAMENT_STATUS_CACHE[t["id"]] = (has_file, has_excel)

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
            # If strip_prefix couldn't remove the prefix (unknown slot format), the
            # result still looks like a slot — skip it to avoid phantom team entries.
            if _SLOT_LIKE_RE.match(name):
                continue
            if "TROJAN" in name.upper():
                key = (name, g["sheet"])
                counts[key] = counts.get(key, 0) + 1

    # A bare "Trojan" entry (organizer typed just the club name into a
    # resolved slot, dropping the color qualifier — see team_matches) isn't
    # a separate team; fold its count into the one more-specific "Trojan ..."
    # candidate in the same sheet, if there's exactly one. If there's more
    # than one (e.g. both Trojan Cardinal and Trojan Gold in the same sheet),
    # we can't tell which it means — leave it showing rather than guess wrong.
    for sheet in {s for (_, s) in counts}:
        bare_key = ("TROJAN", sheet)
        if bare_key not in counts:
            continue
        specific = [k for k in counts if k[1] == sheet and k[0].upper() != "TROJAN"
                    and "TROJAN" in k[0].upper()]
        if len(specific) == 1:
            counts[specific[0]] += counts.pop(bare_key)

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
    r'^(?:\d+(?:st|nd|rd|th)\s*(?:in\s+)?[A-Z]|[A-Z]\d+[-\(]|[WL]\s*#|WIN\s+GM|LOS\s+GM)',
    re.IGNORECASE,
)
# describe_slot's own "Nth in Pool X (A or B or C)" preview for a pool-finish
# slot that can't be decided yet (no pool games played, so every team in the
# pool is still a live candidate -- see describe_slot's _FINISH_SLOT_RE/
# _COMPOSITE_SLOT_RE branch) coincidentally starts with the same "<ordinal> in
# <capital letter>" shape _SLOT_LIKE_RE looks for in a genuinely raw, undecoded
# slot ("2ND G-...") -- "in Pool" begins "in P", an uppercase letter right
# after "in ". Any caller treating a _SLOT_LIKE_RE match as "still broken" must
# carve this one known-good format out first. Found live 2026-07-19: this false
# positive held every team in Junior Olympics' 18U Invite division (including
# Trojan Gold) at yellow confidence, and would have fired on every one of them
# in the automated pre-game sweep notification 12-25h before the tournament.
_POOL_PREVIEW_RE = re.compile(r'^\d+(?:st|nd|rd|th)\s+in\s+Pool\s+[A-Z]\s*\(', re.IGNORECASE)


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
                              div_games: list = None, anchor_date=None,
                              allow_multiple_roots: bool = False) -> list[str]:
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
    # NJO trees legitimately have multiple roots: round-robin pool games (no
    # w_to/l_to at all) and a later placement/consolation phase (a fresh
    # w_to/l_to numbering sequence) are genuinely disconnected segments, not
    # a sign of a broken tree. This check only makes sense for WPL, where the
    # tree builder always produces one connected bracket.
    if not allow_multiple_roots and len(roots) != 1:
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
                           serialized_nodes: list = None, is_njo: bool = False) -> tuple[str, list]:
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
    # NJO is structurally different: _build_njo_game_tree legitimately starts a
    # fresh traversal for every disconnected segment a team has (pool phase,
    # placement phase, sometimes a third consolation phase), each with its own
    # speculative win/lose branching -- confirmed against real 2025 NJO data,
    # a genuine, single-lineage bracket routinely lands at 9-10 nodes with
    # nothing wrong. Raised from 20 to 50 2026-07-25: a real 3-team pool's 3rd-
    # place finisher legitimately enters a deep, fully speculative consolation
    # ladder (every remaining round still branches both win and lose, since
    # none of it is resolved yet) -- confirmed live against Trojan Gold 16U's
    # actual data, 36 real nodes, every one independently verified against the
    # flat schedule as a real scheduled game, not a data error. The node-count
    # ceiling is a blunt safety net; the per-edge foreign-game check just above
    # it in this function is the precise one and is unaffected by this change.
    node_ceiling = 50 if is_njo else 7
    if len(nodes) > node_ceiling:
        red.append(
            f"bracket has {len(nodes)} nodes — expected ≤{node_ceiling}; "
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
    # via WIN GM #N / LOS GM #N (WPL championship style) or plain W#N / L#N
    # (WPL/CCA/NJO's older convention -- e.g. the 2026 JO sheet, which leaves
    # the dedicated w_to/l_to columns blank and encodes advancement this way).
    tree_game_nums = {_game_num(n["game_id"]) for n in nodes} - {None}
    _node_by_id = {n["game_id"]: n for n in nodes}
    def _slot_refs_tree_game(slot: str) -> bool:
        wgm = re.search(r'\b(?:WIN|LOS)\s+GM\s+#(\d+)', slot, re.IGNORECASE)
        if wgm and str(int(wgm.group(1))) in tree_game_nums:
            return True
        wl = re.match(r'^([WL])\s*#\s*([^-\s]+)', slot.strip(), re.IGNORECASE)
        if wl:
            ref = re.search(r'(\d+)$', wl.group(2))
            ref_num = str(int(ref.group(1))) if ref else None
            if ref_num and ref_num in tree_game_nums:
                return True
        return False
    def _slot_refs_parent_advancement_code(node: dict, slot: str) -> bool:
        """True if this node's parent has a w_to/l_to slot code (e.g.
        "ni_D3", see _resolve_slot_code_games) that this slot matches --
        the third legitimate advancement mechanism alongside WIN/LOS GM #N
        and plain W#/L#, for formats that seed a winner/loser directly into
        a later group-stage slot rather than a single numbered game.

        Mirrors _resolve_slot_code_games's own matching exactly -- must
        also accept a code immediately followed by a parenthetical (e.g.
        "AG_T1(W51)"), not just a hyphenated team name. Missing this here
        (even after fixing _resolve_slot_code_games itself) meant the tree
        builder successfully found these games, but this validator still
        did not recognize the same match and downgraded the whole bracket
        to red as a "foreign game" -- found live 2026-07-23, same
        investigation, right after the tree-builder fix above."""
        parent = _node_by_id.get(node.get("src_game_id"))
        if not parent:
            return False
        s = slot.strip().upper()
        for code in (parent.get("w_to"), parent.get("l_to")):
            if isinstance(code, str):
                c = code.strip().upper()
                if c and (s == c or s.startswith(c + "-") or s.startswith(c + "(")):
                    return True
        return False
    for n in nodes:
        if not n.get("src_game_id"):
            continue  # root exempt
        gid = n["game_id"]
        if (n.get("src_path") or "").startswith("pool_"):
            continue  # pool-finish branch -- tree builder already verified the group match
        if n.get("tbd_stub"):
            continue  # synthetic placeholder -- deliberately names neither team
        direct = (team_matches(n.get("white_team", ""), team)
                  or team_matches(n.get("dark_team", ""), team))
        if not direct:
            ref_found = any(_slot_refs_tree_game(slot)
                             or _slot_refs_parent_advancement_code(n, slot)
                             for slot in (n.get("white_team", ""), n.get("dark_team", "")))
            if not ref_found:
                red.append(
                    f"game {gid}: foreign game — neither slot names {team!r} "
                    f"nor references a tree game via WIN/LOS GM # or W#/L# "
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

    # ── Missing later-day node ─────────────────────────────────────────────────
    # Originally hardcoded to Saturday/Sunday, which silently never fired for
    # Junior Olympics (a Thu-Sun event) -- generalized to "day 1 of the
    # upcoming games this team can see" vs "any later day", so it applies to
    # any tournament's actual day span instead of assuming a 2-day weekend.
    if upcoming:
        upcoming_dates = {g.get("date") for g in upcoming if g.get("date")}
        node_dates    = {n.get("date") for n in nodes   if n.get("date")}
        if upcoming_dates:
            first_day = min(upcoming_dates)
            all_first_day = all(d == first_day for d in upcoming_dates)
            has_later_node = any(d > first_day for d in node_dates)
            if all_first_day and not has_later_node:
                yellow.append(
                    f"all {len(upcoming)} upcoming game(s) are on {first_day} with no "
                    "later-day node in bracket — later-round placement data may be missing"
                )

    # ── Unresolved opponent slot strings (yellow — data quality only) ──────────
    # See _POOL_PREVIEW_RE's own docstring for why this carve-out exists.
    if serialized_nodes:
        for sn in serialized_nodes:
            opp = sn.get("opponent", "")
            if opp and _SLOT_LIKE_RE.match(opp) and not _POOL_PREVIEW_RE.match(opp):
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


# Which tree-building algorithm a bracket-tournament format uses, looked up
# once instead of re-evaluating "tournament_id in WPL_TOURNAMENTS / elif
# tournament_id in _NJO_TOURNAMENTS" at every call site that needs to build
# or reason about a team's bracket tree. Both formats' membership sets
# (WPL_TOURNAMENTS, _NJO_TOURNAMENTS) stay the single source of truth for
# "which tournament_ids belong to this format" -- this only consolidates the
# derived "so which tree builder does that imply" decision that was
# otherwise re-derived at each of several call sites in api_games.
_TREE_BUILDERS = {
    "wpl": _build_wpl_game_tree,
    "njo": _build_njo_game_tree,
}


def _tree_format_for(tournament_id: str) -> str | None:
    """'wpl' | 'njo' | None -- which bracket-tree format this tournament_id
    belongs to, or None if it's not a bracket-tournament format at all."""
    if tournament_id in WPL_TOURNAMENTS:
        return "wpl"
    if tournament_id in _NJO_TOURNAMENTS:
        return "njo"
    return None


def _compute_tree_layout(nodes: list, team: str) -> dict:
    """Assign column (bracket-tree round depth), path_condition, and
    eliminated flag to each node.

    column:         longest path from the node's own segment root (1-indexed)
                    -- see below for why this is longest-path, not plain BFS.
    path_condition: "win" | "lose" | None (neutral/always shown) from parent edge
    eliminated:     True if a played result means the team went the other way

    Returns {game_id: {column, path_condition, eliminated}}.
    """
    if not nodes:
        return {}

    node_map = {n["game_id"]: n for n in nodes}
    roots = [n for n in nodes if not n.get("src_game_id")]
    if not roots:
        return {}

    def _next_ids(n: dict) -> list:
        return ((n.get("win_next_ids") or []) + (n.get("lose_next_ids") or [])
                + list((n.get("pool_next") or {}).keys()))

    # NJO trees can have multiple disconnected segments (a round-robin pool
    # phase with no advancement links at all, plus a later placement/
    # consolation phase that starts a fresh w_to/l_to numbering sequence) --
    # each is its own root. Giving every root column 1 would make two
    # genuinely unrelated segments visually overlap in the same columns,
    # reading as one continuous lineage when they are not -- the exact class
    # of bug this column redesign exists to fix in the first place. Lay
    # segments out left to right in the order they actually happen (earliest
    # real game in the segment first), each starting where the previous
    # segment's deepest column left off.
    def _segment_earliest(root_gid: str):
        seen = {root_gid}
        stack = [root_gid]
        best = (date.max, datetime.max.time())
        while stack:
            gid = stack.pop()
            n = node_map.get(gid)
            if not n:
                continue
            key = (n.get("date") or date.max, n.get("time") or datetime.max.time())
            if key < best:
                best = key
            for nid in _next_ids(n):
                if nid not in seen:
                    seen.add(nid)
                    stack.append(nid)
        return best

    roots_sorted = sorted(roots, key=lambda r: _segment_earliest(r["game_id"]))

    column_map: dict[str, int] = {}
    visited: set = set()
    col_offset = 0
    for root in roots_sorted:
        if root["game_id"] in visited:
            continue

        # Collect this segment's nodes first (plain reachability), then
        # assign columns via topological order using EVERY parent, not just
        # whichever one a traversal happens to reach a node through first.
        # A node can have more than one real incoming edge: _append_tbd_stub_chain
        # attaches ONE shared "continues next day" stub to every dead-end
        # branch at the tree's latest known date, and those branches can be
        # at genuinely different depths (a 3-round-deep dead end and a
        # 4-round-deep dead end can both legitimately land on the
        # tournament's last known day). A node like that must sit at
        # max(all parents' columns) + 1, not "one specific parent's column +
        # 1" -- taking just the first-discovered parent (plain BFS) can
        # place a shared node in the SAME column as one of its own other
        # parents, an impossible-looking (same-column parent-child) edge.
        segment_nodes: list = []
        seen_seg = {root["game_id"]}
        stack = [root["game_id"]]
        while stack:
            gid = stack.pop()
            segment_nodes.append(gid)
            n = node_map.get(gid)
            if not n:
                continue
            for nid in _next_ids(n):
                if nid not in seen_seg:
                    seen_seg.add(nid)
                    stack.append(nid)

        parents_of: dict[str, list] = {gid: [] for gid in segment_nodes}
        for gid in segment_nodes:
            n = node_map.get(gid)
            if not n:
                continue
            for nid in _next_ids(n):
                if nid in parents_of:
                    parents_of[nid].append(gid)

        # Every node gets its OWN column now -- never share one with any
        # other node, even true siblings (multiple children of the same
        # parent) or independent nodes that happen to reach the same
        # topological depth. Product decision 2026-07-24: a bracket-tree
        # column used to mean "round depth" (letting a round-robin pool's
        # multiple simultaneous games, or independent branches at the same
        # depth, share one column) -- found live to read as a confusing
        # fork ("either this game or that one") even when both games were
        # actually guaranteed to happen, back to back. Simpler now: a
        # column is just "the Nth game scenario", strictly one per column,
        # so scrolling right always means "the next game", never "the next
        # round that might contain more than one card". A min-heap
        # (ordered by date/time, earliest first) replaces the plain deque
        # so that whenever more than one node is genuinely ready at once
        # (e.g. two round-robin pool games sharing a parent), they still
        # get assigned in a deterministic, chronological order instead of
        # colliding on the same column.
        col_of: dict[str, int] = {}
        remaining = {gid: len(parents_of[gid]) for gid in segment_nodes}
        def _ready_key(gid):
            n = node_map[gid]
            return (n.get("date") or date.max, n.get("time") or datetime.max.time(), gid)
        ready: list = [(_ready_key(gid), gid) for gid in segment_nodes if remaining[gid] == 0]
        heapq.heapify(ready)
        segment_max = col_offset
        next_col = col_offset
        processed = 0
        while ready:
            _, gid = heapq.heappop(ready)
            next_col += 1
            col_of[gid] = next_col
            segment_max = max(segment_max, next_col)
            processed += 1
            n = node_map.get(gid)
            if not n:
                continue
            for nid in _next_ids(n):
                if nid in remaining:
                    remaining[nid] -= 1
                    if remaining[nid] == 0:
                        heapq.heappush(ready, (_ready_key(nid), nid))
        # Defensive: anything left unprocessed means a cycle slipped past
        # _bracket_has_cycle's own check elsewhere -- place it rather than
        # silently drop it, so a real data anomaly still renders (and gets
        # caught by the confidence validator's own cycle check) instead of
        # vanishing from the tree.
        for gid in segment_nodes:
            if gid not in col_of:
                segment_max += 1
                col_of[gid] = segment_max

        column_map.update(col_of)
        visited.update(segment_nodes)
        col_offset = segment_max

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
        for nid, rank in (n.get("pool_next") or {}).items():
            path_cond[nid] = f"pool_{rank}"

    # eliminated: propagate from played results. A node is only truly
    # eliminated once EVERY parent that can reach it is itself eliminated --
    # a node can have more than one real incoming edge (the same multi-
    # parent-merge shape the column computation above already accounts for:
    # _append_tbd_stub_chain attaches one shared "continues next day" stub
    # to every dead-end branch, so a live path and a dead path can both
    # feed the same downstream placeholder). The previous version just
    # recursed "mark this node dead, then mark everything downstream of it
    # dead too" the instant ANY ONE incoming branch died, with no check for
    # a still-live sibling parent also feeding the same node. Confirmed
    # live 2026-07-23: this made Trojan Gold 16U's real, still-upcoming
    # Round 3+ games render as greyed-out "eliminated" placeholders right
    # after their Round 2 loss, because the dead "if we'd won round 2"
    # branch and the real, live next game both merge into the same shared
    # downstream TBD stub -- the dead branch's propagation reached it first
    # and killed it, even though the actual live game reaches it too.
    eliminated: set = set()
    seeded: set = set()
    for n in nodes:
        if not n.get("played"):
            continue
        won = _team_won(team, n)
        win_ids  = set(n.get("win_next_ids")  or [])
        lose_ids = set(n.get("lose_next_ids") or [])
        neutral  = win_ids & lose_ids
        if won is True:
            for nid in lose_ids - neutral:
                eliminated.add(nid)
                seeded.add(nid)
        elif won is False:
            for nid in win_ids - neutral:
                eliminated.add(nid)
                seeded.add(nid)

    all_parents_of: dict[str, list] = {n["game_id"]: [] for n in nodes}
    for n in nodes:
        for nid in _next_ids(n):
            if nid in all_parents_of:
                all_parents_of[nid].append(n["game_id"])

    # Forward propagation in topological order (Kahn's, over the WHOLE
    # graph this time, not per-segment) -- a node's eliminated status can
    # only be finalized once every one of its parents already has theirs.
    from collections import deque as _dq
    in_deg = {gid: len(all_parents_of[gid]) for gid in all_parents_of}
    topo_q = _dq(gid for gid in all_parents_of if in_deg[gid] == 0)
    seen_topo = set(topo_q)
    while topo_q:
        gid = topo_q.popleft()
        if gid not in seeded and all_parents_of[gid] and all(p in eliminated for p in all_parents_of[gid]):
            eliminated.add(gid)
        nd = node_map.get(gid)
        if not nd:
            continue
        for nid in _next_ids(nd):
            if nid not in in_deg:
                continue
            in_deg[nid] -= 1
            if in_deg[nid] == 0 and nid not in seen_topo:
                seen_topo.add(nid)
                topo_q.append(nid)

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


def _fill_missing_branch_stubs(tree: list, tree_sheet: str) -> None:
    """Add a TBD stub for the win or lose branch of a node whose OTHER
    branch is already a real published game -- a node with exactly one of
    win_next_ids/lose_next_ids populated, not both and not neither.

    Real bug, found live 2026-07-16: a tournament organizer building out a
    bracket incrementally can publish one outcome's continuation game
    before the other (confirmed directly against JO's raw sheet data --
    "L#28" and "W#32" existed, "W#28" and "L#32" did not, anywhere in the
    192-game division). Before this fix, a node in that state silently
    showed only its one known branch, with no card at all for the other
    outcome -- reading as "if you win game 28, nothing happens next"
    instead of "we don't know that game yet." Every win/lose split must
    show two cards or none; never exactly one real card with the other
    branch simply absent.

    Deliberately distinct from _append_tbd_stub_chain, which handles a node
    with BOTH branches missing (the tree runs out entirely for the day) by
    chaining stubs across remaining tournament days. This handles a node
    with ONLY one branch missing -- must run before that function so its
    day-chain frontier detection (which requires BOTH branches empty)
    doesn't also need to special-case a partially-filled node.
    """
    if not tree:
        return
    by_id = {n["game_id"]: n for n in tree}
    new_stubs = []
    for n in list(tree):
        if n.get("tbd_stub"):
            continue
        has_win = bool(n.get("win_next_ids"))
        has_lose = bool(n.get("lose_next_ids"))
        if has_win == has_lose:
            continue  # both present (fine) or both absent (handled elsewhere)
        missing_path = "lose" if has_win else "win"

        # Best estimate for the stub's date: whichever real child already
        # exists tells us when this round is actually happening.
        known_next_ids = (n.get("win_next_ids") or []) + (n.get("lose_next_ids") or [])
        known_child = by_id.get(known_next_ids[0]) if known_next_ids else None
        stub_date = (known_child.get("date") if known_child else None) or n.get("date")

        stub_id = f"__tbd_branch_{tree_sheet}_{n['game_id']}"
        stub = {
            "game_id":        stub_id,
            "date":           stub_date,
            "time":           None,
            "location":       None,
            "white_team":     "TBD",
            "dark_team":      "TBD",
            "played":         False,
            "placeholder":    True,
            "tbd_stub":       True,
            "src_game_id":    n["game_id"],
            "src_path":       None,
            "win_next_ids":   [],
            "lose_next_ids":  [],
            "pool_next":      {},
            "sunday_pair_id": None,
            "tree_format":    "bracket",
        }
        if missing_path == "win":
            n["win_next_ids"] = [stub_id]
        else:
            n["lose_next_ids"] = [stub_id]
        new_stubs.append(stub)
    tree.extend(new_stubs)


def _append_tbd_stub_chain(tree: list, division_games: list, tree_sheet: str, team: str) -> None:
    """When a bracket tree runs out of data before the tournament's own
    posted schedule does (an elimination format where the spreadsheet
    doesn't yet specify what happens beyond a certain round -- see
    _build_njo_game_tree's w_to/l_to and W#/L# fallback, and the pool_next
    fallback, all of which can still legitimately dead-end early), mutate
    `tree` in place to append one TBD placeholder node per remaining
    tournament day -- so parents see a real stub card for every day a game
    could still happen, instead of the bracket silently stopping.

    Deliberately day-level, not an attempt to model exact round count or a
    specific future opponent -- that data genuinely isn't in the
    spreadsheet yet. Reuses the frontend's existing tbd_stub card style
    (already built for fixed-count formats -- see the flat_schedule stub
    logic in api_games) rather than inventing a new one. Chains through
    BOTH win_next_ids and lose_next_ids of every "frontier" node (the
    tree's unresolved dead ends at its latest known date), matching
    _compute_tree_layout's existing "reachable via both win and lose =
    neutral, always shown" rule -- no changes needed there.

    No-op if the tree already reaches the division's last posted date
    (nothing more on the schedule to speak of), if there's no tree, or if
    every node at the latest date already has a real continuation.
    """
    if not tree:
        return
    known_dates = [n.get("date") for n in tree if n.get("date")]
    if not known_dates:
        return
    last_known = max(known_dates)

    future_dates = sorted({g.get("date") for g in division_games
                            if g.get("date") and g.get("date") > last_known})
    if not future_dates:
        return

    frontier = [n for n in tree
                if n.get("date") == last_known
                and not n.get("win_next_ids") and not n.get("lose_next_ids")
                and not n.get("pool_next")]
    if not frontier:
        return

    by_id = {n["game_id"]: n for n in tree}
    # Multiple frontier nodes can exist at the same latest-known date (e.g.
    # one real dead-end plus one loss the team was eliminated by earlier
    # that same day). Picking frontier[0] unconditionally, as this used to,
    # attaches the "what happens next" stub to whichever node happened to
    # land first in `tree`'s own build order -- which is not necessarily
    # the team's actual live path. Confirmed live 2026-07-25: after fixing
    # _game_num/_gnum_str (which changed how much of the tree gets
    # discovered, and therefore this list's order), a Friday LOSS
    # (game 71, eliminated) ended up first instead of the real WIN
    # (game 111, still alive), stranding the correct Saturday continuation
    # behind the wrong parent. Prefer a frontier node the team is still
    # alive in (won it, or it has not been played yet) over one that
    # eliminated them; only fall back to an eliminated node if that is all
    # there is (team fully eliminated, no live path exists).
    live_frontier = [n for n in frontier if _team_won(team, n) is not False]
    prev_ids = [n["game_id"] for n in (live_frontier or frontier)]
    for i, d in enumerate(future_dates):
        stub_id = f"__tbd_{tree_sheet}_{i}"
        stub = {
            "game_id":        stub_id,
            "date":           d,
            "time":           None,
            "location":       None,
            "white_team":     "TBD",
            "dark_team":      "TBD",
            "played":         False,
            "placeholder":    True,
            "tbd_stub":       True,
            "src_game_id":    prev_ids[0],
            "src_path":       None,
            "win_next_ids":   [],
            "lose_next_ids":  [],
            "pool_next":      {},
            "sunday_pair_id": None,
            "tree_format":    "bracket",
        }
        for pid in prev_ids:
            by_id[pid]["win_next_ids"].append(stub_id)
            by_id[pid]["lose_next_ids"].append(stub_id)
        tree.append(stub)
        by_id[stub_id] = stub
        prev_ids = [stub_id]


def _same_division(sheet_a: str, sheet_b: str) -> bool:
    """True if two sheet names represent the same age+gender division, even
    across different tournament files with completely different sheet-naming
    conventions (e.g. Quiksilver Cup's "16U BOYS-18 TEAMS" vs Turbo OC's
    "16U BOYS PLATINUM GOLD-11 TEAMS" -- both parse to ('Boys', '16U')).
    A literal sheet-name string match can never succeed across files, since
    every tournament names its sheets differently -- that would silently
    limit every cross-tournament search to zero results."""
    if not sheet_a or not sheet_b:
        return True  # no scoping info available on one side — don't filter
    return _parse_sheet(sheet_a) == _parse_sheet(sheet_b)


def _last_meeting(team: str, opponent: str, all_games: list,
                  before_date=None, sheet: str = None) -> dict | None:
    """Return the most recent played game between team and opponent.

    Searches all_games for a played game where both teams appear.
    sheet: if provided, restricts search to games in the same age+gender
    division (see _same_division) so a 16u Boys lookup doesn't surface 12u
    Boys results for the same team name.
    before_date excludes games on or after that date so the current game
    isn't counted as its own last meeting.
    """
    if not opponent or _SLOT_LIKE_RE.match(opponent):
        return None  # opponent is TBD / unresolved slot — nothing to look up
    best = None
    for g in all_games:
        if not g.get("played"):
            continue
        if sheet and not _same_division(sheet, g.get("sheet")):
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


def _head_to_head(team: str, opponent: str, all_games: list, sheet: str = None) -> dict | None:
    """Tally the full W/L/T record between team and opponent across every
    played meeting found (including the current game, if it's part of
    all_games) — same matching rules as _last_meeting (which only returns the
    single most recent one), aggregated across all local Excel files via
    _all_historical_games() by the caller."""
    if not opponent or _SLOT_LIKE_RE.match(opponent):
        return None
    wins = losses = ties = 0
    for g in all_games:
        if not g.get("played"):
            continue
        if sheet and not _same_division(sheet, g.get("sheet")):
            continue
        wt = strip_prefix(g["white_team"]).strip()
        dt = strip_prefix(g["dark_team"]).strip()
        if not (team_matches(wt, team) or team_matches(dt, team)):
            continue
        if not (team_matches(wt, opponent) or team_matches(dt, opponent)):
            continue
        result = _result_str(g, team)
        if result == "win":
            wins += 1
        elif result == "loss":
            losses += 1
        else:
            ties += 1
    if wins + losses + ties == 0:
        return None
    return {"wins": wins, "losses": losses, "ties": ties}


def _opponent_history(team: str, opponent_label: str, h2h_games: list,
                       before_date=None, sheet: str | None = None) -> dict:
    """Single call site for a game's last_meeting + head-to-head tally.

    The flat played/upcoming list and the bracket tree serializer in
    api_games both need this; route both through this one function instead
    of each calling _last_meeting/_head_to_head independently. The
    flat-vs-tree last_meeting bug shipped TWICE (Quiksilver Cup, then Junior
    Olympics -- the identical bug, because the first fix only touched the
    flat list's call site) because each call site had to independently
    remember to pass the combined h2h_games (this tournament's own live data
    + local archives) instead of the archives alone. One call site left to
    get that right, instead of N.
    """
    return {
        "last_meeting": _last_meeting(team, opponent_label, h2h_games,
                                       before_date=before_date, sheet=sheet),
        "h2h": _head_to_head(team, opponent_label, h2h_games, sheet=sheet),
    }


@app.route("/api/games/<tournament_id>/<path:team>")
def api_games(tournament_id, team):
    excel = find_excel(tournament_id)
    if not excel:
        abort(404)

    _all_games = load_and_parse(excel)
    games      = _filter_by_dates(_all_games, tournament_id)

    # See _h2h_source()'s docstring -- this tournament's own live games must
    # win over the local archive on conflict, so they go first.
    _h2h_games = _h2h_source(_all_games, _all_historical_games())

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
    if tournament_id in _NJO_TOURNAMENTS:
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

    # ── Game numbering: group by decision point, then sort chronologically ──
    # Two games share a game_num ("Game N") iff they are alternative outcomes
    # of the same upstream decision — a win/lose bracket pair, or a multi-way
    # pool-finish branch (1st/2nd/3rd in pool). Rather than inferring that
    # after the fact from column adjacency (the old 7-pass approach — BFS
    # depth, chronological fallback, chronological remap, sibling-merge,
    # parallel-path merge, day-merge, self-heal split), compute the
    # decision-point key directly from each game's own slot text: a slot like
    # "W#12"/"L#12"/"WIN GM #12" names the game it descends from, and a slot
    # like "1stB-" names the pool group it descends from. Games with no such
    # pattern (real, already-resolved games) are singleton groups keyed by
    # their own game_id. Sort all groups by their earliest (date, time) and
    # number them 1..N — one pass, no merge/split heuristics needed.
    _known_gms = {_game_num(g["game_id"]) for g in my_games
                  if _game_num(g["game_id"]) and not re.search(r'-[A-Z]+$', g["game_id"])}
    _gm_to_gid = {_game_num(g["game_id"]): g["game_id"] for g in my_games
                  if _game_num(g["game_id"]) and not re.search(r'-[A-Z]+$', g["game_id"])}

    def _slot_ref(slot: str):
        """('gm', ref_num) | ('pool', group_letter) | None for a bracket slot."""
        s = slot.strip()
        m = _GM_WINLOS_RE.search(s)
        if m:
            return ('gm', str(int(m.group(2))))
        m = _WL_SLOT_RE.match(s)
        if m:
            ref = re.search(r"(\d+)$", m.group(1))
            if ref:
                return ('gm', str(int(ref.group(1))))
        m = re.search(r'\bL(\d+)\s*$', s, re.IGNORECASE)
        if m:
            return ('gm', str(int(m.group(1))))
        m = re.search(r'\bWinner\s+G?(\d+)', s, re.IGNORECASE)
        if m:
            return ('gm', str(int(m.group(1))))
        m = re.search(r'\bLoser\s+G(\d+)', s, re.IGNORECASE)
        if m:
            return ('gm', str(int(m.group(1))))
        fm = _FINISH_SLOT_RE.match(s) or _COMPOSITE_SLOT_RE.search(s)
        if fm:
            return ('pool', fm.group(1).upper())
        return None

    def _decision_ref(g):
        """('gm', ref_num) | ('pool', group_letter) | None -- what g's own
        slot says it descends from. Shared by _decision_key (the grouping
        key) and _predecessor_game_id (the actual predecessor game object),
        so both agree on what "descends from" means."""
        # _expand_bracket_games already tagged pool-finish extras with the
        # team's own group letter (resolved from the team's actual pool seed,
        # scoped to the right weekend) -- trust that over re-parsing slot
        # text, since an unresolved composite slot's OTHER side (the
        # opponent's) can carry an unrelated group letter and there's no way
        # to tell them apart from raw text alone (see 14Bag14's white slot
        # "J1(1stA)-" vs. dark slot "J4(1stH)", only the latter is ours).
        if g.get("pool_rank_group"):
            return ('pool', g['pool_rank_group'])
        own_num = _game_num(g["game_id"])
        white_m = team_matches(g["white_team"], team)
        dark_m  = team_matches(g["dark_team"], team)
        # If exactly one slot is already resolved to our own name, only that
        # slot's pattern can be our decision point -- the other slot (if it
        # also has a pattern) belongs to the opponent's unrelated lineage.
        if white_m and not dark_m:
            slots = (g["white_team"],)
        elif dark_m and not white_m:
            slots = (g["dark_team"],)
        else:
            slots = (g["white_team"], g["dark_team"])
        refs = [r for r in (_slot_ref(s) for s in slots) if r]
        # Prefer a 'gm' ref that resolves to a game already in our own game
        # list — the other slot (if any) belongs to the opponent's unrelated
        # bracket lineage and isn't a decision point of ours.
        for kind, val in refs:
            if kind == 'gm' and val in _known_gms and val != own_num:
                return ('gm', val)
        for kind, val in refs:
            if kind == 'pool':
                return ('pool', val)
        for kind, val in refs:
            if kind == 'gm':
                return ('gm', val)
        return None

    def _decision_key(g):
        ref = _decision_ref(g)
        return f"{ref[0]}:{g['sheet']}:{ref[1]}" if ref else None

    def _predecessor_game_id(g):
        """The actual game_id g's decision point resolves to, or None. Used
        to detect when two DIFFERENT decision keys (e.g. references to two
        different games) are nonetheless the same round for this team --
        see the cascading-merge pass below."""
        ref = _decision_ref(g)
        if not ref:
            return None
        kind, val = ref
        if kind == 'gm':
            return _gm_to_gid.get(val)
        # pool: anchor to the team's own last direct pool-phase game in this group.
        candidates = []
        for m in my_games:
            for slot in (m["white_team"], m["dark_team"]):
                pm = _POOL_SLOT_RE.match(slot.strip())
                if pm and pm.group(1).upper() == val and team_matches(slot, team):
                    candidates.append(m)
                    break
        if not candidates:
            return None
        candidates.sort(key=lambda m: (m.get("date") or date.min, m.get("time") or datetime.min.time()))
        return candidates[-1]["game_id"]

    _groups: dict[str, list] = {}
    for g in my_games:
        dk = _decision_key(g) or f"self:{g['game_id']}"
        _groups.setdefault(dk, []).append(g)

    # A shared decision key only means "mutually exclusive alternatives" when
    # at most one member has already become real — a decision resolves to
    # exactly one outcome, so two or more non-placeholder games can never
    # legitimately share one decision point. Some tournaments reuse the same
    # bracket-seed label (e.g. "2ndB", or "WIN GM #399") across multiple
    # SEQUENTIAL real rounds for a team that has already claimed that seed —
    # those games all resolve to the same decision_key but are not
    # alternatives of each other. Pull every such real game out into its own
    # singleton (ordered purely chronologically); any remaining still-
    # uncertain placeholders stay merged as genuine alternatives.
    for dk in list(_groups.keys()):
        members = _groups[dk]
        real = [m for m in members if not m.get("placeholder")]
        if len(real) > 1:
            placeholders = [m for m in members if m.get("placeholder")]
            del _groups[dk]
            for m in real:
                _groups[f"self:{m['game_id']}"] = [m]
            if placeholders:
                _groups[dk] = placeholders

    # "Game N" means the Nth game this team plays, not "descends from this
    # exact predecessor game" -- so two groups whose predecessors are
    # themselves already in the SAME group (e.g. win-then-lose vs
    # lose-then-win in a multi-round bracket: 028 and 032 already share one
    # group as game 2, so whatever comes after either of them is equally
    # "game 3", regardless of which specific one it descends from) must be
    # merged into one group too. Repeat until no more merges happen -- a
    # merge one level can enable another merge one level up. Bounded by the
    # group count, so this always terminates.
    _game_to_dk: dict[str, str] = {}
    for dk, members in _groups.items():
        for m in members:
            _game_to_dk[m["game_id"]] = dk

    _changed = True
    _iterations = 0
    while _changed and _iterations < len(_groups) + 5:
        _changed = False
        _iterations += 1
        _pred_group: dict[str, str] = {}
        for dk, members in _groups.items():
            for m in members:
                pred_gid = _predecessor_game_id(m)
                if pred_gid and pred_gid in _game_to_dk:
                    _pred_group[dk] = _game_to_dk[pred_gid]
                    break
        _by_pred: dict[str, list] = {}
        for dk, pg in _pred_group.items():
            _by_pred.setdefault(pg, []).append(dk)
        for pg, dks in _by_pred.items():
            if len(dks) <= 1:
                continue
            merged_members = [m for _dk in dks for m in _groups[_dk]]
            real = [m for m in merged_members if not m.get("placeholder")]
            if len(real) > 1:
                continue  # same reused-label guard as above -- don't merge
            target = dks[0]
            for _dk in dks[1:]:
                _groups[target].extend(_groups[_dk])
                del _groups[_dk]
                _changed = True
            for m in _groups[target]:
                _game_to_dk[m["game_id"]] = target

    def _group_earliest(dk):
        members = _groups[dk]
        return min((m.get("date") or date.min, m.get("time") or datetime.min.time())
                   for m in members)

    _sorted_keys = sorted(_groups.keys(), key=lambda dk: (_group_earliest(dk), dk))
    _game_num_map: dict[str, int] = {}
    for _i, _dk in enumerate(_sorted_keys, start=1):
        for _m in _groups[_dk]:
            _game_num_map[_m["game_id"]] = _i

    # Defensive invariant check: game_num should always track the chronological
    # order of each group's earliest occurrence -- guaranteed by construction
    # above, but a trip-wire here catches it immediately (in server logs) if
    # any future change to this function silently breaks that guarantee,
    # rather than a scrambled-looking bracket reaching a parent's phone
    # unnoticed. Same spirit as the "does the LLM give this a common-sense
    # check" ask -- deterministic and free instead of a model call.
    _gn_earliest: dict[int, tuple] = {}
    for _g in my_games:
        _gn = _game_num_map.get(_g["game_id"])
        if _gn is None:
            continue
        _key = (_g.get("date") or date.min, _g.get("time") or datetime.min.time())
        if _gn not in _gn_earliest or _key < _gn_earliest[_gn]:
            _gn_earliest[_gn] = _key
    _gn_sorted = sorted(_gn_earliest)
    for _a, _b in zip(_gn_sorted, _gn_sorted[1:]):
        if _gn_earliest[_a] > _gn_earliest[_b]:
            print(f"[game-num-order] {team!r} | {tournament_id}: GAME {_a} starts "
                  f"{_gn_earliest[_a]} but GAME {_b} starts earlier at {_gn_earliest[_b]} "
                  "— game numbers are out of chronological order", flush=True)

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
            # Never used by the frontend directly -- only carried through so
            # _infer_placement/_estimate_placement (called below on played_out)
            # can read the organizer's ordinal-rank comment on the last played
            # game. Without it, both silently always returned None: found live
            # 2026-07-19 via a round-by-round replay of real 2025 JO results,
            # where the app's own "placement" never matched the real recorded
            # final rank for any team, in any division, ever -- this had never
            # been exercised by a test that checks the VALUE, only that
            # something renders, so the "Finished Nth!" summary banner has
            # been silently falling back to generic "Tournament Record" for
            # every completed tournament.
            "comments":    g.get("comments"),
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
            # last_meeting excludes this game itself (points at the previous
            # meeting, if any -- pointless to reference this same card).
            # h2h INCLUDES this game -- it's a cumulative record, not a
            # pointer, so it should always reflect the full known matchup
            # history through and including this result.
            _hist = _opponent_history(team, opponent_label, _h2h_games,
                                       before_date=g.get("date"), sheet=g.get("sheet"))
            base["last_meeting"] = _hist["last_meeting"]
            base["h2h"] = _hist["h2h"]
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
                # Guard added here to match the tree serializer's identical
                # check (found auditing for other divergent-path bugs after
                # fixing game_num/last_meeting/opponent resolution this same
                # session): a live score can attach to a game_id that's
                # still a placeholder in OUR data (the organizer's sheet
                # cell hasn't been updated with a real opponent name yet,
                # even though the game is factually being played right
                # now) -- CURRENT GAME must never show on a node whose
                # opponent is still an unresolved guess, only the tree had
                # this guard until now.
                if not base.get("placeholder"):
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
            _hist = _opponent_history(team, opponent_label, _h2h_games,
                                       before_date=g.get("date"), sheet=g.get("sheet"))
            base["last_meeting"] = _hist["last_meeting"]
            # For an upcoming game with a real (non-TBD) opponent, explicitly
            # show "0-0" rather than hiding the section when there's no prior
            # history — omitting it here (unlike the played-game case) reads
            # as a missing feature rather than a deliberate "first meeting"
            # signal. A still-unresolved opponent slot has nothing to default.
            base["h2h"] = _hist["h2h"]
            if base["h2h"] is None and opponent_label and not _SLOT_LIKE_RE.match(opponent_label):
                base["h2h"] = {"wins": 0, "losses": 0, "ties": 0}
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
    tree = None
    _tree_format = _tree_format_for(tournament_id)
    if _tree_format and my_games:
        tree_sheet = my_games[0]['sheet']
        div_games_for_tree = [g for g in _all_games if g['sheet'] == tree_sheet]
        latest_team_date = max((g["date"] for g in my_games if g.get("date")), default=None)
        tree = _TREE_BUILDERS[_tree_format](team, div_games_for_tree, anchor_date=latest_team_date)
        if _tree_format == "njo":
            # NJO uses w_to/l_to integer links, not WPL-style WIN GM # slots —
            # skip ground-truth depth checks (_derive_expected_bracket is WPL-specific).
            struct_issues = _check_bracket_structure(team, tree, allow_multiple_roots=True)
            if struct_issues:
                for _si in struct_issues:
                    print(f"[bracket-struct] {team!r} | {tournament_id}: {_si}", flush=True)

    if tree:
        _fill_missing_branch_stubs(tree, tree_sheet)
        _append_tbd_stub_chain(tree, div_games_for_tree, tree_sheet, team)
        # Serialize tree nodes: format dates/times, add opponent label
        def _serialize_tree_node(node, dg):
            gid = node["game_id"]
            if node.get("tbd_stub"):
                # Both slots are the literal string "TBD" by construction
                # (_append_tbd_stub_chain / _fill_missing_branch_stubs) --
                # short-circuit rather than asking _team_opp_slot to resolve
                # it, which would log a spurious "unresolved slot" warning
                # for something that's supposed to be unresolved.
                opp_sl, color = "TBD", "DARK"
            else:
                # _team_opp_slot: the same shared function the flat
                # played/upcoming list already uses, instead of a second,
                # independent inline algorithm. Verified these two
                # previously-parallel algorithms already agreed on every
                # real node across all 7 known JO teams (including deep
                # round-4+ chains) before making this the single
                # implementation -- see test_bracket_tree_opponent_resolution_parity.
                opp_sl = _team_opp_slot(node, team, dg, my_game_ids)
                color  = "WHITE" if opp_sl == node["dark_team"] else "DARK"
            opp_name = describe_slot(opp_sl, dg, ref_date=latest_team_date)
            # Guard: if resolved opponent still equals our own team, flip slots
            if team_matches(opp_name, team):
                other_sl = node["white_team"] if opp_sl == node["dark_team"] else node["dark_team"]
                opp_name = describe_slot(other_sl, dg, ref_date=latest_team_date)
            # our_record/opp_record: same sheet_records lookup the flat list
            # already uses (computed once, above, for the whole team) --
            # this-tournament win/loss record for each side. Was never wired
            # into the tree serializer at all (a missing feature, not a
            # diverging duplicate -- found because Quiksilver, a flat-list
            # tournament, always showed this, and Junior Olympics, a
            # bracket-tree tournament, never did).
            our_rec = opp_rec = None
            if show_records and not node.get("tbd_stub"):
                trec    = sheet_records.get(tree_sheet, {})
                our_key = strip_prefix(node["white_team"] if color == "WHITE" else node["dark_team"]).upper()
                opp_key = strip_prefix(opp_sl).upper()
                our_rec = trec.get(our_key)
                opp_rec = trec.get(opp_key)
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
                "tbd_stub":       node.get("tbd_stub", False),
                "placement_rank": node.get("placement_rank"),
                "src_game_id":    node["src_game_id"],
                "src_path":       node["src_path"],
                "win_next_ids":   node["win_next_ids"],
                "lose_next_ids":  node["lose_next_ids"],
                "sunday_pair_id": node["sunday_pair_id"],
                "is_current":     is_current,
                "our_record":     our_rec,
                "opp_record":     opp_rec,
            }
            if node.get("played") and not node.get("placeholder"):
                ws = node.get("white_score") or 0
                ds = node.get("dark_score")  or 0
                d["score"]     = _fmt_score(node)
                d["our_score"] = ws if color == "WHITE" else ds
                d["opp_score"] = ds if color == "WHITE" else ws
                d["result"]    = _result_str(node, team)
            if not node.get("tbd_stub"):
                # _opponent_history (not a direct _last_meeting/_head_to_head
                # call) -- see its docstring. This exact gap (using
                # _all_historical_games() alone instead of the tournament's
                # own live data combined in) shipped as a real bug twice at
                # this call site before being routed through the one shared
                # function.
                #
                # sheet=tree_sheet: found while wiring this up -- the ORIGINAL
                # code here never passed a sheet at all, so this call was
                # unscoped across every division while the flat list's
                # equivalent call was correctly scoped to the team's own
                # division. That let a same-named opponent in a DIFFERENT age
                # group's history surface here (confirmed live: 12U Trojan
                # Cardinal vs ASPHALT GREEN showed a match in the bracket view
                # that the flat list correctly omitted). Same class of bug
                # this function exists to prevent, just a second instance of
                # it, found as a direct result of consolidating the call.
                #
                # h2h was never computed here at all until now (a missing
                # feature, not a diverging duplicate -- Quiksilver, a
                # flat-list tournament, always showed this; Junior Olympics,
                # a bracket-tree tournament, never did). Same call already
                # made for last_meeting returns both -- just wasn't reading
                # the second half of it before.
                _hist = _opponent_history(team, opp_name, _h2h_games,
                                           before_date=node.get("date"), sheet=tree_sheet)
                d["last_meeting"] = _hist["last_meeting"]
                d["h2h"] = _hist["h2h"]
                if not node.get("played"):
                    # Same explicit-0-0-instead-of-hidden rule as the flat
                    # list: omitting h2h on an upcoming game with a real,
                    # resolved opponent reads as a missing feature rather
                    # than a deliberate "first meeting" signal.
                    if d["h2h"] is None and opp_name and not _SLOT_LIKE_RE.match(opp_name):
                        d["h2h"] = {"wins": 0, "losses": 0, "ties": 0}
            if live:
                d["live_score"] = live
            return d
        wpl_bracket = [_serialize_tree_node(n, div_games_for_tree) for n in tree]

        # Annotate serialized nodes with tree layout: column (round number),
        # path_condition, eliminated. This is the tree's OWN column number --
        # pure BFS depth from root, scoped per disconnected segment so
        # genuinely independent root games or bracket phases never collide
        # in the same column (see _compute_tree_layout) -- deliberately
        # decoupled from _game_num_map, the flat schedule's chronological
        # "Game N" numbering used for played_out/upcoming_out above.
        #
        # Those two numbering schemes used to be unified (this code used to
        # borrow _game_num_map here directly) because chronological order
        # and tree depth agree for a shallow tree, and unifying them fixed a
        # real past bug where the tree had its own separate, untested
        # numbering system that silently merged two different real games
        # under one label. But chronological order and tree depth are
        # fundamentally different sorts once a bracket is wide enough for
        # multiple real branches to be visible at once (this format's
        # 48-team single-elimination-plus-full-placement-ladder shape is
        # exactly that): the tournament schedules different branches'
        # rounds interleaved across day/time slots to fit everyone in, so a
        # game four rounds deep on one branch can easily get a LOWER
        # chronological number than a game two rounds deep on another.
        # Reusing that chronological number for the tree's column then
        # silently violates the tree layout's core assumption that a card's
        # real parent sits in the immediately preceding column -- confirmed
        # live 2026-07-20: Trojan Cardinal 18U's real bracket tree drew a
        # connector line between two completely unrelated branches that
        # happened to land in adjacent chronologically-numbered columns,
        # reading as one long chain of losses when neither game was actually
        # connected to the other. The tree's own "Round N" label (see the
        # frontend header) is intentionally a different number than the
        # schedule's "Game N" for the same reason -- each card still shows
        # its own real game-id badge, which is what actually cross-
        # references between the two views.
        layout = _compute_tree_layout(tree, team)
        for sn in wpl_bracket:
            info = layout.get(sn["game_id"], {})
            sn["column"]         = info.get("column", 1)
            sn["path_condition"] = info.get("path_condition")
            sn["eliminated"]     = info.get("eliminated", False)
            # Aliases kept for validate_tournament.py backwards compatibility
            # (reads len(wpl_bracket) only, not these field values) and the
            # frontend's date-less TBD-stub fallback label.
            sn["game_num"]       = sn["column"]
            sn["path"]           = sn["path_condition"]
            sn["is_alternative"] = False

    # ── Bracket confidence + display mode ────────────────────────────────────
    # WPL tournaments always attempt a bracket. Non-WPL/NJO: no bracket expected.
    is_bracket_tournament = _tree_format is not None

    if wpl_bracket:
        # Pass raw (pre-serialization) upcoming games, not upcoming_out --
        # upcoming_out's "date" field is already a display string
        # (_fmt_date), while `tree`'s nodes carry real date objects.
        # _validate_wpl_bracket compares dates across both; mixing formatted
        # strings with date objects crashes that comparison.
        _raw_upcoming = [g for g in my_games if not g.get("played")]
        bracket_confidence, bracket_warnings = _validate_wpl_bracket(
            team, tree, upcoming=_raw_upcoming, serialized_nodes=wpl_bracket,
            is_njo=(_tree_format == "njo"))
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

    # Additive-only "games remaining" placeholders: for fixed-game-count formats
    # (round-robin + crossover, e.g. Quiksilver Cup) where every team plays the
    # same total number of games, show generic TBD stubs for games not yet
    # reachable via slot resolution — instead of either showing nothing, or a
    # single misleadingly-specific guess (see _expand_bracket_games' "show only
    # one composite slot per group" anti-clutter rule, which picks whichever
    # candidate happens to appear first in the sheet when the team's actual
    # pool rank isn't known yet). Never modifies existing games or resolution
    # logic — purely appends stub entries to fill the known gap in count.
    #
    # Broken down by calendar date (not just a flat total) because every row
    # in this format already carries a real date/time/location even before an
    # opponent is resolvable — so a stub can say WHICH DAY it falls on (e.g.
    # "Saturday, Jul 11") without knowing who or where, which is what parents
    # actually need to plan around.
    if display_mode == "flat_schedule" and my_games:
        _stub_sheet = my_games[0].get("sheet", "")
        _expected_by_day = _expected_games_per_team_by_day(div_map.get(_stub_sheet, []))
        if _expected_by_day:
            _shown_ids = {g["game_id"] for g in played_out + upcoming_out if not g.get("tbd_stub")}
            _known_by_day: dict = {}
            for g in my_games:
                if g["game_id"] in _shown_ids and g.get("date"):
                    _known_by_day[g["date"]] = _known_by_day.get(g["date"], 0) + 1
            _next_gn = max([g["game_num"] for g in upcoming_out if g.get("game_num")]
                            + [g["game_num"] for g in played_out if g.get("game_num")]
                            + [0]) + 1
            _stub_n = 0
            for _d in sorted(_expected_by_day):
                _shortfall = _expected_by_day[_d] - _known_by_day.get(_d, 0)
                for _ in range(max(0, _shortfall)):
                    _stub_n += 1
                    upcoming_out.append({
                        "game_id":        f"__tbd_{_stub_sheet}_{_stub_n}",
                        "date":           _fmt_date(_d),
                        "time":           "TBD",
                        "location":       "TBD",
                        "opponent":       "TBD",
                        "your_color":     None,
                        "game_num":       _next_gn,
                        "is_alternative": False,
                        "path":           None,
                        "placeholder":    True,
                        "tbd_stub":       True,
                        "our_record":     None,
                        "opp_record":     None,
                    })
                    _next_gn += 1

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
    _notify_feedback_email(text)
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
    # Allow pre-game sweep to re-run for the updated file.
    _swept_tournament_ids.discard(tournament_id)
    # Fire immediate smoke test so any parse or team-count errors surface within seconds.
    threading.Thread(target=_run_trojan_smoke_test, args=(tournament_id,), daemon=True).start()
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
    return jsonify([{"game_id": g["game_id"], "date": str(g["date"]), "white": g["white_team"], "dark": g["dark_team"], "ws": g.get("white_score"), "ds": g.get("dark_score"), "w_to": g.get("w_to"), "l_to": g.get("l_to")} for g in games])


@app.route("/api/debug-slots/<tournament_id>/<path:team>")
def api_debug_slots(tournament_id, team):
    """Show raw slot strings for all division games — diagnoses bracket expansion."""
    excel = find_excel(tournament_id)
    if not excel:
        abort(404)
    all_games = load_and_parse(excel)
    games = _filter_by_dates(all_games, tournament_id)
    my_games = [g for g in games if team_matches(g["white_team"], team) or team_matches(g["dark_team"], team)]
    if not my_games:
        return jsonify({"error": "team not found", "sample_teams": list({strip_prefix(g["white_team"]) for g in games[:20]})})
    sheet = my_games[0]["sheet"]
    div_games = [g for g in games if g["sheet"] == sheet]
    return jsonify({
        "team": team, "sheet": sheet,
        "my_game_ids": [g["game_id"] for g in my_games],
        "all_slots": [{"game_id": g["game_id"], "white": g["white_team"], "dark": g["dark_team"],
                       "played": g.get("played"), "w_to": g.get("w_to"), "l_to": g.get("l_to")} for g in div_games],
    })


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
            _mon_tree_format = _tree_format_for(tournament_id) or "wpl"
            tree = _TREE_BUILDERS[_mon_tree_format](team, sheet_games, anchor_date=anchor)

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
            # Sunday check only applies to WPL tournaments (NJO/CCA may be 2-day)
            if tournament_id in WPL_TOURNAMENTS and not sun_ids:
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
                if _SLOT_LIKE_RE.match(opp) and not _POOL_PREVIEW_RE.match(opp):
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
        body = "\n".join(
            f"{x['sheet']}/{x['team']}: {'; '.join(x['issues'][:2])}"
            for x in all_issues[:5]
        )
        _send_ntfy(tournament_id, f"Pre-game check — {tournament_id} issues", body)
    else:
        print(f"[pre-game-check] {tournament_id}: all {n_ok} teams OK ✓")

    return {"tournament": tournament_id, "ok": n_ok, "issues": all_issues}


def _run_trojan_smoke_test(tournament_id: str) -> dict:
    """Immediate smoke test that fires the moment a new URL is loaded.

    Checks every Trojan team across every sheet in the parsed data: game count,
    basic expansion, parse errors. Sends an ntfy push regardless of pass/fail
    so you know immediately whether the new file is usable.
    Unlike _run_pre_game_sweep (which runs 12-25h before game day for all teams),
    this is Trojan-only and fires on demand.
    """
    print(f"[smoke-test] {tournament_id}: starting", flush=True)
    excel = find_excel(tournament_id)
    if not excel:
        msg = "No excel file found after URL save"
        print(f"[smoke-test] {tournament_id}: {msg}", flush=True)
        _send_ntfy(tournament_id, f"Smoke test FAILED — {tournament_id}", msg)
        return {"error": msg}
    try:
        all_games = load_and_parse(excel)
    except Exception as e:
        msg = f"Parse error: {e}"
        print(f"[smoke-test] {tournament_id}: {msg}", flush=True)
        _send_ntfy(tournament_id, f"Smoke test FAILED — {tournament_id}", msg)
        return {"error": msg}

    # Discover all Trojan teams across all sheets
    seen: dict[tuple, tuple] = {}
    for g in all_games:
        for slot in (g["white_team"], g["dark_team"]):
            if is_trojan(slot):
                name = strip_prefix(slot).strip().title()
                key = (name.upper(), g["sheet"])
                seen.setdefault(key, (name, g["sheet"]))

    if not seen:
        msg = f"No Trojan teams found in {len(all_games)} games"
        print(f"[smoke-test] {tournament_id}: {msg}", flush=True)
        _send_ntfy(tournament_id, f"Smoke test WARNING — {tournament_id}", msg)
        return {"error": msg}

    issues: list[str] = []
    ok_teams: list[str] = []

    for (_, sheet), (name, sheet) in sorted(seen.items()):
        sheet_games = [g for g in all_games if g["sheet"] == sheet]
        direct = [g for g in sheet_games
                  if team_matches(g["white_team"], name) or team_matches(g["dark_team"], name)]
        if not direct:
            issues.append(f"{sheet}/{name}: 0 direct games")
            continue
        extras = _expand_bracket_games(name, direct, sheet_games)
        total = len(direct) + len(extras)
        if total < 2:
            issues.append(f"{sheet}/{name}: only {total} game(s) — expansion may have failed")
        elif total > 20:
            issues.append(f"{sheet}/{name}: {total} games — possible over-expansion")
        else:
            ok_teams.append(f"{sheet}/{name}: {total} games")
            print(f"[smoke-test] ✓ {sheet}/{name}: {total} games", flush=True)

    if issues:
        for iss in issues:
            print(f"[smoke-test] ✗ {iss}", flush=True)
        _send_ntfy(
            tournament_id,
            f"Smoke test — {tournament_id} has issues",
            "\n".join(issues[:5]),
        )
    else:
        summary = f"{len(ok_teams)} Trojan teams OK"
        print(f"[smoke-test] {tournament_id}: {summary} ✓", flush=True)
        _send_ntfy(tournament_id, f"Smoke test — {tournament_id} ✓", "\n".join(ok_teams[:8]))

    return {"ok": len(ok_teams), "issues": issues}


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


def _prewarm_url_caches():
    """Fetch + parse every URL-backed tournament once as soon as the server
    boots, instead of leaving the cold-start cost for whichever visitor
    happens to arrive first. Seen live: the JO sheet alone can take 30+s to
    fetch, which briefly showed zero teams for a real visitor before this
    existed. Sequential and best-effort -- a failure here just means the
    first real request pays the normal cold-start cost, same as before."""
    for tid in sorted(PRESET_URL_TOURNAMENTS):
        try:
            excel = find_excel(tid)
            if excel:
                load_and_parse(excel)
                print(f"[prewarm] {tid}: ready", flush=True)
        except Exception as exc:
            print(f"[prewarm] {tid} failed: {exc}", flush=True)

threading.Thread(target=_prewarm_url_caches, daemon=True).start()


def _keep_live_tournaments_warm():
    """Periodically re-fetch each CURRENTLY LIVE tournament's URL in the
    background, so its cache never goes stale purely from a quiet traffic
    lull. _fetch_url's stale-while-revalidate design only refreshes an
    expired cache entry when a real request happens to land after it
    expires -- during a low-traffic window (overnight, early morning) that
    can simply not happen for hours, leaving whoever opens the app next
    stuck looking at a very stale schedule until their own request
    kicks off a refresh. Confirmed live 2026-07-24: Junior Olympics sat
    9+ hours stale overnight (last successful fetch ~9:44 PM, nobody's
    request touched it again until a parent opened the app at 6:39 AM),
    which _prewarm_url_caches alone cannot fix since that only ever runs
    once, at server boot.

    Scoped to tournaments currently within their own date_start/date_end
    window, not every preset tournament forever -- a finished
    tournament's sheet will never change again, so refetching it on a
    schedule would only add unnecessary load and rate-limit risk (the
    same live Google Sheets endpoint got rate-limited earlier today from
    unrelated heavy manual testing) for zero benefit."""
    import time as _time
    _time.sleep(120)  # let the one-time prewarm above finish first
    while True:
        try:
            today = datetime.now(ZoneInfo('America/Los_Angeles')).date()
            for tid in sorted(PRESET_URL_TOURNAMENTS):
                meta = _tournament_meta(tid)
                if not meta or "date_start" not in meta:
                    continue
                if not (meta["date_start"] <= today <= meta["date_end"]):
                    continue
                try:
                    excel = find_excel(tid)
                    if excel:
                        load_and_parse(excel)
                        print(f"[keep-warm] {tid}: refreshed", flush=True)
                except Exception as exc:
                    print(f"[keep-warm] {tid} failed: {exc}", flush=True)
        except Exception as exc:
            print(f"[keep-warm] loop error: {exc}", flush=True)
        # Comfortably inside URL_CACHE_TTL (300s) so a tournament's cache
        # never actually reaches "expired" in the first place.
        _time.sleep(240)

threading.Thread(target=_keep_live_tournaments_warm, daemon=True).start()


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    # threaded=True matches production (railway.json runs gunicorn with
    # --worker-class gthread --threads 8) -- without it, this dev server
    # handles one request at a time, so a single slow /api/tournaments call
    # (live, non-past tournaments re-check find_excel/load_and_parse on
    # every request, no cache) serializes behind every other concurrent
    # request instead of just delaying its own caller. Confirmed live
    # 2026-07-23: this made the CI visual-check tests (tests/*.py, which
    # launch this same "python3 app.py" dev server) intermittently hang on
    # a later navigation even after the initial cold-start wait passed.
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)

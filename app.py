"""
Tournament Translator — Flask app
"""
import os, re, json, glob, io, time, base64
from datetime import datetime, date
from functools import lru_cache
import requests
from flask import Flask, render_template, jsonify, request, abort

from parsers.detect import load_and_parse

app = Flask(__name__)

EXCEL_DIR   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tournaments Excels")
RESULTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "results")
ADMIN_PW    = os.environ.get("ADMIN_PASSWORD", "trojan")  # override via Railway env var

# ── Live URL sources ────────────────────────────────────────────────────────
# Futures/WPL → Google Sheets (organizer keeps this updated throughout the season)
FUTURES_SHEETS_ID  = "1AkX3vwOU9CIc3cymacG2F-uXz-_Gi_A8yR40dEbDpMQ"
FUTURES_SHEETS_URL = f"https://docs.google.com/spreadsheets/d/{FUTURES_SHEETS_ID}/export?format=xlsx"
WPL_TOURNAMENTS    = {"futures-2", "futures-3", "futures-4", "futures-5", "futures-super"}

# Kap7 → OneDrive
KAP7_ONEDRIVE_URL  = "https://1drv.ms/x/c/6f253ef3afcfe1c8/IQBGJ01faI_8Rb24gRIV-VoDAQO0zhZw_eBDO4rxJ-5tASs?e=hHsWYP"
KAP7_TOURNAMENTS   = {"kap7-intl", "kap7-cup"}

# Turbo Cup → OneDrive (separate organizer, separate file)
TURBO_ONEDRIVE_URL = "https://1drv.ms/x/c/6f253ef3afcfe1c8/IQB7PJXtfzNsT74lTYhWpOeXASFcmpB96L1OpYL_E6HBMM0?e=UsRrMb"

# Newport Spring Invite → OneDrive (older resid/authkey format, direct download)
NEWPORT_ONEDRIVE_URL = "https://onedrive.live.com/download?resid=6F253EF3AFCFE1C8!66694&authkey=!AO8pyWY0qwL2sYE"

# Tournaments with no preset URL — user pastes one via the app
URL_ONLY_TOURNAMENTS = {"jo-quals", "junior-olympics"}
USER_URLS_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "user_urls.json")

_URL_CACHE: dict   = {}   # {url: (fetched_at, bytes)}
URL_CACHE_TTL      = 300  # re-fetch at most every 5 minutes

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
    {"id": "kap7-cup",        "name": "Kap7 Cup",               "dates": "Apr 18–19, 2026"},
    {"id": "futures-4",       "name": "Futures Weekend 4",      "dates": "May 2–3, 2026",
     "date_start": date(2026, 5, 2),  "date_end": date(2026, 5, 3)},
    {"id": "futures-5",       "name": "Futures Weekend 5",      "dates": "May 16–17, 2026",
     "date_start": date(2026, 5, 16), "date_end": date(2026, 5, 17)},
    {"id": "jo-quals",        "name": "JO Qualifications",      "dates": "May 29–31, 2026"},
    {"id": "futures-super",   "name": "Futures Superfinal",     "dates": "Jun 26–28, 2026",
     "date_start": date(2026, 6, 26), "date_end": date(2026, 6, 28)},
    {"id": "junior-olympics", "name": "Junior Olympics",        "dates": "Jul 23–26, 2026"},
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
    if os.path.exists(USER_URLS_FILE):
        with open(USER_URLS_FILE) as f:
            return json.load(f)
    return {}

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
    if tournament_id in KAP7_TOURNAMENTS:
        data = _fetch_url(KAP7_ONEDRIVE_URL, onedrive=True)
        return io.BytesIO(data) if data else None
    if tournament_id == "turbo-cup":
        data = _fetch_url(TURBO_ONEDRIVE_URL, onedrive=True)
        return io.BytesIO(data) if data else None
    if tournament_id == "newport-invite":
        data = _fetch_url(NEWPORT_ONEDRIVE_URL)
        return io.BytesIO(data) if data else None
    return None


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

_TIER_ORDER = {"gold": 0, "cardinal": 1, "platinum": 0, "silver": 2, "bronze": 3}
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
    """'TROJAN GOLD' + '16U BOYS PLATINUM GOLD-11 TEAMS' → 'Boys 16U Gold'."""
    gender, age = _parse_sheet(sheet)
    tier = _parse_tier(team_name)
    parts = [p for p in [gender, age, tier] if p]
    return " ".join(parts) if len(parts) >= 2 else None

def _team_sort_key(team: dict):
    """Sort: Boys before Girls, age desc (18U→10U), tier Gold→Cardinal→Silver."""
    friendly = team.get("friendly") or ""
    gender_ord = 0 if "Boys" in friendly else (1 if "Girls" in friendly else 2)
    age_m = _AGE_WORDS.search(friendly)
    age_ord = -(int(age_m.group(1))) if age_m else 0   # negate for descending
    tier = _parse_tier(team["name"]) or ""
    tier_ord = _TIER_ORDER.get(tier.lower(), 9)
    return (gender_ord, age_ord, tier_ord)


# ── Game helpers ───────────────────────────────────────────────────────────────

_PREFIX_RE = re.compile(
    r"^(?:\d+(?:st|nd|rd|th)[A-Z]-|[A-Z]\d+\([^)]+\)-|[WL]#[^-]+-|[A-Z]\d+-|\d+-)(.*)",
    re.IGNORECASE,
)

def strip_prefix(s: str) -> str:
    m = _PREFIX_RE.match(s.strip())
    return m.group(1).strip() if m else s.strip()

def is_trojan(team_slot: str) -> bool:
    return "TROJAN" in strip_prefix(team_slot).upper()

def team_matches(slot: str, name: str) -> bool:
    return name.upper() in strip_prefix(slot).upper()

def describe_slot(slot: str) -> str:
    """Return just the team name, stripping any bracket/pool prefix.
    Only shows bracket notation when no team name is known yet."""
    slot = slot.strip()
    # strip_prefix extracts the team name after any bracket prefix (A1-, W#12-, H3(2ndC)-, etc.)
    name = strip_prefix(slot)
    if name != slot and name:
        # A prefix was stripped and a real team name remains — just show the team name
        return name
    # No prefix was matched (slot IS the name), or prefix matched but no team name yet
    # For unresolved bracket slots, show something readable
    m = re.match(r"^([WL])#([^-]+)-\s*$", slot, re.IGNORECASE)
    if m:
        wl = "Winner" if m.group(1).upper() == "W" else "Loser"
        return f"{wl} of game #{m.group(2).strip()}"
    m = re.match(r"^(\d+(?:st|nd|rd|th))([A-Z])-\s*$", slot, re.IGNORECASE)
    if m:
        return f"{m.group(1)} place Pool {m.group(2)}"
    m = re.match(r"^([A-Z]\d+)\(([^)]+)\)-\s*$", slot, re.IGNORECASE)
    if m:
        return f"{m.group(2)} from {m.group(1)}"
    return slot

def _game_num(game_id: str):
    m = re.search(r"(\d+)$", game_id)
    return str(int(m.group(1))) if m else None

def find_next_games(game, division_games):
    num = _game_num(game["game_id"])
    if not num:
        return None, None
    winner_next = loser_next = None
    for g in division_games:
        if g["game_id"] == game["game_id"]:
            continue
        for slot in (g["white_team"], g["dark_team"]):
            pm = re.match(r"^([WL])#([^-]+)-", slot)
            if pm:
                ref = re.search(r"(\d+)$", pm.group(2))
                ref_num = str(int(ref.group(1))) if ref else pm.group(2)
                if ref_num == num:
                    if pm.group(1).upper() == "W": winner_next = g
                    else:                           loser_next  = g
    return winner_next, loser_next

def confidence(game, division_games, locked_results):
    """Return 'green' | 'yellow' | 'red' and an explanation."""
    if game["played"]:
        return None, None

    game_id = game["game_id"]

    # Already user-confirmed via yes/no button
    if game_id in locked_results:
        return "green", None

    # Explicit W#/L# bracket reference → we can look up both next games
    winner_next, loser_next = find_next_games(game, division_games)
    if winner_next or loser_next:
        return "green", None

    # Pool play — check if we can determine finish from current standings
    white_prefix = re.match(r"^([A-Z])(\d+)-", game["white_team"])
    dark_prefix  = re.match(r"^([A-Z])(\d+)-", game["dark_team"])
    if white_prefix or dark_prefix:
        pool = (white_prefix or dark_prefix).group(1)
        pool_games = [g for g in division_games if
                      re.match(rf"^{pool}\d+-", g["white_team"]) or
                      re.match(rf"^{pool}\d+-", g["dark_team"])]
        played = sum(1 for g in pool_games if g["played"])
        total  = len(pool_games)
        if played > 0:
            return "yellow", f"{played} of {total} pool games played — upload the latest schedule for a more accurate prediction"
        return "red", "No pool results yet — upload the latest schedule to see your bracket path"

    return "red", "Not enough information yet — upload the latest schedule to improve this prediction"


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

def _next_summary(g, team):
    opp = g["dark_team"] if team_matches(g["white_team"], team) else g["white_team"]
    return {
        "opponent": describe_slot(opp),
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
    today = date.today()
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
        # Rough "past" detection from dates string
        year_match = re.search(r"(\d{4})", t["dates"])
        month_match = re.search(
            r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)", t["dates"], re.I
        )
        _MMAP = {"jan":1,"feb":2,"mar":3,"apr":4,"may":5,"jun":6,
                 "jul":7,"aug":8,"sep":9,"oct":10,"nov":11,"dec":12}
        yr = int(year_match.group(1)) if year_match else today.year
        mo = _MMAP.get(month_match.group(1).lower(), 1) if month_match else 1
        t_date = date(yr, mo, 1)
        out.append({**t, "has_excel": has_excel, "has_file": has_file,
                    "past": t_date < today, "_sort_date": t_date})
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

    teams = [{"name": t["name"], "sheet": t["sheet"], "friendly": t["friendly"]}
             for t in deduped.values()]
    teams.sort(key=_team_sort_key)
    return jsonify(teams)


@app.route("/api/games/<tournament_id>/<path:team>")
def api_games(tournament_id, team):
    excel = find_excel(tournament_id)
    if not excel:
        abort(404)

    games    = _filter_by_dates(load_and_parse(excel), tournament_id)
    locked   = load_results(tournament_id)

    sheet    = request.args.get("sheet")
    my_games = [g for g in games
                if (sheet is None or g["sheet"] == sheet)
                and (team_matches(g["white_team"], team) or team_matches(g["dark_team"], team))]
    my_games.sort(key=lambda g: (g["date"] or date.min, g["time"] or datetime.min.time()))

    div_map = {}
    for g in games:
        div_map.setdefault(g["sheet"], []).append(g)

    played_out   = []
    upcoming_out = []

    for g in my_games:
        dg     = div_map.get(g["sheet"], [])
        gid    = g["game_id"]
        opp_sl = g["dark_team"] if team_matches(g["white_team"], team) else g["white_team"]
        color  = "WHITE" if team_matches(g["white_team"], team) else "DARK"

        base = {
            "game_id":    gid,
            "date":       _fmt_date(g["date"]),
            "time":       _fmt_time(g["time"]),
            "location":   g["location"],
            "opponent":   describe_slot(opp_sl),
            "your_color": color,
            "comments":   g["comments"],
        }

        if g["played"]:
            base["score"]  = _fmt_score(g)
            base["result"] = _result_str(g, team)
            played_out.append(base)
        else:
            conf, conf_msg = confidence(g, dg, locked)
            winner_next, loser_next = find_next_games(g, dg)

            # Check if user already locked a result for this game
            lock = locked.get(gid)
            if lock:
                trojan_won = lock["trojan_won"]
                next_game  = winner_next if trojan_won else loser_next
                base["lock"] = {
                    "trojan_won": trojan_won,
                    "team":       lock["team"],
                    "next":       _next_summary(next_game, team) if next_game else None,
                }
            else:
                # Show yes/no prompt with both scenarios
                scenarios = {}
                if winner_next: scenarios["win"]  = _next_summary(winner_next, team)
                if loser_next:  scenarios["lose"] = _next_summary(loser_next,  team)
                base["scenarios"]   = scenarios if scenarios else None
                base["confidence"]  = conf
                base["conf_msg"]    = conf_msg

            upcoming_out.append(base)

    return jsonify({
        "team":     team,
        "played":   played_out,
        "upcoming": upcoming_out,
    })


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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port, debug=False)

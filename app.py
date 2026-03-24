"""
Tournament Translator — Flask app
"""
import os, re, json, glob
from datetime import datetime, date
from functools import lru_cache
from flask import Flask, render_template, jsonify, request, abort

from parsers.detect import load_and_parse

app = Flask(__name__)

EXCEL_DIR   = os.path.join(os.path.dirname(os.path.abspath(__file__)), "Tournaments Excels")
RESULTS_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "results")
ADMIN_PW    = os.environ.get("ADMIN_PASSWORD", "trojan")  # override via Railway env var

os.makedirs(RESULTS_DIR, exist_ok=True)
os.makedirs(EXCEL_DIR, exist_ok=True)

# ── Tournament registry ────────────────────────────────────────────────────────

KNOWN_TOURNAMENTS = [
    {"id": "futures-2",       "name": "Futures Weekend 2",      "dates": "Feb 21–22, 2026",
     "date_start": date(2026, 2, 21), "date_end": date(2026, 2, 22)},
    {"id": "turbo-cup",       "name": "Turbo OC Cup",           "dates": "Mar 7–8, 2026"},
    {"id": "newport-invite",  "name": "Newport Spring Invite",  "dates": "Mar 14–15, 2026"},
    {"id": "futures-3",       "name": "Futures Weekend 3",      "dates": "Mar 21–22, 2026",
     "date_start": date(2026, 3, 21), "date_end": date(2026, 3, 22)},
    {"id": "kap7-intl",       "name": "Kap7 International",     "dates": "Apr 18–19, 2026"},
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


def find_excel(tournament_id: str):
    """Find the Excel file for a tournament id. Returns path or None."""
    keyword = FILE_MAP.get(tournament_id, "")
    if not keyword:
        return None
    for f in os.listdir(EXCEL_DIR):
        if f.endswith(".xlsx") and keyword.lower() in f.lower():
            return os.path.join(EXCEL_DIR, f)
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
        out.append({**t, "has_excel": excel is not None, "past": t_date < today, "_sort_date": t_date})
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


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5001))
    app.run(host="0.0.0.0", port=port, debug=False)

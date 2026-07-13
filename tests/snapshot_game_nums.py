#!/usr/bin/env python3
"""Capture a complete baseline of game_num assignments across every real
Trojan team/tournament combo. Run BEFORE and AFTER any change to the
game-numbering pipeline (api_games, ~app.py:3849 "Game numbering" section
onward); diff the two output files and manually review every difference
before trusting the change.

Usage:
    python3 tests/snapshot_game_nums.py before.json   # on the old commit
    python3 tests/snapshot_game_nums.py after.json     # after the rewrite
    diff <(python3 -m json.tool before.json) <(python3 -m json.tool after.json)
"""
import os, sys, json

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from app import app, KNOWN_TOURNAMENTS

client = app.test_client()
snapshot = {}

for t in KNOWN_TOURNAMENTS:
    tid = t["id"]
    r = client.get(f"/api/trojan-teams/{tid}")
    if r.status_code != 200:
        continue
    teams = r.get_json() or []
    for team in teams:
        name = team.get("name")
        sheet = team.get("sheet")
        qs = f"?sheet={sheet}" if sheet else ""
        gr = client.get(f"/api/games/{tid}/{name}{qs}")
        if gr.status_code != 200:
            continue
        d = gr.get_json()
        key = f"{tid}::{name}::{sheet}"
        rows = []
        for g in (d.get("played", []) + d.get("upcoming", [])):
            rows.append({
                "game_id": g.get("game_id"),
                "game_num": g.get("game_num"),
                "date": g.get("date"),
                "time": g.get("time"),
                "path": g.get("path"),
                "placeholder": g.get("placeholder"),
                "opponent": g.get("opponent"),
            })
        rows.sort(key=lambda r: (r.get("game_num") or 0, r.get("game_id") or ""))
        snapshot[key] = rows

out_path = sys.argv[1] if len(sys.argv) > 1 else "snapshot.json"
with open(out_path, "w") as f:
    json.dump(snapshot, f, indent=2)
print(f"Captured {len(snapshot)} team/tournament combos -> {out_path}")

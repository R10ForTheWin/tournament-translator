#!/usr/bin/env python3
"""
CI gate for .github/workflows/scheduled-visual-check.yml: writes
is_tournament_day=true/false to $GITHUB_OUTPUT depending on whether today
(Pacific time) falls within any KNOWN_TOURNAMENTS date range.

A standalone script rather than an inline `run: python3 -c "..."` one-liner
in the workflow YAML on purpose -- an earlier version embedded the Python
directly in the YAML, and the nested YAML/bash/Python quoting silently
produced a real SyntaxError (an escaped quote inside an f-string
expression) that would only have surfaced the first time the job actually
fired on a real tournament day. Caught by extracting and running the exact
embedded string through bash before shipping it, not by eyeballing the
YAML. A real .py file sidesteps that whole class of bug and is trivially
testable on its own.

Reads KNOWN_TOURNAMENTS directly from app.py rather than duplicating
tournament dates in the workflow -- one place these need to stay correct.
"""
import os
import sys
from datetime import datetime
from zoneinfo import ZoneInfo

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, ROOT)
from app import KNOWN_TOURNAMENTS

today = datetime.now(ZoneInfo("America/Los_Angeles")).date()
hit = next(
    (t for t in KNOWN_TOURNAMENTS
     if t.get("date_start") and t.get("date_end")
     and t["date_start"] <= today <= t["date_end"]),
    None,
)

label = f" ({hit['id']})" if hit else ""
print(f"Pacific date: {today}, tournament day: {bool(hit)}{label}")

github_output = os.environ.get("GITHUB_OUTPUT")
if github_output:
    with open(github_output, "a") as f:
        f.write("is_tournament_day=" + ("true" if hit else "false") + "\n")

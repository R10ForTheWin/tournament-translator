#!/usr/bin/env python3
"""
Visual regression check: pixel-diff screenshots against committed baselines.

Deliberately scoped to the WPL futures-5 fixture (backed by a local static
Excel file, see FILE_MAP in app.py) -- NOT Junior Olympics or any other
live-fetched tournament. A live Google Sheet changes data (scores, dates,
opponents) between every run, which would make a naive pixel diff flag
constant false positives on real, legitimate data changes and train
everyone to ignore it. The geometric/structural checks in
verify_bracket_rendering.py already cover JO; this script covers the class
of bug those checks can't articulate as an explicit assertion (broken CSS,
overlapping elements, a color regression, an unreadable layout) -- which
needs actual pixel comparison, and needs data that holds still to do it.

Usage:
    python3 tests/visual_regression.py --port 5099            # compare against baselines
    python3 tests/visual_regression.py --port 5099 --update   # (re)write baselines

Requires: pip install playwright Pillow && playwright install chromium
Baselines live in tests/visual_baselines/*.png (git-committed).
"""
import argparse
import os
import subprocess
import sys
import time
import urllib.request

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
BASELINE_DIR = os.path.join(ROOT, "tests", "visual_baselines")

# The Trojan-named teams from CHAMPIONSHIP_CHECKS in test_parsers.py, same
# static fixture. Not "imperial"/"socal" (also in CHAMPIONSHIP_CHECKS) --
# those are opponent teams, tested there via direct backend calls, but
# api_trojan_teams only ever surfaces "TROJAN"-named teams to the browsable
# UI team list this script navigates through, so they have no page to visit.
FUTURES5_TEAMS = [
    ("trojan gold",     "16u Boys"),
    ("trojan cardinal", "16u Boys"),
]

FAILURES = []


def check(label: str, condition: bool, detail: str = "") -> bool:
    status = "PASS" if condition else "FAIL"
    suffix = f"  ({detail})" if detail else ""
    print(f"  [{status}] {label}{suffix}")
    if not condition:
        FAILURES.append(label)
    return condition


def wait_for_server(port: int, timeout: int = 20) -> bool:
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            urllib.request.urlopen(f"http://localhost:{port}/", timeout=2)
            return True
        except Exception:
            time.sleep(1)
    return False


def capture(page, port: int, team: str, sheet: str):
    """Navigate to the team's futures-5 bracket/schedule view and return a
    PNG screenshot (bytes), with the volatile "Schedule last fetched Xm
    ago" tag masked out -- the only part of an otherwise-static fixture
    render that changes on every single run.

    Self-verifying navigation (same pattern as verify_bracket_rendering.py):
    team names can share overlapping captured text with sibling cards (the
    proximity heuristic below walks up 2 DOM levels), so the first
    "candidate" match is not reliable on its own -- click it, check the
    page's own header actually says what we expect, and try the next
    candidate via the in-app Back button if not. Also needs a real wait for
    the sheet fetch to complete (a short fixed sleep was caught live
    producing a screenshot of a still-spinning loading state)."""
    page.goto(f"http://localhost:{port}/", wait_until="load")
    page.get_by_text("Futures Weekend 5", exact=False).first.click()
    try:
        page.wait_for_selector("button:has-text('SCHEDULE')", timeout=20000)
    except Exception:
        pass
    page.wait_for_timeout(1500)

    btns = page.get_by_role("button", name="SCHEDULE").all()
    matching = [
        i for i, b in enumerate(btns)
        if team.upper() in b.evaluate(
            "el => el.closest('div')?.parentElement?.parentElement?.innerText || ''"
        ).upper()
    ]
    landed = False
    for idx in matching:
        btns = page.get_by_role("button", name="SCHEDULE").all()
        btns[idx].click()
        page.wait_for_timeout(15000)  # live sheet fetch can be slow on cold cache
        header = page.evaluate("() => document.getElementById('subnav-title')?.innerText || ''")
        if team.upper() in header.upper():
            landed = True
            break
        page.click("#back-btn")
        page.wait_for_timeout(1500)
    if not landed:
        return None

    # Give any final render/layout settle a moment beyond the header
    # appearing (cards/connectors can still be drawing).
    page.wait_for_timeout(1000)
    mask = page.locator(".cache-tag")
    return page.screenshot(mask=[mask] if mask.count() else [])


def compare(name: str, current_png: bytes, update: bool):
    from PIL import Image
    import io

    os.makedirs(BASELINE_DIR, exist_ok=True)
    baseline_path = os.path.join(BASELINE_DIR, f"{name}.png")

    if update or not os.path.exists(baseline_path):
        with open(baseline_path, "wb") as f:
            f.write(current_png)
        print(f"  [BASELINE] {name}: wrote new baseline ({len(current_png)} bytes)")
        return

    baseline = Image.open(baseline_path).convert("RGB")
    current = Image.open(io.BytesIO(current_png)).convert("RGB")

    if baseline.size != current.size:
        check(f"{name}: screenshot dimensions match baseline", False,
              f"baseline={baseline.size} current={current.size} -- layout likely changed")
        return

    bl_px = baseline.load()
    cur_px = current.load()
    w, h = baseline.size
    diff_count = 0
    # Sample every pixel but tolerate small per-channel deltas (anti-aliasing
    # noise from font rendering) -- only count a pixel as "different" if any
    # channel differs by more than a small threshold.
    for y in range(0, h, 2):      # every other row is plenty for this purpose, 4x faster
        for x in range(0, w, 2):
            b = bl_px[x, y]
            c = cur_px[x, y]
            if any(abs(b[i] - c[i]) > 20 for i in range(3)):
                diff_count += 1
    sampled = (w // 2) * (h // 2)
    pct = 100.0 * diff_count / sampled if sampled else 0.0
    check(f"{name}: visual diff under 2% ({pct:.2f}%)", pct < 2.0,
          f"{diff_count}/{sampled} sampled pixels differ -- "
          f"if this is an intentional UI change, rerun with --update")


PORT = 5099

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=PORT)
    ap.add_argument("--keep-server", action="store_true")
    ap.add_argument("--update", action="store_true",
                     help="(re)write baselines instead of comparing against them")
    args = ap.parse_args()
    PORT = args.port

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Playwright not installed: pip install playwright && playwright install chromium")
        sys.exit(2)
    try:
        import PIL  # noqa: F401
    except ImportError:
        print("Pillow not installed: pip install Pillow")
        sys.exit(2)

    proc = None
    if not args.keep_server:
        proc = subprocess.Popen(
            [sys.executable, "app.py"],
            env={"PORT": str(PORT), **os.environ},
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        if not wait_for_server(PORT):
            print("Server never came up")
            sys.exit(1)

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page(viewport={"width": 1400, "height": 1000})
            for team, sheet in FUTURES5_TEAMS:
                name = team.replace(" ", "_")
                print(f"\n{'=' * 60}\nfutures-5: {team} / {sheet}\n{'=' * 60}")
                png = capture(page, PORT, team, sheet)
                if png is None:
                    check(f"{team}/{sheet}: found and navigated to team page", False,
                          "no matching SCHEDULE button")
                    continue
                compare(name, png, args.update)
            browser.close()
    finally:
        if proc:
            proc.terminate()
            proc.wait(timeout=10)

    print(f"\n{'=' * 60}")
    if FAILURES:
        print(f"{len(FAILURES)} check(s) FAILED")
        sys.exit(1)
    print("All checks passed")

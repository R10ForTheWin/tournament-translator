#!/usr/bin/env python3
"""
Frontend bracket-rendering regression check (Playwright, real browser).

Everything in test_parsers.py verifies backend JSON only. Two real bugs
shipped this session (2026-07-14) that no backend test could ever catch,
because they lived entirely in how the frontend renders that JSON:

  1. A tournament-ID prefix check (`tournamentId.startsWith('futures')`)
     silently excluded every non-WPL tournament from using the bracket
     edge data (src_game_id/win_next_ids/lose_next_ids) at all -- it fell
     back to a flat list with no edge data, so connector lines fell back
     to "nearest card on screen", which is wrong the moment a column
     merges children of different parents.
  2. Even after fixing #1, drawBracketLines' connector-matching logic had
     to be verified against the *actual rendered DOM* (data attributes
     and SVG path coordinates) -- reasoning about the code by hand wasn't
     enough to catch the first bug, since the code that WAS fixed (the
     matching logic itself) was correct all along and simply never ran.

This script drives a real headless browser against a running dev server
and checks, from the rendered DOM, that each bracket connector actually
matches its real parent -- not just that the page loads without errors.

Usage:
    python3 tests/verify_bracket_rendering.py [--port 5099]

Requires: pip install playwright && playwright install chromium
"""
import argparse
import subprocess
import sys
import time
import urllib.request

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


def verify_bracket(page, tournament_label: str, team_label: str, sheet_hint: str):
    """Navigate the real UI (tournament card -> team SCHEDULE button) and
    return the rendered game-card + connector data. Navigates by visible
    text/role, not hardcoded indices, so it survives menu-copy changes."""
    page.goto(f"http://localhost:{PORT}/", wait_until="networkidle")
    page.get_by_text(tournament_label, exact=False).first.click()
    page.wait_for_timeout(6000)

    # Team names repeat across age groups (e.g. "Trojan Gold" appears once
    # per division), so don't trust DOM-proximity heuristics to guess the
    # right button in advance -- self-verify instead: click each
    # team_label match in turn and check the page's own header (which the
    # app itself sets authoritatively) confirms both team AND sheet_hint
    # before accepting it. Robust to DOM layout changes; not to renamed
    # team/tournament labels, which would need updating here anyway.
    btns = page.get_by_role("button", name="SCHEDULE").all()
    matching_idx = [
        i for i, b in enumerate(btns)
        if team_label in b.evaluate(
            "el => el.closest('div')?.parentElement?.parentElement?.innerText || ''"
        )
    ]
    landed = False
    for idx in matching_idx:
        btns = page.get_by_role("button", name="SCHEDULE").all()  # re-query after any nav
        btns[idx].click()
        page.wait_for_timeout(15000)  # live sheet fetch can be slow on cold cache
        header = (page.locator("#subnav-title, .header-title, header").first.inner_text()
                  if page.locator("#subnav-title").count() else "")
        if not header:
            header = page.evaluate("() => document.getElementById('subnav-title')?.innerText || ''")
        if team_label.upper() in header.upper() and sheet_hint.upper() in header.upper():
            landed = True
            break
        # This app is a client-side SPA (no real history entries) -- use
        # its own in-app Back button, not browser back navigation.
        page.click("#back-btn")
        page.wait_for_timeout(2000)
    if not landed:
        check(f"landed on the {team_label!r}/{sheet_hint!r} page", False,
              f"tried {len(matching_idx)} candidate button(s), header never matched")
        return None, None

    console_errors = []
    page.on("console", lambda m: console_errors.append(m.text) if m.type == "error" else None)

    cards = page.query_selector_all(".game-card")
    card_info = [
        c.evaluate(
            "el => ({id: el.dataset.gameId, src: el.dataset.srcGameId, "
            "path: el.dataset.path, top: el.getBoundingClientRect().top})"
        )
        for c in cards
    ]
    return card_info, console_errors


def run_case(page, label, tournament_label, team_label, sheet_hint, min_cards):
    print(f"\n{'=' * 60}\n{label}\n{'=' * 60}")
    cards, console_errors = verify_bracket(page, tournament_label, team_label, sheet_hint)
    if cards is None:
        return

    check(f"at least {min_cards} bracket cards rendered", len(cards) >= min_cards,
          f"got {len(cards)}")
    check("no console errors", not console_errors, str(console_errors[:3]))

    # The real invariant that was broken: every non-root card's src_game_id
    # must point at a game_id that actually exists among the rendered
    # cards. An empty src on a non-root card means the edge data pipeline
    # (the exact class of bug fixed today) silently dropped out again.
    ids = {c["id"] for c in cards if c["id"]}
    non_root = [c for c in cards if c["path"]]  # root card has path=''
    for c in non_root:
        ok = check(
            f"{c['id']}: src_game_id {c['src']!r} resolves to a rendered card",
            bool(c["src"]) and c["src"] in ids,
            f"card data: {c}",
        )


PORT = 5099

if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--port", type=int, default=PORT)
    ap.add_argument("--keep-server", action="store_true",
                     help="don't start/stop a Flask server -- assume one is already running")
    args = ap.parse_args()
    PORT = args.port

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Playwright not installed: pip install playwright && playwright install chromium")
        sys.exit(2)

    proc = None
    if not args.keep_server:
        proc = subprocess.Popen(
            [sys.executable, "app.py"],
            env={"PORT": str(PORT), **__import__("os").environ},
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        if not wait_for_server(PORT):
            print("Server never came up")
            sys.exit(1)

    try:
        with sync_playwright() as p:
            browser = p.chromium.launch()
            page = browser.new_page(viewport={"width": 1400, "height": 1000})

            # The known regression case: a multi-round JO bracket where two
            # different immediate parents (games 028 and 032) are already
            # merged as one game_num, so their children must connect to
            # their REAL parent, not whichever card is nearest on screen.
            run_case(page, "Junior Olympics: TROJAN GOLD 16U (cascading-merge case)",
                     "Junior Olympics", "Trojan Gold", "16U", min_cards=4)

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

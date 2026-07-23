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
import json
import re
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


def wait_for_server(port: int, timeout: int = 150, tournament_id: str = "junior-olympics") -> bool:
    """Polls /api/tournaments until (a) the tournament this script needs
    reports has_excel=true AND (b) that same request comes back FAST
    (< 2s) -- a bare 200, or even a has_excel=true match, isn't enough on
    its own. Root cause, confirmed live 2026-07-23: this route builds its
    response by looping over EVERY known tournament in ONE request, so
    even once OUR tournament is warm, a request can still block for a long
    time if a DIFFERENT tournament is simultaneously cold (its own
    find_excel() racing the prewarm thread's first-ever fetch of the same
    live URL -- no de-dup lock covers that first cold fetch, only
    _refresh_url_background's _FETCH_INFLIGHT covers re-fetching an
    already-cached URL). The frontend's own page-load-time fetch can lose
    that race even when a prior probe request (like this one) already saw
    has_excel=true for our tournament specifically -- a slow OTHER
    tournament in the same response is invisible from here unless timed.
    A fast response time is the actual proxy for "everything is warm now",
    not just our one tournament. This produced a real, reproducible
    "Schedule not posted yet" info sheet in place of the team list on a
    just-started process -- a genuine (if narrow-window, self-correcting)
    production gap right after any fresh deploy/restart, see
    project_junior_olympics_prep memory, deferred as low-priority
    post-tournament since it only affects the ~seconds after a restart and
    production has been warm for hours."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            t0 = time.time()
            resp = urllib.request.urlopen(f"http://localhost:{port}/api/tournaments", timeout=timeout)
            data = json.loads(resp.read())
            elapsed = time.time() - t0
            t = next((x for x in data if x.get("id") == tournament_id), None)
            if t and t.get("has_excel") and elapsed < 2.0:
                return True
        except Exception:
            pass
        time.sleep(2)
    return False


def verify_bracket(page, tournament_label: str, team_label: str, sheet_hint: str):
    """Navigate the real UI (tournament card -> team card) and return the
    rendered game-card + connector data. Navigates by visible text, not
    hardcoded indices, so it survives menu-copy changes.

    Team cards for non-futures tournaments (which is everything this script
    tests -- JO_TEAMS below) are a single directly-clickable ".team-btn"
    div, not a separate "SCHEDULE" button -- the 2026-07-20 UI
    simplification ("Team list / subnav UI simplified", commit 6b240ad)
    merged the old Schedule/Rankings button pair into one click for every
    non-futures tournament, and dropped the word "Trojan" from both the
    card label and the destination header (every team in this app is a
    Trojan team, so it was redundant). Confirmed live 2026-07-23: this
    script pre-dates that commit and had never actually run for real since
    (its CI job only fires on a live tournament day, and none occurred
    between 2026-07-20 and today), so the mismatch went uncaught until
    now -- not a live site bug, the real app works fine either way."""
    # NOT wait_until="networkidle" -- this app polls periodically in the
    # background (_startRefreshTimers), so the network is never truly idle
    # and that wait condition can hang indefinitely. "load" plus the
    # explicit wait_for_timeout calls below is what actually works here.
    # Retried, not a single fixed-timeout attempt -- confirmed live
    # 2026-07-23: back-to-back navigations in one long-lived process can
    # occasionally hit a slow /api/tournaments response on just one reload
    # (the live Google Sheets fetch this route depends on has its own real-
    # network variance), even after the whole suite's initial cold-start
    # settling. A second attempt after a fresh reload has always succeeded
    # in practice; this is strictly more forgiving than before, never less.
    for attempt in range(3):
        page.goto(f"http://localhost:{PORT}/", wait_until="load")
        try:
            page.get_by_text(tournament_label, exact=False).first.click(timeout=30000)
            break
        except Exception:
            if attempt == 2:
                raise
    # Wait for the team list to actually render rather than a fixed sleep --
    # under back-to-back runs (many teams checked in one process) the
    # tournament card's fetch can occasionally take longer than a flat
    # timeout, and a fixed sleep that's usually enough becomes an
    # intermittent "0 candidate cards" flake under load, not a real bug.
    try:
        page.wait_for_selector(".team-btn", timeout=20000)
    except Exception:
        pass  # fall through -- the empty-candidate-list check below will report it clearly
    page.wait_for_timeout(1500)

    # "Trojan " is stripped from both the card label and the destination
    # header, so match/verify on the distinguishing part only.
    short_label = re.sub(r'^Trojan\s+', '', team_label, flags=re.IGNORECASE)

    # Team names repeat across age groups (e.g. "Gold" appears once per
    # division), so don't trust DOM-proximity heuristics to guess the right
    # card in advance -- self-verify instead: click each candidate in turn
    # and check the page's own header (which the app itself sets
    # authoritatively) confirms both team AND sheet_hint before accepting
    # it. Robust to DOM layout changes; not to renamed team/tournament
    # labels, which would need updating here anyway.
    cards = page.locator(".team-btn").all()
    matching_idx = [
        i for i, c in enumerate(cards)
        if short_label.upper() in c.inner_text().upper()
        and sheet_hint.upper() in c.inner_text().upper()
    ]
    landed = False
    for idx in matching_idx:
        cards = page.locator(".team-btn").all()  # re-query after any nav
        cards[idx].click()
        page.wait_for_timeout(15000)  # live sheet fetch can be slow on cold cache
        header = (page.locator("#subnav-title, .header-title, header").first.inner_text()
                  if page.locator("#subnav-title").count() else "")
        if not header:
            header = page.evaluate("() => document.getElementById('subnav-title')?.innerText || ''")
        if short_label.upper() in header.upper() and sheet_hint.upper() in header.upper():
            landed = True
            break
        # This app is a client-side SPA (no real history entries) -- use
        # its own in-app Back button, not browser back navigation.
        page.click("#back-btn")
        page.wait_for_timeout(2000)
    if not landed:
        check(f"landed on the {team_label!r}/{sheet_hint!r} page", False,
              f"tried {len(matching_idx)} candidate card(s), header never matched")
        return None, None

    console_errors = []
    page.on("console", lambda m: console_errors.append(m.text) if m.type == "error" else None)

    cards = page.query_selector_all(".game-card")
    card_info = [
        c.evaluate(
            "el => ({id: el.dataset.gameId, src: el.dataset.srcGameId, "
            "path: el.dataset.path, top: el.getBoundingClientRect().top, "
            "left: el.getBoundingClientRect().left})"
        )
        for c in cards
    ]

    # Extract each connector's start/end Y, bend-point X, and source game_id
    # from its SVG path 'd' string / dataset. Straight (same-row) connectors
    # have no bend point.
    paths = page.evaluate(
        r"""
        () => Array.from(document.querySelectorAll('svg path')).map(p => {
            const d = p.getAttribute('d');
            const src = p.dataset.srcGameId || null;
            let m = d.match(/^M ([\d.]+),([\d.]+) H ([\d.]+)$/);
            if (m) return {sy: +m[2], ty: +m[2], gx: null, src};
            m = d.match(/^M ([\d.]+),([\d.]+) H [\d.]+ Q ([\d.]+),[\d.]+ [\d.]+,[\d.]+ V [\d.]+ Q [\d.]+,([\d.]+)/);
            if (m) return {sy: +m[2], ty: +m[4], gx: +m[3], src};
            return null;
        }).filter(Boolean)
        """
    )
    return card_info, console_errors, paths


def run_case(page, label, tournament_label, team_label, sheet_hint, min_cards):
    print(f"\n{'=' * 60}\n{label}\n{'=' * 60}")
    result = verify_bracket(page, tournament_label, team_label, sheet_hint)
    if result[0] is None:
        return
    cards, console_errors, paths = result

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
        check(
            f"{c['id']}: src_game_id {c['src']!r} resolves to a rendered card",
            bool(c["src"]) and c["src"] in ids,
            f"card data: {c}",
        )

    # The second bug, caught only by a human looking at a screenshot: two
    # crossing connectors from DIFFERENT sources that bend through the SAME
    # x point visually collapse into what looks like one merged trunk. Data
    # correctness alone (the check above) can't catch this -- it's a
    # legibility property of the rendered lines. Approximate it
    # geometrically: any two connectors from DIFFERENT sources whose
    # vertical spans overlap (they cross or run parallel on screen) must
    # bend at visually distinct x positions.
    #
    # Connectors from the SAME source are the opposite case and must NOT be
    # required to differ: a single parent's win/lose children are supposed
    # to share one bend point (one trunk line forking into two), the
    # standard tournament-bracket look. An earlier version of this fix
    # fanned out by TARGET index regardless of source, which "fixed" the
    # different-parents case but broke this one -- caught live 2026-07-17
    # via a phone screenshot showing a single source's two children drawn
    # as two independently-curved, visually crossing lines instead of one
    # clean fork.
    bent = [p for p in paths if p["gx"] is not None]
    for i in range(len(bent)):
        for j in range(i + 1, len(bent)):
            a, b = bent[i], bent[j]
            if a["src"] and b["src"] and a["src"] == b["src"]:
                continue  # same source: sharing a bend point is correct
            lo_a, hi_a = sorted((a["sy"], a["ty"]))
            lo_b, hi_b = sorted((b["sy"], b["ty"]))
            overlaps = lo_a < hi_b and lo_b < hi_a
            if overlaps:
                check(
                    f"crossing connectors near y={lo_a:.0f}-{hi_a:.0f} (src {a['src']!r} vs {b['src']!r}) "
                    f"have distinct bend points",
                    abs(a["gx"] - b["gx"]) > 5,
                    f"gx={a['gx']} vs gx={b['gx']} -- would visually overlap",
                )

    # Positive companion to the check above: connectors that DO share a
    # source must share the exact same bend point (one trunk forking into
    # two), not just "close enough."
    by_src: dict = {}
    for p in bent:
        if p["src"]:
            by_src.setdefault(p["src"], []).append(p)
    for src, group in by_src.items():
        if len(group) < 2:
            continue
        gxs = {p["gx"] for p in group}
        check(f"{src!r}: all {len(group)} children share one bend point (single fork, not split lines)",
              len(gxs) == 1, f"gx values={gxs}")

    # Third bug (caught live via a phone screenshot 2026-07-15, after the
    # first two): even with correct src_game_id data and distinct bend
    # points, sorting each column purely by "win before lose" -- ignoring
    # which row the card's own parent actually sits in -- put the win-child
    # of the BOTTOM column N-1 parent above the lose-child of the TOP column
    # N-1 parent. Both cards were individually correct, but the two
    # connectors crossed unnecessarily, rendering as a confusing hook/loop
    # shape instead of a clean staircase. Assert the fix holds: whenever two
    # cards in the same column both have a src_game_id resolvable to a row
    # in the immediately previous column, their own row order must match
    # their parents' row order.
    by_left: dict = {}
    for c in cards:
        by_left.setdefault(round(c["left"] / 50) * 50, []).append(c)
    col_lefts = sorted(by_left.keys())
    for i in range(len(col_lefts) - 1):
        prev_col = sorted(by_left[col_lefts[i]], key=lambda c: c["top"])
        this_col = sorted(by_left[col_lefts[i + 1]], key=lambda c: c["top"])
        prev_row = {c["id"]: idx for idx, c in enumerate(prev_col)}
        resolved = [(c, prev_row.get(c["src"])) for c in this_col if c["src"] in prev_row]
        for a in range(len(resolved)):
            for b in range(a + 1, len(resolved)):
                (ca, pa), (cb, pb) = resolved[a], resolved[b]
                check(
                    f"{ca['id']} (parent row {pa}) vs {cb['id']} (parent row {pb}): "
                    f"no unnecessary connector crossing",
                    pa <= pb,
                    f"{ca['id']} sits above {cb['id']} but its parent is BELOW {cb['id']}'s parent",
                )

    return cards, paths


def run_escape_html_case(page, port: int):
    """_escHtml (templates/index.html) is what stands between organizer-
    controlled spreadsheet text (opponent names, venue locations, team
    names in standings/predictor) and innerHTML -- found via a security
    audit that ~12 interpolation sites inserted that text unescaped, one
    of which (game card location, rendered on every single card) had no
    protection at all. Verify it actually escapes HTML-special characters
    instead of just trusting the source code reads correctly."""
    print(f"\n{'=' * 60}\n_escHtml XSS-escaping sanity check\n{'=' * 60}")
    page.goto(f"http://localhost:{port}/", wait_until="load")
    cases = [
        ("<script>alert(1)</script>", "&lt;script&gt;alert(1)&lt;/script&gt;"),
        ("Team & Co.", "Team &amp; Co."),
        ("<img src=x onerror=alert(1)>", "&lt;img src=x onerror=alert(1)&gt;"),
        ("O'Brien's Team", "O&#39;Brien&#39;s Team"),
        ("Normal Team Name", "Normal Team Name"),
        (None, ""),
    ]
    results = page.evaluate(
        "(cases) => cases.map(([input, expected]) => "
        "({input, expected, got: _escHtml(input)}))",
        cases,
    )
    for r in results:
        check(f"_escHtml({r['input']!r}) escapes correctly",
              r["got"] == r["expected"], f"got {r['got']!r}")


def run_tbd_stub_case(page, tournament_label, team_label, sheet_hint):
    """When a team's bracket tree runs out of real data before the
    tournament's own posted schedule does, the app appends TBD placeholder
    cards (one per remaining day) directly into the tree -- not a text
    footer. Confirm they actually render as cards with the expected
    minimal content, not just that the backend returns tbd_stub=true."""
    print(f"\n{'=' * 60}\nTBD stub cards render correctly\n{'=' * 60}")
    result = verify_bracket(page, tournament_label, team_label, sheet_hint)
    if result[0] is None:
        return
    stub_cards = page.query_selector_all(".game-card.placeholder .tbd-stub-msg")
    check("at least one TBD stub card rendered", len(stub_cards) > 0,
          f"got {len(stub_cards)}")
    if stub_cards:
        text = stub_cards[0].inner_text()
        check("TBD stub card shows the expected placeholder message",
              "TBD" in text, f"got {text!r}")


# All 7 known Trojan JO teams (same list used by test_parsers.py's
# generic tournament-health sweep and game_num-collision test) -- run the
# full geometric + TBD-stub battery against every one of them, not just the
# one case originally used to develop these checks. Generalizing this is
# what caught the missing-branch bug affecting 4 of these 7 teams
# (2026-07-16): the single hardcoded case below only ever exercised 16U
# Trojan Gold, so the same bug sitting on the other 3 teams' brackets
# would have shipped invisibly to this suite.
JO_TEAMS = [
    ("Trojan Cardinal", "18U"),
    ("Trojan Gold",     "18U"),
    ("Trojan Cardinal", "16U"),
    ("Trojan Gold",     "16U"),
    ("Trojan Cardinal", "14U"),
    ("Trojan Cardinal", "12U"),
    ("Trojan Gold",     "12U"),
]

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

            for team_label, sheet_hint in JO_TEAMS:
                run_case(page, f"Junior Olympics: {team_label} {sheet_hint}",
                         "Junior Olympics", team_label, sheet_hint, min_cards=1)
                run_tbd_stub_case(page, "Junior Olympics", team_label, sheet_hint)

            run_escape_html_case(page, PORT)

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

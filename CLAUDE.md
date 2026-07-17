# Context for anyone (or anything) working on this codebase

This is the real history, pulled from commit logs and this project's own
incident notes, not a guess from a diagram. It's context, not orders. Read
it, weigh it against whatever you're being asked to do, and use judgment.

## What this app is

iPhone-first web app that turns messy tournament Excel/Google Sheets
exports into a schedule/bracket view for parents following their kids'
water polo games in real time. Flask backend (single app.py, ~5000+ lines),
vanilla JS/HTML frontend (single templates/index.html, no framework),
deployed on Railway, CI on GitHub.

Pipeline: Fetch -> Parse Cascade (AI-cached, then Format A / Format B /
CCA deterministic parsers, then AI API as last resort) -> Validate
(deterministic hard rules, not an LLM judge, that was tried and replaced)
-> Expand (slot resolution) -> Game Numbering (decision-key grouping) ->
Tree Build (format-specific builders) -> Bracket Confidence Validator
(green/yellow/red) -> Serialize into two parallel outputs (flat schedule,
canonical bracket graph) -> API JSON -> frontend renders one of the two
views depending on confidence.

Each tournament format (Format A, Format B, CCA/JO, WPL Championship) has
its own parser module but shares the downstream expand/numbering/tree/
validation code. That shared code is the single biggest recurring source
of bugs in this project's history: a new format exercises an assumption
baked into shared code that was written for an older one.

## The one pattern behind almost every real incident so far

Two parallel code paths compute the same underlying concept, one gets
fixed when a bug surfaces, the other, doing the same job somewhere else
in the codebase, doesn't get touched and breaks the same way later. This
project's own self-review names it as the most common root cause across
its incident history, ahead of "new format breaks shared code" and
"visual bugs invisible to backend-only tests."

It's hit "last meeting / opponent history" three separate times and
"game numbering" twice, most recently this week (2026-07-15/16), when the
bracket-tree view turned out to be running a completely different,
untested numbering system than the one already fixed for the flat
schedule list, and two genuinely separate real games got silently merged
under one "GAME 1" label as a result.

**Practical implication:** before writing new logic for something that
sounds like it should already exist (numbering, day-ordering, head-to-head
history, anything computed from game state), search for whether it's
already implemented somewhere else first. If it is, fix or reuse that one
implementation instead of adding a second.

## Other things this project has already learned the hard way

- **Never push architecture changes on tournament day.** The 2026-05-02/03
  WPL Futures Week 4 outage happened because a new deterministic validator
  was pushed on a live tournament day; it correctly caught a pre-existing
  format_b bug, then rejected format_b output entirely and took the app
  down for parents mid-tournament. Targeted bug fixes on tournament day
  are fine, architecture changes are not.

- **Audit shared code before a new format goes live.** The 2026-05-30/31
  CCA/JO Quals sprint (five production bugs in one day) all traced back to
  adding the CCA format without checking its new patterns (spaced slot
  syntax, sequential same-day bracket rounds) against existing shared
  parsing/day-merge/slot-regex code. There's a partially-manual "new
  tournament checklist" for this now, worth following rather than skipping.

- **A fix for one tournament can quietly break another.** During the
  2026-07-09/10 Quiksilver Cup prep, a format_a fix for one tournament
  broke Newport, which depended on the old, technically-buggy behavior.
  Worth checking other live tournaments' data against a shared-code change,
  not just the one that prompted it.

- **Silently-wrong logic can hide for a long time if data doesn't exercise
  it.** The composite pool-rank slot bug found the same week had been
  inverting a win/loss result in shared code for a while, undetected only
  because past tournament data hadn't hit that case.

- **Visual bugs don't show up in backend tests.** The "dead zone" on the
  horizontal-scroll bracket view took several wrong-theory attempts
  (touch-action CSS, padding tuning) before the real cause, the scrollable
  container's height not reaching the last row, was found via live phone
  testing, a week after it was first noticed. Same category this week: the
  bracket connector lines were visually crossing even though the
  underlying data was correct, fixed by making card row-order follow
  parent row-order instead of sorting by win/lose label alone. Real phone
  testing on the actual view catches things unit tests won't.

- **Scope lookups to the right boundary.** This week's third bug: the
  bracket tree's "last meeting" lookup wasn't scoped by age division,
  leaking a 16U opponent's history onto a 12U team's card. Worth checking
  that any lookup keyed loosely (by team name alone, etc.) is actually
  scoped to the right tournament, division, or bracket context.

## Reliability layer, still in force

Deterministic validator (not an LLM) checks bracket structure and scores
green/yellow/red confidence. RED always falls back to the flat schedule,
never show a bracket the code doesn't trust. Background jobs: smoke test
on every URL/data change, pre-game monitor sweeping all teams 12-25h
before each tournament, and (added this project cycle) a full parser test
suite plus real-browser Playwright checks on every push to GitHub.

## Worth doing before trusting a change

- Check whether the thing being built already exists somewhere else in
  the codebase under a different name or view, given how often that's
  been the actual bug.
- Don't bundle an architecture change into a tournament-day fix.
- Run the new-format audit against shared code before a new format format
  goes live, not after.
- Test the actual frontend view on a phone for anything visual, not just
  the API output.

---
name: ig-partnership-timeline
description: Enrich an existing Instagram creator/profile-metrics Excel file (one tab per brand, rows of creators with a Username column) with partnership/collab timeline data — when each creator's collab post went live, the post URL, and the caption. Also cross-checks Meta's Ad Library for paid creator ads missing from the organic list. Use this whenever the user asks to add a partnership/collab timeline to a creator spreadsheet, wants to know when a creator's partnership post or ad "went live," asks to enrich or update a profile-metrics Excel with collab dates, mentions cross-checking Meta Ad Library / TrendTrack against an existing creator list, or hands over an Instagram creator-metrics spreadsheet expecting collab/ad history added — even if they don't name every step or say "skill."
---

# Instagram partnership timeline enrichment

Adds "when did this creator's brand partnership go live" data onto an
existing creator-metrics spreadsheet, by combining two independent scans:

- **Method A — organic Instagram scan**: finds regular collab posts on the
  brand's own feed.
- **Method B — Meta Ad Library cross-check**: finds creators the brand is
  currently running as a *paid ad*, including ones that never showed up
  organically.

Both methods, why they're needed, and their gotchas are documented in
detail below and in `references/ad_library_method.md`. Read that reference
before running Method B — it contains a critical accuracy warning that
changes how you should report results.

## Before you start: read the target spreadsheet

Open the user's existing Excel file first (or ask for it if not provided).
Confirm:
- Which sheet/tab corresponds to which brand's Instagram handle.
- Which row has headers, and which column holds the `Username` (or
  `@handle`) values — this workflow defaults to header row 2, username
  column 2 (matching this project's standard "Profiles Overview" layout),
  but confirm rather than assume if the file looks different.
- Whether the user wants results kept in **separate tabs per brand** (the
  default and generally the safer assumption) or merged — don't combine
  accounts unless explicitly told to.

Never create new sheets for this data. It goes into the existing tabs as
appended columns, per how this workflow has always been used.

## Step 1 — Organic scan (Method A) for each brand

Run the bundled script once per brand handle:

```bash
python scripts/scan_partnerships.py <brand_handle> --project-dir <path-to-project> --out scratch_<brand_handle>.json
```

- `--project-dir` must point at the folder containing a `scrape_profiles.py`
  with a live `COOKIES` dict — this workflow reuses that project's existing
  session rather than asking the user for credentials again.
- Default `--max-posts 400` is usually enough to reach well past a brand's
  collab-heavy recent history; raise it if the brand posts very frequently
  or the user wants deeper history.
- Watch the script's own stdout: it prints a warning if one "partner"
  dominates the hit list (>50% share), which usually means a sister/
  sub-brand account cross-tagging the parent rather than a real third-party
  creator (this happened with jockeywomanindia ↔ jockeyindia mutually
  tagging each other). Sanity-check that name with the user before treating
  it as a genuine partner.

This produces one JSON file per brand with every detected collab post, the
partner(s) tagged, and the caption — this is the input to Step 3.

## Step 2 — Ad Library cross-check (Method B)

Read `references/ad_library_method.md` in full before doing this step — it
covers the exact browser-automation procedure (persistent accumulator
object across repeated scroll calls, the blob-download extraction trick,
and quirks with timeouts/dropped downloads) and, importantly, **a proven
completeness trap**: naive scroll capture reliably undercounts, sometimes
by roughly half, in an unpredictable way per brand. Never report a Method B
number as "confirmed complete" without either the double-zero-pass
accumulation check or a cross-check against a real external source
(TrendTrack or similar, if the user provides one).

After capturing Ad Library ads for a brand, compare creator names against
that brand's Method A `unique` partner list (from Step 1's JSON output).
For any Ad Library creator with no match in the organic list, add an entry
to an extras JSON file (see the reference doc for the exact shape) — this
becomes the input to Step 3's `--ad-library-extra` flag. Never fabricate a
post URL or caption for these rows; leave them explicitly marked
unavailable, which the merge script already does automatically.

If the user hasn't asked about live/paid ads specifically (i.e. they only
want the organic collab timeline), it's fine to skip Method B entirely and
go straight from Step 1 to Step 3 — don't do unrequested work.

## Step 3 — Merge into the existing Excel

```bash
python scripts/merge_partnerships_to_excel.py <source.xlsx> <output.xlsx> \
  --map "SheetName1=scratch_brand1.json" \
  --map "SheetName2=scratch_brand2.json" \
  --ad-library-extra "SheetName1=extras_brand1.json"
```

This appends four columns to each mapped sheet, starting right after its
existing last column (never inserting or reordering existing columns):
`Went Live Date(s)`, `# Partnership Posts`, `Post URL(s)`, `Caption /
Transcript`. Matching is done by `Username` (case-insensitive, `@`
stripped) against the scan's partner list; creators with multiple posts get
all dates/URLs/captions stacked in one cell (`\n`-joined, wrapped, row
height auto-grown). Styling matches this project's established convention:
navy `#1F4E78` bold-white headers, Calibri font, thin light-gray borders,
hyperlinked URLs when a creator has exactly one post.

If any Ad Library-only extras were passed, they're appended as a clearly
separated, red-flagged block below each sheet's data rows — not mixed into
the main rows — with an explanatory merged note so it's obvious to anyone
opening the file later why those rows have blank URL/caption cells.

The script never overwrites the source file (writes to `<output.xlsx>`) —
confirm the output filename with the user before running, especially if
they want it saved over the original.

## A note on why the code is bundled as scripts, not just described

Both the organic-scan logic and the Excel-column-append logic were
hand-written from scratch several times across brands before this skill
existed, and each rewrite risked reintroducing already-solved bugs — a
`web_profile_info` 400 error on certain business accounts, an
`UnicodeEncodeError: surrogates not allowed` crash from stray UTF-16
surrogates in scraped captions, and mismatched header/username row
assumptions. Use the bundled scripts as-is rather than re-deriving this
logic inline; if a brand's data needs handling the scripts don't cover,
extend the scripts rather than writing one-off replacement code, so the
fix benefits the next run too.

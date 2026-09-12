# Method B: Meta Ad Library cross-check

This half of the workflow can't be a standalone script — it requires driving
a real browser (Meta blocks headless/API-style access), and it's used to
answer a *different* question than the organic scan:

- **Method A (organic scan)** finds creator collabs that are visible as
  regular Instagram posts. This covers most partnerships.
- **Method B (Ad Library)** finds creator collabs that the brand is
  *currently running as a paid Meta ad* — some of these are boosted
  versions of organic posts (already caught by Method A), but some are
  ad-only creator content that never appeared organically, or that appeared
  on the creator's page only, never the brand's. Method B is how you catch
  those.

Use Method B when the user asks whether partnerships are "live" or "active"
as paid ads specifically — not just still-visible on Instagram (which
Method A already answers on its own).

## When a creator ad is distinguishable from a plain brand ad

Search the Ad Library UI:

```
https://www.facebook.com/ads/library/?active_status=active&ad_type=all&country={cc}&q={brand}&search_type=keyword_unordered&media_type=all
```

Branded-content/creator ads show the advertiser field as
`"<Creator Name> with <Brand>"`. Plain brand ads just show the brand name
alone. Filter for the `with` pattern to isolate creator partnership ads
from the brand's own general advertising.

## The completeness trap — read this before reporting any numbers

**A single scroll-to-bottom pass, or even several, reliably UNDER-COUNTS —
often by roughly half.** The Ad Library UI lazy-loads results and hits a
false plateau: the page looks like it's stopped loading new ads when it
hasn't actually exhausted the result set. This was proven, not theorized —
cross-checking a scroll-based capture against real TrendTrack data for one
brand found 108 ads captured vs 208 actual, and 75 vs 88 for another metric
on the same brand. A second brand showed a much smaller gap. **The size of
the undercount is unpredictable per brand — you cannot assume "we scrolled
a lot" means "we got them all."**

Because of this, never report a Method B capture as "confirmed complete."
Always state how confident the number is (e.g. "captured via N accumulation
passes with 2 consecutive zero-growth confirmations" vs "single scroll
pass, likely undercounted") and prefer cross-checking against a real
source (TrendTrack or similar) when the user provides one or asks for
verification.

## Reliable capture procedure

1. Navigate to the Ad Library search URL above for the brand/country.
2. In the browser, run repeated `javascript_exec` calls that accumulate
   results into a **persistent global object** rather than trusting one
   scroll pass. State persists across separate `javascript_exec` calls as
   long as there's no page navigation in between — this is what makes the
   accumulation trick work.

```javascript
// First call: initialize the accumulator if it doesn't exist yet.
window.__adLibSeen = window.__adLibSeen || {};

// Each call: scroll to bottom repeatedly (~14 iterations, ~500ms apart)
// to force more lazy-loaded ads in, then harvest whatever is currently
// in the DOM into the accumulator keyed by Meta's own "Library ID" so
// re-harvesting never double-counts.
for (let i = 0; i < 14; i++) {
  window.scrollTo(0, document.body.scrollHeight);
  await new Promise(r => setTimeout(r, 500));
}

document.querySelectorAll('[role="article"], .x1dr75xp').forEach(card => {
  const text = card.innerText || "";
  const idMatch = text.match(/Library ID:\s*(\d+)/);
  if (!idMatch) return;
  const id = idMatch[1];
  if (!window.__adLibSeen[id]) {
    window.__adLibSeen[id] = {
      library_id: id,
      text: text.slice(0, 2000), // parse advertiser name / dates / caption from this
    };
  }
});

Object.keys(window.__adLibSeen).length; // report this count after each call
```

3. Repeat step 2 (re-run the same call) until **two consecutive calls add
   zero new ads** to `window.__adLibSeen`. Only then treat the capture as
   plateaued — and still label it "scroll-plateau capture," not
   "confirmed complete," unless cross-checked against an external source.
4. If a `javascript_exec` call times out mid-scroll (can happen on
   image/video-heavy pages), don't assume the data was lost — re-read
   `window.__adLibSeen.length` afterward; the accumulation frequently
   continued despite the tool-level timeout. Flag the result as
   partial/unconfirmed rather than discarding it.

## Extracting large result sets: the blob-download trick

Manually transcribing hundreds of accumulated ad records is slow and
error-prone. Instead, trigger a real file download and read it from disk:

```javascript
const data = window.__adLibSeen;
const blob = new Blob([JSON.stringify(data)], { type: "application/json" });
const url = URL.createObjectURL(blob);
const a = document.createElement("a");
a.href = url;
a.download = "adlib_export.json";
document.body.appendChild(a);
a.click();
a.remove();
```

The downloaded file lands in the user's actual Downloads/Desktop folder,
which the shell tool can read directly (browser and shell share the same
filesystem). Known quirks:
- The first read sometimes fails ("no such file") because the download
  hasn't flushed to disk yet — retry once after a short pause.
- Chrome throttles or silently drops repeated *automatic* downloads
  without an intervening user gesture — if a second export goes missing,
  re-trigger the download call.

## Turning a capture into `--ad-library-extra` input

Once you have accumulated ad records, diff them against the Method A
`unique` partner list for the same brand (see `scripts/scan_partnerships.py`
output). For any Ad Library creator name that has **no match** in the
organic scan's partner list, add it to the extras JSON consumed by
`scripts/merge_partnerships_to_excel.py`:

```json
[
  {
    "name": "Creator Name",
    "count": 3,
    "launch_date": "2026-01-10",
    "note": "found via Meta Ad Library search, not in organic scan"
  }
]
```

**Never invent a post URL or caption for these rows.** The Ad Library UI
doesn't expose the original Instagram post URL for most creator ads, and
guessing one is worse than leaving it blank — the merge script already
renders these cells as an explicit "N/A — not captured" rather than empty,
so there's no need to fill them with a plausible-looking fake.

## Matching names between the two sources

Ad Library shows a display name ("Priya Sharma"), the organic scan matches
on Instagram @handles. These often don't line up automatically — do a
manual/fuzzy pass (search the display name + brand on Instagram, or check
if it's an obvious variant of a handle already in the organic list) rather
than assuming a name that doesn't textually match is definitely a new
creator. When genuinely unsure, ask the user rather than guessing — a
wrong merge is harder to spot later than a question now.

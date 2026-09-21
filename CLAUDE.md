# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this is

A Google Apps Script project bound to a Google Spreadsheet that tracks Taylor Swift's Spotify
stream counts. Data is scraped from Apify, written into the spreadsheet, and turned into daily
stats, milestone logs, and Twitter-ready text summaries.

Prod spreadsheet: https://docs.google.com/spreadsheets/d/1ANVfUER8MInJ_y-oRp7Nguwk0yDI9LtBt36PWBMmZbU

## Environments

The script is **container-bound**, so an Apps Script project cannot be re-pointed at another
spreadsheet. A second environment is therefore a *copy of the spreadsheet*, which carries its
own forked copy of the script with a new script ID.

| | Config file | Script |
|---|---|---|
| prod | `.clasp.json` | bound to the live spreadsheet |
| dev | `.clasp.dev.json` | bound to a copy of it |

Both configs share `rootDir: src`, so the same source pushes to either target. clasp 3's
global `-P` flag selects which:

```
npm run push:dev      # clasp -P .clasp.dev.json push
npm run pull:dev
npm run open:dev
npm run push:prod     # clasp -P .clasp.json push
npm run pull:prod
npm run open:prod
```

There is deliberately no bare `npm run push` — the environment always has to be named.

`src/additional/env.js` works out at runtime which one it is in, by comparing
`SpreadsheetApp.getActiveSpreadsheet().getId()` against `CONFIG.ENV.PROD_SPREADSHEET_ID`.
Anything that is not the live sheet is dev, so a fresh copy needs no setup. A script property
`ENV` (`prod`/`dev`) overrides it. Consequences in dev:

- The spreadsheet menu is titled **Update [DEV]** and confirmation dialogs carry a `[DEV]` tag.
- `importSpotifyData()` refuses to run, because Apify credits are limited (~17 runs per token)
  and shared. Override with `CONFIG.ENV.ALLOW_APIFY_IN_DEV`.

The dev copy arrives with production's last raw import already in `Tools!F2:G1000`; that is the
fixture to develop against, so the whole pipeline can be exercised without touching Apify.
Script properties do **not** copy across, so the dev project has no Apify tokens at all —
which is the intended state.

Dev drifts from prod (prod gains an archive column per day), so refresh it by re-copying the
spreadsheet rather than trying to reconcile it.

## Commands

There is no build, lint, or test suite. `package.json` only pulls in
`@types/google-apps-script` for editor autocomplete. Code cannot be run locally — it only
executes inside Apps Script, triggered from the spreadsheet's **Update** menu.

`node --check` on the source files catches syntax errors before a push; beyond that, testing
means `npm run push:dev`, then running the menu item in the dev copy and inspecting the sheet.

Apps Script has no modules — every `.js` file under `src/` is concatenated into one global
namespace, so all top-level functions, `const CONFIG` and `var ApifyService` are globally
visible. There are no imports/exports; do not add any.

## Architecture

### Entry points

`onOpen()` in `src/update/main.js` builds the spreadsheet's **Update** menu. Every menu item
names a global function — renaming a function requires updating the menu string.

`main()` is the full daily pipeline and the usual path:

1. Guards on `Tools!C1` (sum of today's daily streams). If `<= 0`, Spotify hasn't refreshed
   yet and the run aborts — this guard exists because running twice would corrupt the archive.
2. `transferStats()` — advances the date in `Latest!Q1` by one day, **inserts a new column B**
   in the current Daily Archive sheet, and copies today's numbers in.
3. `updateStats()` → `transferBestSinceRows()` → `findBestSince()`.
4. `checkMilestones()`, `updateUpcomingMilestones()`.
5. `generateSummaries()`, `generateDiscographySummary()`.

Step 2 is destructive and non-idempotent (column insert + date increment). Never invoke it
speculatively.

### Data flow

```
Apify actor ──> Tools!F2:G1000 (raw name + streamCount)
                      │  (spreadsheet formulas do the matching/ordering)
                      ▼
              Tools!L2:M549 (total, daily)
                 ├──> Latest!F2:G549   (today's totals + dailies)
                 └──> Daily Archive!B2:B549, new column per day
                            │
                            ▼
              Latest!M (yesterday, Archive col C), Latest!N (week ago, Archive col I),
              Latest!L / Latest!Y ("best since" dates)
                            │
                            ▼
              Albums sheet text summaries, Milestone Log, Tracks!N45+ upcoming milestones
```

The Daily Archive is a **column-per-day** layout with the newest day in column B: column C is
yesterday, column I is a week ago. This is why `updateStats` reads fixed columns C and I, and
why inserting the new column must happen before anything else reads them.

### The row-layout contract

Rows are fixed across sheets and hardcoded in several places that must be changed together:

- Rows 2–549 is the song **range** (`CONFIG.SONGS`, `START_ROW: 2`, `COUNT: 548`). That 548 is
  a range size, **not** a song count: only rows 2–491 hold real tracks (490 of them) and
  rows 492–549 are empty headroom for future releases. Code that loops over `COUNT` relies on
  the `if (!songName) continue;` / `if (!title) continue;` guards to skip the blanks, so the
  padding is harmless — but never read 548 as "number of songs".
- Rows 550–574 = 25 album/category aggregate rows (`CONFIG.ALBUMS`), with row 574 the
  artist-wide total. These are sums of song-row blocks above, not data.
- `setAlbumFormulas()` in `src/update/daily_update.js` hardcodes the `=SUM(Bx:By)` ranges for
  those 25 aggregate rows. **Adding or reordering songs invalidates these sums** as well as
  `CONFIG.SONGS.COUNT`, `CONFIG.ALBUMS.START_ROW`, and every `songRow`/`count` in
  `CONFIG.STATS`.
- `CONFIG.STATS` maps each album to its song block in `Latest` (`songRow` + `count`), its
  aggregate row in `Latest` (`summaryRow`), and the cell in the `Albums` sheet where its text
  summary is written (`destCell`).

### Column conventions on `Latest`

Summary and milestone code reads columns by numeric index, so these meanings matter:
E = song name, F = total streams, G = daily, H = daily %, L = best-since date, M = yesterday,
N = a week ago. Album block starts at column T: T = name, U = total, V = daily, W = daily %,
Y = best since, Z = daily change, AA = weekly change. Album aggregate rows start at row 2,
with row 26 (`CONFIG.LATEST.OVERALL_ROW`) being the whole discography.

### `CONFIG` is the single source of truth

`src/additional/config.js` holds the environment IDs, all sheet names, A1 ranges, row counts,
the per-album `STATS` table, and the Apify actor ID + album URL map. Range changes belong here,
not inline —
the only significant exceptions are the hardcoded sums in `setAlbumFormulas()` and the numeric
column indices described above. Adding a new release means adding its Spotify album URL to
`CONFIG.APIFY.ALBUMS` *and* fixing up the row layout above.

### Apify integration

`src/apify/apify_service.js` is a thin API wrapper (start actor, poll status every 5s for up
to 5 minutes, fetch dataset). `src/apify/apify_import.js` builds the input from
`CONFIG.APIFY.ALBUMS`, flattens `album.tracks` into `[name, streamCount]` pairs, and writes
them to `Tools`.

Tokens are **never** stored in the repo. `src/apify/token_manager.js` keeps three tokens in
`PropertiesService.getScriptProperties()` under `APIFY_TOKEN_1..3`, with
`ACTIVE_TOKEN_INDEX` and per-token `RUN_COUNT_n`. Each token allows roughly 17 runs before
its free credits are exhausted; `logRunAndCheckLimits()` warns at 15 and 18, and
`switchApifyToken()` cycles 1→2→3→1. `initializeTokensOneTime()` is commented out on purpose
— it is uncommented, filled in, run once, then re-commented.

### Summaries

`src/update/summary.js` opens with a long comment block specifying the exact text format
(header, biggest gainer / most stable, best-since line, closing delta line). Treat that spec
as the requirements doc — change it alongside the code. A "best since" date is only reported
when it is at least 7 days old. Percentages are taken from `getDisplayValues()` so the sheet's
own formatting is reused rather than re-derived.

### Discontinued code

`src/website_discontinued/` was a public web app (`doGet` + `index.html`) and is no longer
maintained. Its hardcoded row/column assumptions (e.g. `numRows = 483`) are stale and it is
deliberately not kept in sync with `CONFIG`. Don't update it unless asked.

## Repo notes

- `node_modules/` and `*.xlsx` exports are gitignored; the spreadsheet exports are 35MB+ and
  must not be committed.
- Source comments are in English with occasional Ukrainian; both are fine.
- The "Description of the sheet from Gemini" section below is a useful map of what each sheet
  column holds, but its counts are approximate (it says ~726 tracks / ~577 rows where `CONFIG`
  uses 548 song rows). Where they disagree, `CONFIG` and the code win.

## Sheet reference

What each sheet holds, and which parts of it the script reads or writes. Anything marked
**[sheet-side]** is maintained by spreadsheet formulas, not by this codebase — the script never
writes it. Row/column facts below were reconciled against `CONFIG` and `setAlbumFormulas()`;
where a description and the code disagree, the code wins.

### The song-row blocks

Rows 2–549 mean the same thing in `Tools!J:M`, `Latest` (left table), and every `Daily Archive`
sheet, so one song sits on one row everywhere. The blocks come from the `=SUM()` formulas in
`setAlbumFormulas()`, which is the only authoritative listing of them:

| Song rows | # | Category | Aggregate row | `Latest` album row |
| --- | --- | --- | --- | --- |
| 2–16 | 15 | Taylor Swift (Debut) | 550 | 2 |
| 17–35 | 19 | Fearless (2008) | 551 | 3 |
| 36–56 | 21 | Speak Now (2010) | 552 | 4 |
| 57–78 | 22 | Red (2012) | 553 | 5 |
| 79–97 | 19 | 1989 (2014) | 554 | 6 |
| 98–112 | 15 | reputation | 555 | 7 |
| 113–130 | 18 | Lover | 556 | 8 |
| 131–164 | 34 | folklore | 557 | 9 |
| 165–181 | 17 | evermore | 558 | 10 |
| 182–207 | 26 | Fearless (Taylor's Version) | 561 | 13 |
| 208–237 | 30 | Red (Taylor's Version) | 562 | 14 |
| 238–260 | 23 | Midnights | 559 | 11 |
| 261–282 | 22 | Speak Now (Taylor's Version) | 563 | 15 |
| 283–304 | 22 | 1989 (Taylor's Version) | 564 | 16 |
| 305–335 | 31 | The Tortured Poets Department | 560 | 12 |
| 336–354 | 19 | The Life of a Showgirl | 565 | 17 |
| 355–362 | 8 | Live From Clear Channel Stripped 2008 | 568 | 20 |
| 363–368 | 6 | The Taylor Swift Holiday Collection | 567 | 19 |
| 369–376 | 8 | Live From Paris | 570 | 22 |
| 377–392 | 16 | Speak Now World Tour Live | 569 | 21 |
| **393–404** | **12** | **The Life of a Showgirl — Track by Track ⚠ not summed by any category** | — | — |
| 405–416 | 12 | Soundtracks | 571 | 23 |
| 417–428 | 12 | Features | 573 | 25 |
| 429–435 | 7 | Droplets | 566 | 18 |
| 436–518 | 83 | Remixes and etc. | 572 | 24 |
| 519–549 | 31 | empty headroom | — | — |
| | | **Total artist streams** = `SUM(B2:B549)` | 574 | 26 |

Note that the aggregate rows are **not** in song-block order — `Midnights` and `TTPD` sit at
rows 559/560 while their song blocks fall after the Taylor's Versions. Use the table rather
than assuming the orders line up.

Real tracks occupy rows 2–491 (490 songs); 492–549 are blank. That total only reconciles if
rows 393–404 are populated, which is how we know that block holds 12 real tracks.

**Adding new tracks** currently means appending at the next free row (492, then 493, …), which
puts them outside whichever category block they belong to and leaves them summed only by the
row 574 artist total — the same trap rows 393–404 fell into. Making new releases land in the
right block is known outstanding work; do not design around the append-at-the-end behaviour as
if it were intended.

### `Tools` — ingestion and normalisation **[sheet-side, except the raw dump]**

| Range | Contents |
| --- | --- |
| `C1` | `SUM_OF_DAILYS` — the freshness guard `main()` aborts on when `<= 0` |
| `F2:G1000` | **Raw Apify dump**, written by `importSpotifyData()`. Cleared and rewritten each import |
| `J2:M491` | **Mapping block** — the 490 tracked songs. J = album/era, K = title, L = total, M = daily |

The raw dump holds **more tracks than are tracked** (roughly 700+, varying per import). Apify
scrapes whole albums, and several albums in `CONFIG.APIFY.ALBUMS` (compilations, soundtracks,
other artists' records carrying a Swift feature) contain tracks she is not credited on. The
mapping block is what filters that down: sheet formulas look each tracked title up in `F:G` and
pull its total into `L`. Extra rows in the dump are simply never looked up.

The script only ever *writes* `F2:G1000` and *reads* `L2:M549` — it never touches the matching
logic in between. The import warns if the dump would exceed 998 rows.

### `Latest` — current-day snapshot

Two side-by-side tables. Numeric column indices matter, because the summary and milestone code
reads by number, not by header.

Left table, one row per song (rows 2–549):

| Col | # | Contents |
| --- | --- | --- |
| A | 1 | Album/era shorthand (`debutOG`, `fearlessTV`, …) |
| B / C | 2 / 3 | Rank, rank change vs. yesterday |
| D | 4 | Cover art placeholder |
| E | 5 | Title — `CONFIG.LATEST.SONG_NAMES`; a blank here means "skip this row" |
| F / G | 6 / 7 | Total streams, daily streams — **written** by `transferStats()` |
| H / I | 8 / 9 | % change vs. yesterday / vs. a week ago |
| J / K | 10 / 11 | Absolute change vs. yesterday / vs. a week ago |
| L | 12 | Best-since date — **written** by `findBestSince()` |
| M / N | 13 / 14 | Daily a day ago / a week ago — **written** by `updateStats()` |
| O | 15 | Previous day's rank |
| P | 16 | Cover art URL |

Right table, one row per album/category (rows 2–25, overall total at row 26):

| Col | # | Contents |
| --- | --- | --- |
| T | 20 | Album/category title |
| U / V | 21 / 22 | Total, daily |
| W / X | 23 / 24 | % change vs. yesterday / vs. a week ago |
| Y | 25 | Best-since date — **written** by `findBestSince()` |
| Z / AA | 26 / 27 | Absolute change vs. yesterday / vs. a week ago |
| AB | 28 | Cover art URL |

`Q1` (`DATE_CELL`) holds the update date and is incremented by `transferStats()`.

### `Daily Archive 2023`–`2026` — the time series

Column-per-day, **newest first**: `transferStats()` does `insertColumnAfter(1)`, so today lands
in column B, yesterday is C, and a week ago is I. Row 1 holds the dates, column A the titles.
Rows follow the song-row blocks above. `findBestSince()` walks the years newest-to-oldest via
`CONFIG.SHEETS.ARCHIVE_YEARS`, skipping column B on the current year (that is today's value,
the threshold being beaten) and including it on past years (that is Dec 31st).

### Total archive — manual cold backup **[not touched by the script]**

Structurally a Daily Archive, but holding **cumulative totals** instead of daily counts, and
only two columns per month: the first and last day. It runs back to 2022-12-31. Song rows match
the standard block layout, so a row means the same track here as everywhere else.

It is maintained by hand — the owner pastes the totals column across on those dates — as a
second copy of the data and so that the total for any given day can be reconstructed if the
Daily Archives are ever damaged. No code reads or writes it, and it is absent from `CONFIG`.
Leave it alone unless asked; if it ever is automated, note that its columns are sparse and
non-contiguous, so none of the fixed-offset tricks the Daily Archive relies on (column C is
yesterday, column I is a week ago) apply.

### `Tracks` — sortable leaderboard **[sheet-side]** + one written block

Sorting controls in A–B drive a formula-sorted leaderboard in C–M. The script's only stake is
`N45:Q549` (`CONFIG.TRACKS`), cleared and rewritten by `updateUpcomingMilestones()`: headers on
row 45 (`N45:O45` merged), then one row per predicted milestone with title (N), milestone
value (P) and the generated announcement sentence (Q).

### `Albums` — era analysis **[sheet-side]** + the generated summaries

Summary tables and per-era track breakdowns are laid out in column groups, with each era's
generated text summary written into the column just right of its breakdown. Destinations are
`destCell` in `CONFIG.STATS` (F23, F45, F79, L23, L54, L93, L119, R23, R49, R78, R99, W23,
W51, W78, W102, AB23, AB89) plus `G5` for the whole-discography summary
(`CONFIG.ALBUMS.TOTAL_SUMMARY`). Each write clears 10 rows first, so summaries must stay at
least 10 rows apart.

## Known inconsistencies

Found by reconciling the code against the layout; none are currently breaking anything, but
they will bite whoever next edits these areas.

1. **Rows 393–404 fall through every category sum.** These are the 12 *Track by Track* versions
   of The Life of a Showgirl, scraped from the `"The Life of a Showgirl (Track by Track
   Version)"` entry in `CONFIG.APIFY.ALBUMS`. The 24 category formulas cover rows 2–392 and
   405–518, so these 12 are in none of them — yet they are included in the `SUM(B2:B549)`
   artist total on row 574. The 24 category rows therefore do not add up to row 574. The
   omission was not deliberate, but **the decision (2026-09-22) is to leave it as it is** —
   the tracks are low-value and not worth restructuring the aggregates for. Treat this as a
   known and accepted gap, not a bug to fix; do not "correct" it without being asked.

   Note this scatters the Showgirl era across three blocks: 336–347 (standard, the only part
   `CONFIG.STATS` scans for summaries), 348–354 (the Acoustic Collection tail) and 393–404
   (Track by Track). Only the first two are in the era's aggregate on row 565.

2. **`CONFIG.STATS.count` is narrower than the sum block for four albums.** Speak Now (17 vs
   21), Red (19 vs 22), 1989 (16 vs 19) and The Life of a Showgirl (12 vs 19). These are the
   deluxe/bonus-track eras, so it looks intentional — the summary scans a narrower tracklist
   than the album total sums — but the effect is that tracks in the excluded tail can never be
   named as an era's biggest gainer. Which tracks those are has not been confirmed.

3. ~~**`generateDiscographySummary()` derives a row span from an entry count.**~~ **Fixed
   2026-09-22.** It used `Object.keys(CONFIG.STATS).length` (17) as the number of album rows to
   read, which scanned `Latest` rows 2–18 and so let Droplets (row 18) be named as the
   discography's biggest gainer. Worse, the coupling was accidental: adding any entry to
   `CONFIG.STATS` would silently widen the scan by a row. It now reads
   `CONFIG.LATEST.DISCOGRAPHY_ALBUMS_COUNT` (16), covering rows 2–17 — the studio albums only.

   Provenance, because the symptom is misleading: with the 16 entries `CONFIG.STATS` held
   before 2026-07-02 the scan was rows 2–17 and happened to be exactly right. Commit `d6ed8c5`
   added the `Soundtracks` entry, widening it to row 18. So adding Soundtracks *is* what broke
   the discography summary — but it never injected Soundtracks itself (row 23 was never in
   range); it pulled in Droplets at row 18. `generateDiscographySummary()` reads titles from
   column T of the sheet and only ever used `CONFIG.STATS` for its `.length`.

   The discography summary is deliberately about **albums**, so it must never reach Droplets
   (18), the live/compilation rows (19–22), Soundtracks (23), Remixes (24) or Features (25).
   Those categories still get their own per-era summaries where they have a `CONFIG.STATS`
   entry (Soundtracks does, writing to `R99`); that is separate and intended.

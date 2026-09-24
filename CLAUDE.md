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

The dev copy arrives with production's last raw import already in `Import!E2:I1000`; that is the
fixture to develop against, so the whole pipeline can be exercised without touching Apify.
A dump written before track-ID matching (2026-09-22) has no track IDs, and `readTodayFromImport()`
refuses it. Paste a dump that has IDs at `Import!E2` instead. Right after a real update the dump
equals Latest, so every daily is 0 and Update Daily Stats refuses to run; to try it in dev, raise
some counts in `Import!G` by hand first.
Script properties do **not** copy across, so the dev project has no Apify tokens at all —
which is the intended state.

Dev drifts from prod (prod gains an archive column per day), so refresh it by re-copying the
spreadsheet rather than trying to reconcile it.

## Commands

There is no build or lint. `package.json` pulls in `@types/google-apps-script` for editor
autocomplete, and has one script besides the clasp ones:

```
npm test              # node --test tests/*.test.js, offline, Node 22+
```

`tests/lib/fake_sheets.js` is an in-memory stand-in for `SpreadsheetApp`: values, formulas and
display text per cell; simple formulas (`SUM`, references, arithmetic) evaluated on read; and
formula references re-pointed on row inserts, deletes, `moveTo` and `moveRows` the way Sheets
does it, so row-moving code can be tested. `tests/lib/fixture.js` builds a spreadsheet from
`tests/fixtures` (Latest, the raw import, the Tracklist and Covers from 2026-09-22; synthetic
archives) and loads every `src/` file into one `vm` context like Apps Script. Each `f.run(expr)`
is a fresh execution, like a menu click: the files run again from the top and caches start empty.
`f.nextDayImport()` raises the raw counts so an update has something to do. Add a test beside
any change to the pipeline or the row-moving code. The fake doesn't model formatting, charts,
data validation or the dialog's browser side, so a change still gets tried in the dev copy:
`npm run push:dev`, then the menu item, then look at the sheet.

Apps Script has no modules — every `.js` file under `src/` is concatenated into one global
namespace, so all top-level functions, `const CONFIG` and `var ApifyService` are globally
visible. There are no imports/exports; do not add any.

## Architecture

### Entry points

`onOpen()` in `src/update/main.js` builds the spreadsheet's **Update** menu. Every menu item
names a global function — renaming a function requires updating the menu string.

`main()` is the full daily pipeline and the usual path:

1. Guards, all before anything is written: a Tracklist or Categories sheet that fails
   validation (`buildAlbumFormulas()`); then `readTodayFromImport()` works out every song's
   total (by track ID) and daily (total − `Latest!F`), and the run stops if an active song's ID
   is missing from the import, or if the dailies add up to `<= 0`. The latter means Spotify
   hasn't refreshed yet, and is also what stops a second run on the same import (Latest then
   already holds those totals) — running twice would corrupt the archive.
2. `transferStats(today)` — advances the date in `Latest!Q1` by one day, **inserts a new column B**
   in the current Daily Archive sheet, and writes today's dailies there and the totals and
   dailies into `Latest!F:G`.
3. `updateStats()` → `transferBestSinceRows()` → `findBestSince()`.
4. `checkMilestones()`, `updateUpcomingMilestones()`.
5. `generateSummaries()`, `generateDiscographySummary()`.

Step 2 is destructive and non-idempotent (column insert + date increment). Never invoke it
speculatively.

### Data flow

```
Apify actor ──> Import!E2:I1000 (album, raw name, streamCount, track ID, album cover)
                      │  readTodayFromImport(): Tracklist track ID -> row;
                      │  daily = total − Latest!F (nothing is stored per row in Import)
                      ▼
              today = { totals[], dailies[] }, rows 50+
                 ├──> Latest!F:G, rows 50+    (today's totals + dailies)
                 └──> Daily Archive!B, rows 50+, new column per day
                      (+ generated aggregate formulas in B2:B27)
                            │
                            ▼
              Latest!M (yesterday, Archive col C), Latest!N (week ago, Archive col I),
              Latest!L ("best since" dates) - rows 2 to the last song, albums included
                            │
                            ▼
              Albums sheet text summaries, Milestone Log, Tracks!N45+ upcoming milestones
```

The Daily Archive is a **column-per-day** layout with the newest day in column B: column C is
yesterday, column I is a week ago. This is why `updateStats` reads fixed columns C and I, and
why inserting the new column must happen before anything else reads them.

### The row-layout contract

**Layout since phase 1b (2026-09-22).** Both prod and dev were migrated from the old layout
(songs at 2–549, archive aggregates at 550–574, Latest's album table in T:AB) by a one-off script,
since removed; it's in git history (commit `30546b2`, `src/migration/migrate_layout.js`) if ever
needed. The code assumes the new layout. A row means the same thing in Latest, every Daily
Archive and the Total Archive (Tools J:M was part of this until 2.1; the Import sheet isn't):

| Rows | Holds | Config |
| --- | --- | --- |
| 1 | Headers / dates | |
| 2 | Total Artist Streams (`=SUM` over every song row) | `LAYOUT.TOTAL_ROW` |
| 3 | Total Artist Solo Streams = row 2 − Features | `LAYOUT.SOLO_ROW`, `SOLO_EXCLUDES` |
| 4–27 | The 24 categories | `row` in the Categories sheet |
| 28–48 | Spare rows for new categories (hidden in Latest from 36) | `LAYOUT.LAST_AGGREGATE_ROW` |
| 49 | The songs' header row in Latest (merged A:O), blank in the other sheets. Nothing is written to it: updates write the aggregate rows and the song rows as two blocks (`layoutBlocks()`) | |
| 50 → | Songs, open-ended | `LAYOUT.FIRST_SONG_ROW` |

- **There is no fixed song count.** The last song row is the highest Tracklist `row`
  (`getLastSongRow()`); ranges are built from `FIRST_SONG_ROW` and `getSongRowCount()`.
  Never hardcode an A1 range over the songs.
- **Which rows belong to which category comes from the `Tracklist` sheet** (see
  `src/additional/songs.js`), not from code. `buildAlbumFormulas(column)` turns its `category`
  column into the aggregate formulas for rows 2–27 of one archive column, and
  `generateSummaries()` takes each album's song rows from it.
- **The Categories sheet** (`getCategories()` in `src/additional/categories.js`) maps each
  category name, exactly as written in the Tracklist, to its `row` (the same in Latest and every
  archive), its `type`, and optionally the Albums cell for its text summary (`summaryCell`) and
  `summaryLimit` (see Known inconsistency 2). An unknown category in the Tracklist is an error.
  It replaced `CONFIG.CATEGORIES` in 2.0.0.3.
- **Latest's album F/G (rows 2–27) are generated by code** (`setLatestAggregateFormulas()`): the
  same `buildAlbumFormulas()` sums as the archives, written for columns F and G on every update
  and after songs are added. Never hand-edit them; they are overwritten. H–L on those rows are
  still sheet formulas on the same row.
- **Latest** also keeps 8 special-edition totals (tlpss, folklore standard, … showgirl era) as a
  small sheet-side table in **T:AB**, beside the album rows. Only the Albums sheet reads them; the
  code never does, and inserts never move them (they shift A:P only).
- **Merged cells must not cross A:P in Latest.** Only Latest shifts part of a row when a row is
  inserted (so the side table beside the songs stays put), and Sheets refuses to shift part of a
  merged cell. A merge inside A:P is fine (the songs' header row 49). `mergesInTheWay()` checks it
  before anything is changed, and the Merged cells check reports it. The archives and the
  Total Archive take whole rows, so merges there don't matter.
- **New songs go in through the Pending sheet** and **Update → Add Pending Songs**
  (`src/songs/add_songs.js`). Never insert song rows by hand. It inserts whole rows in Latest,
  every Daily Archive and the Total Archive, then adds 1 to every Tracklist `row` at or below the
  insert.
- **Songs are moved with Update → Move a Song** (`src/songs/move_song.js`, dialog HTML in
  `move_song_dialog.js`). Never move song rows by hand. It moves Latest A:P (`insertCells` at
  the target, `moveTo`, `deleteCells` at the old row) and whole rows in the archives
  (`Sheet.moveRows`), so references to the song follow it, renumbers the Tracklist rows in
  between, and can change the category in the same step, then calls `rebuildAllAggregates()`.
- **Past archive columns hold the album figures (rows 2–27) as plain numbers, on purpose**: 365 × 26
  `SUM`s per sheet made the sheet lag. Only today's column B has formulas (written by the update);
  the owner converts older ones to values. So history can't be recalculated from the songs.
  Category history is **retroactive** by decision (2026-09-24), done as a **delta**:
  `rebuildAllAggregates()` finds the songs whose Tracklist `category` differs from their
  `historyCategory` (the category the history counts them under), and in every dated column of
  every archive subtracts the song's streams that day from the old category's figure and adds them
  to the new one's (the solo row moves the opposite way when Features is involved). Every other
  number stays exactly as it was; formula cells get today's `buildAlbumFormulas()`. It then sets
  `historyCategory` to match (creating the column if missing), rewrites Latest's album formulas,
  and runs `updateStats()` so Latest's M/N/L re-read the adjusted history. **Update → Rebuild
  Album History** runs it alone, for after a category is changed by hand in the Tracklist; the
  **Category history** check warns until it has run. Never set `historyCategory` by hand except to
  tell the rebuild about a change made before the column existed.

### Column conventions on `Latest`

Albums and songs share columns (`CONFIG.LATEST.COLS`): E = title, F = total streams, G = daily,
H = daily %, I = weekly %, J = daily change, K = weekly change, L = best-since date, M = daily a
day ago, N = a week ago, P = cover URL. Summary and milestone code reads these by column number.
The whole-discography row is 2 (`LAYOUT.TOTAL_ROW`); the discography summary scans the
categories of type `studio` (16 today, rows 4–19), wherever they are.

### `CONFIG` is the single source of truth

`src/additional/config.js` holds the environment IDs, all sheet names, A1 ranges, row counts,
and the Apify actor ID + album URL map. Categories live in the Categories sheet, not here. Range changes belong here,
not inline. The
per-song facts (category, status, track ID) live in the `Tracklist` sheet instead. Adding a new
release means adding its Spotify album URL to `CONFIG.APIFY.ALBUMS`, its songs to the Tracklist
sheet with their track IDs, and the matching rows in the other sheets.

### Apify integration

`src/apify/apify_service.js` is a thin API wrapper (start actor, poll status every 5s for up
to 5 minutes, fetch dataset). `src/apify/apify_import.js` builds the input from
`CONFIG.APIFY.ALBUMS`, flattens `album.tracks` into `[album name, name, streamCount, id]`
rows, writes them to `Import!E:I` (album first, for readability), then calls `matchTotalsById()` (`src/apify/match_totals.js`).
The dataset is fetched unfiltered, so every field the actor returns (`coverArt`, album
`artists`, `_url`, …) is available if needed.

`readTodayFromImport()` (read-only) looks each song's total up by its Tracklist track ID and
works its daily out against `Latest!F`. Retired songs get 0, and upcoming songs get 0 until their ID
appears. An active song whose ID is absent gets no figure and is listed in `missing`, which makes
`main()` abort — deliberate, since a silent 0 would corrupt the archive. `matchTotalsById()` is
the same plus writing the sum of dailies into `Import!C1` for reference; the **Check Import**
menu item runs it for free after fixing an ID. `getImportSheet()` also accepts the sheet's
pre-2.1 name `Tools` until **Setup → Tidy up the Import sheet** renames it and clears J:M.
Track IDs survive Spotify renaming a track (the old title-based `XLOOKUP` did not). One
recording on several albums has several IDs sharing one stream count; each song stores the ID
from its own album.

Tokens are **never** stored in the repo. `src/apify/token_manager.js` keeps three tokens in
`PropertiesService.getScriptProperties()` under `APIFY_TOKEN_1..3`, with
`ACTIVE_TOKEN_INDEX` and per-token `RUN_COUNT_n`. Each token allows roughly 17 runs before
its free credits are exhausted; `logRunAndCheckLimits()` warns at 15 and 18, and
`switchApifyToken()` cycles 1→2→3→1. `initializeTokensOneTime()` is commented out on purpose
— it is uncommented, filled in, run once, then re-commented.

### Health checks

`src/checks/checks.js` builds a second **Checks** menu (**Checks [DEV]** in dev). Every check is
**read-only**: it reports and never repairs. Each takes the shared loader from
`loadCheckData()`, which reads each range once per run, and returns
`{ status: 'ok' | 'warn' | 'fail', summary, details[] }`. The checks in `HEALTH_CHECKS` are:
Tracklist validity, row alignment (titles in Tracklist = `Latest!E` = column A of every Daily
Archive and the Total archive), Latest album-table totals/dailies against the sum of their
songs, today's aggregate formulas against `buildAlbumFormulas()`, import health, new tracks,
sources, merged cells, Latest vs Import totals, archive date sequence, covers, spare category
rows, the Categories sheet (including the studio → other → fixed row order), and category
blocks (warns when a category's songs aren't on consecutive rows).

`main()` ends with `runChecksAfterUpdate()` (the `AFTER_UPDATE_CHECKS` subset). It shows a
dialog only if something isn't ✅ and never throws. Checks that depend on the Tracklist report
"can't read" rather than listing every row when the Tracklist is unreadable. A new check goes
in `HEALTH_CHECKS`; keep it read-only.

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
writes it. Row/column facts below were reconciled against `CONFIG` and the Tracklist sheet;
where a description and the code disagree, the code wins.

### `Tracklist` — the song registry

One row per song. It is a **list, not a grid**: each song's `row` value is the row it lives on
in every other sheet, so the Tracklist's own order doesn't matter and it may be sorted. In code it is `CONFIG.SHEETS.SONGS` (read by `getSongs()` in `src/additional/songs.js`),
so renaming the tab needs only that one string changed. The code finds columns by header, ignoring case and spaces (`CONFIG.SONGS_SHEET.HEADERS`),
so headers may be renamed as long as the words stay the same, and columns may be reordered.
Extra columns are ignored.

| Header | Read by code | Contents |
| --- | --- | --- |
| `row` | yes | The row the song lives on everywhere else (50 or below, used once) |
| `status` | yes | `active`; `upcoming` (added before release: ID optional, held at 0 until its ID is in an import, never stops the update); or `retired` (kept only for its archive history) |
| `category` | yes | Aggregate it sums into; must be a name in the Categories sheet. Blank = none |
| `coverKey` | checks only | Key into the Covers sheet (same value as `Latest!A`) |
| `title` | yes | Display title, same as `Latest!E`. A row with no title is ignored |
| `Spotify Title` | no | Spotify's own name for the track, for humans |
| `trackId` | yes | Spotify track ID — what `matchTotalsById()` matches on |
| `sourceAlbum` | no | Which scraped album the ID came from |
| `historyCategory` | yes, optional | The category the archive history counts the song under. Written by the code (Rebuild Album History, Move a Song, Add Pending Songs); differs from `category` only until the history has been adjusted |

**Retired songs** are *Our Song (International Mix)* and *Love Story (Pop Mix)*: rows 486 and 488 since
*Patient Zero* was added at 484 (485 and 487 before that, 437 and 439 before phase 1b; rows shift, so go by the Tracklist):
Spotify merged their counts into the originals around August 2026. They are held at 0, left out
of milestones and summaries, and should be hidden from Tracks. The rows remain because the
archives hold ~3 years of their history. Retire a song rather than deleting its row.

**Display titles differ from Spotify's on purpose** ("TV" for Taylor's Version etc.). Spotify's
names exist only in the raw import and the `Spotify Title` column; never "correct" `title` to match them.

### `Categories` — one row per album/category

Created once by **Update → Setup → Create Categories sheet** (seeded from `CATEGORY_SEED` in
`src/setup/setup_sheets.js`, which is what `CONFIG.CATEGORIES` held). Headers are matched like
the Tracklist's (`CONFIG.CATEGORIES_SHEET.HEADERS`); the sheet's own order doesn't matter.

| Header | Contents |
| --- | --- |
| `name` | As used in the Tracklist's `category` column |
| `row` | Its row in Latest and every archive: 4 up to `LAYOUT.LAST_AGGREGATE_ROW`, used once |
| `type` | `studio` (albums and re-recordings: in the discography summary), `other` (compilations, live, Droplets), `fixed` (Soundtracks, Remixes, Features) |
| `summaryCell` | Albums-sheet cell for its text summary; blank = no summary |
| `summaryLimit` | Summary considers only the first N songs; blank = all |

A row with a name but **no `row`** is one waiting to be added: **Update → 🗂️ Add New Categories**
(`src/songs/add_category.js`) gives it a row after the last category of its own type (studio after
the last studio album, other after the last compilation, fixed at the end), opens that row in Latest
(A:P cells only, so the T:AB side table stays put), every Daily Archive and the Total Archive, and
deletes one spare row just above the songs so the songs never move. It renumbers the Categories
sheet, rebuilds today's archive column, Latest's album formulas and the Total Archive's, then runs
the checks. It refuses unless the rows line up and the spare rows really are empty. What it can't
do: widen the Albums sheet's studio-albums chart, place the summary, or add the album's URL.

Rows must run studio, then other, then fixed; the Categories check enforces it. A new category
(phase 5, not built yet) is meant to go after the last of its type.

### `Pending` — songs waiting to be added

Created once by **Update → Setup → Create Pending sheet**. One row per song: `title`, `category`
(dropdown from Categories), `coverKey`, `trackId`, `status` (`active` / `upcoming` / `ignore`; blank
means active if there's an ID, else upcoming; `ignore` needs only the ID and sends it to Ignored), optional `Spotify Title` and `album` (filled in from the
import by ID when blank), and `result`. **Add Pending Songs** validates the whole batch first
(category exists, active songs need an ID that is in the import, no ID or title already tracked,
rows aligned) and changes nothing if any row fails, writing each row's problems into `result`.
Otherwise it adds each song after the last song of its category (a category with no songs yet goes
after the nearest category above it), runs Check Import, regenerates Latest's album formulas,
re-checks alignment and totals, and marks each row `✅ Added at row N`. Rows marked ✅ are skipped
from then on. What it can't do: the Albums-sheet breakdowns, and Covers entries for new keys (it
lists both).

### `Ignored` — track IDs deliberately not tracked

Created once by **Update → Setup → Create Ignored sheet**, which starts it with every track in the
current import that isn't tracked (other artists on soundtracks, other editions). **Create it
before adding a new album's URL**, or that album's tracks are ignored along with the rest.
Columns: `trackId`, `Spotify Title`, `album`, `reason`, `date`. Rows are only ever appended
(by Find New Tracks for same-recording duplicates, and by Add Pending Songs for `ignore` rows).
To start tracking an ignored track, delete its row here and add it through Pending.

### `Sources` — the album URLs to scrape

Created once by **Update → Setup → Create Sources sheet**, seeded from `CONFIG.APIFY.ALBUMS`
(which stays in the code only as that seed). Columns `name` and `url`; the importer reads it
through `readSources()` (`src/apify/sources.js`), so **a new release needs a row here, not a code
change**. Only `open.spotify.com/album/…` links are accepted, and the same album twice is refused.
Until the sheet exists the config map is used and the Sources check warns.

### Finding new tracks

`findNewTracks()` (`src/songs/detect_new.js`) runs at the end of every import and from
**Update → Find New Tracks** (reads the import already in the Import sheet; no Apify cost). Every ID in the
raw import that is in none of Tracklist, Pending and Ignored is sorted:
- **Same recording on another edition**: same stream count (> 0) and a matching name as a tracked
  song. It goes to Ignored with the reason. This is what a new deluxe edition's standard tracks
  look like; so do mixes whose counts Spotify merged into the original.
- **An announced song now out**: its name matches exactly one `upcoming` Tracklist song without an
  ID. The ID is written into the Tracklist, and the song counts from that import. Names are compared
  normalised: case, punctuation, "(Taylor's Version)", "(From The Vault)", "(TV)" and "(FTV)" are
  ignored.
- **Anything else**: a new Pending row (`status` active) with category and cover key suggested
  from the most common among tracked songs on the same album (blank for an unknown album).

It never touches Latest or the archives; adding songs stays with Add Pending Songs. The
**New tracks** health check warns while the import holds IDs that are unaccounted for.

### Covers for new cover keys

The Covers sheet has no header row: key in column A, the image formula (`=IMAGE` of the URL) in B,
the image URL in C (`CONFIG.COVERS`). Latest!D shows `XLOOKUP(key, Covers!A:A, Covers!B:B)`. `ensureCovers()` (`src/songs/covers.js`) adds a row for
any key that isn't there yet, straight under the last key, copying column B's formula from the
row above. It uses the cover of the album the
track was scraped from (Import!I), or else Spotify's oEmbed endpoint for the track
(`open.spotify.com/oembed`, no login, no credits). It runs after Add Pending Songs and when
Find New Tracks links an upcoming song, and **Update → Fill Missing Covers** runs it for the whole
Tracklist. Existing keys are never changed; to use a different image (e.g. the standard green
TLOAS rather than the scraped Acoustic edition), edit the URL in Covers by hand. The Covers check
warns if a key in use has no image formula in Covers!B.

### The song-row blocks

The song rows mean the same thing in `Latest`, every `Daily Archive` and the
Total Archive, so one song sits on one row everywhere. The blocks are defined by the Tracklist's
`category` column; as of 2026-09-22 (after phase 1b) they are:

| Song rows | # | Category | Aggregate row |
| --- | --- | --- | --- |
| 50–64 | 15 | Taylor Swift (Debut) | 4 |
| 65–83 | 19 | Fearless (2008) | 5 |
| 84–104 | 21 | Speak Now (2010) | 6 |
| 105–126 | 22 | Red (2012) | 7 |
| 127–145 | 19 | 1989 (2014) | 8 |
| 146–160 | 15 | reputation | 9 |
| 161–178 | 18 | Lover | 10 |
| 179–212 | 34 | folklore | 11 |
| 213–229 | 17 | evermore | 12 |
| 230–255 | 26 | Fearless (Taylor's Version) | 15 |
| 256–285 | 30 | Red (Taylor's Version) | 16 |
| 286–308 | 23 | Midnights | 13 |
| 309–330 | 22 | Speak Now (Taylor's Version) | 17 |
| 331–352 | 22 | 1989 (Taylor's Version) | 18 |
| 353–383 | 31 | The Tortured Poets Department | 14 |
| 384–402 | 19 | The Life of a Showgirl | 19 |
| 403–410 | 8 | Live From Clear Channel Stripped 2008 | 22 |
| 411–416 | 6 | The Taylor Swift Holiday Collection | 21 |
| 417–424 | 8 | Live From Paris | 24 |
| 425–440 | 16 | Speak Now World Tour Live | 23 |
| **441–452** | **12** | **The Life of a Showgirl — Track by Track ⚠ not summed by any category** | — |
| 453–464 | 12 | Soundtracks | 25 |
| 465–476 | 12 | Features | 27 |
| 477–483 | 7 | Droplets | 20 |
| 484–539 | 56 | Remixes and etc. (485, 487 retired) | 26 |
| | | **Total artist streams** = every song row | 2 |
| | | **Solo streams** = total − Features | 3 |

Subtract 48 for the row numbers before phase 1b (e.g. Debut was 2–16, Remixes 436–491).

The aggregate rows are **not** in song-block order: `Midnights` and `TTPD` sit at rows 13/14,
while their song blocks come after the Taylor's Versions. Use the table rather than assuming the
orders line up.

A song is summed only by the category the Tracklist gives it. **A track added to Latest but
not to the Tracklist gets no total from the import and is in no category.** Add them through the
Pending sheet (see the row-layout contract), which keeps every block contiguous.

### `Import` (called `Tools` before 2.1) — the raw scrape

Stand-alone: nothing outside it refers to it, and since 2.1 it has no per-song rows, so it never
has to line up with the other sheets.

| Range | Contents |
| --- | --- |
| `C1` | `SUM_OF_DAILYS` — today's dailies added up, **written** by `matchTotalsById()` and `main()` for reference. The update never reads it; it works the sum out itself |
| `A2:B2` | A link to the repository |
| `E2:I1000` | **Raw Apify dump**, written by `importSpotifyData()`: E album, F name, G stream count, H track ID, I album cover URL (300px, from the actor's `coverArt`). Cleared and rewritten each import |
| `J:M` | Empty since 2.1. Until then: a merged separator (J), and per song row the Spotify name (K), total (L, written by the script) and daily (M, `=(L50-Latest!F50)`). **Setup → Tidy up the Import sheet** clears it and renames the tab |

The raw dump holds **more tracks than are tracked** (roughly 700+, varying per import). Apify
scrapes whole albums, and several albums in `CONFIG.APIFY.ALBUMS` (compilations, soundtracks,
other artists' records carrying a Swift feature) contain tracks she is not credited on. Tracks
whose ID is in no Tracklist row are simply never used. The actor gives no per-track artist, only
album-level `artists`, so non-Swift tracks can't be filtered out automatically.

Before 2.0, `Tools!L` was `=XLOOKUP(K,F:F,G:G)` on the Spotify name, with `MIN`/`MAX` variants
for the same-named pairs on (old) rows 42/56 (*The Story Of Us*) and 86/441 (*Bad Blood*).
The import warns if the dump would exceed 998 rows.

### `Latest` — current-day snapshot

One table: aggregate rows 2–27 on top, songs from row 50, sharing columns. Numeric column
indices matter, because the summary and milestone code reads by number, not by header.

| Col | # | Song rows | Aggregate rows (2–27) |
| --- | --- | --- | --- |
| A | 1 | Album/era cover key (`debutOG`, `fearlessTV`, …) | — |
| B / C | 2 / 3 | Rank, rank change vs. yesterday | — |
| D | 4 | Cover image: `XLOOKUP(A, Covers!A:A, Covers!B:B)` | (as for songs) |
| E | 5 | Title; a blank here means "skip this row" | Album/category title |
| F / G | 6 / 7 | Total, daily — **written** by `transferStats()` | Total, daily (sheet formulas) |
| H / I | 8 / 9 | % change vs. yesterday / vs. a week ago | same |
| J / K | 10 / 11 | Absolute change vs. yesterday / vs. a week ago | same |
| L | 12 | Best-since date — **written** by `findBestSince()` | same |
| M / N | 13 / 14 | Daily a day ago / a week ago — **written** by `updateStats()` | same |
| O | 15 | Previous day's rank | — |
| P | 16 | Unused: it held a second cover lookup, redundant with D, and has been cleared. New rows leave it blank (`LATEST.COPIED_COLUMNS` copies A:O only) | same |

The special-edition side table (tlpss, folklore standard, … showgirl era) sits in **T:AB** beside
the album rows, from row 2 down (T title, U total, V daily, W/X %, Y best since, Z/AA change,
AB URL). It is sheet-side and read only by the Albums sheet. Inserts only ever shift A:P in Latest,
so the table stays where it is even when categories move - it does not follow its albums.

`Q1` (`DATE_CELL`) holds the update date and is incremented by `transferStats()`.

### `Daily Archive 2023`–`2026` — the time series

Column-per-day, **newest first**: `transferStats()` does `insertColumnAfter(1)`, so today lands
in column B, yesterday is C, and a week ago is I. Row 1 holds the dates, column A the titles.
Rows follow the layout above (aggregates 2–27, songs 50+). `findBestSince()` does one pass over rows 2 to the last song, walking the years newest-to-oldest via
`CONFIG.SHEETS.ARCHIVE_YEARS`, skipping column B on the current year (that is today's value,
the threshold being beaten) and including it on past years (that is Dec 31st).

### Total archive — manual cold backup

Structurally a Daily Archive, but holding **cumulative totals** instead of daily counts, and
only two columns per month: the first and last day. It runs back to 2022-12-31. Song rows match
the standard block layout, so a row means the same track here as everywhere else.

It is maintained by hand — the owner pastes the totals column across on those dates — as a
second copy of the data and so that the total for any given day can be reconstructed if the
Daily Archives are ever damaged. The daily code never writes it. The row-alignment check reads
its column A (via `CONFIG.SHEETS.TOTAL_ARCHIVE`), and the phase 1b migration gave it aggregate
rows 2–27 (formulas) in every column, which it had never had.
Leave it alone unless asked; if it ever is automated, note that its columns are sparse and
non-contiguous, so none of the fixed-offset tricks the Daily Archive relies on (column C is
yesterday, column I is a week ago) apply.

### `Tracks` — sortable leaderboard **[sheet-side]** + one written block

Sorting controls in A–B drive a formula-sorted leaderboard in C–M over the song rows of Latest.
After phase 1b it should be a `FILTER` from row 50 that drops blank rows and looks each row's
status up in the Tracklist by `row` (see `PLAN.md` for the formula). The script's only stake is
`N45:Q549` (`CONFIG.TRACKS`), cleared and rewritten by `updateUpcomingMilestones()`: headers on
row 45 (`N45:O45` merged), then one row per predicted milestone with title (N), milestone
value (P) and the generated announcement sentence (Q).

### `Albums` — era analysis **[sheet-side]** + the generated summaries

Summary tables and per-era track breakdowns are laid out in column groups, with each era's
generated text summary written into the column just right of its breakdown. The breakdowns
reference specific `Latest` cells (and the album chart a range, `SORT(Latest!D4:H19,…)` after phase 1b), which Sheets shifts along when rows are inserted in Latest,
but a newly inserted song does not appear in them by itself. Destinations are
`summaryCell` in the Categories sheet (F23, F45, F79, L23, L54, L93, L119, R23, R49, R78, R99, W23,
W51, W78, W102, AB23, AB89) plus `G5` for the whole-discography summary
(`CONFIG.ALBUMS.TOTAL_SUMMARY`). Each write clears 10 rows first, so summaries must stay at
least 10 rows apart.

## Known inconsistencies

Found by reconciling the code against the layout; none are currently breaking anything, but
they will bite whoever next edits these areas.

1. **Rows 393–404 fall through every category sum.** These are the 12 *Track by Track* versions
   of The Life of a Showgirl, scraped from the `"The Life of a Showgirl (Track by Track
   Version)"` entry in `CONFIG.APIFY.ALBUMS`. Their `category` in the Tracklist sheet is blank,
   so they are in none of the 24 category sums — yet they are included in the `SUM(B2:B549)`
   artist total on row 574. The 24 category rows therefore do not add up to row 574. The
   (Rows given here and below are pre-phase-1b; add 48 for today's.) The
   omission was not deliberate, but **the decision (2026-09-22) is to leave it as it is** —
   the tracks are low-value and not worth restructuring the aggregates for. Treat this as a
   known and accepted gap, not a bug to fix; do not "correct" it without being asked.

   Note this scatters the Showgirl era across three blocks: 336–347 (standard, the only part
   the summary scans), 348–354 (the Acoustic Collection tail) and 393–404 (Track by Track).
   Only the first two are in the era's aggregate (old row 565, now 19). If the gap ever needs closing,
   **Move a Song** can give each track a category (and move it next to its era), with the history
   rebuilt to match.

2. **Four summaries scan fewer songs than their album sums.** `summaryLimit` in
   the Categories sheet (formerly `CONFIG.STATS.count`): Speak Now (17 of 21), Red (19 of 22),
   1989 (16 of 19) and The Life of a Showgirl (12 of 19). These are the deluxe/bonus-track
   eras, so it looks intentional, but tracks in the excluded tail can never be named as an
   era's biggest gainer. The limit counts from the start of the category, so songs appended to
   one of these eras are summed but not summarised.

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
   (old Latest row 18, now 20), the live/compilation rows, Soundtracks, Remixes or Features. Today
   it scans rows 4–19.
   Those categories still get their own per-era summaries where they have a `summaryCell` in
   the Categories sheet (Soundtracks does, writing to `R99`); that is separate and intended.
   (`CONFIG.STATS` became `CONFIG.CATEGORIES` and then the Categories sheet, 2026-09-22.) Today the
   discography summary picks its rows by type `studio`, not by a row count.

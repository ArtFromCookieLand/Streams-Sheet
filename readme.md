 # TStreams

Project for storing and managing Taylor Swift's streams on Spotify in a google spreadsheet.

## Description

In order to keep track of Taylor Swift's streams on Spotify, this project was created. Google Spreadsheets is used to store the data, Apify App "Spotify Play Count Scraper" is used to scrape the data, and Google Apps Script is used to automate the process of updating the spreadsheet with the latest stream counts. Clasp is used to connect Google Apps Script with a coding environment such as VSCode. 

## Usage

### Environments

There are two Apps Script projects sharing this one codebase: **prod**, bound to the live
spreadsheet, and **dev**, bound to a copy of it for testing changes safely.

| Environment | clasp config | Target |
| --- | --- | --- |
| prod | `.clasp.json` | the live spreadsheet |
| dev | `.clasp.dev.json` | a copy of the spreadsheet |

### Pushing and pulling

Push and pull always name the environment:
```
npm run push:dev      npm run push:prod
npm run pull:dev      npm run pull:prod
npm run open:dev      npm run open:prod
```

The script detects at runtime which spreadsheet it is attached to. In the dev copy the menu is
titled **Update [DEV]**, and the Apify import is blocked so that testing never spends scraper
credits — the last raw import copied over in `Tools!E2:H1000` is used as the test data instead,
via **Match Totals by ID**. It must be one that includes track IDs, i.e. one made after 2.0.0.1.
The **Checks** menu runs read-only health checks on either copy at any time.

### Setting up the dev copy

1. In the live spreadsheet: **File → Make a copy**, name it e.g. `TStreams Dashboard (DEV)`.
   The bound Apps Script project is copied along with it.
2. In the copy: **Extensions → Apps Script → Project Settings**, copy the **Script ID**.
3. Paste it into `.clasp.dev.json` as `scriptId`.
4. `npm run push:dev`, then reload the copy and use the **Update [DEV]** menu.

Refresh the dev copy by repeating step 1 whenever it has drifted too far from production.

## Authors

Haunted_Spotify (also known as Haunted_Jade or Jade) [Jade is not the actual name]

## Version History

* 2.0.0.6 (22/9/26)
    * New albums and categories: put the name and type in the **Categories** sheet, leave the row empty, and run **Update → Add New Categories**. It gives the category a row after the last one of its type — a studio album after the last album, a compilation after the last live album — opens that row in Latest and every archive, moves the categories below down one, uses up a spare row so the songs never move, and rebuilds every total, including the Total Archive's history.
    * The album URLs to scrape moved from the code into a **Sources** sheet, so a new release needs a row in a sheet rather than a code change.
    * The layout now knows about the header row above the songs (row 49): updates write the album rows and the song rows as two blocks and never touch it.
    * Inserting a row no longer shifts part of a row in Tools, so merged cells there (the separator column J) can't break it. Tools' raw import just moves down and is rewritten on the next import. In Latest, where part of a row still shifts, merged cells in the way are reported before anything changes, and a new **Merged cells** check warns in advance.
    * New rows no longer fill in column P, which is no longer used.
* 2.0.0.5 (22/9/26)
    * New cover keys get their Covers row automatically, with the key, the image formula and the URL. The URL is the cover of the album the track was scraped from, which the import now keeps in Tools!I, or else Spotify's embed lookup for the track. This happens when songs are added, and when an upcoming song turns up on Spotify. **Update → Fill Missing Covers** does the same for the whole Tracklist.
    * Existing covers are never changed; to use a different image, edit the URL in Covers.
    * The Covers check also warns if a cover key in use has no image formula.
* 2.0.0.4 (22/9/26)
    * New tracks are found automatically after every import, or with **Update → Find New Tracks** against the current import. Tracks that are the same recording on another edition of an album (same stream count and name, for example a deluxe's standard tracks) are ignored automatically. A track matching an announced **upcoming** song gets its ID filled in and counts straight away. Anything else lands in **Pending** with the album and cover key suggested from songs already tracked on that album.
    * New **Ignored** sheet for tracks that are deliberately not tracked, such as other artists on soundtracks and other editions. It is started with everything untracked in the import, and Pending rows set to **ignore** are moved there.
    * New **New tracks** health check, which warns while the import holds tracks nobody has looked at yet.
* 2.0.0.3 (22/9/26)
    * New songs are added through a **Pending** sheet and **Update → Add Pending Songs**. Each song goes at the end of its album's block, at the same row in Latest, Tools, every Daily Archive and the Total Archive, and the Tracklist is renumbered to match. The whole batch is checked first; if any row has a problem nothing changes, and the problem is written next to that row.
    * Songs can be added before they're released, with the new **upcoming** status. The track ID can be left blank, the song stays at 0 until it appears in an import, and it never stops the daily update. The import says when an upcoming song has gone live.
    * Album and category definitions moved from the code into a **Categories** sheet (row, type, summary cell, summary limit), so changing them needs no code push. Type is studio, other or fixed; the discography summary uses every studio album, wherever it sits.
    * Latest's album totals and dailies are now written by the script from the Tracklist, so a new song is always counted in its album.
    * Removed the one-off layout migration from 2.0.0.2, now that both spreadsheets are migrated.
* 2.0.0.2 (22/9/26)
    * New spreadsheet layout, the same in Latest, Tools, every Daily Archive and the Total Archive: the whole discography on row 2, the solo total on row 3, the 24 albums and categories on rows 4-27, spare rows for new categories up to 49, and songs from row 50 with no upper limit. Before, songs were capped at row 549, with the album totals below them in the archives and beside them in Latest.
    * Latest's album table now uses the same columns as the songs, and an album's row is the same number in Latest and in every archive. Albums also get "a day ago" and "a week ago" figures, and "best since" is worked out in one pass for albums and songs.
    * The solo total (all streams minus Features) now has a row in the archives, with its full history, so it gets a "best since" date too.
    * The Total Archive has album and category totals for the first time, in every column back to the start.
    * The Tracklist's `row` column now says where each song lives, so the Tracklist can be sorted freely.
    * Added a **Checks** menu with ten read-only health checks: row alignment across every sheet, album totals against their songs, archive dates (gaps or repeated days), the import, covers, spare rows and config. The quick ones also run at the end of every Update Daily Stats and only show a dialog if something is wrong.
    * Included a one-off **Migrate Layout** script that converted the sheets. It refuses to start unless the sheets are exactly in the old layout, and can't run twice.
* 2.0.0.1 (22/9/26)
    * Added the Tracklist sheet: one row per song with its category, status, cover key, display title and Spotify track ID. It is now the single place that decides which rows belong to which album.
    * Stream totals are matched by Spotify track ID instead of by title, so Spotify renaming a track no longer loses its streams. The import now also saves each track's album and ID in Tools (E:H).
    * If a track ID is missing from an import, the update refuses to run and names the songs, instead of failing silently. The new **Match Totals by ID** menu item re-matches after a fix without spending Apify credits.
    * The album and category totals in the Daily Archive are now built from the Tracklist instead of hardcoded row ranges. Album summaries also take their songs from it.
    * Songs can be retired: they keep their row and archive history, but are held at 0 and left out of milestones and summaries. Our Song (International Mix) and Love Story (Pop Mix) are the first, after Spotify merged their counts into the originals.
* 1.1.4 (21/9/26)
    * Added a separate dev environment, so changes can be tested on a copy of the spreadsheet before they touch the live one. Apify imports are blocked there to avoid spending scraper credits.
    * Fixed the discography summary including Droplets among the albums. It scanned one row too many because the number of rows was taken from the length of CONFIG.STATS, which grew when Soundtracks was added in 1.1.
    * Documented the project and the spreadsheet layout in CLAUDE.md, including the song-row blocks and a few known inconsistencies.
    * Stopped tracking node_modules and spreadsheet exports in git.
* 1.1.3 (2/7/26)
    * I Knew It, I Knew You remixes are now being tracked
* 1.1.2 (8/6/26)
    * I Knew It, I Knew You is now being tracked
    * Added third Apify token
* 1.1 (13/5/26)
    * Added generation of summaries for albums and the discrography.
* 1.0
    * The first tracked version of the project from May 2026 (initially the project was created in August 2025). Includes such features as:
        * Saving the data in a Google Spreadsheet and managing it
        * Scraping Taylor Swift's stream counts on Spotify from Apify (added late 2025)
        * Tracking milestones surpassed and to be surpassed 
        * Automating the process of updating the spreadsheet with the latest stream counts using Google Apps Script.
        * Running a webpage with daily streams update, which pulls the data from the spreadsheet and displays it in a visually appealing way (added early 2026 and deprecated shortly after due to lack of time and release of alternatives)

## Links

* [The Spreadsheet](https://docs.google.com/spreadsheets/d/1ANVfUER8MInJ_y-oRp7Nguwk0yDI9LtBt36PWBMmZbU)
* [The Apify App](https://console.apify.com/actors/YZhD6hYc8daYSWXKs/input?addFromActorId=YZhD6hYc8daYSWXKs)
* [Twitter Page](https://x.com/Haunted_Spotify)
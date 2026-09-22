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
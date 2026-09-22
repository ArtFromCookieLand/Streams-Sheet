# How to add songs, albums and categories

Plain-text playbook, meant to be pasted into the spreadsheet. Menu items are under **Update**
and **Checks**. Every scenario ends the same way: run **Checks → Run All Checks**.

---

## THE SHEETS

- **Tracklist** — every tracked song: its row, status, category, cover key, title and Spotify
  track ID. The row column says where the song lives in Latest, Tools and the archives. It can
  be sorted freely.
- **Categories** — one row per album/category: row, type (studio / other / fixed), summary cell,
  summary limit.
- **Pending** — songs waiting to be added. Filled in by you, or by Find New Tracks.
- **Ignored** — track IDs that are deliberately not tracked.
- **Sources** — the Spotify album links the import scrapes.
- **Covers** — cover key (A), image (B), URL (C).

## STATUS OF A SONG

- **active** — normal. Its track ID must be in every import, or the daily update stops.
- **upcoming** — added before release. The ID may be blank; the song stays at 0 and never blocks
  the update. Switch it to active once it is out.
- **retired** — kept only for its history, held at 0 (e.g. a mix Spotify merged into the original).

---

## 1. A NEW SINGLE, ANNOUNCED BUT NOT OUT YET

1. **Covers**: add the cover key in A, the URL in C, and copy the image formula down into B.
2. **Pending**: title, category (a standalone single is usually `Droplets`), cover key,
   status `upcoming`, track ID empty.
3. **Update → Add Pending Songs**. It gets its row everywhere and sits at 0.
4. **On release day**: paste the single's Spotify album link into **Sources**, then
   **Update → Import Data**. The dialog says "1 upcoming song(s) are now on Spotify and counted".
5. Set its status in **Tracklist** to `active`.
6. Its first day counts all of its streams, which is correct for a new release.

**If the import does not link it** (Spotify's title differs from yours by more than case,
punctuation or the TV/FTV tags), the track appears in Pending instead. Then: copy its track ID
into the song's Tracklist row, set that row to `active`, set the Pending row's status to `ignore`
and run Add Pending Songs.

## 2. A NEW ALBUM OR ERA (including a deluxe edition that becomes its own category)

1. **Categories**: new row with the name and type. Leave `row` **empty**.
   - `studio` = album or re-recording (counts in the discography summary, placed after the last
     studio album)
   - `other` = compilation or live album (placed after the last of those)
   - `fixed` = Soundtracks, Remixes, Features — leave these alone
   Optionally set the summary cell (where its text summary goes on the Albums sheet) and the
   summary limit.
2. **Update → Add New Categories**. It gets a row in Latest and every archive; the categories
   below move down one.
3. **Covers**: add the album's cover key (A, B, C as above) if it has its own.
4. **Pending**: one row per song, with the new category, status `upcoming` before release.
5. **Update → Add Pending Songs**.
6. **On release day**: album link into **Sources**, then **Import Data**, then set the songs to
   `active`.
7. By hand afterwards, if you want them: the album's breakdown and summary on the **Albums**
   sheet, and widening the studio-albums chart range by one row.

## 3. NEW SONGS ON AN ALBUM THAT ALREADY EXISTS (a deluxe of an existing era)

1. **Pending**: one row per song, category = the existing album, cover key, status `upcoming`
   (or `active` with its track ID if it is already out).
2. **Update → Add Pending Songs**. They land at the end of that album's block, so everything
   below moves down a row — that is fine and expected.
3. If the album's summary has a **summary limit** (Speak Now, Red, 1989, The Life of a Showgirl),
   the new songs fall outside it. Raise the limit in **Categories** if you want them in the
   summary.
4. Release day as above: link into **Sources**, import, set to `active`.

## 4. A SONG THAT HAS BEEN ON SPOTIFY FOR A WHILE

1. Make sure its album is in **Sources**, then **Update → Import Data** (or use the last import).
2. If it is in **Ignored**, delete its row there.
3. **Update → Find New Tracks** — it appears in **Pending** with a suggested category and cover
   key. (Or write the Pending row yourself, with its track ID and status `active`.)
4. Check the title, category and cover key, then **Update → Add Pending Songs**.
5. **Its first day will count its whole total as one day's streams.** After the next daily
   update, correct that day by hand: put the real daily figure in the archive column and in
   Latest, taken from wherever you track it.

## 5. A TRACK YOU DO NOT WANT TO TRACK

In **Pending**, set the row's status to `ignore` and run **Add Pending Songs**. It moves to
**Ignored** and never comes back. To undo, delete its row from Ignored.

## 6. SPOTIFY MERGED ONE SONG'S COUNT INTO ANOTHER

Set the song's status in **Tracklist** to `retired`. It keeps its row and its history, is held
at 0, and is left out of milestones and summaries. Never delete a song's row.

## 7. A TRACK ID CHANGED, OR THE UPDATE SAYS "#MISSING"

1. Find the song's new ID: **Update → Find New Tracks** usually puts the new track in Pending,
   or take it from the Spotify link (the part after /track/).
2. Put it in the song's **Tracklist** row.
3. **Update → Match Totals by ID** — no new import needed, so no Apify credits.
4. Run **Update Daily Stats** again.

---

## THE DAILY ROUTINE

1. **Update → Import Data** (also matches totals and looks for new tracks).
2. **Update → Update Daily Stats**. A dialog appears at the end only if a check fails.

Adding songs is fine at any time, including between the import and the update. Just do not leave
it half-done in the middle of a daily update.

## GOOD TO KNOW

- Nothing is changed unless the whole batch is valid: if one Pending row has a problem, nothing
  is added and the reason is written next to that row.
- Rows marked ✅ in Pending are done and are skipped from then on. Delete them whenever you like.
- **Do not insert or delete song rows by hand** — Add Pending Songs keeps all the sheets in step.
- **Do not merge cells across the edge of A:P in Latest** below the albums. That block shifts on
  its own when a row is inserted. A merge inside A:P is fine.
- **Never write into Tools column J** — it is the merged separator.
- If anything ever looks out of step, **Checks → Check Row Alignment** tells you the first row
  that disagrees, and Version history is the safety net.

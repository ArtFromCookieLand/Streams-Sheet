/**
 * ===================================================================
 * FIND NEW TRACKS
 * -------------------------------------------------------------------
 * Looks through the raw import for track IDs that are in none of the
 * Tracklist, Pending or Ignored sheets, and sorts each one:
 *
 *   same recording  - same stream count and name as a tracked song
 *                     (another edition of the album): Ignored, with the
 *                     reason
 *   upcoming song   - its name matches exactly one `upcoming` Tracklist
 *                     song that has no ID yet: the ID is filled in there,
 *                     so the song starts counting
 *   anything else   - a new Pending row, with the category and cover key
 *                     suggested from tracked songs on the same album
 *
 * Runs after every import, and from Update → Find New Tracks against the
 * import already in Tools (no Apify cost). It never touches Latest, the
 * archives or Tools - adding songs is still Add Pending Songs' job.
 * ===================================================================
 */

/**
 * [MENU] Runs the detection against the current raw import and reports.
 */
function findNewTracksMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const found = findNewTracks();
    if (found.linked.length) matchTotalsById();  // the linked songs' totals, from the same import
    ui.alert('Find New Tracks' + envTag(), describeNewTracks(found), ui.ButtonSet.OK);
  } catch (error) {
    console.error(error);
    ui.alert('Error', error.toString(), ui.ButtonSet.OK);
  }
}


/**
 * @return {{pending: Array, linked: Array, duplicates: Array}} What was done, one entry per track.
 */
function findNewTracks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName(CONFIG.SHEETS.IGNORED)) {
    throw new Error(`The "${CONFIG.SHEETS.IGNORED}" sheet is missing - run Update → Setup → Create Ignored sheet first, or every untracked track would land in Pending.`);
  }
  const raw = readRawImport();
  if (!raw.length) throw new Error('The raw import in Tools has no track IDs - import first.');

  const songs = getSongs();
  const known = knownTrackIds(songs);
  const byId = {};
  raw.forEach(t => { byId[t.id] = t; });

  // Tracked songs by stream count, to spot the same recording on another edition.
  const trackedByCount = {};
  songs.forEach(s => {
    const t = s.trackId && byId[s.trackId];
    if (t && t.count > 0) (trackedByCount[t.count] = trackedByCount[t.count] || []).push({ song: s, name: t.name });
  });

  // Upcoming songs still waiting for an ID, by normalised title.
  const upcomingByTitle = {};
  songs.filter(s => s.status === SONG_STATUS.UPCOMING && !s.trackId).forEach(s => {
    const k = normalizeTitle(s.title);
    (upcomingByTitle[k] = upcomingByTitle[k] || []).push(s);
  });

  // Suggestions: the most common category and cover key among tracked songs on the same album.
  const byAlbum = {};
  songs.forEach(s => {
    const t = s.trackId && byId[s.trackId];
    if (!t) return;
    const a = byAlbum[t.album] = byAlbum[t.album] || { categories: {}, covers: {} };
    if (s.category) a.categories[s.category] = (a.categories[s.category] || 0) + 1;
    if (s.coverKey) a.covers[s.coverKey] = (a.covers[s.coverKey] || 0) + 1;
  });
  const mostCommon = counts => Object.keys(counts).sort((x, y) => counts[y] - counts[x])[0] || '';

  const result = { pending: [], linked: [], duplicates: [] };
  const today = formatDateString(new Date());
  const fresh = raw.filter(t => !known[t.id]);
  const namesOfFresh = {};
  fresh.forEach(t => { const k = normalizeTitle(t.name); namesOfFresh[k] = (namesOfFresh[k] || 0) + 1; });

  fresh.forEach(t => {
    // 1. Another edition of a tracked recording.
    const same = (trackedByCount[t.count] || []).filter(x => sameRecordingName(x.name, t.name));
    if (t.count > 0 && same.length) {
      result.duplicates.push({ trackId: t.id, spotifyTitle: t.name, album: t.album,
        reason: `Same recording as "${same[0].song.title}" (row ${same[0].song.row}), on another edition` });
      return;
    }
    // 2. An announced song that is now out - only when the match is unambiguous both ways.
    const k = normalizeTitle(t.name);
    const waiting = upcomingByTitle[k] || [];
    if (waiting.length === 1 && namesOfFresh[k] === 1) {
      const song = waiting[0];
      result.linked.push({ song: song, trackId: t.id, spotifyTitle: t.name, album: t.album });
      delete upcomingByTitle[k];
      // On a brand-new album this may be its first known track - let it guide the suggestions
      // for the album's other new tracks.
      const a = byAlbum[t.album] = byAlbum[t.album] || { categories: {}, covers: {} };
      if (song.category) a.categories[song.category] = (a.categories[song.category] || 0) + 1;
      if (song.coverKey) a.covers[song.coverKey] = (a.covers[song.coverKey] || 0) + 1;
      return;
    }
    // 3. Something new, for you to look at.
    const album = byAlbum[t.album];
    result.pending.push({
      title: t.name, trackId: t.id, spotifyTitle: t.name, album: t.album, status: SONG_STATUS.ACTIVE,
      category: album ? mostCommon(album.categories) : '',
      coverKey: album ? mostCommon(album.covers) : '',
      result: `🆕 Found on ${today}: check the title, category and cover key (or set status to ignore), then Add Pending Songs`
    });
  });

  if (result.duplicates.length) addToIgnored(result.duplicates);
  if (result.linked.length) {
    linkUpcomingSongs(result.linked);
    // Their cover keys may be new too; now there is an album to take the cover from.
    result.covers = ensureCovers(result.linked.map(l => ({ coverKey: l.song.coverKey, trackId: l.trackId, album: l.album, title: l.song.title })));
  }
  if (result.pending.length) appendPendingRows(result.pending);
  return result;
}


function describeNewTracks(r) {
  if (!r.pending.length && !r.linked.length && !r.duplicates.length) return 'No new tracks in the import.';
  const lines = [];
  if (r.linked.length) {
    lines.push(`🎉 ${r.linked.length} upcoming song(s) found on Spotify - their IDs are now in the ${CONFIG.SHEETS.SONGS} sheet, so they count from this import:`);
    r.linked.forEach(l => lines.push(`   • ${l.song.title} (row ${l.song.row})`));
    lines.push(`   Switch them to "active" once you're happy.`);
  }
  if (r.pending.length) {
    lines.push(`🆕 ${r.pending.length} new track(s) added to ${CONFIG.SHEETS.PENDING} - check them, then run Add Pending Songs:`);
    r.pending.slice(0, 15).forEach(p => lines.push(`   • ${p.title} (${p.album})${p.category ? '' : ' - no category suggested'}`));
    if (r.pending.length > 15) lines.push(`   … and ${r.pending.length - 15} more`);
  }
  if (r.duplicates.length) {
    lines.push(`🔁 ${r.duplicates.length} track(s) are the same recording on another edition - ignored automatically.`);
  }
  if (r.covers && describeCovers(r.covers)) lines.push(describeCovers(r.covers));
  return lines.join('\n');
}


// --- helpers ---

/** @return {Array<{album: string, name: string, count: number, id: string, cover: string}>} Rows that have an ID. */
function readRawImport() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.TOOLS)
    .getRange(CONFIG.TOOLS.RAW_DATA).getValues()
    .map(r => ({ album: String(r[0]), name: String(r[1]), count: Number(r[2]) || 0, id: String(r[3]).trim(), cover: String(r[4] || '').trim() }))
    .filter(t => t.id);
}

/** Every track ID already accounted for: in the Tracklist, Pending (any row) or Ignored. */
function knownTrackIds(songs) {
  const known = {};
  songs.forEach(s => { if (s.trackId) known[s.trackId] = true; });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [[CONFIG.SHEETS.PENDING, CONFIG.PENDING_SHEET.HEADERS], [CONFIG.SHEETS.IGNORED, CONFIG.IGNORED_SHEET.HEADERS]].forEach(([name, headers]) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 2) return;
    const values = sheet.getDataRange().getValues();
    const col = findHeaderColumns(values[0], { trackId: headers.trackId }, name).trackId;
    values.slice(1).forEach(r => { const id = String(r[col]).trim(); if (id) known[id] = true; });
  });
  return known;
}

/**
 * Lower-case letters and digits only, with the version tags removed - Spotify's "(Taylor's Version)"
 * and "(From The Vault)" as well as the sheet's own "(TV)" and "(FTV)" - so the two spellings meet.
 */
function normalizeTitle(title) {
  return String(title).toLowerCase()
    .replace(/[’‘]/g, "'")
    .replace(/\(?\s*taylor'?s version\s*\)?/g, '')
    .replace(/\(?\s*from the vault\s*\)?/g, '')
    .replace(/\((f?tv)\)/g, '')
    .replace(/[^a-z0-9]/g, '');
}

/** Same stream count is the strong signal; this just makes sure the names agree too. */
function sameRecordingName(a, b) {
  const x = normalizeTitle(a), y = normalizeTitle(b);
  return !!x && !!y && (x === y || x.indexOf(y) === 0 || y.indexOf(x) === 0);
}

/**
 * Appends rows to the Ignored sheet.
 * @param {Array<{trackId, spotifyTitle, album, reason}>} items
 */
function addToIgnored(items) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.IGNORED);
  if (!sheet) throw new Error(`The "${CONFIG.SHEETS.IGNORED}" sheet is missing - run Update → Setup → Create Ignored sheet.`);
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = findHeaderColumns(header, CONFIG.IGNORED_SHEET.HEADERS, CONFIG.SHEETS.IGNORED);
  const now = new Date();
  const rows = items.map(it => {
    const row = header.map(() => '');
    row[col.trackId] = it.trackId;
    row[col.spotifyTitle] = it.spotifyTitle || '';
    row[col.album] = it.album || '';
    row[col.reason] = it.reason || '';
    row[col.date] = now;
    return row;
  });
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, header.length).setValues(rows);
}

/** Writes each found ID (and the Spotify name/album, where those columns exist) into the Tracklist. */
function linkUpcomingSongs(links) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SONGS);
  const values = sheet.getDataRange().getValues();
  const col = findSongColumns(values[0]);
  const header = values[0].map(h => String(h).toLowerCase().replace(/\s+/g, ''));
  const spotifyCol = header.indexOf('spotifytitle'), albumCol = header.indexOf('sourcealbum');
  links.forEach(l => {
    const i = values.findIndex((r, n) => n > 0 && Number(r[col.row]) === l.song.row && String(r[col.title]).trim() === l.song.title);
    if (i < 1) return;
    sheet.getRange(i + 1, col.trackId + 1).setValue(l.trackId);
    if (spotifyCol !== -1) sheet.getRange(i + 1, spotifyCol + 1).setValue(l.spotifyTitle);
    if (albumCol !== -1) sheet.getRange(i + 1, albumCol + 1).setValue(l.album);
  });
  _songsCache = null;
}

/** Appends detected tracks to Pending, and makes sure its status dropdown offers "ignore". */
function appendPendingRows(items) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PENDING);
  if (!sheet) throw new Error(`The "${CONFIG.SHEETS.PENDING}" sheet is missing - run Update → Setup → Create Pending sheet.`);
  const header = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = findHeaderColumns(header, CONFIG.PENDING_SHEET.HEADERS, CONFIG.SHEETS.PENDING);
  const rows = items.map(it => {
    const row = header.map(() => '');
    Object.keys(CONFIG.PENDING_SHEET.HEADERS).forEach(k => { if (it[k] !== undefined) row[col[k]] = it[k]; });
    return row;
  });
  const start = Math.max(sheet.getLastRow(), 1) + 1;
  sheet.getRange(start, 1, rows.length, header.length).setValues(rows);

  // Sheets created before "ignore" existed have a dropdown without it.
  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(CONFIG.PENDING_SHEET.STATUSES, true).build();
  sheet.getRange(2, col.status + 1, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(statusRule);
}

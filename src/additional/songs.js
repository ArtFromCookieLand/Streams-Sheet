/**
 * ===================================================================
 * SONGS REGISTRY
 * -------------------------------------------------------------------
 * The `Tracklist` sheet is the master list of tracked songs. Its sheet row
 * N is song row N everywhere else (Latest, Tools, every Daily Archive),
 * so it must never be sorted. The "row" column holds each song's own
 * row number, and reading stops with an error if the two ever disagree.
 *
 * Columns are found by their header (CONFIG.SONGS_SHEET.HEADERS),
 * ignoring case and spaces. Extra columns are ignored.
 *
 * Status:
 *   active  - tracked normally, its total is matched by track ID
 *   retired - kept only for its archive history (e.g. a mix Spotify
 *             merged into the original). Held at 0 and left out of
 *             milestones and summaries.
 * ===================================================================
 */

const SONG_STATUS = { ACTIVE: 'active', RETIRED: 'retired' };

var _songsCache = null;

/**
 * @return {Array<{row: number, status: string, category: string, title: string, trackId: string}>}
 *   One entry per song, in row order. Rows with no title (the headroom) are left out.
 */
function getSongs() {
  if (_songsCache) return _songsCache;

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SONGS);
  if (!sheet) throw new Error(`The "${CONFIG.SHEETS.SONGS}" sheet is missing.`);

  const values = sheet.getDataRange().getValues();
  const col = findSongColumns(values[0]);
  const lastSongRow = CONFIG.SONGS.START_ROW + CONFIG.SONGS.COUNT - 1;
  const statuses = Object.values(SONG_STATUS);

  const songs = [];
  const rowOfId = {};

  for (let i = CONFIG.SONGS.START_ROW - 1; i < values.length; i++) {
    const r = values[i];
    const sheetRow = i + 1;
    const title = String(r[col.title]).trim();
    if (!title) continue;

    if (sheetRow > lastSongRow) {
      throw new Error(`${CONFIG.SHEETS.SONGS} sheet row ${sheetRow} ("${title}") is past the song range, which ends at row ${lastSongRow}.`);
    }
    if (Number(r[col.row]) !== sheetRow) {
      throw new Error(
        `${CONFIG.SHEETS.SONGS} sheet row ${sheetRow} ("${title}") says it belongs on row ${r[col.row]}. ` +
        `The sheet has been sorted or had rows moved, and every song must stay on its own row.`
      );
    }

    const status = String(r[col.status]).trim().toLowerCase();
    if (statuses.indexOf(status) === -1) {
      throw new Error(`${CONFIG.SHEETS.SONGS} sheet row ${sheetRow} ("${title}") has status "${r[col.status]}"; expected one of: ${statuses.join(', ')}.`);
    }

    const trackId = String(r[col.trackId]).trim();
    if (trackId) {
      if (rowOfId[trackId]) {
        throw new Error(`Track ID ${trackId} is used on both row ${rowOfId[trackId]} and row ${sheetRow} of the ${CONFIG.SHEETS.SONGS} sheet.`);
      }
      rowOfId[trackId] = sheetRow;
    }

    songs.push({
      row: sheetRow,
      status: status,
      category: String(r[col.category]).trim(),
      title: title,
      trackId: trackId
    });
  }

  _songsCache = songs;
  return songs;
}

/**
 * Maps each configured header to its column index in the Tracklist sheet.
 * @param {Array} headerRow - Row 1 of the sheet.
 * @return {Object} e.g. { row: 0, status: 1, ... }
 */
function findSongColumns(headerRow) {
  const normalize = h => String(h).toLowerCase().replace(/\s+/g, '');
  const headers = headerRow.map(normalize);
  const col = {};

  for (const [key, header] of Object.entries(CONFIG.SONGS_SHEET.HEADERS)) {
    const index = headers.indexOf(normalize(header));
    if (index === -1) {
      // Listing what is there makes a typo in a header obvious.
      const found = headerRow.filter(h => String(h).trim()).map(h => `"${h}"`).join(', ');
      throw new Error(`The ${CONFIG.SHEETS.SONGS} sheet has no "${header}" column in row 1. Row 1 has: ${found}.`);
    }
    col[key] = index;
  }
  return col;
}

/**
 * @return {Object} Sheet row -> true, for every retired song.
 */
function getRetiredRows() {
  const retired = {};
  getSongs().forEach(s => {
    if (s.status === SONG_STATUS.RETIRED) retired[s.row] = true;
  });
  return retired;
}

/**
 * @return {Object} Category name -> ascending list of its song rows (retired songs included).
 *   Songs with a blank category are in none.
 */
function getRowsByCategory() {
  const rowsByCategory = {};
  getSongs().forEach(s => {
    if (!s.category) return;
    if (!rowsByCategory[s.category]) rowsByCategory[s.category] = [];
    rowsByCategory[s.category].push(s.row);
  });
  return rowsByCategory;
}

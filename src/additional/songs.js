/**
 * ===================================================================
 * SONGS REGISTRY
 * -------------------------------------------------------------------
 * The `Tracklist` sheet is the master list of tracked songs: a list, not
 * a grid. Each song's "row" column says which row it lives on in Latest,
 * Tools and every archive (CONFIG.LAYOUT.FIRST_SONG_ROW onwards), so the
 * Tracklist's own order doesn't matter and it can be sorted freely.
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
 * @return {Array<{row: number, status: string, category: string, coverKey: string, title: string, trackId: string}>}
 *   One entry per song, in row order. Rows with no title (the headroom) are left out.
 *   Throws on the first problem found; readTracklist() lists them all instead.
 */
function getSongs() {
  if (_songsCache) return _songsCache;

  const result = readTracklist();
  if (result.problems.length) throw new Error(result.problems[0]);

  _songsCache = result.songs;
  return _songsCache;
}

/**
 * Reads and validates the Tracklist without throwing.
 * @return {{songs: Array, problems: Array<string>}} Every song that could be read, and every
 *   problem found. A missing sheet or header is the only problem that leaves `songs` empty.
 */
function readTracklist() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SONGS);
  if (!sheet) return { songs: [], problems: [`The "${CONFIG.SHEETS.SONGS}" sheet is missing.`] };

  const values = sheet.getDataRange().getValues();
  let col;
  try {
    col = findSongColumns(values[0]);
  } catch (error) {
    return { songs: [], problems: [error.message] };
  }

  const firstSongRow = CONFIG.LAYOUT.FIRST_SONG_ROW;
  const statuses = Object.values(SONG_STATUS);
  const name = CONFIG.SHEETS.SONGS;

  const songs = [];
  const problems = [];
  const rowOfId = {};
  const titleOfRow = {};

  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const sheetRow = i + 1;
    const title = String(r[col.title]).trim();
    if (!title) continue;

    // The song's row everywhere else. It must be a whole number in the song range, used once.
    const row = Number(r[col.row]);
    if (!Number.isInteger(row) || row < firstSongRow) {
      problems.push(`${name} sheet row ${sheetRow} ("${title}") has row "${r[col.row]}"; songs live on row ${firstSongRow} or below.`);
      continue;
    }
    if (titleOfRow[row]) {
      problems.push(`${name}: "${titleOfRow[row]}" and "${title}" both say they live on row ${row}.`);
      continue;
    }
    titleOfRow[row] = title;

    const status = String(r[col.status]).trim().toLowerCase();
    if (statuses.indexOf(status) === -1) {
      problems.push(`${name}: row ${row} ("${title}") has status "${r[col.status]}"; expected one of: ${statuses.join(', ')}.`);
    }

    const trackId = String(r[col.trackId]).trim();
    if (trackId) {
      if (rowOfId[trackId]) {
        problems.push(`${name}: track ID ${trackId} is used by both row ${rowOfId[trackId]} and row ${row}.`);
      }
      rowOfId[trackId] = row;
    }

    songs.push({
      row: row,
      status: status,
      category: String(r[col.category]).trim(),
      coverKey: String(r[col.coverKey]).trim(),
      title: title,
      trackId: trackId
    });
  }

  // Row order, whatever order the sheet is in - category blocks and summaryLimit depend on it.
  songs.sort((a, b) => a.row - b.row);
  return { songs: songs, problems: problems };
}

/**
 * @return {number} The last song row in use (the highest Tracklist `row`).
 */
function getLastSongRow() {
  const songs = getSongs();
  return songs.length ? songs[songs.length - 1].row : CONFIG.LAYOUT.FIRST_SONG_ROW - 1;
}

/**
 * @return {number} How many rows the song range spans, first song row to last, gaps included.
 */
function getSongRowCount() {
  return getLastSongRow() - CONFIG.LAYOUT.FIRST_SONG_ROW + 1;
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

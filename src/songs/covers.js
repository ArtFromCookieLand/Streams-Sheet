/**
 * ===================================================================
 * COVERS FOR NEW COVER KEYS
 * -------------------------------------------------------------------
 * The Covers sheet maps a cover key (column A) to an image URL
 * (column C), with column B showing it (=IMAGE of C); Latest!D looks
 * each song's key up and shows column B. When a song comes
 * in with a key that isn't in Covers yet, a row is added for it:
 *
 *   1. the cover of the album the track was scraped from - the import
 *      keeps it in Import!I - or, failing that,
 *   2. the track's cover from Spotify's embed lookup
 *      (open.spotify.com/oembed: no login, no Apify credits).
 *
 * Existing keys are never changed - edit a URL in Covers by hand to use
 * a different image (e.g. the standard green TLOAS cover).
 * ===================================================================
 */

/**
 * [MENU] Adds a Covers row for every Tracklist song whose cover key is missing.
 */
function fillMissingCoversMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const songs = getSongs().filter(s => s.coverKey);
    const result = ensureCovers(songs.map(s => ({ coverKey: s.coverKey, trackId: s.trackId, title: s.title })));
    ui.alert('Fill Missing Covers' + envTag(), describeCovers(result) || 'Every cover key in the Tracklist is already in Covers.', ui.ButtonSet.OK);
  } catch (error) {
    console.error(error);
    ui.alert('Error', error.toString(), ui.ButtonSet.OK);
  }
}


/**
 * Adds a Covers row for each cover key that doesn't have one yet.
 * @param {Array<{coverKey: string, trackId: string, album: string, title: string}>} items
 *   album is optional; without it, the album is looked up in the import by track ID.
 * @return {{added: Array<{key, url, source}>, missing: Array<{key, title}>}}
 */
function ensureCovers(items) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.COVERS);
  if (!sheet) throw new Error(`The "${CONFIG.SHEETS.COVERS}" sheet is missing.`);
  const have = coverKeysInCoversSheet() || {};

  const raw = readRawImport();
  const albumOfTrack = {}, coverOfAlbum = {};
  raw.forEach(t => {
    albumOfTrack[t.id] = t.album;
    if (t.cover && !coverOfAlbum[t.album]) coverOfAlbum[t.album] = t.cover;
  });

  const result = { added: [], missing: [] };
  const done = {};
  items.forEach(it => {
    const key = String(it.coverKey || '').trim();
    const lower = key.toLowerCase();
    if (!key || have[lower] || done[lower]) return;
    done[lower] = true;

    const album = it.album || albumOfTrack[it.trackId] || '';
    let url = coverOfAlbum[album] || '', source = url ? `cover of "${album}"` : '';
    if (!url && it.trackId) {
      url = spotifyCoverOfTrack(it.trackId);
      if (url) source = 'Spotify embed lookup';
    }
    if (url) result.added.push({ key: key, url: url, source: source });
    else result.missing.push({ key: key, title: it.title || '' });
  });

  if (result.added.length) {
    const c = CONFIG.COVERS;
    const width = Math.max(c.KEY_COLUMN, c.URL_COLUMN);
    const rows = result.added.map(a => {
      const row = new Array(width).fill('');
      row[c.KEY_COLUMN - 1] = a.key;
      row[c.URL_COLUMN - 1] = a.url;
      return row;
    });
    // Straight under the last key, so a lookup over the Covers range picks them up.
    const lastKeyRow = lastNonBlankRow(sheet, c.KEY_COLUMN);
    const first = lastKeyRow + 1;
    sheet.getRange(first, 1, rows.length, width).setValues(rows);

    // Column B shows the image. Use the same formula as the row above (adjusted to each new row),
    // or =IMAGE() of the URL if there is none to copy.
    const above = lastKeyRow >= 1 ? sheet.getRange(lastKeyRow, c.IMAGE_COLUMN).getFormulaR1C1() : '';
    const urlOffset = c.URL_COLUMN - c.IMAGE_COLUMN;
    const imageFormula = above || `=IMAGE(RC[${urlOffset}])`;
    sheet.getRange(first, c.IMAGE_COLUMN, rows.length, 1).setFormulasR1C1(rows.map(() => [imageFormula]));
  }
  return result;
}


function describeCovers(result) {
  const lines = [];
  if (result.added.length) {
    lines.push(`🖼️ ${result.added.length} cover(s) added to ${CONFIG.SHEETS.COVERS}:`);
    result.added.forEach(a => lines.push(`   • ${a.key} - ${a.source}`));
  }
  if (result.missing.length) {
    lines.push(`No cover found for ${result.missing.length} key(s) - add them to ${CONFIG.SHEETS.COVERS} by hand:`);
    result.missing.forEach(m => lines.push(`   • ${m.key}${m.title ? ` (${m.title})` : ''}`));
  }
  return lines.join('\n');
}


/**
 * @return {string} The 300px cover of the track's album from Spotify's oEmbed endpoint, or '' if
 *   the lookup fails. Needs no login and costs no Apify credits.
 */
function spotifyCoverOfTrack(trackId) {
  try {
    const url = 'https://open.spotify.com/oembed?url=' + encodeURIComponent('https://open.spotify.com/track/' + trackId);
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return '';
    return JSON.parse(response.getContentText()).thumbnail_url || '';
  } catch (error) {
    console.warn('oEmbed lookup failed for ' + trackId + ': ' + error);
    return '';
  }
}


/** @return {Object|null} Lower-cased cover keys in the Covers sheet, or null if it has none. */
function coverKeysInCoversSheet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.COVERS);
  if (!sheet) return null;
  const last = lastNonBlankRow(sheet, CONFIG.COVERS.KEY_COLUMN);
  if (last < 1) return null;
  const keys = {};
  sheet.getRange(1, CONFIG.COVERS.KEY_COLUMN, last, 1).getValues().forEach(r => {
    const k = String(r[0]).trim().toLowerCase();
    if (k) keys[k] = true;
  });
  return keys;
}


/** @return {number} The last row with something in the given column (0 if none). */
function lastNonBlankRow(sheet, column) {
  const last = sheet.getLastRow();
  if (last < 1) return 0;
  const values = sheet.getRange(1, column, last, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) if (String(values[i][0]).trim() !== '') return i + 1;
  return 0;
}

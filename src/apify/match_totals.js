/**
 * ===================================================================
 * TODAY'S FIGURES FROM THE IMPORT
 * -------------------------------------------------------------------
 * The Import sheet (called Tools before 2.1) holds only the raw Apify
 * dump in E:I. Each song's total is looked up there by the Spotify track
 * ID from the Tracklist, and its daily is that total minus yesterday's
 * total in Latest!F. Nothing is kept per row in the Import sheet any
 * more, so it never has to line up with the other sheets.
 *
 *   active song, ID found    -> its stream count
 *   active song, ID missing  -> no figure; listed in `missing`, and
 *                               Update Daily Stats refuses to run
 *   upcoming song            -> its stream count once the ID is in the
 *                               import, 0 until then (never missing)
 *   retired song             -> 0
 *
 * The sum of the dailies is the freshness guard: it is 0 or less when
 * Spotify hasn't refreshed since the last update, and also right after
 * an update (Latest!F then holds these same totals), so the same import
 * can't be used twice.
 *
 * Everything here only reads the import, so it can be re-run at no
 * Apify cost - e.g. after correcting a track ID in the Tracklist sheet.
 * ===================================================================
 */

/**
 * @return {Sheet} The sheet holding the raw import. Falls back to its old name, so nothing breaks
 *   between pushing 2.1 and running Update → Setup → Tidy up the Import sheet.
 */
function getImportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.IMPORT) || ss.getSheetByName(CONFIG.SHEETS.IMPORT_OLD_NAME);
  if (!sheet) throw new Error(`The "${CONFIG.SHEETS.IMPORT}" sheet is missing.`);
  return sheet;
}

/** @return {Array<{album: string, name: string, count: number, id: string, cover: string}>} Rows that have an ID. */
function readRawImport() {
  return getImportSheet().getRange(CONFIG.IMPORT.RAW_DATA).getValues()
    .map(r => ({ album: String(r[0]), name: String(r[1]), count: Number(r[2]) || 0, id: String(r[3]).trim(), cover: String(r[4] || '').trim() }))
    .filter(t => t.id);
}


/**
 * [MENU] Re-checks the import already in the Import sheet: what matches, what's missing, and the
 * sum of today's dailies.
 */
function matchTotalsMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = matchTotalsById();
    ui.alert('Check Import' + envTag(), describeMatchResult(result), ui.ButtonSet.OK);
  } catch (error) {
    console.error(error);
    ui.alert('Error', error.toString(), ui.ButtonSet.OK);
  }
}


/**
 * Works out today's figures (readTodayFromImport) and writes their sum into Import!C1, where it is
 * shown for reference. The update never reads C1; it works the figures out again itself.
 * @return {Object} As readTodayFromImport().
 */
function matchTotalsById() {
  const today = readTodayFromImport();
  writeSumOfDailies(today.missing.length ? 'missing IDs' : today.sumOfDailies);
  return today;
}

function writeSumOfDailies(value) {
  getImportSheet().getRange(CONFIG.IMPORT.SUM_OF_DAILYS).setValue(value);
}


/**
 * Reads only. Throws if the import has no track IDs at all.
 * @return {{
 *   totals: Array, dailies: Array,   one entry per row from FIRST_SONG_ROW to the last song:
 *                                    a number, or '' for a row with no song or a missing ID
 *   sumOfDailies: number,
 *   matched: number, retired: number, upcoming: number,
 *   nowLive: Array<{row, title}>, missing: Array<{row, title}>
 * }}
 */
function readTodayFromImport() {
  const countById = {};
  readRawImport().forEach(t => { countById[t.id] = t.count; });
  if (Object.keys(countById).length === 0) {
    throw new Error(
      `The raw import in ${CONFIG.SHEETS.IMPORT}!${CONFIG.IMPORT.RAW_DATA} has no track IDs (column H is empty). ` +
      `Import again, or paste a dump that has IDs.`
    );
  }

  const first = CONFIG.LAYOUT.FIRST_SONG_ROW;
  const count = getSongRowCount();
  const previous = count > 0
    ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LATEST)
      .getRange(first, CONFIG.LATEST.COLS.TOTAL, count, 1).getValues().map(r => Number(r[0]) || 0)
    : [];

  const result = {
    totals: new Array(count).fill(''), dailies: new Array(count).fill(''), sumOfDailies: 0,
    matched: 0, retired: 0, upcoming: 0, nowLive: [], missing: []
  };

  getSongs().forEach(s => {
    const i = s.row - first;
    const found = s.trackId && countById.hasOwnProperty(s.trackId);
    let total;
    if (s.status === SONG_STATUS.RETIRED) {
      total = 0;
      result.retired++;
    } else if (s.status === SONG_STATUS.UPCOMING) {
      // Not out yet: 0 until its ID turns up, then counted like any other song.
      total = found ? countById[s.trackId] : 0;
      if (found) result.nowLive.push({ row: s.row, title: s.title });
      else result.upcoming++;
    } else if (found) {
      total = countById[s.trackId];
      result.matched++;
    } else {
      result.missing.push({ row: s.row, title: s.title });
      return;
    }
    result.totals[i] = total;
    result.dailies[i] = total - previous[i];
    result.sumOfDailies += total - previous[i];
  });

  return result;
}


/**
 * @param {Object} result - From readTodayFromImport().
 * @return {string} A message for a dialog.
 */
function describeMatchResult(result) {
  const held = [];
  if (result.retired) held.push(`${result.retired} retired`);
  if (result.upcoming) held.push(`${result.upcoming} upcoming`);
  let text = `Matched ${result.matched + result.nowLive.length} songs by track ID` +
    (held.length ? ` (${held.join(', ')}, held at 0).` : '.');

  if (!result.missing.length) {
    text += `\n\nSum of today's dailies: ${Math.round(result.sumOfDailies).toLocaleString('en-US')}` +
      (result.sumOfDailies > 0 ? '.' : ' - Spotify hasn\'t refreshed since the last update, so Update Daily Stats would refuse to run.');
  }

  if (result.nowLive.length) {
    const list = result.nowLive.map(m => `• row ${m.row}: ${m.title}`).join('\n');
    text += `\n\n🎉 ${result.nowLive.length} upcoming song(s) are now on Spotify and counted:\n${list}\n` +
      `Switch them to "active" in the ${CONFIG.SHEETS.SONGS} sheet once you're happy, so a missing ID is caught again.`;
  }

  if (result.missing.length) {
    const list = result.missing.map(m => `• row ${m.row}: ${m.title}`).join('\n');
    text += `\n\n⚠️ ${result.missing.length} track ID(s) were NOT in the import:\n${list}\n\n` +
      `"Update Daily Stats" will refuse to run until they are. ` +
      `Correct the track ID in the ${CONFIG.SHEETS.SONGS} sheet, then use "Check Import" - it needs no new import.`;
  }
  return text;
}

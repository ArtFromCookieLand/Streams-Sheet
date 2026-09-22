/**
 * ===================================================================
 * MATCH TOTALS BY TRACK ID
 * -------------------------------------------------------------------
 * Fills Tools!L (each song's total) from the raw import in Tools!E:H,
 * matching on the Spotify track ID from the Tracklist sheet. IDs survive
 * Spotify renaming a track, which matching on its title did not.
 *
 *   active song, ID found    -> its stream count
 *   active song, ID missing  -> CONFIG.TOOLS.MISSING_MARKER, which makes
 *                               C1 an error so main() refuses to run
 *   upcoming song            -> its stream count once the ID is in the
 *                               import, 0 until then (never a marker)
 *   retired song             -> 0
 *
 * It only reads the raw import, so it can be re-run at no Apify cost -
 * e.g. after correcting a track ID in the Tracklist sheet.
 * ===================================================================
 */

/**
 * [MENU] Re-runs the matching against the raw import already in Tools.
 */
function matchTotalsMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    const result = matchTotalsById();
    ui.alert('Match Totals by ID' + envTag(), describeMatchResult(result), ui.ButtonSet.OK);
  } catch (error) {
    console.error(error);
    ui.alert('Error', error.toString(), ui.ButtonSet.OK);
  }
}


/**
 * @return {{matched: number, retired: number, missing: Array<{row: number, title: string}>}}
 */
function matchTotalsById() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const toolsSheet = ss.getSheetByName(CONFIG.SHEETS.TOOLS);

  // --- 1. Index the raw import by track ID ---
  // Columns of the raw range: album, name, stream count, track ID.
  const raw = toolsSheet.getRange(CONFIG.TOOLS.RAW_DATA).getValues();
  const countById = {};
  raw.forEach(r => {
    const id = String(r[3]).trim();
    if (id) countById[id] = Number(r[2]);
  });

  if (Object.keys(countById).length === 0) {
    throw new Error(
      `The raw import in Tools!${CONFIG.TOOLS.RAW_DATA} has no track IDs (its last column is empty). ` +
      `It was probably written by the old importer - import again, or load a fixture that has IDs.`
    );
  }

  // --- 2. Work out each song's total ---
  const songs = getSongs();
  const totalByRow = {};
  const result = { matched: 0, retired: 0, upcoming: 0, nowLive: [], missing: [] };

  songs.forEach(s => {
    const found = s.trackId && countById.hasOwnProperty(s.trackId);
    if (s.status === SONG_STATUS.RETIRED) {
      totalByRow[s.row] = 0;
      result.retired++;
    } else if (s.status === SONG_STATUS.UPCOMING) {
      // Not out yet: 0 until its ID turns up, then counted like any other song.
      totalByRow[s.row] = found ? countById[s.trackId] : 0;
      if (found) result.nowLive.push({ row: s.row, title: s.title });
      else result.upcoming++;
    } else if (found) {
      totalByRow[s.row] = countById[s.trackId];
      result.matched++;
    } else {
      totalByRow[s.row] = CONFIG.TOOLS.MISSING_MARKER;
      result.missing.push({ row: s.row, title: s.title });
    }
  });

  // --- 3. Write Col L in one go ---
  // Rows that are not songs keep whatever they hold now, formulas included.
  const startRow = CONFIG.LAYOUT.FIRST_SONG_ROW;
  const range = toolsSheet.getRange(startRow, CONFIG.TOOLS.TOTALS_COLUMN, getSongRowCount(), 1);
  const currentFormulas = range.getFormulas();
  const currentValues = range.getValues();

  const output = currentValues.map((current, i) => {
    const row = startRow + i;
    if (totalByRow.hasOwnProperty(row)) return [totalByRow[row]];
    return [currentFormulas[i][0] || current[0]];
  });
  range.setValues(output);

  return result;
}


/**
 * @param {{matched: number, retired: number, missing: Array}} result - From matchTotalsById().
 * @return {string} A message for a dialog.
 */
function describeMatchResult(result) {
  const held = [];
  if (result.retired) held.push(`${result.retired} retired`);
  if (result.upcoming) held.push(`${result.upcoming} upcoming`);
  let text = `Matched ${result.matched + result.nowLive.length} songs by track ID` +
    (held.length ? ` (${held.join(', ')}, held at 0).` : '.');

  if (result.nowLive.length) {
    const list = result.nowLive.map(m => `• row ${m.row}: ${m.title}`).join('\n');
    text += `\n\n🎉 ${result.nowLive.length} upcoming song(s) are now on Spotify and counted:\n${list}\n` +
      `Switch them to "active" in the ${CONFIG.SHEETS.SONGS} sheet once you're happy, so a missing ID is caught again.`;
  }

  if (result.missing.length) {
    const list = result.missing.map(m => `• row ${m.row}: ${m.title}`).join('\n');
    text += `\n\n⚠️ ${result.missing.length} track ID(s) were NOT in the import:\n${list}\n\n` +
      `Their totals are marked ${CONFIG.TOOLS.MISSING_MARKER}, so "Update Daily Stats" will refuse to run. ` +
      `Correct the track ID in the ${CONFIG.SHEETS.SONGS} sheet, then use "Match Totals by ID" - it needs no new import.`;
  }
  return text;
}

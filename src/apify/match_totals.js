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
  const result = { matched: 0, retired: 0, missing: [] };

  songs.forEach(s => {
    if (s.status === SONG_STATUS.RETIRED) {
      totalByRow[s.row] = 0;
      result.retired++;
    } else if (s.trackId && countById.hasOwnProperty(s.trackId)) {
      totalByRow[s.row] = countById[s.trackId];
      result.matched++;
    } else {
      totalByRow[s.row] = CONFIG.TOOLS.MISSING_MARKER;
      result.missing.push({ row: s.row, title: s.title });
    }
  });

  // --- 3. Write Col L in one go ---
  // Rows that are not songs keep whatever they hold now, formulas included.
  const startRow = CONFIG.SONGS.START_ROW;
  const range = toolsSheet.getRange(startRow, CONFIG.TOOLS.TOTALS_COLUMN, CONFIG.SONGS.COUNT, 1);
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
  let text = `Matched ${result.matched} songs by track ID` +
    (result.retired ? ` (${result.retired} retired, held at 0).` : '.');

  if (result.missing.length) {
    const list = result.missing.map(m => `• row ${m.row}: ${m.title}`).join('\n');
    text += `\n\n⚠️ ${result.missing.length} track ID(s) were NOT in the import:\n${list}\n\n` +
      `Their totals are marked ${CONFIG.TOOLS.MISSING_MARKER}, so "Update Daily Stats" will refuse to run. ` +
      `Correct the track ID in the ${CONFIG.SHEETS.SONGS} sheet, then use "Match Totals by ID" - it needs no new import.`;
  }
  return text;
}

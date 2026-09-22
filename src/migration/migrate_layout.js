/**
 * ===================================================================
 * ONE-OFF LAYOUT MIGRATION (phase 1b)
 * -------------------------------------------------------------------
 * Moves every sheet from the old layout to the one in CONFIG.LAYOUT:
 *
 *            old                          new
 *   songs    rows 2-549                   rows 50 -> (CONFIG.LAYOUT.FIRST_SONG_ROW)
 *   totals   archive rows 550-574,        row 2 artist total, row 3 solo total,
 *            Latest album table T2:AB35   rows 4-27 categories (Latest: in the
 *                                         song columns D:P)
 *
 * It refuses to start unless the sheets are exactly in the old layout,
 * and records that it ran, so it can never run twice. It only ever runs
 * from the menu item, which disappears once the migration is done.
 *
 * Remove this file (and its menu item) once prod has been migrated.
 * ===================================================================
 */

const MIGRATION = {
  DONE_PROPERTY: 'LAYOUT_MIGRATED',
  OLD: {
    FIRST_SONG_ROW: 2,
    LAST_SONG_ROW: 549,
    FIRST_CATEGORY_ROW: 550,   // archives: 24 categories at 550-573...
    CATEGORY_COUNT: 24,
    TOTAL_ROW: 574,            // ...and the artist total at 574
    LATEST_FIRST_ROW: 2,       // Latest album table: categories T2:T25,
    LATEST_TOTAL_ROW: 26,      // artist total T26, solo total T27,
    LATEST_SOLO_ROW: 27,
    LATEST_LAST_ROW: 35        // special editions T28:T35
  },
  SOLO_LABEL: 'Total Artist Solo Streams'
};
// Rows inserted at the top: songs move from row 2 to CONFIG.LAYOUT.FIRST_SONG_ROW.
MIGRATION.SHIFT = CONFIG.LAYOUT.FIRST_SONG_ROW - MIGRATION.OLD.FIRST_SONG_ROW;

function isLayoutMigrated() {
  return !!PropertiesService.getDocumentProperties().getProperty(MIGRATION.DONE_PROPERTY);
}

/**
 * Run by hand from the Apps Script editor, only after restoring the sheets to the old layout from
 * Version history: that restores the sheets but not this record, so the migration would otherwise
 * refuse to run again. The migration's own checks still refuse if the sheets aren't the old layout.
 */
function forgetLayoutMigration() {
  PropertiesService.getDocumentProperties().deleteProperty(MIGRATION.DONE_PROPERTY);
  console.log('Migration record cleared. Reload the spreadsheet to see the menu item again.');
}


/**
 * [MENU] Checks, confirms, migrates, then runs the health checks on the new layout.
 */
function migrateLayoutMenu() {
  const ui = SpreadsheetApp.getUi();

  if (isLayoutMigrated()) {
    ui.alert('Already migrated', 'This spreadsheet was migrated on ' +
      PropertiesService.getDocumentProperties().getProperty(MIGRATION.DONE_PROPERTY) + '. It must not run twice.', ui.ButtonSet.OK);
    return;
  }

  const problems = findMigrationBlockers();
  if (problems.length) {
    ui.alert('Migration not started' + envTag(),
      'Nothing was changed. The sheets are not exactly in the old layout:\n\n• ' + problems.slice(0, 12).join('\n• ') +
      (problems.length > 12 ? `\n… and ${problems.length - 12} more` : ''), ui.ButtonSet.OK);
    return;
  }

  const response = ui.alert('Migrate layout' + envTag(),
    `This moves songs to row ${CONFIG.LAYOUT.FIRST_SONG_ROW} and the album totals to rows 2-27 in Latest, Tools, every Daily Archive and the Total Archive, and adds ${MIGRATION.SHIFT} to every Tracklist row.\n\n` +
    'It cannot be undone except from Version history. Have you created a named version (File → Version history → Name current version)?\n\n' +
    'Run it now?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast('Migrating... do not edit the spreadsheet.', 'Migration', -1);
  migrateLayout();
  PropertiesService.getDocumentProperties().setProperty(MIGRATION.DONE_PROPERTY, new Date().toISOString());
  SpreadsheetApp.flush();

  ss.toast('Migration done. Running checks...', 'Migration', 5);
  showCheckReport('Migration done - checks on the new layout', runChecks(HEALTH_CHECKS));
}


/**
 * Everything that has to be true before the migration may touch anything.
 * @return {Array<string>} Problems; empty means it is safe to run.
 */
function findMigrationBlockers() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = MIGRATION.OLD;
  const problems = [];
  const songCount = old.LAST_SONG_ROW - old.FIRST_SONG_ROW + 1;

  // --- Every sheet involved exists ---
  const names = [CONFIG.SHEETS.SONGS, CONFIG.SHEETS.LATEST, CONFIG.SHEETS.TOOLS, CONFIG.SHEETS.TOTAL_ARCHIVE]
    .concat(CONFIG.SHEETS.ARCHIVE_YEARS);
  names.forEach(n => { if (!ss.getSheetByName(n)) problems.push(`Sheet "${n}" not found.`); });
  if (problems.length) return problems;

  // --- Tracklist: still the old grid, where row N sits on sheet row N ---
  const trackValues = ss.getSheetByName(CONFIG.SHEETS.SONGS).getDataRange().getValues();
  let col;
  try { col = findSongColumns(trackValues[0]); } catch (e) { return [e.message]; }
  const titleOfRow = {};
  for (let i = 1; i < trackValues.length; i++) {
    const title = String(trackValues[i][col.title]).trim();
    if (!title) continue;
    const row = Number(trackValues[i][col.row]);
    if (row !== i + 1) problems.push(`${CONFIG.SHEETS.SONGS} sheet row ${i + 1} ("${title}") has row ${trackValues[i][col.row]}; before the migration every song must still sit on its own row.`);
    if (row > old.LAST_SONG_ROW) problems.push(`${CONFIG.SHEETS.SONGS}: "${title}" has row ${row}, past the old song range.`);
    titleOfRow[i + 1] = title;
  }

  // --- Rows line up: Tracklist = Latest!E = column A of every archive ---
  const sources = [{ name: CONFIG.SHEETS.LATEST, col: CONFIG.LATEST.COLS.TITLE }]
    .concat(CONFIG.SHEETS.ARCHIVE_YEARS.concat([CONFIG.SHEETS.TOTAL_ARCHIVE]).map(n => ({ name: n, col: 1 })));
  sources.forEach(src => {
    const titles = ss.getSheetByName(src.name).getRange(old.FIRST_SONG_ROW, src.col, songCount, 1).getValues();
    let bad = 0;
    titles.forEach((t, i) => {
      const row = old.FIRST_SONG_ROW + i;
      if (String(t[0]).trim() !== (titleOfRow[row] || '')) {
        if (bad++ < 3) problems.push(`Row ${row}: ${src.name} has "${String(t[0]).trim()}", ${CONFIG.SHEETS.SONGS} has "${titleOfRow[row] || ''}".`);
      }
    });
    if (bad > 3) problems.push(`${src.name}: ${bad - 3} more row(s) out of line.`);
  });

  // --- Daily Archives: the aggregate rows are where they should be, labelled the same everywhere ---
  const aggregateCount = old.CATEGORY_COUNT + 1;
  let reference = null;
  CONFIG.SHEETS.ARCHIVE_YEARS.forEach(n => {
    const labels = ss.getSheetByName(n).getRange(old.FIRST_CATEGORY_ROW, 1, aggregateCount, 1).getValues().map(r => String(r[0]).trim());
    if (labels.some(l => !l)) problems.push(`${n}: column A rows ${old.FIRST_CATEGORY_ROW}-${old.TOTAL_ROW} should all hold aggregate labels, but some are blank.`);
    if (!reference) reference = { name: n, labels: labels };
    else if (labels.join('|') !== reference.labels.join('|')) {
      problems.push(`${n}: the aggregate labels in rows ${old.FIRST_CATEGORY_ROW}-${old.TOTAL_ROW} differ from ${reference.name}.`);
    }
  });

  // --- Latest: the album table is in T2:AB35, with the solo total = artist total minus Features ---
  const latest = ss.getSheetByName(CONFIG.SHEETS.LATEST);
  const table = latest.getRange(old.LATEST_FIRST_ROW, 20, old.LATEST_SOLO_ROW - old.LATEST_FIRST_ROW + 1, 3).getValues(); // T:V
  table.forEach((r, i) => {
    if (!String(r[0]).trim()) problems.push(`Latest!T${old.LATEST_FIRST_ROW + i} is blank; the album table should run unbroken from T2 to T27.`);
  });
  const excluded = CONFIG.CATEGORIES.filter(c => c.name === CONFIG.LAYOUT.SOLO_EXCLUDES)[0];
  if (excluded) {
    const oldExcludedRow = excluded.row - 2;  // new categories 4-27 were Latest rows 2-25
    const at = r => table[r - old.LATEST_FIRST_ROW];
    [[1, 'total'], [2, 'daily']].forEach(([k, what]) => {
      const solo = Number(at(old.LATEST_SOLO_ROW)[k]);
      const expected = Number(at(old.LATEST_TOTAL_ROW)[k]) - Number(at(oldExcludedRow)[k]);
      if (solo !== expected) {
        problems.push(`Latest row ${old.LATEST_SOLO_ROW} ${what} (${solo}) is not row ${old.LATEST_TOTAL_ROW} minus ${excluded.name} (${expected}), so the archives' new solo row would disagree with Latest.`);
      }
    });
  }

  return problems;
}


/**
 * The migration itself. Only call it after findMigrationBlockers() came back empty.
 */
function migrateLayout() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const old = MIGRATION.OLD;
  const shift = MIGRATION.SHIFT;
  const layout = CONFIG.LAYOUT;
  const firstCategoryRow = layout.SOLO_ROW + 1;

  // --- 1. Tracklist: every song's row moves down by the shift. No rows move in the sheet itself. ---
  const tracklist = ss.getSheetByName(CONFIG.SHEETS.SONGS);
  const values = tracklist.getDataRange().getValues();
  const col = findSongColumns(values[0]);
  const rowColumn = [];
  for (let i = 1; i < values.length; i++) {
    const v = values[i][col.row];
    rowColumn.push([String(values[i][col.title]).trim() && v !== '' ? Number(v) + shift : v]);
  }
  if (rowColumn.length) tracklist.getRange(2, col.row + 1, rowColumn.length, 1).setValues(rowColumn);
  _songsCache = null;

  // --- 2. Insert the shift's worth of rows at the top of Latest, every Daily Archive and the Total Archive. ---
  // Sheets shifts every reference to the rows below, in every sheet.
  const archives = CONFIG.SHEETS.ARCHIVE_YEARS.map(n => ss.getSheetByName(n));
  const totalArchive = ss.getSheetByName(CONFIG.SHEETS.TOTAL_ARCHIVE);
  const latest = ss.getSheetByName(CONFIG.SHEETS.LATEST);
  [latest, totalArchive].concat(archives).forEach(sheet => sheet.insertRowsBefore(2, shift));

  // --- 3. Tools: shift only J:M down, so the raw import in E:H stays at row 2. ---
  const tools = ss.getSheetByName(CONFIG.SHEETS.TOOLS);
  tools.getRange(2, 10, shift, 4).insertCells(SpreadsheetApp.Dimension.ROWS);

  // --- 4. Daily Archives: aggregate rows from the bottom to rows 2-27, history included. ---
  const excluded = CONFIG.CATEGORIES.filter(c => c.name === layout.SOLO_EXCLUDES)[0];
  archives.forEach(sheet => {
    const lastCol = sheet.getLastColumn();
    sheet.getRange(old.FIRST_CATEGORY_ROW + shift, 1, old.CATEGORY_COUNT, lastCol).moveTo(sheet.getRange(firstCategoryRow, 1));
    sheet.getRange(old.TOTAL_ROW + shift, 1, 1, lastCol).moveTo(sheet.getRange(layout.TOTAL_ROW, 1));

    // The solo row is new: artist total minus Features, in every column that has a date.
    sheet.getRange(layout.SOLO_ROW, 1).setValue(MIGRATION.SOLO_LABEL);
    if (lastCol > 1) {
      const soloFormulas = [];
      for (let c = 2; c <= lastCol; c++) {
        const L = columnToLetter(c);
        soloFormulas.push(`=${L}${layout.TOTAL_ROW}-${L}${excluded.row}`);
      }
      sheet.getRange(layout.SOLO_ROW, 2, 1, lastCol - 1).setFormulas([soloFormulas]).setNumberFormat('#,##0');
    }
  });
  // Today's column gets freshly generated formulas, so it matches exactly what future days get.
  setAlbumFormulas(ss.getSheetByName(CONFIG.SHEETS.ARCHIVE), buildAlbumFormulas('B'));

  // --- 5. Latest: the album table from T:AB (pushed down by the insert) into the song columns. ---
  // Moved in blocks, so ranges lying inside a block (the Albums chart's S2:W17) follow it.
  const from = r => r + shift;  // old row -> where the insert put it
  const moveAlbumRows = (oldFirst, count, newFirst) => {
    const r = from(oldFirst);
    latest.getRange(r, 19, count, 6).moveTo(latest.getRange(newFirst, 4));    // S:X  -> D:I
    latest.getRange(r, 26, count, 2).moveTo(latest.getRange(newFirst, 10));   // Z:AA -> J:K
    latest.getRange(r, 25, count, 1).moveTo(latest.getRange(newFirst, 12));   // Y    -> L
    latest.getRange(r, 28, count, 1).moveTo(latest.getRange(newFirst, 16));   // AB   -> P
  };
  moveAlbumRows(old.LATEST_FIRST_ROW, old.CATEGORY_COUNT, firstCategoryRow);  // categories -> 4-27
  moveAlbumRows(old.LATEST_TOTAL_ROW, 2, layout.TOTAL_ROW);                   // total, solo -> 2-3
  // The special-edition rows go straight back to where they were, beside the spare rows.
  const specials = old.LATEST_LAST_ROW - old.LATEST_SOLO_ROW;
  latest.getRange(from(old.LATEST_SOLO_ROW + 1), 19, specials, 10).moveTo(latest.getRange(old.LATEST_SOLO_ROW + 1, 19));

  // --- 6. Total Archive: aggregate rows for the first time, in every existing column. ---
  const lastTotalCol = totalArchive.getLastColumn();
  const labels = archives[0].getRange(layout.TOTAL_ROW, 1, firstCategoryRow + old.CATEGORY_COUNT - layout.TOTAL_ROW, 1).getValues();
  totalArchive.getRange(layout.TOTAL_ROW, 1, labels.length, 1).setValues(labels);
  if (lastTotalCol > 1) {
    const perColumn = [];
    for (let c = 2; c <= lastTotalCol; c++) perColumn.push(buildAlbumFormulas(columnToLetter(c)).map(r => r[0]));
    // perColumn is [column][row]; the range wants [row][column].
    const matrix = perColumn[0].map((_, r) => perColumn.map(colFormulas => colFormulas[r]));
    totalArchive.getRange(layout.TOTAL_ROW, 2, matrix.length, lastTotalCol - 1).setFormulas(matrix).setNumberFormat('#,##0');
  }
}


/**
 * @param {number} n - 1-based column number.
 * @return {string} e.g. 1 -> "A", 28 -> "AB".
 */
function columnToLetter(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

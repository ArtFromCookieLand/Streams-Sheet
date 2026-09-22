/**
 * ===================================================================
 * ADD NEW CATEGORIES
 * -------------------------------------------------------------------
 * To add an album or category: put its name, type and (optionally) its
 * summary cell and limit in the Categories sheet, leave `row` empty, and
 * run Update → Add New Categories.
 *
 * It opens a row for it in the aggregate block (rows 4 up to
 * LAYOUT.LAST_AGGREGATE_ROW) of Latest, every Daily Archive and the
 * Total Archive, placing it after the last category of its own type:
 * a studio album after the last studio album, an "other" after the last
 * compilation or live album, a "fixed" at the very end.
 *
 * Every category below it moves down one row, and one spare row above
 * the songs is used up, so the songs never move. The Categories sheet's
 * row numbers are renumbered to match.
 * ===================================================================
 */

/**
 * [MENU] Adds every Categories row that has no row number yet.
 */
function addCategoriesMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const read = readCategories();
  if (read.problems.length) {
    ui.alert('Add New Categories' + envTag(), `Fix the ${CONFIG.SHEETS.CATEGORIES} sheet first:\n\n• ` + read.problems.join('\n• '), ui.ButtonSet.OK);
    return;
  }
  if (!read.pending.length) {
    ui.alert('Add New Categories' + envTag(),
      `Nothing waiting. To add one, put its name and type in the ${CONFIG.SHEETS.CATEGORIES} sheet and leave "${CONFIG.CATEGORIES_SHEET.HEADERS.row}" empty, then run this again.`,
      ui.ButtonSet.OK);
    return;
  }

  // --- 1. Check the batch ---
  const plan = planCategories(read);
  if (plan.problems.length) {
    ui.alert('Nothing was added' + envTag(), 'Fix these first:\n\n• ' + plan.problems.join('\n• '), ui.ButtonSet.OK);
    return;
  }

  const list = plan.entries.map(e => `• ${e.name} (${e.type}) → row ${e.row}`).join('\n');
  const response = ui.alert(`Add ${plan.entries.length} categor${plan.entries.length === 1 ? 'y' : 'ies'}?` + envTag(),
    `${list}\n\nEvery category below moves down a row in Latest and the archives, using up ${plan.entries.length} of the ${plan.spare} spare row(s). Songs stay where they are.`,
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  // --- 2. Add them, top row first so the later ones' rows still hold ---
  ss.toast('Adding categories... do not edit the spreadsheet.', 'Add New Categories', -1);
  let failed = null;
  const added = [];
  for (const entry of plan.entries) {
    try {
      insertCategoryRow(entry);
      added.push(entry);
    } catch (error) {
      console.error(error);
      failed = { entry: entry, error: error };
      break;
    }
  }
  _categoriesCache = null;

  // --- 3. Formulas that depend on the categories ---
  let formulaText = '';
  try {
    setAlbumFormulas(ss.getSheetByName(CONFIG.SHEETS.ARCHIVE), buildAlbumFormulas('B'));
    setLatestAggregateFormulas();
    refreshTotalArchiveFormulas();
  } catch (error) {
    formulaText = '\n\n⚠️ Rebuilding the total formulas failed: ' + error.message;
  }
  SpreadsheetApp.flush();

  if (failed) {
    ui.alert('Stopped partway' + envTag(),
      `Adding "${failed.entry.name}" failed: ${failed.error.message}\n\nRun Checks → Run All Checks. If rows are out of line, restore from Version history.`,
      ui.ButtonSet.OK);
    return;
  }

  // --- 4. Report ---
  const checks = runChecks(HEALTH_CHECKS.filter(c => ['Row alignment', 'Categories', 'Totals add up'].indexOf(c.name) !== -1));
  const todo = added.map(e => `"${e.name}" (row ${e.row}): add its songs through the ${CONFIG.SHEETS.PENDING} sheet.`);
  todo.push(`Add each new album's Spotify URL to the ${CONFIG.SHEETS.SOURCES} sheet, so the import scrapes it.`);
  if (added.some(e => e.type === 'studio')) {
    todo.push('The Albums sheet chart over the studio rows does not grow by itself - widen its range by one row per new studio album.');
  }
  added.filter(e => e.summaryCell).forEach(e => todo.push(`Make room on the Albums sheet for "${e.name}"'s summary at ${e.summaryCell} (10 rows).`));

  ui.alert(`Added ${added.length} categor${added.length === 1 ? 'y' : 'ies'}` + envTag(),
    added.map(e => `✅ ${e.name} → row ${e.row}`).join('\n') + formulaText + '\n\n' +
    formatCheckReport(checks) + '\n\nStill to do by hand:\n• ' + todo.join('\n• '),
    ui.ButtonSet.OK);
}


/**
 * Works out each pending category's row, as if the earlier ones were already in. Reads only.
 * @return {{entries: Array, problems: Array<string>, spare: number}}
 */
function planCategories(read) {
  const layout = CONFIG.LAYOUT;
  const problems = [];
  const categories = read.categories.slice();
  const lastRow = () => Math.max.apply(null, categories.map(c => c.row));
  const spare = layout.LAST_AGGREGATE_ROW - lastRow();

  if (read.pending.length > spare) {
    problems.push(`Only ${spare} spare aggregate row(s) left (up to row ${layout.LAST_AGGREGATE_ROW}), and ${read.pending.length} categor(y/ies) are waiting.`);
  }
  const alignment = runChecks(HEALTH_CHECKS.filter(c => c.name === 'Row alignment'))[0];
  if (alignment.status === CHECK_STATUS.FAIL) problems.push('Rows are out of line - run Checks → Check Row Alignment and fix that first.');

  // The spare rows are about to be used up, so make sure they really are empty everywhere.
  const needed = Math.min(read.pending.length, spare);
  if (needed > 0) {
    const first = layout.LAST_AGGREGATE_ROW - needed + 1;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const latest = ss.getSheetByName(CONFIG.SHEETS.LATEST);
    if (latest.getRange(first, 1, needed, 16).getValues().some(r => r.some(v => v !== ''))) {
      problems.push(`Latest rows ${first}-${layout.LAST_AGGREGATE_ROW} (A:P) should be empty spare rows, but something is in them.`);
    }
    CONFIG.SHEETS.ARCHIVE_YEARS.concat([CONFIG.SHEETS.TOTAL_ARCHIVE]).forEach(name => {
      const sheet = ss.getSheetByName(name);
      if (!sheet) { problems.push(`Sheet "${name}" not found.`); return; }
      if (sheet.getRange(first, 1, needed, Math.max(sheet.getLastColumn(), 1)).getValues().some(r => r.some(v => v !== ''))) {
        problems.push(`${name} rows ${first}-${layout.LAST_AGGREGATE_ROW} should be empty spare rows, but something is in them.`);
      }
    });
  }

  const entries = [];
  read.pending.forEach(p => {
    if (CATEGORY_TYPES.indexOf(p.type) === -1) {
      problems.push(`"${p.name}" has type "${p.type}"; expected one of: ${CATEGORY_TYPES.join(', ')}.`);
      return;
    }
    if (p.summaryCell && !/^[A-Z]+\d+$/.test(p.summaryCell)) {
      problems.push(`"${p.name}" has summary cell "${p.summaryCell}", which is not a single cell like F23.`);
      return;
    }
    // After the last category of its own type, or of any type above it.
    const order = CATEGORY_TYPES.indexOf(p.type);
    const above = categories.filter(c => CATEGORY_TYPES.indexOf(c.type) <= order);
    const row = above.length ? Math.max.apply(null, above.map(c => c.row)) + 1 : layout.SOLO_ROW + 1;

    // Everything from that row down shifts by one.
    categories.forEach(c => { if (c.row >= row) c.row++; });
    categories.push({ name: p.name, row: row, type: p.type, summaryCell: p.summaryCell, summaryLimit: p.summaryLimit });
    entries.push(Object.assign({}, p, { row: row }));
  });

  if (entries.length) {
    const top = Math.min.apply(null, entries.map(e => e.row));
    mergesInTheWay(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.LATEST), top, 1, 16)
      .forEach(p => problems.push(p));
  }

  return { entries: entries, problems: problems, spare: spare };
}


/**
 * Opens the row in every sheet, writes the category's title, and renumbers the Categories sheet.
 */
function insertCategoryRow(entry) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const row = entry.row;
  const spareRow = CONFIG.LAYOUT.LAST_AGGREGATE_ROW;   // the blank row used up, just above the songs
  const latest = ss.getSheetByName(CONFIG.SHEETS.LATEST);
  const archives = CONFIG.SHEETS.ARCHIVE_YEARS.map(n => ss.getSheetByName(n)).concat([ss.getSheetByName(CONFIG.SHEETS.TOTAL_ARCHIVE)]);

  // Archives: a whole row in, the spare row out, so the songs end up where they started.
  archives.forEach(sheet => {
    sheet.insertRowBefore(row);
    sheet.deleteRow(spareRow + 1);
    sheet.getRange(row, 1).setValue(entry.name);
  });

  // Latest: only the song columns move, so the special-edition table in T:AB stays put.
  latest.getRange(row, 1, 1, 16).insertCells(SpreadsheetApp.Dimension.ROWS);
  latest.getRange(spareRow + 1, 1, 1, 16).deleteCells(SpreadsheetApp.Dimension.ROWS);
  const template = row > CONFIG.LAYOUT.TOTAL_ROW ? row - 1 : row + 1;
  copyFormulasOnly(latest, template, row, 1, CONFIG.LATEST.COPIED_COLUMNS);
  latest.getRange(row, CONFIG.LATEST.COLS.TITLE).setValue(entry.name);

  // The Categories sheet: this one's row, and +1 for everything at or below it.
  const sheet = ss.getSheetByName(CONFIG.SHEETS.CATEGORIES);
  const values = sheet.getDataRange().getValues();
  const col = findHeaderColumns(values[0], CONFIG.CATEGORIES_SHEET.HEADERS, CONFIG.SHEETS.CATEGORIES);
  const rows = values.slice(1).map((r, i) => {
    if (i + 2 === entry.sheetRow) return [row];
    const v = Number(r[col.row]);
    return [Number.isInteger(v) && v >= row ? v + 1 : r[col.row]];
  });
  sheet.getRange(2, col.row + 1, rows.length, 1).setValues(rows);
  _categoriesCache = null;
}


/**
 * Rewrites the aggregate formulas in every column of the Total Archive, so a new category (or a
 * song added to one) is counted in its cumulative totals as well, back through the history.
 */
function refreshTotalArchiveFormulas() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.TOTAL_ARCHIVE);
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  if (lastCol < 2) return;

  const perColumn = [];
  for (let c = 2; c <= lastCol; c++) perColumn.push(buildAlbumFormulas(columnToLetter(c)).map(r => r[0]));
  const matrix = perColumn[0].map((_, r) => perColumn.map(colFormulas => colFormulas[r]));
  sheet.getRange(CONFIG.LAYOUT.TOTAL_ROW, 2, matrix.length, lastCol - 1).setFormulas(matrix);
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

/**
 * ===================================================================
 * MOVE A SONG
 * -------------------------------------------------------------------
 * Update → Move a Song opens a dialog: pick a song by its row, the row
 * it should end up on, and (optionally) a new category. The song moves
 * in Latest, every Daily Archive and the Total Archive, and the songs
 * in between shift by one row, exactly as if it had been cut and pasted
 * in each sheet - so formulas that point at it (the Albums breakdowns,
 * for instance) follow it.
 *
 *   Latest            A:P only (insert cells at the target, cut and
 *                     paste the song there, delete the cells it left),
 *                     the same block Add Pending Songs shifts
 *   archives          whole rows (Sheet.moveRows), history included
 *   Tracklist         the rows in between are renumbered, the song gets
 *                     its new row and category
 *
 * When the category changes, the song's past streams move with it in
 * every archive ("retroactive", decided 2026-09-24): a category's history
 * always means "the songs it has today". See rebuildAllAggregates(); the
 * same step is on the menu as Update → Rebuild Album History, for after
 * a category is changed by hand in the Tracklist.
 * ===================================================================
 */

/**
 * [MENU] Opens the Move a Song dialog.
 */
function moveSongMenu() {
  const ui = SpreadsheetApp.getUi();
  const html = HtmlService.createHtmlOutput(moveSongDialogHtml())
    .setWidth(560)
    .setHeight(520);
  ui.showModalDialog(html, 'Move a Song' + envTag());
}


/**
 * [DIALOG] Everything the dialog needs to preview a move without asking again.
 * @return {{songs: Array, categories: Array<string>, firstRow: number, lastRow: number, env: string}}
 */
function getMoveSongData() {
  const songs = getSongs().map(s => ({ row: s.row, title: s.title, category: s.category, status: s.status }));
  return {
    songs: songs,
    categories: getCategories().map(c => c.name),
    firstRow: CONFIG.LAYOUT.FIRST_SONG_ROW,
    lastRow: getLastSongRow(),
    env: envTag()
  };
}


/**
 * [DIALOG] Validates again, moves, rebuilds, checks. Never throws: the dialog shows the text.
 * @param {number} from - The song's row now.
 * @param {number} to - The row it should end up on.
 * @param {string} category - Its category afterwards ('' for none).
 * @return {{ok: boolean, message: string}}
 */
function moveSongFromDialog(from, to, category) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) return { ok: false, message: 'Another change is running on the spreadsheet. Try again in a moment.' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let toasted = false;
  try {
    const plan = planMoveSong(Number(from), Number(to), String(category || '').trim());
    if (plan.problems.length) return { ok: false, message: 'Nothing was moved:\n\n• ' + plan.problems.join('\n• ') };

    ss.toast('Moving the song... do not edit the spreadsheet.', 'Move a Song', -1);
    toasted = true;
    try {
      moveSongRows(plan.from, plan.to);
      updateTracklistForMove(plan);
    } catch (error) {
      console.error(error);
      return {
        ok: false,
        message: `Moving stopped partway: ${error.message}\n\nRun Checks → Check Row Alignment. If rows are out of line, restore from Version history.`
      };
    }
    _songsCache = null;

    const lines = [plan.from === plan.to
      ? `✅ "${plan.song.title}" stays on row ${plan.to}, now in ${plan.category || 'no category'}.`
      : `✅ "${plan.song.title}" moved from row ${plan.from} to row ${plan.to}` +
        (plan.category !== plan.song.category ? `, now in ${plan.category || 'no category'}` : '') + '.'];
    lines.push(describeHistoryRebuild(rebuildAllAggregates()));
    plan.notes.forEach(n => lines.push('• ' + n));
    const checks = runChecks(HEALTH_CHECKS.filter(c => ['Row alignment', 'Totals add up', 'Aggregate formulas', 'Category blocks'].indexOf(c.name) !== -1));
    lines.push(formatCheckReport(checks));
    SpreadsheetApp.flush();
    return { ok: true, message: lines.join('\n\n') };
  } catch (error) {
    console.error(error);
    return { ok: false, message: 'Error: ' + error.message };
  } finally {
    // A toast with no timeout stays until another one replaces it.
    if (toasted) ss.toast('Finished - see the dialog.', 'Move a Song', 3);
    lock.releaseLock();
  }
}


/**
 * Checks a move and works out what it means. Reads only.
 * @return {{from, to, category, song, problems: Array<string>, notes: Array<string>}}
 */
function planMoveSong(from, to, category) {
  const plan = { from: from, to: to, category: category, song: null, problems: [], notes: [] };
  const problems = plan.problems;
  let songs, categories;
  try { songs = getSongs(); } catch (e) { problems.push(`${CONFIG.SHEETS.SONGS}: ${e.message}`); }
  try { categories = getCategories(); } catch (e) { problems.push(`${CONFIG.SHEETS.CATEGORIES}: ${e.message}`); }
  if (problems.length) return plan;

  const first = CONFIG.LAYOUT.FIRST_SONG_ROW, last = getLastSongRow();
  if (!Number.isInteger(from) || !Number.isInteger(to)) {
    problems.push('Both rows must be whole numbers.');
    return plan;
  }
  plan.song = songs.filter(s => s.row === from)[0] || null;
  if (!plan.song) problems.push(`Row ${from} has no song in the ${CONFIG.SHEETS.SONGS} sheet.`);
  if (to < first || to > last) problems.push(`Row ${to} is outside the songs (rows ${first}-${last}).`);
  if (category && !categories.some(c => c.name === category)) problems.push(`Category "${category}" is not in the ${CONFIG.SHEETS.CATEGORIES} sheet.`);
  if (plan.song && from === to && category === plan.song.category) problems.push('That is where the song already is, in that category.');
  if (problems.length) return plan;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [CONFIG.SHEETS.LATEST, CONFIG.SHEETS.TOTAL_ARCHIVE].concat(CONFIG.SHEETS.ARCHIVE_YEARS)
    .forEach(n => { if (!ss.getSheetByName(n)) problems.push(`Sheet "${n}" not found.`); });
  if (problems.length) return plan;

  // The rows must line up before anything moves, or the move would make it worse.
  const alignment = runChecks(HEALTH_CHECKS.filter(c => c.name === 'Row alignment'))[0];
  if (alignment.status === CHECK_STATUS.FAIL) problems.push('Rows are out of line - run Checks → Check Row Alignment and fix that first.');
  if (from !== to) {
    mergesInTheWay(ss.getSheetByName(CONFIG.SHEETS.LATEST), Math.min(from, to), 1, CONFIG.LATEST.SHIFTED_COLUMNS).forEach(p => problems.push(p));
  }

  // Worth knowing, but not a reason to stop.
  const after = songsAfterMove(songs, from, to, category);
  if (category && !isContiguous(after.filter(s => s.category === category).map(s => s.row))) {
    plan.notes.push(`${category}'s songs are no longer on consecutive rows. That works (its sum covers them wherever they are), but new songs are added after its last row.`);
  }
  const cat = categories.filter(c => c.name === category)[0];
  if (cat && cat.summaryLimit) {
    const position = after.filter(s => s.category === category).map(s => s.row).indexOf(to) + 1;
    if (position > cat.summaryLimit) plan.notes.push(`It is song ${position} of ${category}, past its summary limit (${cat.summaryLimit}), so the album summary won't name it.`);
  }
  return plan;
}


/**
 * Where every song will be after the move, as a plain list (reads nothing).
 */
function songsAfterMove(songs, from, to, category) {
  return songs.map(s => {
    if (s.row === from) return Object.assign({}, s, { row: to, category: category });
    return Object.assign({}, s, { row: shiftedRow(s.row, from, to) });
  }).sort((a, b) => a.row - b.row);
}

/** @return {number} Where a row that is not the moved one ends up. */
function shiftedRow(row, from, to) {
  if (from < to && row > from && row <= to) return row - 1;
  if (to < from && row >= to && row < from) return row + 1;
  return row;
}

function isContiguous(rows) {
  const sorted = rows.slice().sort((a, b) => a - b);
  return sorted.every((r, i) => i === 0 || r === sorted[i - 1] + 1);
}


/**
 * Moves one row in Latest (A:P), every Daily Archive and the Total Archive. Every sheet's row
 * `from` ends up on row `to`, and the rows in between shift one towards `from`.
 */
function moveSongRows(from, to) {
  if (from === to) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const width = CONFIG.LATEST.SHIFTED_COLUMNS;

  // Latest: only A:P shifts, as with Add Pending Songs. Open a gap where the song goes, cut and
  // paste it in, then close the gap it left. References to it follow it, as they would by hand.
  const latest = ss.getSheetByName(CONFIG.SHEETS.LATEST);
  if (from < to) {
    latest.getRange(to + 1, 1, 1, width).insertCells(SpreadsheetApp.Dimension.ROWS);
    latest.getRange(from, 1, 1, width).moveTo(latest.getRange(to + 1, 1, 1, width));
    latest.getRange(from, 1, 1, width).deleteCells(SpreadsheetApp.Dimension.ROWS);
  } else {
    latest.getRange(to, 1, 1, width).insertCells(SpreadsheetApp.Dimension.ROWS);
    latest.getRange(from + 1, 1, 1, width).moveTo(latest.getRange(to, 1, 1, width));
    latest.getRange(from + 1, 1, 1, width).deleteCells(SpreadsheetApp.Dimension.ROWS);
  }

  // Archives: whole rows, so every day of the song's history goes with it.
  CONFIG.SHEETS.ARCHIVE_YEARS.concat([CONFIG.SHEETS.TOTAL_ARCHIVE]).forEach(name => {
    const sheet = ss.getSheetByName(name);
    sheet.moveRows(sheet.getRange(from + ':' + from), from < to ? to + 1 : to);
  });
}


/**
 * Renumbers the Tracklist for a move and sets the song's new row and category.
 */
function updateTracklistForMove(plan) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SONGS);
  const values = sheet.getDataRange().getValues();
  const col = findSongColumns(values[0]);
  const rows = [], cats = [];
  values.slice(1).forEach(r => {
    const title = String(r[col.title]).trim();
    const row = Number(r[col.row]);
    if (!title || !Number.isInteger(row)) { rows.push([r[col.row]]); cats.push([r[col.category]]); return; }
    if (row === plan.from) { rows.push([plan.to]); cats.push([plan.category]); return; }
    rows.push([shiftedRow(row, plan.from, plan.to)]);
    cats.push([r[col.category]]);
  });
  if (!rows.length) return;
  sheet.getRange(2, col.row + 1, rows.length, 1).setValues(rows);
  sheet.getRange(2, col.category + 1, cats.length, 1).setValues(cats);
  _songsCache = null;
}


// --- album history ---
//
// Past archive columns hold the album figures as plain numbers, not formulas: 365 x 26 SUMs per
// sheet made it lag (only today's column B keeps its formulas). So a figure can't be recalculated
// from the songs without risk; instead, when a song changes category, its streams are moved from
// the old category's figure to the new one's, day by day, and every other number stays as it was.
// The Tracklist's historyCategory column records which category the history counts each song in.

/**
 * [MENU] Brings the album history in line with the Tracklist's categories.
 */
function rebuildAlbumHistoryMenu() {
  const ui = SpreadsheetApp.getUi();
  try {
    buildAlbumFormulas();   // throws on a Tracklist or Categories problem, before anything is written
  } catch (error) {
    ui.alert('Rebuild Album History' + envTag(), error.message, ui.ButtonSet.OK);
    return;
  }
  const changes = categoryHistoryChanges();
  const what = changes === null
    ? `The ${CONFIG.SHEETS.SONGS} sheet has no "${CONFIG.SONGS_SHEET.HISTORY_CATEGORY}" column yet. This run creates it, filled with each song's current category, and adjusts nothing. ` +
      'If you already changed a category before, set that song\'s old category in the new column and run this again.'
    : changes.length
      ? `${changes.length} song(s) changed category since the history was last adjusted:\n` + describeChanges(changes) +
        '\n\nTheir streams move from the old category to the new one on every past day. Nothing else changes.'
      : 'No song has changed category since the history was last adjusted, so no figures change.';
  const response = ui.alert('Rebuild Album History' + envTag(),
    what + `\n\nToday's album formulas in ${CONFIG.SHEETS.ARCHIVE} column B and in Latest are rewritten either way. Go ahead?`,
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  const result = rebuildAllAggregates();
  const checks = runChecks(HEALTH_CHECKS.filter(c => ['Totals add up', 'Aggregate formulas', 'Category history'].indexOf(c.name) !== -1));
  ui.alert('Rebuild Album History' + envTag(), describeHistoryRebuild(result) + '\n\n' + formatCheckReport(checks), ui.ButtonSet.OK);
}


/**
 * @return {Array<{row, title, from, to}>|null} Songs whose category differs from the category the
 *   history counts them in; null when the Tracklist has no historyCategory column.
 */
function categoryHistoryChanges() {
  const songs = getSongs();
  if (songs.length && songs[0].historyCategory === null) return null;
  return songs.filter(s => s.historyCategory !== s.category)
    .map(s => ({ row: s.row, title: s.title, from: s.historyCategory, to: s.category }));
}

function describeChanges(changes) {
  const lines = changes.slice(0, 12).map(c => `• row ${c.row} "${c.title}": ${c.from || 'no category'} → ${c.to || 'no category'}`);
  if (changes.length > 12) lines.push(`… and ${changes.length - 12} more`);
  return lines.join('\n');
}


/**
 * Brings the aggregates everywhere in line with the Tracklist: moves the history of every song
 * that changed category, rewrites the formulas (today's column, Latest's album rows), records the
 * new categories as the history's, and refreshes Latest's day-ago, week-ago and best-since figures.
 * @return {{changes: Array|null, created: boolean, sheets: Array<Object>}}
 */
function rebuildAllAggregates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const changes = categoryHistoryChanges();
  const categories = getCategories();

  // Work everything out before writing anything, so a problem stops the run with nothing half-done.
  const plans = CONFIG.SHEETS.ARCHIVE_YEARS.concat([CONFIG.SHEETS.TOTAL_ARCHIVE])
    .map(name => ss.getSheetByName(name))
    .filter(sheet => sheet)
    .map(sheet => planAggregateHistory(sheet, changes || [], categories));
  plans.forEach(p => { if (p.range) p.range.setValues(p.out); });

  const created = changes === null;
  recordHistoryCategories();
  setLatestAggregateFormulas();
  // Latest's M, N and L for the albums were read from the old history.
  if (changes && changes.length) updateStats();
  return { changes: changes, created: created, sheets: plans.map(p => p.result) };
}


/**
 * Works out one archive's new aggregate block (rows 2 to the last category, every dated column).
 * A formula cell gets the formula buildAlbumFormulas() gives today. A number is adjusted by the
 * streams of the songs that changed category, on that day; a blank stays blank.
 * @return {{range: Range|null, out: Array, result: {sheet, columns, adjusted}}}
 */
function planAggregateHistory(sheet, changes, categories) {
  const result = { sheet: sheet.getName(), columns: 0, adjusted: 0 };
  const lastCol = sheet.getLastColumn();
  if (lastCol < 2) return { range: null, result: result };

  // The dated columns: from B up to the first blank date.
  const header = sheet.getRange(1, 2, 1, lastCol - 1).getValues()[0];
  let count = header.findIndex(v => v === '' || v === null);
  if (count === -1) count = header.length;
  if (count === 0) return { range: null, result: result };

  const layout = CONFIG.LAYOUT;
  const top = layout.TOTAL_ROW;
  const height = getLastCategoryRow() - top + 1;
  const range = sheet.getRange(top, 2, height, count);
  const formulas = range.getFormulas();
  const values = range.getValues();

  // Where each change's streams leave and arrive, as offsets into the block.
  const rowOf = {};
  categories.forEach(c => { rowOf[c.name] = c.row - top; });
  const soloAt = layout.SOLO_ROW - top;
  const moves = changes.map(ch => ({
    from: ch.from ? rowOf[ch.from] : undefined,
    to: ch.to ? rowOf[ch.to] : undefined,
    // The solo total is the artist total minus SOLO_EXCLUDES, so it moves the opposite way.
    solo: (ch.from === layout.SOLO_EXCLUDES ? 1 : 0) - (ch.to === layout.SOLO_EXCLUDES ? 1 : 0),
    streams: sheet.getRange(ch.row, 2, 1, count).getValues()[0].map(v => Number(v) || 0)
  }));

  const out = [];
  for (let r = 0; r < height; r++) out.push(new Array(count));
  for (let c = 0; c < count; c++) {
    const built = buildAlbumFormulas(columnToLetter(c + 2));
    const delta = new Array(height).fill(0);
    moves.forEach(m => {
      const s = m.streams[c];
      if (!s) return;
      if (m.from !== undefined) delta[m.from] -= s;
      if (m.to !== undefined) delta[m.to] += s;
      delta[soloAt] += m.solo * s;
    });
    for (let r = 0; r < height; r++) {
      const v = values[r][c];
      if (formulas[r][c] || v === '') {
        out[r][c] = formulas[r][c] ? built[r][0] : v;   // a blank cell stays blank
      } else if (delta[r] && typeof v === 'number') {
        out[r][c] = v + delta[r];
        result.adjusted++;
      } else {
        out[r][c] = v;
      }
    }
  }
  result.columns = count;
  // setValues writes a string starting with "=" as a formula, so formulas and numbers go in together.
  return { range: range, out: out, result: result };
}


/**
 * Sets every song's historyCategory to its category, creating the column if it isn't there.
 */
function recordHistoryCategories() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SONGS);
  const values = sheet.getDataRange().getValues();
  const col = findSongColumns(values[0]);
  let at = findOptionalColumn(values[0], CONFIG.SONGS_SHEET.HISTORY_CATEGORY);
  if (at === -1) {
    at = values[0].length;
    sheet.getRange(1, at + 1).setValue(CONFIG.SONGS_SHEET.HISTORY_CATEGORY)
      .setNote('Filled in by the code: the category the archive history counts this song under. ' +
        'Leave it alone, except to tell Rebuild Album History about a category change made before this column existed.');
  }
  if (values.length < 2) return;
  const cats = values.slice(1).map(r => [String(r[col.title]).trim() ? String(r[col.category]).trim() : '']);
  sheet.getRange(2, at + 1, cats.length, 1).setValues(cats);
  _songsCache = null;
}


function describeHistoryRebuild(result) {
  const lines = [];
  if (result.created) {
    lines.push(`Created the "${CONFIG.SONGS_SHEET.HISTORY_CATEGORY}" column in the ${CONFIG.SHEETS.SONGS} sheet, filled with each song's current category. No past figures were changed.`);
  } else if (result.changes.length) {
    const figures = result.sheets.reduce((n, s) => n + s.adjusted, 0);
    lines.push(`Album history adjusted: ${figures} past figure(s) changed, for ${result.changes.length} song(s):\n` + describeChanges(result.changes));
  } else {
    lines.push('No song had changed category, so no past figures changed.');
  }
  lines.push("Today's album formulas were rewritten.");
  return lines.join('\n');
}

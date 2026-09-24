/**
 * ===================================================================
 * ADD PENDING SONGS
 * -------------------------------------------------------------------
 * Adds every not-yet-added row of the Pending sheet to the spreadsheet:
 * each song gets a row at the end of its category's block, inserted at
 * the same row in Latest, every Daily Archive and the Total Archive,
 * plus an entry in the Tracklist. Songs below it move down one row everywhere; their
 * Tracklist rows are renumbered to match.
 *
 * The whole batch is validated first. If any row has a problem, nothing
 * is changed and the problems are written into Pending's result column.
 *
 * A category with no songs yet (a new album) is placed after the songs
 * of the nearest category above it in the Categories sheet.
 * ===================================================================
 */

/**
 * [MENU] Validates, confirms, adds, then reports.
 */
function addPendingSongsMenu() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let pending;
  try {
    pending = readPending();
  } catch (error) {
    ui.alert('Add Pending Songs', error.message, ui.ButtonSet.OK);
    return;
  }
  if (!pending.entries.length) {
    ui.alert('Add Pending Songs' + envTag(), `There are no songs waiting in the ${CONFIG.SHEETS.PENDING} sheet.`, ui.ButtonSet.OK);
    return;
  }

  // --- 1. Validate everything first ---
  const plan = planPendingSongs(pending.entries);
  if (plan.blockers.length || plan.entries.some(e => e.errors.length)) {
    writePendingResults(pending, plan.entries.map(e => e.errors.length ? '❌ ' + e.errors.join(' ') : ''));
    const lines = plan.blockers.concat(plan.entries.filter(e => e.errors.length).map(e => `"${e.title}": ${e.errors.join(' ')}`));
    ui.alert('Nothing was added' + envTag(), 'Fix these first (details are in the result column):\n\n• ' + lines.join('\n• '), ui.ButtonSet.OK);
    return;
  }

  // --- 2. Confirm ---
  const toAdd = plan.entries.filter(e => !e.ignore);
  const toIgnore = plan.entries.filter(e => e.ignore);
  let list = toAdd.map(e => `• ${e.title} → ${e.category}, row ${e.targetRow}${e.status === SONG_STATUS.UPCOMING ? ' (upcoming)' : ''}`).join('\n');
  if (toIgnore.length) list += (list ? '\n\n' : '') + `Ignore ${toIgnore.length} track(s):\n` + toIgnore.map(e => `• ${e.spotifyTitle || e.title}`).join('\n');
  const response = ui.alert(`Add ${toAdd.length} song(s)${toIgnore.length ? `, ignore ${toIgnore.length}` : ''}?` + envTag(),
    list + (toAdd.length ? '\n\nSongs below each one move down a row in every sheet.' : '') +
    (toAdd.length > 3 ? ' For a batch this size, a named version in Version history first is a good idea.' : ''),
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  // --- 3. Add, one at a time, in the planned order ---
  ss.toast('Adding songs... do not edit the spreadsheet.', 'Add Pending Songs', -1);
  const results = plan.entries.map(() => '');
  if (toIgnore.length) {
    addToIgnored(toIgnore.map(e => ({ trackId: e.trackId, spotifyTitle: e.spotifyTitle || e.title, album: e.album, reason: 'Set to ignore in Pending' })));
    plan.entries.forEach((e, i) => { if (e.ignore) results[i] = `${CONFIG.PENDING_SHEET.DONE_PREFIX} Ignored on ${formatDateString(new Date())}`; });
  }
  let failed = null;
  for (let i = 0; i < plan.entries.length; i++) {
    const e = plan.entries[i];
    if (e.ignore) continue;
    try {
      insertSongRow(e);
      results[i] = `${CONFIG.PENDING_SHEET.DONE_PREFIX} Added at row ${e.targetRow} on ${formatDateString(new Date())}`;
    } catch (error) {
      console.error(error);
      failed = { entry: e, error: error };
      results[i] = '❌ Stopped here: ' + error.message;
      break;
    }
  }

  // --- 4. Totals, Latest's album formulas, results ---
  _songsCache = null;
  let matchText = '';
  try {
    const match = matchTotalsById();
    setLatestAggregateFormulas();
    matchText = describeMatchResult(match);
  } catch (error) {
    matchText = 'Matching totals failed: ' + error.message;
  }
  writePendingResults(pending, results);
  SpreadsheetApp.flush();
  ss.toast('Finished - see the report.', 'Add Pending Songs', 3);   // replaces the one with no timeout

  if (failed) {
    ui.alert('Stopped partway' + envTag(),
      `Adding "${failed.entry.title}" failed: ${failed.error.message}\n\n` +
      'The songs before it were added; this one may be only partly in. Run Checks → Check Row Alignment. If rows are out of line, restore from Version history.',
      ui.ButtonSet.OK);
    return;
  }

  // --- 5. Check the rows still line up, and report ---
  const checks = runChecks(HEALTH_CHECKS.filter(c => c.name === 'Row alignment' || c.name === 'Totals add up'));
  const todo = [];
  let coverText = '';
  try {
    const covers = ensureCovers(toAdd.map(e => ({ coverKey: e.coverKey, trackId: e.trackId, album: e.album, title: e.title })));
    coverText = describeCovers({ added: covers.added, missing: [] });   // missing ones become to-dos below
    covers.missing.forEach(m => todo.push(`Add cover key "${m.key}" to the ${CONFIG.SHEETS.COVERS} sheet${m.title ? ` (${m.title})` : ''} - no cover to copy yet${toAdd.some(e => e.coverKey === m.key && !e.trackId) ? '; it can be filled once the song is out (Update → Fill Missing Covers)' : ''}.`));
  } catch (error) {
    coverText = 'Adding covers failed: ' + error.message;
  }
  toAdd.forEach(e => {
    if (e.beyondSummaryLimit) todo.push(`"${e.title}" is past ${e.category}'s summary limit, so it won't be named in that album's summary.`);
  });
  if (toAdd.length) todo.push('Add the new song(s) to their album breakdowns on the Albums sheet, if they belong there.');

  ui.alert(`Added ${toAdd.length} song(s)${toIgnore.length ? `, ignored ${toIgnore.length}` : ''}` + envTag(),
    toAdd.map(e => `✅ ${e.title} → row ${e.targetRow}`).join('\n') +
    (toIgnore.length ? `\n🚫 ${toIgnore.length} track(s) moved to the ${CONFIG.SHEETS.IGNORED} sheet.` : '') + '\n\n' + matchText + '\n\n' +
    (coverText ? coverText + '\n\n' : '') +
    formatCheckReport(checks) + (todo.length ? '\n\nStill to do by hand:\n• ' + todo.join('\n• ') : ''),
    ui.ButtonSet.OK);
}


/**
 * @return {{sheet: Sheet, col: Object, entries: Array}} Pending rows not yet added, in sheet order.
 */
function readPending() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.PENDING);
  if (!sheet) throw new Error(`The "${CONFIG.SHEETS.PENDING}" sheet is missing - run Update → Setup → Create Pending sheet.`);
  const values = sheet.getDataRange().getValues();
  const col = findHeaderColumns(values[0], CONFIG.PENDING_SHEET.HEADERS, CONFIG.SHEETS.PENDING);

  const entries = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const trackId = String(r[col.trackId]).trim();
    const title = String(r[col.title]).trim() || trackId;   // an "ignore" row may have only an ID
    if (!title) continue;
    if (String(r[col.result]).indexOf(CONFIG.PENDING_SHEET.DONE_PREFIX) === 0) continue;
    entries.push({
      sheetRow: i + 1,
      title: title,
      category: String(r[col.category]).trim(),
      coverKey: String(r[col.coverKey]).trim(),
      trackId: trackId,
      // Blank status: active if there's an ID, otherwise it can only be upcoming.
      status: String(r[col.status]).trim().toLowerCase() || (trackId ? SONG_STATUS.ACTIVE : SONG_STATUS.UPCOMING),
      spotifyTitle: String(r[col.spotifyTitle]).trim(),
      album: String(r[col.album]).trim()
    });
  }
  return { sheet: sheet, col: col, entries: entries };
}


/**
 * Validates every entry and works out where each one goes, as if the earlier ones were already in.
 * Reads only; changes nothing.
 * @return {{blockers: Array<string>, entries: Array}} Entries gain errors[], targetRow, import info.
 */
function planPendingSongs(entries) {
  const blockers = [];
  let songs, categories;
  try { songs = getSongs(); } catch (e) { blockers.push(`${CONFIG.SHEETS.SONGS}: ${e.message}`); }
  try { categories = getCategories(); } catch (e) { blockers.push(`${CONFIG.SHEETS.CATEGORIES}: ${e.message}`); }
  if (blockers.length) return { blockers: blockers, entries: entries.map(e => Object.assign(e, { errors: [] })) };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [CONFIG.SHEETS.LATEST, CONFIG.SHEETS.TOTAL_ARCHIVE].concat(CONFIG.SHEETS.ARCHIVE_YEARS)
    .forEach(n => { if (!ss.getSheetByName(n)) blockers.push(`Sheet "${n}" not found.`); });
  let raw = [];
  try { raw = readRawImport(); } catch (e) { blockers.push(e.message); }

  // The rows must line up before anything is inserted, or the inserts would make it worse.
  const alignment = runChecks(HEALTH_CHECKS.filter(c => c.name === 'Row alignment'))[0];
  if (alignment.status === CHECK_STATUS.FAIL) blockers.push('Rows are out of line - run Checks → Check Row Alignment and fix that first.');

  // What the import knows about each track ID.
  const imported = {};
  raw.forEach(t => { imported[t.id] = { album: t.album, name: t.name }; });

  const byName = {};
  categories.forEach(c => { byName[c.name] = c; });
  const knownIds = {};
  const knownTitles = {};
  songs.forEach(s => {
    if (s.trackId) knownIds[s.trackId] = s.title;
    knownTitles[s.category + '|' + s.title.toLowerCase()] = true;
  });

  // Current rows of each category, updated as the plan inserts songs.
  const rowsByCategory = {};
  songs.forEach(s => { if (s.category) (rowsByCategory[s.category] = rowsByCategory[s.category] || []).push(s.row); });
  let lastSongRow = songs.length ? songs[songs.length - 1].row : CONFIG.LAYOUT.FIRST_SONG_ROW - 1;
  const idsInBatch = {};

  entries.forEach(e => {
    e.errors = [];
    const info = imported[e.trackId] || {};
    e.spotifyTitle = e.spotifyTitle || info.name || '';
    e.album = e.album || info.album || '';

    // "ignore": only the track ID matters; it goes to the Ignored sheet, not into any row.
    if (e.status === CONFIG.PENDING_SHEET.IGNORE_STATUS) {
      e.ignore = true;
      if (!e.trackId) e.errors.push('Ignoring needs the track ID.');
      else if (knownIds[e.trackId]) e.errors.push(`Track ID is tracked as "${knownIds[e.trackId]}" - retire that song instead.`);
      else if (idsInBatch[e.trackId]) e.errors.push('Track ID appears twice in Pending.');
      idsInBatch[e.trackId] = true;
      return;
    }

    const cat = byName[e.category];
    if (!e.category) e.errors.push('No category.');
    else if (!cat) e.errors.push(`Category "${e.category}" is not in the ${CONFIG.SHEETS.CATEGORIES} sheet.`);
    if (!e.coverKey) e.errors.push('No cover key.');
    if (CONFIG.PENDING_SHEET.STATUSES.indexOf(e.status) === -1) e.errors.push(`Status "${e.status}" must be one of: ${CONFIG.PENDING_SHEET.STATUSES.join(', ')}.`);
    if (e.status === SONG_STATUS.ACTIVE && !e.trackId) e.errors.push('An active song needs its track ID (use "upcoming" if it isn\'t out yet).');
    if (e.trackId) {
      if (knownIds[e.trackId]) e.errors.push(`Track ID is already used by "${knownIds[e.trackId]}".`);
      if (idsInBatch[e.trackId]) e.errors.push('Track ID appears twice in Pending.');
      idsInBatch[e.trackId] = true;
      if (e.status === SONG_STATUS.ACTIVE && !imported[e.trackId]) e.errors.push('Track ID is not in the latest import - import first, or use "upcoming".');
    }
    if (cat && knownTitles[e.category + '|' + e.title.toLowerCase()]) e.errors.push(`"${e.title}" is already in ${e.category}.`);
    if (e.errors.length || !cat) return;

    // Where it goes: after the category's last song, or - for a category with none yet - after the
    // songs of the nearest category above it.
    let after = null;
    if ((rowsByCategory[cat.name] || []).length) {
      after = Math.max.apply(null, rowsByCategory[cat.name]);
    } else {
      const above = categories.filter(c => c.row < cat.row && (rowsByCategory[c.name] || []).length);
      if (above.length) after = Math.max.apply(null, rowsByCategory[above[above.length - 1].name]);
    }
    e.targetRow = after === null ? CONFIG.LAYOUT.FIRST_SONG_ROW : after + 1;
    e.beyondSummaryLimit = !!(cat.summaryLimit && (rowsByCategory[cat.name] || []).length >= cat.summaryLimit);

    // Apply it to the plan, so the next entry sees this one in place.
    Object.keys(rowsByCategory).forEach(k => {
      rowsByCategory[k] = rowsByCategory[k].map(r => (r >= e.targetRow ? r + 1 : r));
    });
    (rowsByCategory[cat.name] = rowsByCategory[cat.name] || []).push(e.targetRow);
    lastSongRow++;
    knownTitles[e.category + '|' + e.title.toLowerCase()] = true;
  });

  // Sheets can't shift a block of columns past a merged cell that crosses its edge.
  const targets = entries.filter(e => e.targetRow).map(e => e.targetRow);
  if (targets.length) {
    const top = Math.min.apply(null, targets);
    mergesInTheWay(ss.getSheetByName(CONFIG.SHEETS.LATEST), top, 1, 16).forEach(p => blockers.push(p));
  }

  return { blockers: blockers, entries: entries };
}


/**
 * Inserts one planned song into every sheet. Assumes planPendingSongs() placed it.
 */
function insertSongRow(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const row = e.targetRow;
  const first = CONFIG.LAYOUT.FIRST_SONG_ROW;
  const latest = ss.getSheetByName(CONFIG.SHEETS.LATEST);
  const archives = CONFIG.SHEETS.ARCHIVE_YEARS.map(n => ss.getSheetByName(n)).concat([ss.getSheetByName(CONFIG.SHEETS.TOTAL_ARCHIVE)]);

  // --- 1. Open the row everywhere ---
  [latest].concat(archives).forEach(sheet => insertRowAt(sheet, row));

  // A neighbouring song row to copy formulas and formatting from.
  const template = row - 1 >= first ? row - 1 : row + 1;

  // --- 2. Latest: formulas and formats from the neighbour, values cleared, then the song's own ---
  copyFormulasOnly(latest, template, row, 1, CONFIG.LATEST.COPIED_COLUMNS);
  latest.getRange(row, 1).setValue(e.coverKey);
  latest.getRange(row, CONFIG.LATEST.COLS.TITLE).setValue(e.title);
  latest.getRange(row, CONFIG.LATEST.COLS.TOTAL, 1, 2).setValues([[0, 0]]);

  // --- 3. Archives: just the title; there is no history yet ---
  archives.forEach(sheet => sheet.getRange(row, 1).setValue(e.title));

  // --- 4. Tracklist: move everything at or below the row down one, then add this song ---
  const tracklist = ss.getSheetByName(CONFIG.SHEETS.SONGS);
  const values = tracklist.getDataRange().getValues();
  const col = findSongColumns(values[0]);
  if (values.length > 1) {
    const rows = values.slice(1).map(r => {
      const v = Number(r[col.row]);
      return [String(r[col.title]).trim() && Number.isInteger(v) && v >= row ? v + 1 : r[col.row]];
    });
    tracklist.getRange(2, col.row + 1, rows.length, 1).setValues(rows);
  }
  const header = values[0].map(h => String(h).toLowerCase().replace(/\s+/g, ''));
  const newRow = values[0].map(() => '');
  newRow[col.row] = row;
  newRow[col.status] = e.status;
  newRow[col.category] = e.category;
  newRow[col.coverKey] = e.coverKey;
  newRow[col.title] = e.title;
  newRow[col.trackId] = e.trackId;
  // Optional reference columns, filled when the Tracklist has them.
  // A new song has no history yet, so its history category is simply its category.
  const optional = { spotifytitle: e.spotifyTitle, sourcealbum: e.album, historycategory: e.category };
  header.forEach((h, i) => { if (optional.hasOwnProperty(h)) newRow[i] = optional[h]; });
  tracklist.getRange(values.length + 1, 1, 1, newRow.length).setValues([newRow]);
  _songsCache = null;
}


/**
 * Inserts a blank row so that the new row is `row`, even when that is past the end of the sheet.
 */
function insertRowAt(sheet, row) {
  const max = sheet.getMaxRows();
  if (row <= max) sheet.insertRowBefore(row);
  else sheet.insertRowsAfter(max, row - max);
}


/**
 * Copies a row's formulas (adjusted to the new row) and formatting to another row, leaving the
 * neighbour's plain values behind - a new song must not inherit its neighbour's figures.
 */
function copyFormulasOnly(sheet, fromRow, toRow, col, width) {
  const source = sheet.getRange(fromRow, col, 1, width);
  const target = sheet.getRange(toRow, col, 1, width);
  try {
    source.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  } catch (error) {
    // Only formatting is copied here, so a merge in the way is worth a note, not a failed run.
    console.warn('Copying formatting from ' + sheet.getName() + ' row ' + fromRow + ' to ' + toRow + ' failed: ' + error.message);
  }
  const formulas = source.getFormulasR1C1()[0];
  const out = formulas.map(f => f || '');
  target.setFormulasR1C1([out]);
}


function writePendingResults(pending, results) {
  const col = pending.col.result + 1;
  pending.entries.forEach((e, i) => pending.sheet.getRange(e.sheetRow, col).setValue(results[i] || ''));
}



/**
 * Sheets refuses to shift part of a row when a merged cell crosses the edge of the block being
 * shifted ("cannot cut or paste part of a merged cell"). This finds those merges before anything
 * is changed, so a run stops with an explanation instead of halfway through.
 * @param {Sheet} sheet
 * @param {number} fromRow - The topmost row that will be shifted down.
 * @param {number} col - First column of the block being shifted.
 * @param {number} width - How many columns it covers.
 * @return {Array<string>} One message per merge in the way.
 */
function mergesInTheWay(sheet, fromRow, col, width) {
  const lastRow = sheet.getLastRow();
  if (!sheet || lastRow < fromRow) return [];
  const lastCol = Math.max(sheet.getLastColumn(), col + width - 1);
  const problems = [];
  sheet.getRange(fromRow, 1, lastRow - fromRow + 1, lastCol).getMergedRanges().forEach(m => {
    const first = m.getColumn(), last = first + m.getNumColumns() - 1;
    const overlaps = last >= col && first <= col + width - 1;
    const inside = first >= col && last <= col + width - 1;
    if (overlaps && !inside) {
      problems.push(sheet.getName() + ': the merged cells at ' + m.getA1Notation() + ' stick out of columns ' +
        columnToLetter(col) + ':' + columnToLetter(col + width - 1) + ', the block that has to shift down. ' +
        'Unmerge them, or keep the merge inside that block, and try again.');
    }
  });
  return problems;
}

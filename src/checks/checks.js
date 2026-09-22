/**
 * ===================================================================
 * HEALTH CHECKS
 * -------------------------------------------------------------------
 * Read-only checks of the live spreadsheet, run from the Checks menu
 * and (a quick subset) at the end of every Update Daily Stats.
 * Nothing in this file writes to the spreadsheet: a check reports a
 * problem, it never repairs one.
 *
 * Each check takes the shared data loader and returns
 *   { status: 'ok' | 'warn' | 'fail', summary: string, details: string[] }
 * ===================================================================
 */

const CHECK_STATUS = { OK: 'ok', WARN: 'warn', FAIL: 'fail' };
const CHECK_ICONS = { ok: '✅', warn: '⚠️', fail: '❌' };

// Sheet dates come back as Date objects; this test does not depend on which realm made them.
const isDate = v => Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());

const HEALTH_CHECKS = [
  { name: 'Tracklist',          run: checkTracklist },
  { name: 'Row alignment',      run: checkRowAlignment },
  { name: 'Totals add up',      run: checkTotals },
  { name: 'Aggregate formulas', run: checkAggregateFormulas },
  { name: 'Import',             run: checkImport },
  { name: 'Latest vs Tools',    run: checkLatestVsTools },
  { name: 'Date sequence',      run: checkDateSequence },
  { name: 'Covers',             run: checkCovers },
  { name: 'Spare category rows', run: checkSpareRows },
  { name: 'Config',             run: checkConfig }
];

// Cheap, and each one catches damage a daily update could do. Run after every update.
const AFTER_UPDATE_CHECKS = ['Row alignment', 'Totals add up', 'Latest vs Tools', 'Date sequence'];


// --- MENU ENTRY POINTS ---

function runAllChecksMenu() {
  showCheckReport('Health checks', runChecks(HEALTH_CHECKS));
}

function checkRowAlignmentMenu() {
  showCheckReport('Row alignment', runChecks(HEALTH_CHECKS.filter(c => c.name === 'Row alignment')));
}

/**
 * Called at the end of main(). Silent when everything passes; a dialog otherwise.
 * Never throws - the update itself has already finished by the time this runs.
 */
function runChecksAfterUpdate() {
  try {
    const results = runChecks(HEALTH_CHECKS.filter(c => AFTER_UPDATE_CHECKS.indexOf(c.name) !== -1));
    if (results.some(r => r.status !== CHECK_STATUS.OK)) {
      showCheckReport('The update finished, but a check did not pass', results);
    }
  } catch (error) {
    console.error(error);
    SpreadsheetApp.getUi().alert('Checks could not run', error.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


// --- RUNNER ---

/**
 * @param {Array<{name: string, run: Function}>} checks
 * @return {Array<{name: string, status: string, summary: string, details: string[]}>}
 */
function runChecks(checks) {
  const data = loadCheckData();
  return checks.map(check => {
    try {
      const result = check.run(data);
      return Object.assign({ name: check.name, details: [] }, result);
    } catch (error) {
      console.error(check.name, error);
      return { name: check.name, status: CHECK_STATUS.FAIL, summary: 'The check itself failed: ' + error.message, details: [] };
    }
  });
}

function showCheckReport(title, results) {
  const text = formatCheckReport(results);
  console.log(text);
  const ui = SpreadsheetApp.getUi();
  ui.alert(title + envTag(), text, ui.ButtonSet.OK);
}

function formatCheckReport(results) {
  const max = CONFIG.CHECKS.MAX_DETAILS;
  const counts = { ok: 0, warn: 0, fail: 0 };
  const lines = [];

  results.forEach(r => {
    counts[r.status]++;
    lines.push(`${CHECK_ICONS[r.status]} ${r.name}: ${r.summary}`);
    r.details.slice(0, max).forEach(d => lines.push(`      • ${d}`));
    if (r.details.length > max) lines.push(`      … and ${r.details.length - max} more`);
  });

  const head = `${counts.ok} passed, ${counts.warn} warning(s), ${counts.fail} failed`;
  return head + '\n\n' + lines.join('\n');
}


// --- SHARED DATA ---

/**
 * Reads each range at most once per run, however many checks use it.
 */
function loadCheckData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cache = {};
  const once = (key, read) => () => {
    if (!cache.hasOwnProperty(key)) cache[key] = read();
    return cache[key];
  };

  const data = {
    ss: ss,
    sheet: name => ss.getSheetByName(name),
    tracklist: once('tracklist', readTracklist),
    firstSongRow: CONFIG.LAYOUT.FIRST_SONG_ROW,
    // Worked out from readTracklist(), which never throws, so a broken Tracklist can't stop the checks.
    lastSongRow: once('lastSongRow', () => {
      const songs = data.tracklist().songs;
      return songs.length ? songs[songs.length - 1].row : CONFIG.LAYOUT.FIRST_SONG_ROW - 1;
    }),
    // Last aggregate row in use: the highest category row.
    lastAggregateRow: Math.max.apply(null, CONFIG.CATEGORIES.map(c => c.row)),
    rawImport: once('rawImport', () => ss.getSheetByName(CONFIG.SHEETS.TOOLS).getRange(CONFIG.TOOLS.RAW_DATA).getValues())
  };
  const songRange = (sheet, col, width) =>
    sheet.getRange(data.firstSongRow, col, data.lastSongRow() - data.firstSongRow + 1, width).getValues();

  // Latest A:P for every song row
  data.latestSongs = once('latestSongs', () => songRange(ss.getSheetByName(CONFIG.SHEETS.LATEST), 1, 16));
  // Latest E:G (title, total, daily) for the aggregate rows, from the artist total down
  data.latestAggregates = once('latestAggregates', () => ss.getSheetByName(CONFIG.SHEETS.LATEST)
    .getRange(CONFIG.LAYOUT.TOTAL_ROW, CONFIG.LATEST.COLS.TITLE, data.lastAggregateRow - CONFIG.LAYOUT.TOTAL_ROW + 1, 3).getValues());
  data.toolsTotals = once('toolsTotals', () => songRange(ss.getSheetByName(CONFIG.SHEETS.TOOLS), CONFIG.TOOLS.TOTALS_COLUMN, 1).map(r => r[0]));
  data.songsByRow = once('songsByRow', () => {
    const byRow = {};
    data.tracklist().songs.forEach(s => { byRow[s.row] = s; });
    return byRow;
  });
  return data;
}


// --- THE CHECKS ---

// For checks that are meaningless without a readable Tracklist.
function tracklistUnreadable() {
  return { status: CHECK_STATUS.FAIL, summary: `Can't read the ${CONFIG.SHEETS.SONGS} sheet - see the Tracklist check.` };
}

function checkTracklist(data) {
  const t = data.tracklist();
  const problems = t.problems.slice();

  const known = CONFIG.CATEGORIES.map(c => c.name);
  t.songs.forEach(s => {
    if (s.category && known.indexOf(s.category) === -1) {
      problems.push(`Row ${s.row} ("${s.title}") has category "${s.category}", which is not in CONFIG.CATEGORIES.`);
    }
    if (s.status === SONG_STATUS.ACTIVE && !s.trackId) {
      problems.push(`Row ${s.row} ("${s.title}") is active but has no track ID.`);
    }
  });

  if (problems.length) {
    return { status: CHECK_STATUS.FAIL, summary: `${problems.length} problem(s) in the ${CONFIG.SHEETS.SONGS} sheet.`, details: problems };
  }
  const retired = t.songs.filter(s => s.status === SONG_STATUS.RETIRED).length;
  return { status: CHECK_STATUS.OK, summary: `${t.songs.length} songs (${t.songs.length - retired} active, ${retired} retired).` };
}


function checkRowAlignment(data) {
  if (data.tracklist().songs.length === 0) return tracklistUnreadable();
  const byRow = data.songsByRow();
  const expected = i => (byRow[data.firstSongRow + i] || {}).title || '';

  // Every sheet whose rows must line up with the Tracklist, and the column holding the title.
  const sources = [{ name: CONFIG.SHEETS.LATEST, titles: data.latestSongs().map(r => r[4]) }];
  const others = CONFIG.SHEETS.ARCHIVE_YEARS.concat([CONFIG.SHEETS.TOTAL_ARCHIVE]);
  const missing = [];
  others.forEach(name => {
    const sheet = data.sheet(name);
    if (!sheet) { missing.push(name); return; }
    const titles = sheet.getRange(data.firstSongRow, 1, data.lastSongRow() - data.firstSongRow + 1, 1).getValues().map(r => r[0]);
    sources.push({ name: name, titles: titles });
  });

  const details = [];
  sources.forEach(src => {
    src.titles.forEach((value, i) => {
      const have = String(value).trim();
      const want = expected(i);
      if (have !== want) {
        details.push(`Row ${data.firstSongRow + i}: ${src.name} has "${have}", ${CONFIG.SHEETS.SONGS} has "${want}".`);
      }
    });
  });

  const missingNote = missing.map(n => `Sheet "${n}" not found, so it was not compared.`);
  if (details.length) {
    return { status: CHECK_STATUS.FAIL, summary: `${details.length} row(s) out of line.`, details: details.concat(missingNote) };
  }
  if (missing.length) {
    return { status: CHECK_STATUS.WARN, summary: `The ${sources.length} sheet(s) found agree on every row.`, details: missingNote };
  }
  return { status: CHECK_STATUS.OK, summary: `${sources.length} sheets agree with the ${CONFIG.SHEETS.SONGS} sheet on every row.` };
}


function checkTotals(data) {
  if (data.tracklist().songs.length === 0) return tracklistUnreadable();
  const layout = CONFIG.LAYOUT;
  const songs = data.latestSongs();             // F = index 5, G = index 6
  const aggregates = data.latestAggregates();   // E, F, G from the artist total down
  const rowsByCategory = {};
  data.tracklist().songs.forEach(s => {
    if (!s.category) return;
    (rowsByCategory[s.category] = rowsByCategory[s.category] || []).push(s.row);
  });

  const songSum = (rows, index) => rows.reduce((sum, row) => sum + (Number(songs[row - data.firstSongRow][index]) || 0), 0);
  const inLatest = (row, index) => Number(aggregates[row - layout.TOTAL_ROW][index]);
  const fmt = n => Math.round(n).toLocaleString('en-US');
  const details = [];
  const compare = (label, row, expectedTotal, expectedDaily, what) => {
    [[1, expectedTotal, 'total'], [2, expectedDaily, 'daily']].forEach(([index, expected, figure]) => {
      const actual = inLatest(row, index);
      if (actual !== expected) {
        details.push(`${label} (Latest row ${row}) ${figure}: Latest says ${fmt(actual)}, ${what} add up to ${fmt(expected)}.`);
      }
    });
  };

  CONFIG.CATEGORIES.forEach(c => {
    const rows = rowsByCategory[c.name] || [];
    compare(c.name, c.row, songSum(rows, 5), songSum(rows, 6), 'its songs');
  });
  const allRows = data.tracklist().songs.map(s => s.row);
  compare('Total Artist Streams', layout.TOTAL_ROW, songSum(allRows, 5), songSum(allRows, 6), 'all songs');

  const excluded = CONFIG.CATEGORIES.filter(c => c.name === layout.SOLO_EXCLUDES)[0];
  if (excluded) {
    compare('Total Artist Solo Streams', layout.SOLO_ROW,
      inLatest(layout.TOTAL_ROW, 1) - inLatest(excluded.row, 1),
      inLatest(layout.TOTAL_ROW, 2) - inLatest(excluded.row, 2),
      `the total minus ${excluded.name}`);
  }

  if (details.length) {
    return { status: CHECK_STATUS.FAIL, summary: `${details.length} album figure(s) in Latest don't match their songs.`, details: details };
  }
  return { status: CHECK_STATUS.OK, summary: `All ${CONFIG.CATEGORIES.length} categories, the artist total and the solo total match their songs, totals and dailies.` };
}


function checkAggregateFormulas(data) {
  if (data.tracklist().problems.length) {
    return { status: CHECK_STATUS.FAIL, summary: `Can't work out the formulas while the ${CONFIG.SHEETS.SONGS} sheet has problems - see the Tracklist check.` };
  }
  const sheet = data.sheet(CONFIG.SHEETS.ARCHIVE);
  const want = buildAlbumFormulas().map(r => r[0]);
  const have = sheet.getRange(CONFIG.LAYOUT.TOTAL_ROW, 2, want.length, 1).getFormulas().map(r => r[0]);
  const norm = f => String(f).replace(/\s+/g, '').toUpperCase();

  const details = [];
  want.forEach((f, i) => {
    if (norm(have[i]) !== norm(f)) {
      details.push(`Row ${CONFIG.LAYOUT.TOTAL_ROW + i}: has ${have[i] || '(no formula)'}, the ${CONFIG.SHEETS.SONGS} sheet gives ${f}.`);
    }
  });

  if (details.length) {
    return {
      status: CHECK_STATUS.WARN,
      summary: `${details.length} of today's aggregate formulas in ${CONFIG.SHEETS.ARCHIVE} column B differ. ` +
        `They were written before the latest ${CONFIG.SHEETS.SONGS} change, or edited by hand; the next update writes fresh ones.`,
      details: details
    };
  }
  return { status: CHECK_STATUS.OK, summary: `Today's aggregate formulas (rows ${CONFIG.LAYOUT.TOTAL_ROW}-${CONFIG.LAYOUT.TOTAL_ROW + want.length - 1}) match the ${CONFIG.SHEETS.SONGS} sheet.` };
}


function checkImport(data) {
  const raw = data.rawImport();              // E album, F name, G count, H track ID
  const rows = raw.filter(r => String(r[1]).trim() || String(r[3]).trim()).length;
  const ids = {};
  raw.forEach(r => { const id = String(r[3]).trim(); if (id) ids[id] = true; });
  const idCount = Object.keys(ids).length;

  if (rows === 0) return { status: CHECK_STATUS.WARN, summary: 'The raw import in Tools is empty.' };
  if (idCount === 0) return { status: CHECK_STATUS.FAIL, summary: 'The raw import has no track IDs - it was written by the old importer.' };

  const details = [];
  const active = data.tracklist().songs.filter(s => s.status === SONG_STATUS.ACTIVE);
  active.forEach(s => {
    if (s.trackId && !ids[s.trackId]) details.push(`Row ${s.row} ("${s.title}"): track ID ${s.trackId} is not in the import.`);
  });

  data.toolsTotals().forEach((v, i) => {
    if (v === CONFIG.TOOLS.MISSING_MARKER) details.push(`Tools row ${data.firstSongRow + i} is marked ${CONFIG.TOOLS.MISSING_MARKER}.`);
  });

  const c1 = data.sheet(CONFIG.SHEETS.TOOLS).getRange(CONFIG.TOOLS.SUM_OF_DAILYS).getValue();
  if (typeof c1 !== 'number') details.push(`Tools!${CONFIG.TOOLS.SUM_OF_DAILYS} (sum of dailies) is "${c1}", not a number.`);

  if (details.length) {
    return { status: CHECK_STATUS.FAIL, summary: `${details.length} problem(s) with the import.`, details: details };
  }
  if (rows >= CONFIG.CHECKS.RAW_ROWS_WARN) {
    return { status: CHECK_STATUS.WARN, summary: `${rows} tracks imported - close to the ${raw.length}-row limit of Tools!${CONFIG.TOOLS.RAW_DATA}.` };
  }
  return { status: CHECK_STATUS.OK, summary: `${rows} tracks imported; all ${active.length} active track IDs are present.` };
}


function checkLatestVsTools(data) {
  if (data.tracklist().songs.length === 0) return tracklistUnreadable();
  const latest = data.latestSongs();
  const tools = data.toolsTotals();
  const byRow = data.songsByRow();
  const details = [];

  latest.forEach((r, i) => {
    const row = data.firstSongRow + i;
    if (!byRow[row]) return;
    if (r[5] !== tools[i]) details.push(`Row ${row} ("${byRow[row].title}"): Latest total ${r[5]}, Tools total ${tools[i]}.`);
  });

  if (details.length) {
    return {
      status: CHECK_STATUS.WARN,
      summary: `${details.length} song total(s) in Latest differ from Tools. That's normal between Import Data and Update Daily Stats; right after an update they must match.`,
      details: details
    };
  }
  return { status: CHECK_STATUS.OK, summary: 'Every song total in Latest matches Tools.' };
}


function checkDateSequence(data) {
  const tz = data.ss.getSpreadsheetTimeZone();
  const fmt = d => Utilities.formatDate(d, tz, 'yyyy/MM/dd');
  // Whole days between two dates; rounding absorbs the 23/25-hour days around DST changes.
  const daysBetween = (a, b) => Math.round((a.getTime() - b.getTime()) / 86400000);
  const details = [];

  // Latest!Q1 must be the date of today's archive column.
  const current = data.sheet(CONFIG.SHEETS.ARCHIVE);
  const latestDate = data.sheet(CONFIG.SHEETS.LATEST).getRange(CONFIG.LATEST.DATE_CELL).getValue();
  const archiveDate = current.getRange('B1').getValue();
  if (!isDate(latestDate) || !isDate(archiveDate)) {
    details.push(`Latest!${CONFIG.LATEST.DATE_CELL} or ${CONFIG.SHEETS.ARCHIVE}!B1 is not a date.`);
  } else if (daysBetween(latestDate, archiveDate) !== 0) {
    details.push(`Latest!${CONFIG.LATEST.DATE_CELL} is ${fmt(latestDate)}, but ${CONFIG.SHEETS.ARCHIVE}!B1 is ${fmt(archiveDate)}.`);
  }

  // Each archive runs newest to oldest, one day per column, and the years join up.
  let previousOldest = null;   // oldest date of the newer year's sheet
  let days = 0;
  CONFIG.SHEETS.ARCHIVE_YEARS.forEach(name => {
    const sheet = data.sheet(name);
    if (!sheet) { details.push(`Sheet "${name}" not found.`); previousOldest = null; return; }
    const lastCol = sheet.getLastColumn();
    if (lastCol < 2) return;
    const header = sheet.getRange(1, 2, 1, lastCol - 1).getValues()[0];

    const dates = [];
    for (let i = 0; i < header.length; i++) {
      const v = header[i];
      if (v === '' || v === null) break;
      if (!isDate(v)) { details.push(`${name} column ${i + 2}: "${v}" is not a date.`); break; }
      dates.push(v);
    }
    if (!dates.length) return;

    if (previousOldest && daysBetween(previousOldest, dates[0]) !== 1) {
      details.push(`The newer year starts at ${fmt(previousOldest)}, but ${name} ends at ${fmt(dates[0])}.`);
    }
    for (let i = 1; i < dates.length; i++) {
      const gap = daysBetween(dates[i - 1], dates[i]);
      if (gap === 0) details.push(`${name}: ${fmt(dates[i])} appears twice in a row.`);
      else if (gap !== 1) details.push(`${name}: ${fmt(dates[i - 1])} is followed by ${fmt(dates[i])} (${gap > 1 ? (gap - 1) + ' day(s) missing' : 'out of order'}).`);
    }
    days += dates.length;
    previousOldest = dates[dates.length - 1];
  });

  if (details.length) {
    return { status: CHECK_STATUS.FAIL, summary: `${details.length} date problem(s).`, details: details };
  }
  return { status: CHECK_STATUS.OK, summary: `${days} archive days, one per column with no gaps or repeats; Latest matches today's column.` };
}


function checkCovers(data) {
  if (data.tracklist().songs.length === 0) return tracklistUnreadable();
  const coversSheet = data.sheet(CONFIG.SHEETS.COVERS);
  const songs = data.tracklist().songs;
  const latest = data.latestSongs();
  const details = [];
  const notes = [];

  let keys = null;
  if (coversSheet && coversSheet.getLastRow() > 0) {
    keys = {};
    // The cover lookup in the sheet is case-insensitive, so this is too.
    coversSheet.getRange(1, 1, coversSheet.getLastRow(), 1).getValues()
      .forEach(r => { const k = String(r[0]).trim().toLowerCase(); if (k) keys[k] = true; });
  } else {
    notes.push(`Sheet "${CONFIG.SHEETS.COVERS}" not found, so cover keys were not looked up.`);
  }

  songs.forEach(s => {
    const latestKey = String(latest[s.row - data.firstSongRow][0]).trim();
    if (!s.coverKey) {
      details.push(`Row ${s.row} ("${s.title}") has no cover key.`);
    } else if (keys && !keys[s.coverKey.toLowerCase()]) {
      details.push(`Row ${s.row} ("${s.title}"): cover key "${s.coverKey}" is not in ${CONFIG.SHEETS.COVERS}.`);
    }
    if (latestKey !== s.coverKey) {
      details.push(`Row ${s.row} ("${s.title}"): Latest!A has "${latestKey}", ${CONFIG.SHEETS.SONGS} has "${s.coverKey}".`);
    }
  });

  if (details.length) {
    return { status: CHECK_STATUS.FAIL, summary: `${details.length} cover problem(s).`, details: details.concat(notes) };
  }
  if (notes.length) return { status: CHECK_STATUS.WARN, summary: 'Latest!A matches every cover key.', details: notes };
  return { status: CHECK_STATUS.OK, summary: `Every song's cover key is in ${CONFIG.SHEETS.COVERS} and matches Latest!A.` };
}


function checkSpareRows(data) {
  const spare = CONFIG.LAYOUT.LAST_AGGREGATE_ROW - data.lastAggregateRow;
  const summary = `${spare} spare row(s) for new categories (rows ${data.lastAggregateRow + 1}-${CONFIG.LAYOUT.LAST_AGGREGATE_ROW}). Songs have no limit.`;

  if (spare <= 0) return { status: CHECK_STATUS.FAIL, summary: summary + ' No room for another category.' };
  if (spare < CONFIG.CHECKS.SPARE_ROWS_WARN) return { status: CHECK_STATUS.WARN, summary: summary };
  return { status: CHECK_STATUS.OK, summary: summary };
}


function checkConfig() {
  const details = [];
  const layout = CONFIG.LAYOUT;
  const cats = CONFIG.CATEGORIES;
  const duplicates = (key) => {
    const seen = {};
    cats.forEach(c => {
      if (seen.hasOwnProperty(c[key])) details.push(`CONFIG.CATEGORIES: "${c.name}" and "${seen[c[key]]}" share ${key} ${c[key]}.`);
      else seen[c[key]] = c.name;
    });
  };
  duplicates('name');
  duplicates('row');

  cats.forEach(c => {
    if (c.row <= layout.SOLO_ROW || c.row > layout.LAST_AGGREGATE_ROW) {
      details.push(`CONFIG.CATEGORIES: "${c.name}" has row ${c.row}, outside ${layout.SOLO_ROW + 1}-${layout.LAST_AGGREGATE_ROW}.`);
    }
  });
  if (!cats.some(c => c.name === layout.SOLO_EXCLUDES)) {
    details.push(`CONFIG.LAYOUT.SOLO_EXCLUDES is "${layout.SOLO_EXCLUDES}", which is not a category.`);
  }
  if (layout.FIRST_SONG_ROW <= layout.LAST_AGGREGATE_ROW) {
    details.push(`CONFIG.LAYOUT: songs start at row ${layout.FIRST_SONG_ROW}, inside the aggregate rows (up to ${layout.LAST_AGGREGATE_ROW}).`);
  }

  // Each summary clears 10 rows before writing, so two in one column need 10 rows between them.
  const byColumn = {};
  const cells = cats.filter(c => c.summaryCell).map(c => ({ name: c.name, cell: c.summaryCell }))
    .concat([{ name: 'Discography summary', cell: CONFIG.ALBUMS.TOTAL_SUMMARY }]);
  cells.forEach(({ name, cell }) => {
    const m = String(cell).match(/^([A-Z]+)(\d+)$/);
    if (!m) { details.push(`"${name}" has summary cell "${cell}", which is not a single cell.`); return; }
    (byColumn[m[1]] = byColumn[m[1]] || []).push({ name: name, row: Number(m[2]) });
  });
  Object.keys(byColumn).forEach(col => {
    const list = byColumn[col].sort((a, b) => a.row - b.row);
    for (let i = 1; i < list.length; i++) {
      if (list[i].row - list[i - 1].row < 10) {
        details.push(`Albums column ${col}: "${list[i - 1].name}" (row ${list[i - 1].row}) and "${list[i].name}" (row ${list[i].row}) are less than 10 rows apart and would overwrite each other.`);
      }
    }
  });

  if (details.length) return { status: CHECK_STATUS.FAIL, summary: `${details.length} problem(s) in config.js.`, details: details };
  return { status: CHECK_STATUS.OK, summary: `${cats.length} categories, rows and summary cells are consistent.` };
}

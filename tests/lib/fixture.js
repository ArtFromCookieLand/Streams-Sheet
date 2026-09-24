/**
 * Builds a fresh copy of the spreadsheet from tests/fixtures and loads every file in src/ into it,
 * the way Apps Script does: one shared global scope, files in path order.
 *
 * The fixture is the 2026-09-22 state after that day's update:
 *   Latest      - the real export (values), with H = G/M-1 as a formula like the real sheet
 *   Import      - the raw Apify dump from the same day (E:I), so every song's daily is 0 until
 *                 nextDayImport() bumps the counts
 *   Tracklist   - rows, statuses, categories and track IDs lined up with Latest
 *   Categories  - from CATEGORY_SEED
 *   Daily Archive 2026 (Jan 1 - Sep 22) and 2025 (Dec 12 - Dec 31): every song's daily is its
 *                 Latest daily, aggregates are the generated formulas. 2024 and 2023 hold titles only.
 *   Total Archive - 12 columns of totals
 *
 * Usage:
 *   const f = loadFixture();
 *   f.run('main()');                  // any expression, as a fresh execution (like a menu click);
 *                                     // the result comes back as plain data
 *   f.sheet('Latest').value(50, 5);   // a cell's (evaluated) value
 *   f.alerts                          // every ui.alert call: [title, message, buttons]
 *   f.answer('NO')                    // what the next ui.alert returns (default YES)
 */

const fs = require('fs');
const path = require('path');
const vm = require('vm');
const { Spreadsheet } = require('./fake_sheets');

const ROOT = path.join(__dirname, '..', '..');
const FIX = path.join(__dirname, '..', 'fixtures');

function parseCsv(text) {
  const rows = []; let row = [], field = '', quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else quoted = false; } else field += ch;
    } else if (ch === '"') quoted = true;
    else if (ch === ',') { row.push(field); field = ''; }
    else if (ch === '\n') { row.push(field.replace(/\r$/, '')); rows.push(row); row = []; field = ''; }
    else field += ch;
  }
  if (field || row.length) { row.push(field); rows.push(row); }
  return rows;
}
const csv = name => parseCsv(fs.readFileSync(path.join(FIX, name), 'utf8'));
// "1,234" / "+48,470" / "12.5%" -> number; anything else stays text.
const num = v => {
  if (v === '') return '';
  const s = String(v).replace(/[,+]/g, '');
  if (/^-?\d+(\.\d+)?$/.test(s)) return Number(s);
  return v;
};
const day = (y, m, d) => new Date(Date.UTC(y, m - 1, d, 12));
const addDays = (d, n) => new Date(d.getTime() + n * 86400000);

function srcFiles() {
  const out = [];
  const walk = dir => fs.readdirSync(dir, { withFileTypes: true }).forEach(e => {
    const p = path.join(dir, e.name);
    if (e.isDirectory()) { if (e.name !== 'website_discontinued') walk(p); } else if (e.name.endsWith('.js')) out.push(p);
  });
  walk(path.join(ROOT, 'src'));
  return out.sort();
}

function loadFixture(options) {
  options = Object.assign({ importSheetName: 'Import' }, options || {});
  const ss = new Spreadsheet({ id: 'dev-copy', timeZone: 'Etc/UTC' });
  const put = (sheet, r, c, v, display) => { const x = sheet.cell(r, c); x.v = v; x.d = display; };

  // --- Tracklist ---
  const tracklist = ss.insertSheet('Tracklist');
  csv('tracklist.csv').forEach((row, i) => row.forEach((v, j) => put(tracklist, i + 1, j + 1, i && j === 0 ? Number(v) : v)));
  // historyCategory (created by the first Rebuild Album History): the history matches the categories.
  if (options.historyCategory !== false) {
    const n = tracklist.getLastRow(), col = tracklist.getLastColumn() + 1;
    put(tracklist, 1, col, 'historyCategory');
    for (let r = 2; r <= n; r++) put(tracklist, r, col, tracklist.peek(r, 3).v);
  }

  // --- Latest ---
  const latest = ss.insertSheet('Latest');
  const latestRows = csv('latest.csv');
  latestRows.forEach((row, i) => row.forEach((v, j) => {
    if (v === '') return;
    // Row 1 is headers; Q1 is the date. Keep text columns as text.
    put(latest, i + 1, j + 1, i === 0 || j === 4 || j === 0 || j === 19 ? v : num(v), v);
  }));
  put(latest, 1, 17, day(2026, 9, 22), 'September 22, 2026');
  const lastRow = latestRows.length;
  for (let r = 2; r <= lastRow; r++) {
    if (r >= 28 && r <= 49) continue;
    if (latest.peek(r, 5).v === '') continue;
    const x = latest.cell(r, 8); x.f = `=G${r}/M${r}-1`;
  }

  // --- Import (the raw dump in E:I) ---
  const imp = ss.insertSheet(options.importSheetName);
  put(imp, 1, 1, 'Sum of dailys:');
  ['Album', 'Track', 'Streams', 'ID', 'Cover'].forEach((h, j) => put(imp, 1, 5 + j, h));
  csv('raw_import.csv').forEach((row, i) => row.forEach((v, j) => put(imp, i + 2, 5 + j, j === 2 ? num(v) : v)));

  // --- Covers: A key, B =IMAGE(C), C url ---
  const covers = ss.insertSheet('Covers');
  csv('covers.csv').forEach((row, i) => {
    put(covers, i + 1, 1, row[0]);
    if (row[2]) { put(covers, i + 1, 3, row[2]); covers.cell(i + 1, 2).f = `=IMAGE(C${i + 1})`; }
  });

  ['Albums', 'Tracks', 'Milestone Log'].forEach(n => ss.insertSheet(n));

  // --- the runtime ---
  const alerts = [];
  const answers = [];
  const props = { script: {}, document: {} };
  const propStore = store => ({
    getProperty: k => (store.hasOwnProperty(k) ? store[k] : null),
    setProperty: (k, v) => { store[k] = String(v); return this; },
    deleteProperty: k => { delete store[k]; },
    getProperties: () => Object.assign({}, store)
  });
  const dialogs = [];
  const ui = {
    alert: (...a) => { alerts.push(a); return answers.length ? answers.shift() : 'YES'; },
    ButtonSet: { OK: 'OK', YES_NO: 'YES_NO', OK_CANCEL: 'OK_CANCEL' },
    Button: { YES: 'YES', NO: 'NO', OK: 'OK', CANCEL: 'CANCEL' },
    createMenu: () => { const m = { addItem: () => m, addSeparator: () => m, addSubMenu: () => m, addToUi: () => m }; return m; },
    showModalDialog: (html, title) => { dialogs.push({ html, title }); }
  };
  const htmlOutput = content => {
    const o = { content, title: '', setWidth: () => o, setHeight: () => o, setTitle: t => { o.title = t; return o; }, getContent: () => o.content,
      setXFrameOptionsMode: () => o, append: s => { o.content += s; return o; } };
    return o;
  };
  const globals = () => ({
    console: { log() {}, warn() {}, error() {}, info() {} },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ss,
      getUi: () => ui,
      flush: () => {},
      newDataValidation: () => { const b = { requireValueInList: () => b, requireValueInRange: () => b, setAllowInvalid: () => b, build: () => ({}) }; return b; },
      CopyPasteType: { PASTE_FORMAT: 'PASTE_FORMAT', PASTE_VALUES: 'PASTE_VALUES', PASTE_NORMAL: 'PASTE_NORMAL' },
      Dimension: { ROWS: 'ROWS', COLUMNS: 'COLUMNS' }
    },
    PropertiesService: {
      getScriptProperties: () => propStore(props.script),
      getDocumentProperties: () => propStore(props.document)
    },
    Utilities: {
      formatDate: (d, tz, fmt) => fmt
        .replace('yyyy', d.getUTCFullYear())
        .replace('MM', String(d.getUTCMonth() + 1).padStart(2, '0'))
        .replace('dd', String(d.getUTCDate()).padStart(2, '0')),
      sleep: () => {}
    },
    HtmlService: {
      createHtmlOutput: s => htmlOutput(s || ''),
      createHtmlOutputFromFile: name => htmlOutput(fs.readFileSync(path.join(ROOT, 'src', name.replace(/\.html$/, '') + '.html'), 'utf8')),
      createTemplateFromFile: name => {
        const t = { evaluate: () => htmlOutput(fs.readFileSync(path.join(ROOT, 'src', name.replace(/\.html$/, '') + '.html'), 'utf8')) };
        return t;
      },
      XFrameOptionsMode: { ALLOWALL: 'ALLOWALL' }
    },
    LockService: { getDocumentLock: () => ({ tryLock: () => true, waitLock: () => {}, releaseLock: () => {} }) },
    UrlFetchApp: { fetch: () => { throw new Error('No network in tests'); } },
    Logger: { log() {} }
  });
  const script = new vm.Script(srcFiles().map(f => `// ${path.relative(ROOT, f)}\n` + fs.readFileSync(f, 'utf8')).join('\n;\n'),
    { filename: 'src (concatenated)' });
  // Each call is a new execution, as in Apps Script: every file runs again from the top (so globals
  // like main.js's latestSheet see the sheets as they are now, and caches start empty), then expr.
  const runRaw = expr => {
    const ctx = vm.createContext(globals());
    script.runInContext(ctx);
    return vm.runInContext(expr, ctx);
  };
  // Results cross from the script's realm as plain data, so assert.deepEqual works on them.
  const run = expr => { const r = runRaw(expr); return r === undefined ? undefined : JSON.parse(JSON.stringify(r)); };
  const resetCaches = () => {};   // every run() starts with empty caches anyway

  // --- sheets that the code itself creates ---
  run('createCategoriesSheet()');

  // --- Daily Archives and the Total Archive, lined up with Latest ---
  const titles = [];
  for (let r = 2; r <= lastRow; r++) titles[r] = r >= 28 && r <= 49 ? '' : String(latest.peek(r, 5).v);
  const dailyOf = r => { const v = latest.peek(r, 7).v; return typeof v === 'number' ? v : 0; };
  const firstSong = 50;
  const makeArchive = (name, newest, count) => {
    const s = ss.insertSheet(name);
    for (let r = 2; r <= lastRow; r++) if (titles[r]) put(s, r, 1, titles[r]);
    for (let i = 0; i < count; i++) {
      const c = i + 2;
      put(s, 1, c, addDays(newest, -i));
      for (let r = firstSong; r <= lastRow; r++) if (titles[r]) put(s, r, c, dailyOf(r));
    }
    return s;
  };
  const a2026 = makeArchive('Daily Archive 2026', day(2026, 9, 22), 265);   // back to Jan 1
  const a2025 = makeArchive('Daily Archive 2025', day(2025, 12, 31), 20);   // back to Dec 12
  makeArchive('Daily Archive 2024', day(2024, 12, 31), 0);
  makeArchive('Daily Archive 2023', day(2023, 12, 31), 0);
  const total = ss.insertSheet('Total Archive');
  for (let r = 2; r <= lastRow; r++) if (titles[r]) put(total, r, 1, titles[r]);
  for (let i = 0; i < 12; i++) {
    put(total, 1, i + 2, addDays(day(2026, 9, 1), -30 * i));
    for (let r = firstSong; r <= lastRow; r++) if (titles[r]) put(total, r, i + 2, Number(latest.peek(r, 6).v) || 0);
  }
  // Aggregate formulas in every archive column, as the updates wrote them.
  // Then, as in the real sheet, every column but today's B holds them as plain numbers (365 x 26
  // SUMs per sheet made it lag), so only B keeps its formulas.
  [a2026, a2025].forEach(s => {
    const cols = s.getLastColumn();
    const perColumn = run(`(() => { const out = []; for (let c = 2; c <= ${cols}; c++) out.push(buildAlbumFormulas(columnToLetter(c))); return out; })()`);
    perColumn.forEach((f, i) => s.getRange(2, i + 2, f.length, 1).setFormulas(f));
    const firstValueCol = s === a2026 ? 3 : 2;
    if (cols >= firstValueCol) {
      const block = s.getRange(2, firstValueCol, perColumn[0].length, cols - firstValueCol + 1);
      block.setValues(block.getValues());
    }
  });
  run('refreshTotalArchiveFormulas()');
  run('setLatestAggregateFormulas()');
  resetCaches();
  alerts.length = 0;

  return {
    ss, run, runRaw, alerts, dialogs, props, resetCaches,
    sheet: name => ss.getSheetByName(name),
    answer: (...a) => answers.push(...a),
    /** Moves every raw import count up by the song's usual daily, as if Spotify had refreshed. */
    nextDayImport(factor) {
      factor = factor || 1;
      const imp2 = ss.getSheetByName(options.importSheetName);
      const idToDaily = {};
      const tl = tracklist.getDataRange().getValues();
      tl.slice(1).forEach(r => { if (r[6]) idToDaily[r[6]] = dailyOf(Number(r[0])); });
      const n = imp2.getLastRow();
      for (let r = 2; r <= n; r++) {
        const x = imp2.cell(r, 7);
        if (typeof x.v === 'number') x.v += Math.round((idToDaily[imp2.peek(r, 8).v] || 100) * factor);
      }
      ss.touch();
    },
    lastAlert: () => alerts[alerts.length - 1] || [],
    alertText: () => alerts.map(a => a[0] + '\n' + (a[1] || '')).join('\n---\n')
  };
}

module.exports = { loadFixture, parseCsv };

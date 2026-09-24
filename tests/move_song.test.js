// Update → Move a Song (2.1): the same row moves in Latest (A:P) and every archive, the Tracklist
// is renumbered, and the album history is rebuilt from the Tracklist.
const test = require('node:test');
const assert = require('node:assert/strict');
const vm = require('vm');
const { loadFixture } = require('./lib/fixture');

const ARCHIVES = ['Daily Archive 2026', 'Daily Archive 2025', 'Daily Archive 2024', 'Daily Archive 2023', 'Total Archive'];
const titleAt = (f, sheet, row) => f.sheet(sheet).value(row, sheet === 'Latest' ? 5 : 1);
const tracklistRow = (f, title) => f.run('getSongs()').find(s => s.title === title);
const checks = (f, names) => f.run(`runChecks(HEALTH_CHECKS.filter(c => ${JSON.stringify(names)}.indexOf(c.name) !== -1))`);
const assertChecksOk = (f, names) => checks(f, names).forEach(r =>
  assert.equal(r.status, 'ok', `${r.name}: ${r.summary} ${r.details.slice(0, 4).join(' | ')}`));
const move = (f, from, to, category) => f.run(`moveSongFromDialog(${from}, ${to}, ${JSON.stringify(category)})`);
const snapshot = (f, sheet, row, cols) => f.sheet(sheet).getRange(row, 1, 1, cols).getValues()[0];

test('moving a song down inside its album: every sheet, history included', () => {
  const f = loadFixture();
  const title = titleAt(f, 'Latest', 52);                 // Teardrops on My Guitar (Debut, 50-64)
  const next = titleAt(f, 'Latest', 53);
  const latestRow = snapshot(f, 'Latest', 52, 16);
  const history = snapshot(f, 'Daily Archive 2026', 52, 40);
  const cat = tracklistRow(f, title).category;

  const r = move(f, 52, 60, cat);
  assert.ok(r.ok, r.message);
  ARCHIVES.concat(['Latest']).forEach(s => {
    assert.equal(titleAt(f, s, 60), title, s);
    assert.equal(titleAt(f, s, 52), next, s);
  });
  // Values travel with the song; H's formula still points at its own row.
  assert.deepEqual(snapshot(f, 'Latest', 60, 7), latestRow.slice(0, 7));
  assert.equal(f.sheet('Latest').peek(60, 8).f, '=G60/M60-1');
  assert.deepEqual(snapshot(f, 'Daily Archive 2026', 60, 40), history);
  assert.equal(tracklistRow(f, title).row, 60);
  assert.equal(tracklistRow(f, next).row, 52);
  assertChecksOk(f, ['Row alignment', 'Totals add up', 'Aggregate formulas', 'Category blocks', 'Covers']);
});

test('moving a song up into another category rewrites that category\'s history', () => {
  const f = loadFixture();
  // A remix (row 500, Remixes and etc.) to the end of Droplets, which ends at row 484 (Patient Zero).
  const title = titleAt(f, 'Latest', 500);
  const archive = f.sheet('Daily Archive 2026');
  const col = 30;                                          // some day in the past
  const droplets = f.run("getCategories().find(c => c.name === 'Droplets').row");
  const remixes = f.run("getCategories().find(c => c.name === 'Remixes and etc.').row");
  const dropletsBefore = archive.value(droplets, col);
  const remixesBefore = archive.value(remixes, col);
  const songDay = archive.value(500, col);

  const r = move(f, 500, 485, 'Droplets');
  assert.ok(r.ok, r.message);
  assert.match(r.message, /moved from row 500 to row 485, now in Droplets/);
  assert.equal(tracklistRow(f, title).row, 485);
  assert.equal(tracklistRow(f, title).category, 'Droplets');
  assert.equal(archive.value(droplets, col), dropletsBefore + songDay);   // retroactive
  assert.equal(archive.value(remixes, col), remixesBefore - songDay);
  assert.equal(archive.value(2, col), archive.value(2, col));            // the artist total is unchanged
  const totalBefore = f.sheet('Total Archive').value(2, 3);
  assert.ok(totalBefore > 0);
  assertChecksOk(f, ['Row alignment', 'Totals add up', 'Aggregate formulas', 'Category blocks']);
});

test('formulas elsewhere that point at the song follow it', () => {
  const f = loadFixture();
  const albums = f.sheet('Albums');
  albums.getRange('B5').setFormula('=Latest!F52');
  albums.getRange('B6').setFormula('=Latest!F53');
  const value52 = f.sheet('Latest').value(52, 6);
  const r = move(f, 52, 60, tracklistRow(f, titleAt(f, 'Latest', 52)).category);
  assert.ok(r.ok, r.message);
  assert.equal(albums.peek(5, 2).f, '=Latest!F60');
  assert.equal(albums.value(5, 2), value52);
  assert.equal(albums.peek(6, 2).f, '=Latest!F52');         // the song below moved up, and its reference with it
});

test('only the category changes when the row stays the same', () => {
  const f = loadFixture();
  const title = titleAt(f, 'Latest', 441);                  // Track by Track: no category today
  assert.equal(tracklistRow(f, title).category, '');
  const r = move(f, 441, 441, 'The Life of a Showgirl');
  assert.ok(r.ok, r.message);
  assert.match(r.message, /stays on row 441, now in The Life of a Showgirl/);
  assert.match(r.message, /no longer on consecutive rows|split/);
  assert.equal(tracklistRow(f, title).category, 'The Life of a Showgirl');
  const sum = f.run("buildAlbumFormulas('B')")[19 - 2][0];
  assert.match(sum, /B441/);
  const blocks = checks(f, ['Category blocks'])[0];
  assert.equal(blocks.status, 'warn');
});

test('nothing changes when the move is refused', () => {
  const f = loadFixture();
  const before = JSON.stringify(f.sheet('Latest').getRange(50, 1, 491, 16).getValues());
  const cases = [
    [49, 60, 'Lover', /Row 49 has no song/],
    [52, 9999, 'Lover', /outside the songs/],
    [52, 60, 'Not A Category', /not in the Categories sheet/],
    [52, 52, 'Taylor Swift (Debut)', /already is/]
  ];
  cases.forEach(([from, to, cat, re]) => {
    const r = move(f, from, to, cat);
    assert.equal(r.ok, false);
    assert.match(r.message, re);
  });
  f.sheet('Latest').addMerge(300, 15, 3);                 // O:Q, sticking out of A:P
  const r = move(f, 52, 60, 'Taylor Swift (Debut)');
  assert.equal(r.ok, false);
  assert.match(r.message, /stick out of columns A:P/);
  assert.equal(JSON.stringify(f.sheet('Latest').getRange(50, 1, 491, 16).getValues()), before);
});

test('past figures are numbers: only the moved song\'s streams change hands, everything else stays exact', () => {
  const f = loadFixture();
  const a2025 = f.sheet('Daily Archive 2025');
  const archive = f.sheet('Daily Archive 2026');
  a2025.getRange(10, 4).setValue(123456);                   // Lover on one day: a figure that isn't its songs' sum
  const before2026 = archive.getRange(2, 3, 26, 200).getValues();
  const song = archive.getRange(500, 3, 1, 200).getValues()[0];
  const r = move(f, 500, 485, 'Droplets');                   // Remixes (26) -> Droplets (20)
  assert.ok(r.ok, r.message);
  assert.match(r.message, /Album history adjusted: \d+ past figure\(s\) changed, for 1 song/);
  assert.equal(a2025.value(10, 4), 123456);
  assert.equal(a2025.peek(10, 4).f, '');
  const after2026 = archive.getRange(2, 3, 26, 200).getValues();
  for (let r2 = 0; r2 < 26; r2++) for (let c = 0; c < 200; c++) {
    const row = r2 + 2;
    const expected = row === 20 ? before2026[r2][c] + song[c] : row === 26 ? before2026[r2][c] - song[c] : before2026[r2][c];
    assert.equal(after2026[r2][c], expected, `row ${row}, column ${c + 3}`);
    assert.equal(archive.peek(row, c + 3).f, '');            // still numbers, no formulas added
  }
  assert.equal(archive.peek(20, 2).f, '=SUM(B477:B485)');    // today's column keeps formulas
});

test('moving a song out of Features raises the solo total\'s history', () => {
  const f = loadFixture();
  const archive = f.sheet('Daily Archive 2026');
  const features = f.run("getCategories().find(c => c.name === 'Features').row");
  const soloBefore = archive.value(3, 10), featBefore = archive.value(features, 10), totalBefore = archive.value(2, 10);
  const s = archive.value(470, 10);
  assert.ok(s > 0);
  const r = move(f, 470, 470, 'Remixes and etc.');
  assert.ok(r.ok, r.message);
  assert.equal(archive.value(features, 10), featBefore - s);
  assert.equal(archive.value(3, 10), soloBefore + s);
  assert.equal(archive.value(2, 10), totalBefore);
});

test('the daily update still runs and checks out after a move', () => {
  const f = loadFixture();
  assert.ok(move(f, 500, 485, 'Droplets').ok);
  f.nextDayImport();
  f.run('main()');
  assert.ok(!f.alerts.some(a => /Aborted|did not pass/.test(a[0])), f.alertText());
  assertChecksOk(f, ['Row alignment', 'Totals add up', 'Aggregate formulas', 'Latest vs Import', 'Date sequence']);
});

const setTracklist = (f, row, col, value) => {
  const tl = f.sheet('Tracklist');
  for (let r = 2; r <= tl.getLastRow(); r++) if (tl.value(r, 1) === row) tl.getRange(r, col).setValue(value);
};

test('a category changed by hand: the check warns, and Rebuild Album History moves the history', () => {
  const f = loadFixture();
  const a2025 = f.sheet('Daily Archive 2025');
  const showgirlBefore = a2025.value(19, 5), s = a2025.value(441, 5);
  setTracklist(f, 441, 3, 'The Life of a Showgirl');         // Track by Track: no category -> Showgirl
  const warn = checks(f, ['Category history'])[0];
  assert.equal(warn.status, 'warn');
  assert.match(warn.details[0], /Row 441 .*no category → The Life of a Showgirl/);

  f.run('rebuildAlbumHistoryMenu()');
  assert.match(f.alerts[0][1], /1 song\(s\) changed category/);
  assert.match(f.lastAlert()[1], /Album history adjusted: \d+ past figure\(s\) changed, for 1 song/);
  assert.equal(a2025.value(19, 5), showgirlBefore + s);
  assertChecksOk(f, ['Totals add up', 'Aggregate formulas', 'Category history']);
  // Running it again changes nothing.
  f.run('rebuildAlbumHistoryMenu()');
  assert.equal(a2025.value(19, 5), showgirlBefore + s);
  assert.match(f.lastAlert()[1], /No song had changed category/);
});

test('without a historyCategory column: the first rebuild creates it; an old change can then be told to it', () => {
  // Your situation: categories changed and rebuilt before 2.1's fix, history still in the old ones.
  const f = loadFixture({ historyCategory: false });
  const a2025 = f.sheet('Daily Archive 2025');
  const droplets = 20, remixes = 26;
  const dropletsBefore = a2025.value(droplets, 5), remixesBefore = a2025.value(remixes, 5), s = a2025.value(483, 5);
  setTracklist(f, 483, 3, 'Remixes and etc.');               // changed by hand earlier

  assert.equal(checks(f, ['Category history'])[0].status, 'warn');
  f.run('rebuildAlbumHistoryMenu()');
  assert.match(f.lastAlert()[1], /Created the "historyCategory" column/);
  assert.equal(a2025.value(droplets, 5), dropletsBefore);    // nothing moved: it couldn't know
  const header = f.sheet('Tracklist').getRange(1, 1, 1, 12).getValues()[0];
  const col = header.indexOf('historyCategory') + 1;
  assert.ok(col > 0);

  setTracklist(f, 483, col, 'Droplets');                     // tell it where the history still is
  f.run('rebuildAlbumHistoryMenu()');
  assert.equal(a2025.value(droplets, 5), dropletsBefore - s);
  assert.equal(a2025.value(remixes, 5), remixesBefore + s);
  assertChecksOk(f, ['Category history']);
});

test('a new song added through Pending gets its historyCategory', () => {
  const f = loadFixture();
  f.run('createPendingSheet()');
  f.run('createIgnoredSheet()');
  f.sheet('Pending').getRange(2, 1, 1, 5).setValues([['Another Lover Song', 'Lover', 'lover', '', 'upcoming']]);
  f.run('addPendingSongsMenu()');
  assertChecksOk(f, ['Category history']);
});

test('the dialog loads its data, and its script is valid JavaScript', () => {
  const f = loadFixture();
  const d = f.run('getMoveSongData()');
  assert.equal(d.firstRow, 50);
  assert.equal(d.lastRow, 540);
  assert.equal(d.songs.length, 491);
  assert.ok(d.categories.includes('Droplets'));
  const html = f.run('moveSongDialogHtml()');
  const script = html.match(/<script>([\s\S]*)<\/script>/)[1];
  assert.doesNotThrow(() => new vm.Script(script));
  f.run('moveSongMenu()');
  assert.equal(f.dialogs.length, 1);
  assert.match(f.dialogs[0].title, /Move a Song/);
});

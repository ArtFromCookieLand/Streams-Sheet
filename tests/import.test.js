// The daily update working from the raw import alone (2.1): totals by track ID, dailies against
// Latest, and the guards that used to live in Tools!C1 and the #MISSING marker.
const test = require('node:test');
const assert = require('node:assert/strict');
const { loadFixture } = require('./lib/fixture');

const FIRST = 50;
const statuses = f => f.run('runChecks(HEALTH_CHECKS)').map(r => `${r.status} ${r.name}: ${r.summary}`);
const archiveColumns = f => f.sheet('Daily Archive 2026').getLastColumn();
const rawCount = (f, id) => {
  const imp = f.sheet('Import');
  for (let r = 2; r <= imp.getLastRow(); r++) if (imp.value(r, 8) === id) return imp.value(r, 7);
  return undefined;
};
const songAt = (f, row) => f.run('getSongs()').find(s => s.row === row);

test('the fixture passes every check before anything is run', () => {
  const f = loadFixture();
  const failed = statuses(f).filter(s => s.startsWith('fail'));
  assert.deepEqual(failed, []);
});

test('an import no newer than Latest is refused, and nothing changes', () => {
  const f = loadFixture();
  const cols = archiveColumns(f);
  const date = f.sheet('Latest').value(1, 17).getTime();
  f.run('main()');
  assert.match(f.alerts[0][0], /Update Aborted/);
  assert.match(f.alerts[0][1], /sum of today's daily streams is 0/);
  assert.equal(archiveColumns(f), cols);
  assert.equal(f.sheet('Latest').value(1, 17).getTime(), date);
});

test('a fresh import is archived: totals by track ID, dailies against Latest', () => {
  const f = loadFixture();
  f.nextDayImport();
  const latest = f.sheet('Latest');
  const song = songAt(f, 52);                  // Teardrops on My Guitar
  const before = latest.value(52, 6);
  const expectedTotal = rawCount(f, song.trackId);
  const cols = archiveColumns(f);

  f.run('main()');
  const titles = f.alerts.map(a => a[0]);
  assert.ok(titles.some(t => /Confirm Update/.test(t)), f.alertText());
  assert.ok(!titles.some(t => /Aborted|did not pass|Error/.test(t)), f.alertText());

  const archive = f.sheet('Daily Archive 2026');
  assert.equal(archive.getLastColumn(), cols + 1);
  assert.equal(archive.value(1, 2).toISOString().slice(0, 10), '2026-09-23');
  assert.equal(latest.value(52, 6), expectedTotal);
  assert.equal(latest.value(52, 7), expectedTotal - before);
  assert.equal(archive.value(52, 2), expectedTotal - before);
  // Aggregates are formulas over today's column, and they add up.
  assert.equal(archive.peek(4, 2).f, '=SUM(B50:B64)');
  let debut = 0; for (let r = 50; r <= 64; r++) debut += archive.value(r, 2);
  assert.equal(archive.value(4, 2), debut);
  // Nothing left to add up afterwards, so a second run is refused.
  assert.equal(f.sheet('Import').value(1, 3), 0);
  f.alerts.length = 0;
  f.run('main()');
  assert.match(f.alerts[0][1], /sum of today's daily streams is 0/);
  assert.equal(archive.getLastColumn(), cols + 1);
});

test('after the update the quick checks pass, Latest vs Import included', () => {
  const f = loadFixture();
  f.nextDayImport();
  f.run('main()');
  f.resetCaches();
  const after = f.run("runChecks(HEALTH_CHECKS.filter(c => AFTER_UPDATE_CHECKS.indexOf(c.name) !== -1))");
  after.forEach(r => assert.equal(r.status, 'ok', `${r.name}: ${r.summary} ${r.details.slice(0, 3).join(' | ')}`));
});

test('between import and update, Latest vs Import is a warning, not a failure', () => {
  const f = loadFixture();
  f.nextDayImport();
  const r = f.run("runChecks(HEALTH_CHECKS.filter(c => c.name === 'Latest vs Import'))")[0];
  assert.equal(r.status, 'warn');
});

test('a track ID missing from the import stops the update before anything is written', () => {
  const f = loadFixture();
  f.nextDayImport();
  const tl = f.sheet('Tracklist');
  let sheetRow = 0;
  for (let r = 2; r <= tl.getLastRow(); r++) if (tl.value(r, 1) === 60) sheetRow = r;
  tl.getRange(sheetRow, 7).setValue('NotARealTrackId000000');
  f.resetCaches();
  const cols = archiveColumns(f);
  f.run('main()');
  assert.match(f.alerts[0][0], /Update Aborted/);
  assert.match(f.alerts[0][1], /row 60:/);
  assert.equal(archiveColumns(f), cols);
  const check = f.run("runChecks(HEALTH_CHECKS.filter(c => c.name === 'Import'))")[0];
  assert.equal(check.status, 'fail');
});

test('retired songs stay at 0; an upcoming song counts from the day its ID appears', () => {
  const f = loadFixture();
  // Give the upcoming song (Patient Zero, row 484) the ID of a track in the import nobody tracks.
  const tracked = {}; f.run('getSongs()').forEach(s => { tracked[s.trackId] = true; });
  const imp = f.sheet('Import');
  let id = null, count = 0;
  for (let r = 2; r <= imp.getLastRow() && !id; r++) if (!tracked[imp.value(r, 8)]) { id = imp.value(r, 8); count = imp.value(r, 7); }
  const tl = f.sheet('Tracklist');
  for (let r = 2; r <= tl.getLastRow(); r++) if (tl.value(r, 1) === 484) tl.getRange(r, 7).setValue(id);
  f.resetCaches();
  f.nextDayImport();
  count = rawCount(f, id);

  const today = f.run('readTodayFromImport()');
  assert.equal(today.nowLive.length, 1);
  f.run('main()');
  const latest = f.sheet('Latest');
  assert.equal(latest.value(484, 6), count);
  assert.equal(latest.value(484, 7), count);          // its whole total on day one, as before
  f.run('getSongs()').filter(s => s.status === 'retired').forEach(s => {
    assert.equal(latest.value(s.row, 6), 0);
    assert.equal(f.sheet('Daily Archive 2026').value(s.row, 2), 0);
  });
});

test('before the tidy-up the old Tools tab still works, and the tidy-up clears J:M', () => {
  const f = loadFixture({ importSheetName: 'Tools' });
  const tools = f.sheet('Tools');
  // The old block: a merged separator in J, Spotify names, totals and daily formulas.
  tools.addMerge(50, 10, 1);
  for (let r = FIRST; r <= 60; r++) {
    tools.getRange(r, 11).setValue('name ' + r);
    tools.getRange(r, 12).setValue(123);
    tools.getRange(r, 13).setFormula(`=(L${r}-Latest!F${r})`);
  }
  tools.getRange('C1').setFormula('=SUM(M50:M60)');
  const rawBefore = JSON.stringify(tools.getRange('E2:I20').getValues());

  f.nextDayImport();
  f.run('main()');
  assert.ok(f.alerts.some(a => /Confirm Update/.test(a[0])), f.alertText());

  f.alerts.length = 0;
  f.run('tidyImportSheet()');
  assert.equal(f.sheet('Tools'), null);
  const imp = f.sheet('Import');
  assert.ok(imp);
  assert.deepEqual(imp.getRange(1, 10, 70, 4).getValues().flat().filter(v => v !== ''), []);
  assert.equal(imp.merges.length, 0);
  assert.notEqual(JSON.stringify(imp.getRange('E2:I20').getValues()), rawBefore);  // the next-day counts, untouched by the tidy-up
  assert.equal(imp.peek(1, 3).f, '');
  assert.equal(imp.value(1, 3), 0);
  f.alerts.length = 0;
  f.run('tidyImportSheet()');
  assert.match(f.alerts[0][0], /Nothing changed/);
});

test('Add Pending Songs no longer touches the Import sheet', () => {
  const f = loadFixture();
  f.run('createPendingSheet()');
  f.run('createIgnoredSheet()');
  const pending = f.sheet('Pending');
  pending.getRange(2, 1, 1, 5).setValues([['A New Lover Song', 'Lover', 'lover', '', 'upcoming']]);
  const imp = f.sheet('Import');
  const before = JSON.stringify(imp.getDataRange().getValues());
  f.alerts.length = 0;
  f.run('addPendingSongsMenu()');
  assert.ok(f.alerts.some(a => /^Added 1 song/.test(a[0])), f.alertText());
  const song = f.run('getSongs()').find(s => s.title === 'A New Lover Song');
  assert.equal(song.row, 179);                         // after Lover's last song (178)
  assert.equal(f.sheet('Latest').value(179, 5), 'A New Lover Song');
  const after = imp.getDataRange().getValues();
  after[0][2] = JSON.parse(before)[0][2];              // C1, the sum of dailies, is rewritten
  assert.equal(JSON.stringify(after), before);
  f.resetCaches();
  const align = f.run("runChecks(HEALTH_CHECKS.filter(c => c.name === 'Row alignment' || c.name === 'Totals add up'))");
  align.forEach(r => assert.equal(r.status, 'ok', `${r.name}: ${r.summary} ${r.details.slice(0, 3).join(' | ')}`));
});

test('Totals add up ignores rounding noise under 1, and names fractional songs in a real mismatch', () => {
  const f = loadFixture();
  const latest = f.sheet('Latest');
  const remixes = 26;
  const check = () => f.run("runChecks(HEALTH_CHECKS.filter(c => c.name === 'Totals add up'))")[0];
  latest.getRange(500, 6).setValue(latest.value(500, 6) + 0.25);   // a song total with a fraction (e.g. random test data)
  latest.getRange(remixes, 6).setValue(latest.value(remixes, 6) + 0.0000001);
  assert.equal(check().status, 'ok');
  latest.getRange(remixes, 6).setValue(latest.value(remixes, 6) + 5);
  const r = check();
  assert.equal(r.status, 'fail');
  const line = r.details.find(d => /Remixes and etc\. \(Latest row 26\) total/.test(d));
  assert.match(line, /row 500 is \d+\.25/);
});

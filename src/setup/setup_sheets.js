/**
 * ===================================================================
 * SETUP: THE CATEGORIES, PENDING, IGNORED, SOURCES AND IMPORT SHEETS
 * -------------------------------------------------------------------
 * One-off helpers, under Update → Setup, shown only while the sheet is
 * missing. Each refuses to touch a sheet that already exists.
 *
 * CATEGORY_SEED is what CONFIG.CATEGORIES held before categories moved
 * into their own sheet (2.0.0.3). Once both spreadsheets have the sheet,
 * the Categories sheet is the only source; edit it there, not here.
 * ===================================================================
 */

// name, row, type, summaryCell, summaryLimit
const CATEGORY_SEED = [
  ['Taylor Swift (Debut)',                    4,  'studio', 'F23',  ''],
  ['Fearless (2008)',                         5,  'studio', 'F79',  ''],
  ['Speak Now (2010)',                        6,  'studio', 'L93',  17],
  ['Red (2012)',                              7,  'studio', 'L119', 19],
  ['1989 (2014)',                             8,  'studio', 'R23',  16],
  ['reputation',                              9,  'studio', 'R78',  ''],
  ['Lover',                                   10, 'studio', 'W23',  ''],
  ['folklore',                                11, 'studio', 'W51',  ''],
  ['evermore',                                12, 'studio', 'W78',  ''],
  ['Midnights',                               13, 'studio', 'W102', ''],
  ['The Tortured Poets Department',           14, 'studio', 'AB23', ''],
  ["Fearless (Taylor's Version)",             15, 'studio', 'F45',  ''],
  ["Red (Taylor's Version)",                  16, 'studio', 'L54',  ''],
  ["Speak Now (Taylor's Version)",            17, 'studio', 'L23',  ''],
  ["1989 (Taylor's Version)",                 18, 'studio', 'R49',  ''],
  ['The Life of a Showgirl',                  19, 'studio', 'AB89', 12],
  ['Droplets',                                20, 'other',  '',     ''],
  ['The Taylor Swift Holiday Collection',     21, 'other',  '',     ''],
  ['Live From Clear Channel Stripped 2008',   22, 'other',  '',     ''],
  ['Speak Now World Tour Live',               23, 'other',  '',     ''],
  ['Live From Paris',                         24, 'other',  '',     ''],
  ['Soundtracks',                             25, 'fixed',  'R99',  ''],
  ['Remixes and etc.',                        26, 'fixed',  '',     ''],
  ['Features',                                27, 'fixed',  '',     '']
];


/**
 * [MENU] Creates the Categories sheet from CATEGORY_SEED.
 */
function createCategoriesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const name = CONFIG.SHEETS.CATEGORIES;
  if (ss.getSheetByName(name)) {
    ui.alert('Nothing changed', `The "${name}" sheet already exists.`, ui.ButtonSet.OK);
    return;
  }

  const h = CONFIG.CATEGORIES_SHEET.HEADERS;
  const header = [h.name, h.row, h.type, h.summaryCell, h.summaryLimit];
  const sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sheet.getRange(2, 1, CATEGORY_SEED.length, header.length).setValues(CATEGORY_SEED);
  sheet.setFrozenRows(1);

  // Type is one of three words; a dropdown keeps typos out.
  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(CATEGORY_TYPES, true).build();
  sheet.getRange(2, 3, sheet.getMaxRows() - 1, 1).setDataValidation(typeRule);
  sheet.autoResizeColumns(1, header.length);

  ui.alert('Categories sheet created' + envTag(),
    `${CATEGORY_SEED.length} categories copied in. From now on this sheet is where categories are defined.\n\n` +
    'Type: studio = albums and re-recordings (in the discography summary), other = compilations and live albums, fixed = Soundtracks, Remixes, Features.',
    ui.ButtonSet.OK);
}


/**
 * [MENU] Creates an empty Pending sheet with dropdowns for category and status.
 */
function createPendingSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const name = CONFIG.SHEETS.PENDING;
  if (ss.getSheetByName(name)) {
    ui.alert('Nothing changed', `The "${name}" sheet already exists.`, ui.ButtonSet.OK);
    return;
  }
  const categories = ss.getSheetByName(CONFIG.SHEETS.CATEGORIES);
  if (!categories) {
    ui.alert('Create the Categories sheet first', 'The category dropdown in Pending reads from it.', ui.ButtonSet.OK);
    return;
  }

  const h = CONFIG.PENDING_SHEET.HEADERS;
  const header = [h.title, h.category, h.coverKey, h.trackId, h.status, h.spotifyTitle, h.album, h.result];
  const sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sheet.setFrozenRows(1);

  const rows = sheet.getMaxRows() - 1;
  const catCol = findHeaderColumns(categories.getRange(1, 1, 1, categories.getLastColumn()).getValues()[0],
    { name: CONFIG.CATEGORIES_SHEET.HEADERS.name }, CONFIG.SHEETS.CATEGORIES).name + 1;
  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(categories.getRange(2, catCol, categories.getMaxRows() - 1, 1), true).build();
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(CONFIG.PENDING_SHEET.STATUSES, true).build();
  sheet.getRange(2, 2, rows, 1).setDataValidation(catRule);
  sheet.getRange(2, 5, rows, 1).setDataValidation(statusRule);
  sheet.getRange(1, 8).setNote('Filled in by Add Pending Songs. A row whose result starts with ' + CONFIG.PENDING_SHEET.DONE_PREFIX + ' has been added and is skipped.');
  sheet.autoResizeColumns(1, header.length);

  ui.alert('Pending sheet created' + envTag(),
    'One row per song to add. Title, category and cover key are required; the track ID is required for "active" and optional for "upcoming" (a song not out yet). Spotify Title and album are optional - they are filled in from the import when the ID is there.\n\n' +
    'Then run Update → Add Pending Songs.', ui.ButtonSet.OK);
}


/**
 * [MENU] Creates the Sources sheet from CONFIG.APIFY.ALBUMS, after which the importer reads the
 * album URLs from there and a new release needs no code change.
 */
function createSourcesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const name = CONFIG.SHEETS.SOURCES;
  if (ss.getSheetByName(name)) {
    ui.alert('Nothing changed', `The "${name}" sheet already exists.`, ui.ButtonSet.OK);
    return;
  }

  const h = CONFIG.SOURCES_SHEET.HEADERS;
  const rows = Object.keys(CONFIG.APIFY.ALBUMS).map(albumName => [albumName, CONFIG.APIFY.ALBUMS[albumName]]);
  const sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, 2).setValues([[h.name, h.url]]).setFontWeight('bold');
  sheet.getRange(2, 1, rows.length, 2).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 320);
  sheet.setColumnWidth(2, 420);

  ui.alert('Sources sheet created' + envTag(),
    `${rows.length} albums copied in. The next import reads its URLs from here, so a new release just needs a row - no code change.\n\n` +
    'Only open.spotify.com/album/... links; the same album twice is refused.', ui.ButtonSet.OK);
}


/**
 * [MENU] Creates the Ignored sheet, starting it with every track in the current import that isn't
 * tracked or waiting in Pending - other artists on soundtracks, other editions of an album - so
 * that Find New Tracks only reports what is genuinely new from then on.
 */
function createIgnoredSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const name = CONFIG.SHEETS.IGNORED;
  if (ss.getSheetByName(name)) {
    ui.alert('Nothing changed', `The "${name}" sheet already exists.`, ui.ButtonSet.OK);
    return;
  }
  const raw = readRawImport();
  if (!raw.length) {
    ui.alert('Import first', `The raw import in ${CONFIG.SHEETS.IMPORT} has no track IDs, so there is nothing to start the Ignored sheet from.`, ui.ButtonSet.OK);
    return;
  }

  const known = knownTrackIds(getSongs());
  const untracked = raw.filter(t => !known[t.id]);
  const response = ui.alert('Create Ignored sheet' + envTag(),
    `${untracked.length} of the ${raw.length} tracks in the current import aren't tracked. They'll all be marked as ignored, so Find New Tracks only reports tracks that appear after today.\n\n` +
    'Anything in there you do want to track can be added later through Pending.\n\nGo ahead?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const h = CONFIG.IGNORED_SHEET.HEADERS;
  const header = [h.trackId, h.spotifyTitle, h.album, h.reason, h.date];
  const sheet = ss.insertSheet(name);
  sheet.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  if (untracked.length) {
    addToIgnored(untracked.map(t => ({ trackId: t.id, spotifyTitle: t.name, album: t.album, reason: 'Not tracked when detection started' })));
  }
  sheet.autoResizeColumns(1, header.length);
  ui.alert('Ignored sheet created' + envTag(), `${untracked.length} track(s) marked as ignored.`, ui.ButtonSet.OK);
}


/**
 * [MENU] 2.1: renames Tools to Import and clears its old per-song block in J:M (separator, Spotify
 * name, total, daily). Since 2.1 the update works each song's total and daily out from the raw
 * import itself, so nothing reads that block, and the sheet no longer has to line up with the
 * others. The raw import in E:I and whatever is in A:D stay as they are.
 */
function tidyImportSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const name = CONFIG.SHEETS.IMPORT, oldName = CONFIG.SHEETS.IMPORT_OLD_NAME;
  if (ss.getSheetByName(name)) {
    ui.alert('Nothing changed', `The "${name}" sheet already exists.`, ui.ButtonSet.OK);
    return;
  }
  const sheet = ss.getSheetByName(oldName);
  if (!sheet) {
    ui.alert('Nothing changed', `There is no "${oldName}" sheet to tidy up.`, ui.ButtonSet.OK);
    return;
  }

  const b = CONFIG.IMPORT.OLD_BLOCK;
  const block = columnToLetter(b.COLUMN) + ':' + columnToLetter(b.COLUMN + b.WIDTH - 1);
  const response = ui.alert('Tidy up the Import sheet' + envTag(),
    `This renames "${oldName}" to "${name}" and clears columns ${block} (the old separator, Spotify name, total and daily per song). ` +
    `Nothing reads them any more: the update now works each song's total and daily out from the raw import by track ID.\n\n` +
    `The raw import (${CONFIG.IMPORT.RAW_DATA}) and columns A:D stay as they are. Go ahead?`,
    ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const range = sheet.getRange(1, b.COLUMN, sheet.getMaxRows(), b.WIDTH);
  range.breakApart();
  range.clear();
  sheet.setName(name);

  let sumText = '';
  try {
    const today = matchTotalsById();   // C1 held a formula over the old block; it becomes a plain figure
    sumText = '\n\n' + describeMatchResult(today);
  } catch (error) {
    sumText = '\n\nChecking the import failed: ' + error.message;
  }
  ui.alert('Import sheet tidied up' + envTag(), `"${oldName}" is now "${name}", and ${block} is empty.` + sumText, ui.ButtonSet.OK);
}

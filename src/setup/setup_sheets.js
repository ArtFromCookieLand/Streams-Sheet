/**
 * ===================================================================
 * SETUP: THE CATEGORIES AND PENDING SHEETS
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
    .requireValueInList([SONG_STATUS.ACTIVE, SONG_STATUS.UPCOMING], true).build();
  sheet.getRange(2, 2, rows, 1).setDataValidation(catRule);
  sheet.getRange(2, 5, rows, 1).setDataValidation(statusRule);
  sheet.getRange(1, 8).setNote('Filled in by Add Pending Songs. A row whose result starts with ' + CONFIG.PENDING_SHEET.DONE_PREFIX + ' has been added and is skipped.');
  sheet.autoResizeColumns(1, header.length);

  ui.alert('Pending sheet created' + envTag(),
    'One row per song to add. Title, category and cover key are required; the track ID is required for "active" and optional for "upcoming" (a song not out yet). Spotify Title and album are optional - they are filled in from the import when the ID is there.\n\n' +
    'Then run Update → Add Pending Songs.', ui.ButtonSet.OK);
}

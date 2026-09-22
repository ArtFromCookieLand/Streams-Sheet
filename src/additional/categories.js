/**
 * ===================================================================
 * CATEGORIES
 * -------------------------------------------------------------------
 * The `Categories` sheet defines every album/category: its name (as
 * used in the Tracklist's "category" column), its row in Latest and in
 * every archive (CONFIG.LAYOUT: rows 4 up to LAST_AGGREGATE_ROW), its
 * type, and optionally where its text summary goes.
 *
 * Type decides two things:
 *   studio - an album or re-recording: in the discography summary, and
 *            a new one is placed after the last studio album
 *   other  - compilations, live albums, Droplets: a new one is placed
 *            after the last "other"
 *   fixed  - Soundtracks, Remixes, Features: always kept at the end
 *
 * Columns are found by header (CONFIG.CATEGORIES_SHEET.HEADERS),
 * ignoring case and spaces; the row order of the sheet doesn't matter.
 * ===================================================================
 */

const CATEGORY_TYPES = ['studio', 'other', 'fixed'];

var _categoriesCache = null;

/**
 * @return {Array<{name: string, row: number, type: string, summaryCell: string, summaryLimit: number}>}
 *   In row order. Throws on the first problem; readCategories() lists them all.
 */
function getCategories() {
  if (_categoriesCache) return _categoriesCache;
  const result = readCategories();
  if (result.problems.length) throw new Error(result.problems[0]);
  _categoriesCache = result.categories;
  return _categoriesCache;
}

/**
 * Reads and validates the Categories sheet without throwing.
 * @return {{categories: Array, problems: Array<string>}}
 */
function readCategories() {
  const name = CONFIG.SHEETS.CATEGORIES;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sheet) return { categories: [], problems: [`The "${name}" sheet is missing - run Update → Setup → Create Categories sheet.`] };

  const values = sheet.getDataRange().getValues();
  let col;
  try {
    col = findHeaderColumns(values[0], CONFIG.CATEGORIES_SHEET.HEADERS, name);
  } catch (error) {
    return { categories: [], problems: [error.message] };
  }

  const layout = CONFIG.LAYOUT;
  const categories = [];
  const problems = [];
  const nameOfRow = {};
  const seenNames = {};

  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    const catName = String(r[col.name]).trim();
    if (!catName) continue;

    const row = Number(r[col.row]);
    const type = String(r[col.type]).trim().toLowerCase();
    const summaryCell = String(r[col.summaryCell]).trim().toUpperCase();
    const limitCell = r[col.summaryLimit];
    const summaryLimit = limitCell === '' ? 0 : Number(limitCell);

    if (seenNames[catName]) { problems.push(`${name}: "${catName}" is listed twice.`); continue; }
    seenNames[catName] = true;
    if (!Number.isInteger(row) || row <= layout.SOLO_ROW || row > layout.LAST_AGGREGATE_ROW) {
      problems.push(`${name}: "${catName}" has row "${r[col.row]}"; categories go in rows ${layout.SOLO_ROW + 1}-${layout.LAST_AGGREGATE_ROW}.`);
      continue;
    }
    if (nameOfRow[row]) { problems.push(`${name}: "${nameOfRow[row]}" and "${catName}" both use row ${row}.`); continue; }
    nameOfRow[row] = catName;
    if (CATEGORY_TYPES.indexOf(type) === -1) {
      problems.push(`${name}: "${catName}" has type "${r[col.type]}"; expected one of: ${CATEGORY_TYPES.join(', ')}.`);
    }
    if (summaryCell && !/^[A-Z]+\d+$/.test(summaryCell)) {
      problems.push(`${name}: "${catName}" has summary cell "${r[col.summaryCell]}", which is not a single cell like F23.`);
    }
    if (!Number.isInteger(summaryLimit) || summaryLimit < 0) {
      problems.push(`${name}: "${catName}" has summary limit "${limitCell}"; leave it blank or use a whole number.`);
    }

    categories.push({ name: catName, row: row, type: type, summaryCell: summaryCell, summaryLimit: summaryLimit });
  }

  if (!categories.length && !problems.length) problems.push(`The "${name}" sheet has no categories.`);
  if (categories.length && !categories.some(c => c.name === layout.SOLO_EXCLUDES)) {
    problems.push(`CONFIG.LAYOUT.SOLO_EXCLUDES is "${layout.SOLO_EXCLUDES}", which is not in the ${name} sheet.`);
  }

  categories.sort((a, b) => a.row - b.row);
  return { categories: categories, problems: problems };
}

/**
 * @return {number} The last aggregate row in use: the highest category row.
 */
function getLastCategoryRow() {
  return Math.max.apply(null, getCategories().map(c => c.row));
}

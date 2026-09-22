function transferStats() {
  // --- 0. Build the album formulas first ---
  // Reads the Tracklist sheet, so any problem with it throws here, before the date is advanced
  // or a column is inserted - never halfway through.
  var albumFormulas = buildAlbumFormulas();
  var firstSongRow = CONFIG.LAYOUT.FIRST_SONG_ROW;
  var songRows = getSongRowCount();

  // --- 1. Advance date in Latest by +1 ---
  var latestDateCell = latestSheet.getRange(CONFIG.LATEST.DATE_CELL);
  var currentDate = new Date(latestDateCell.getValue());
  var newDate = new Date(currentDate.setDate(currentDate.getDate() + 1));
  latestDateCell.setValue(newDate);

  // --- 2. Insert new column in Daily Archive ---
  dailyArchiveSheet.insertColumnAfter(1);
  dailyArchiveSheet.setColumnWidth(2, 100);

  // Set new date header in B1
  dailyArchiveSheet.getRange('B1')
    .setValue(newDate)
    .setNumberFormat("yyyy/mm/dd")
    .setFontWeight("bold");

  // --- 3. Copy Tools -> Daily Archive (New Song Streams) ---
  var toolsDailyData = toolsSheet.getRange(firstSongRow, CONFIG.TOOLS.DAILY_COLUMN, songRows, 1).getValues();
  dailyArchiveSheet.getRange(firstSongRow, 2, songRows, 1)
    .setValues(toolsDailyData)
    .setNumberFormat("#,##0")
    .setFontColor("black")
    .setFontWeight("normal");

  // --- 4. Set Album Formulas ---
  // Calculates SUMs for albums in the new Col B
  setAlbumFormulas(dailyArchiveSheet, albumFormulas);

  // --- 5. Copy Tools -> Latest (Today's Data) ---
  // Tools L:M (total, daily) -> Latest F:G, song rows only. The album rows in Latest are formulas.
  var toolsData = toolsSheet.getRange(firstSongRow, CONFIG.TOOLS.TOTALS_COLUMN, songRows, 2).getValues();
  latestSheet.getRange(firstSongRow, CONFIG.LATEST.COLS.TOTAL, songRows, 2).setValues(toolsData);
}


/**
 * Copies yesterday's and last week's daily figures from the archive into Latest M and N, for
 * every row - albums and songs alike - then works out the "best since" dates.
 */
function updateStats(dailyArchiveSheet, latestSheet) {
  if (!dailyArchiveSheet) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    dailyArchiveSheet = ss.getSheetByName(CONFIG.SHEETS.ARCHIVE);
  }
  if (!latestSheet) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    latestSheet = ss.getSheetByName(CONFIG.SHEETS.LATEST);
  }

  var firstRow = CONFIG.LAYOUT.TOTAL_ROW;
  var numRows = getLastSongRow() - firstRow + 1;

  // Copy yesterday's streams (Col C)
  var archiveColC = dailyArchiveSheet.getRange(firstRow, CONFIG.ARCHIVE.YESTERDAY_COLUMN, numRows, 1).getValues();
  latestSheet.getRange(firstRow, CONFIG.LATEST.COLS.DAY_AGO, numRows, 1).setValues(archiveColC);

  // Copy streams from a week ago (Col I)
  var archiveColI = dailyArchiveSheet.getRange(firstRow, CONFIG.ARCHIVE.WEEK_AGO_COLUMN, numRows, 1).getValues();
  latestSheet.getRange(firstRow, CONFIG.LATEST.COLS.WEEK_AGO, numRows, 1).setValues(archiveColI);

  // Continue workflow
  transferBestSinceRows(dailyArchiveSheet, latestSheet);
}


/**
 * [WORKFLOW 3] Triggers the "Best Since" calculation across multiple Archive sheets.
 * Albums and songs share one pass: rows 2 to the last song, written to Latest column L.
 */
function transferBestSinceRows(dailyArchiveSheet, latestSheet) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Determine the "Active" sheet (Where today's value lives)
  if (!dailyArchiveSheet) {
    // This MUST be the current year sheet (e.g., 'Daily Archive 2026')
    dailyArchiveSheet = ss.getSheetByName(CONFIG.SHEETS.ARCHIVE);
  }

  if (!latestSheet) {
    latestSheet = ss.getSheetByName(CONFIG.SHEETS.LATEST);
  }

  // 2. Get the list of all historical sheets to check
  // Ensure this list in Config is ordered: ['2026', '2025', '2024']
  var archiveSheetNames = CONFIG.SHEETS.ARCHIVE_YEARS;

  var firstRow = CONFIG.LAYOUT.TOTAL_ROW;
  findBestSince(
    dailyArchiveSheet,   // Source of Truth (Today's value)
    archiveSheetNames,   // List of sheets to search through
    latestSheet,         // Where to write results
    firstRow,
    getLastSongRow() - firstRow + 1,
    CONFIG.LATEST.COLS.BEST_SINCE
  );
}


/**
 * [REUSABLE] Cascading Best Since Finder.
 * Iterates through multiple yearly sheets until a higher value is found.
 * @param {number} destColumn - Column in Latest to write the dates to, from sourceStartRow down.
 */
function findBestSince(activeSheet, sheetNames, latestSheet, sourceStartRow, numRows, destColumn) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- A. GET THRESHOLDS (Today's Values) ---
  // Always get the value to beat from Col B of the ACTIVE sheet (2026)
  var thresholds = activeSheet.getRange(sourceStartRow, 2, numRows, 1).getValues().flat();

  // Initialize results array with empty strings
  var finalResults = new Array(numRows).fill("");

  // --- B. CYCLE THROUGH ARCHIVE SHEETS ---
  for (var s = 0; s < sheetNames.length; s++) {

    // Optimization: If we found a date for EVERY song, stop looking.
    var allFound = finalResults.every(function(r) { return r !== ""; });
    if (allFound) break;

    // Load the sheet
    var currentSheetName = sheetNames[s];
    var sheet = ss.getSheetByName(currentSheetName);

    if (!sheet) {
      console.warn("Sheet not found: " + currentSheetName);
      continue;
    }

    // --- C. DYNAMIC COLUMN LOGIC ---
    // Check if this sheet is the "Current Active Year"
    var isCurrentYear = (sheet.getName() === activeSheet.getName());

    // If it's the Current Year (2026), skip Col B (Today). Start at Col 3 (C).
    // If it's a Past Year (2025), include Col B (Dec 31st). Start at Col 2 (B).
    var startCol = isCurrentYear ? 3 : 2;

    var lastCol = sheet.getLastColumn();

    // Calculate width based on the dynamic start column
    // Width = End - Start + 1
    var width = lastCol - startCol + 1;

    if (width < 1) continue; // Skip if no history exists yet

    // --- D. FETCH DATA ---
    // Fetch Data Matrix starting at startCol
    var valuesMatrix = sheet.getRange(sourceStartRow, startCol, numRows, width).getValues();
    var datesRow = sheet.getRange(1, startCol, 1, width).getValues()[0];

    // --- E. COMPARE ---
    for (var i = 0; i < numRows; i++) {
      // If we already found a date in a newer year, SKIP this song
      if (finalResults[i] !== "") continue;

      var threshold = thresholds[i];
      var rowValues = valuesMatrix[i];

      // Run the specific math helper
      var result = BESTSINCE(threshold, rowValues, datesRow);

      if (result !== "") {
        finalResults[i] = result;
      }
    }
  }

  // --- F. WRITE RESULTS ---
  var output = finalResults.map(function(r) { return [r]; });

  latestSheet.getRange(sourceStartRow, destColumn, numRows, 1).setValues(output);
}

/**
 * Writes the aggregate formulas into the new column B (rows 2 to the last category).
 * @param {Sheet} dailyArchiveSheet
 * @param {Array<Array<string>>} formulas - From buildAlbumFormulas(). Built by the caller
 *   before anything destructive happens, so a Tracklist sheet problem aborts the run cleanly.
 */
function setAlbumFormulas(dailyArchiveSheet, formulas) {
  if (!formulas) formulas = buildAlbumFormulas();

  dailyArchiveSheet.getRange(CONFIG.LAYOUT.TOTAL_ROW, 2, formulas.length, 1)
    .setFormulas(formulas)
    .setNumberFormat("#,##0")
    .setFontColor("black")
    .setFontWeight("normal");
}


/**
 * Builds the aggregate formulas for one archive column, from the total row (2) down to the
 * last category. Each category sums exactly the rows the Tracklist sheet assigns to it; songs
 * with a blank category (the Track by Track rows) are in none, but still in the artist total.
 * Spare rows between categories get an empty string, which clears them.
 * @param {string} [column] - Column letter the formulas are for. Defaults to B, today's column.
 * @return {Array<Array<string>>} One [formula] per row, starting at CONFIG.LAYOUT.TOTAL_ROW.
 */
function buildAlbumFormulas(column) {
  column = column || 'B';
  const layout = CONFIG.LAYOUT;
  const rowsByCategory = getRowsByCategory();

  const configured = CONFIG.CATEGORIES.map(c => c.name);
  Object.keys(rowsByCategory).forEach(name => {
    if (configured.indexOf(name) === -1) {
      throw new Error(`The ${CONFIG.SHEETS.SONGS} sheet uses category "${name}", which is not in CONFIG.CATEGORIES.`);
    }
  });

  const lastRow = Math.max.apply(null, CONFIG.CATEGORIES.map(c => c.row));
  const formulas = [];
  for (let r = layout.TOTAL_ROW; r <= lastRow; r++) formulas.push(['']);
  const put = (row, formula) => {
    if (formulas[row - layout.TOTAL_ROW][0] !== '') {
      throw new Error(`Two aggregates are configured for row ${row}.`);
    }
    formulas[row - layout.TOTAL_ROW] = [formula];
  };

  // Artist total over the whole song range, and the solo total (the artist total minus one category).
  put(layout.TOTAL_ROW, `=SUM(${column}${layout.FIRST_SONG_ROW}:${column}${getLastSongRow()})`);
  const excluded = CONFIG.CATEGORIES.filter(c => c.name === layout.SOLO_EXCLUDES)[0];
  if (!excluded) throw new Error(`CONFIG.LAYOUT.SOLO_EXCLUDES names "${layout.SOLO_EXCLUDES}", which is not a category.`);
  put(layout.SOLO_ROW, `=${column}${layout.TOTAL_ROW}-${column}${excluded.row}`);

  CONFIG.CATEGORIES.forEach(c => {
    if (c.row <= layout.SOLO_ROW || c.row > layout.LAST_AGGREGATE_ROW) {
      throw new Error(`Category "${c.name}" has row ${c.row}; categories go in rows ${layout.SOLO_ROW + 1}-${layout.LAST_AGGREGATE_ROW}.`);
    }
    put(c.row, sumOfRows(rowsByCategory[c.name] || [], column));
  });
  return formulas;
}


/**
 * @param {Array<number>} rows - Ascending row numbers.
 * @param {string} column - Column letter.
 * @return {string} e.g. "=SUM(B100:B114)", or "=SUM(B434:B452,B600)" for a scattered category.
 */
function sumOfRows(rows, column) {
  if (!rows.length) return '=0';

  const parts = [];
  let start = rows[0];
  for (let i = 1; i <= rows.length; i++) {
    if (i < rows.length && rows[i] === rows[i - 1] + 1) continue;
    const end = rows[i - 1];
    parts.push(start === end ? `${column}${start}` : `${column}${start}:${column}${end}`);
    start = rows[i];
  }
  return `=SUM(${parts.join(',')})`;
}

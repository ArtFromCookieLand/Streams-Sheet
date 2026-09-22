function transferStats() {
  // --- 0. Build the album formulas first ---
  // Reads the Tracklist sheet, so any problem with it throws here, before the date is advanced
  // or a column is inserted - never halfway through.
  var albumFormulas = buildAlbumFormulas();

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
  var toolsDailyData = toolsSheet.getRange(CONFIG.TOOLS.DAILY_STREAMS).getValues(); 
  dailyArchiveSheet.getRange(CONFIG.ARCHIVE.DAILY_STREAMS_DEST)
    .setValues(toolsDailyData)
    .setNumberFormat("#,##0")
    .setFontColor("black")
    .setFontWeight("normal");

  // --- 4. Set Album Formulas ---
  // Calculates SUMs for albums in the new Col B
  setAlbumFormulas(dailyArchiveSheet, albumFormulas);

  // --- 5. Copy Tools -> Latest (Today's Data) ---
  var toolsData = toolsSheet.getRange(CONFIG.TOOLS.SONG_DATA).getValues();
  latestSheet.getRange(CONFIG.LATEST.SONG_DATA_DEST).setValues(toolsData); 
}


function updateStats(dailyArchiveSheet, latestSheet) {
  if (!dailyArchiveSheet) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    dailyArchiveSheet = ss.getSheetByName(CONFIG.SHEETS.ARCHIVE);
  }
  if (!latestSheet) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    latestSheet = ss.getSheetByName(CONFIG.SHEETS.LATEST);
  }

  // Copy yesterday's streams (Col C)
  var archiveColC = dailyArchiveSheet.getRange(CONFIG.ARCHIVE.YESTERDAY_STREAMS).getValues(); 
  latestSheet.getRange(CONFIG.LATEST.YESTERDAY_STREAMS_DEST).setValues(archiveColC);

  // Copy streams from a week ago (Col I)
  var archiveColI = dailyArchiveSheet.getRange(CONFIG.ARCHIVE.WEEK_DATA).getValues();
  latestSheet.getRange(CONFIG.LATEST.WEEK_STREAMS_DEST).setValues(archiveColI);

  // Continue workflow
  transferBestSinceRows(dailyArchiveSheet, latestSheet);
}


/**
 * [WORKFLOW 3] Triggers the "Best Since" calculations across multiple Archive sheets.
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

  // 3. Calculate for Songs
  findBestSince(
    dailyArchiveSheet,   // Source of Truth (Today's value)
    archiveSheetNames,   // List of sheets to search through
    latestSheet,         // Where to write results
    CONFIG.SONGS.START_ROW,
    CONFIG.SONGS.COUNT,
    CONFIG.LATEST.SONG_BEST_SINCE_DEST
  );
  
  // 4. Calculate for Albums
  findBestSince(
    dailyArchiveSheet,
    archiveSheetNames,
    latestSheet,
    CONFIG.ALBUMS.START_ROW,
    CONFIG.ALBUMS.COUNT,
    CONFIG.LATEST.ALBUM_BEST_SINCE_DEST
  );
}


/**
 * [REUSABLE] Cascading Best Since Finder.
 * Iterates through multiple yearly sheets until a higher value is found.
 */
function findBestSince(activeSheet, sheetNames, latestSheet, sourceStartRow, numRows, destA1Notation) {
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

  latestSheet.getRange(destA1Notation)
    .offset(0, 0, numRows, 1)
    .setValues(output);
}

/**
 * Writes the aggregate formulas into the new column B (rows 550-574).
 * @param {Sheet} dailyArchiveSheet
 * @param {Array<Array<string>>} formulas - From buildAlbumFormulas(). Built by the caller
 *   before anything destructive happens, so a Tracklist sheet problem aborts the run cleanly.
 */
function setAlbumFormulas(dailyArchiveSheet, formulas) {
  if (!formulas) formulas = buildAlbumFormulas();

  dailyArchiveSheet.getRange(CONFIG.ALBUMS.START_ROW, 2, formulas.length, 1)
    .setFormulas(formulas)
    .setNumberFormat("#,##0")
    .setFontColor("black")
    .setFontWeight("normal");
}


/**
 * Builds the aggregate formulas for Daily Archive rows 550-574 from the Tracklist sheet.
 * Each category sums exactly the rows the Tracklist sheet assigns to it; songs with a blank
 * category (the Track by Track rows) are in none, but still in the row-574 artist total.
 * @return {Array<Array<string>>} One [formula] per aggregate row, top to bottom.
 */
function buildAlbumFormulas() {
  const rowsByCategory = getRowsByCategory();
  const firstRow = CONFIG.ALBUMS.START_ROW;
  const lastSongRow = CONFIG.SONGS.START_ROW + CONFIG.SONGS.COUNT - 1;

  const configured = CONFIG.CATEGORIES.map(c => c.name);
  Object.keys(rowsByCategory).forEach(name => {
    if (configured.indexOf(name) === -1) {
      throw new Error(`The ${CONFIG.SHEETS.SONGS} sheet uses category "${name}", which is not in CONFIG.CATEGORIES.`);
    }
  });

  const formulas = new Array(CONFIG.ALBUMS.COUNT).fill(null);
  CONFIG.CATEGORIES.forEach(c => {
    formulas[c.archiveRow - firstRow] = [sumOfRows(rowsByCategory[c.name] || [])];
  });
  // The last aggregate row is the artist-wide total over the whole song range.
  formulas[CONFIG.ALBUMS.COUNT - 1] = [`=SUM(B${CONFIG.SONGS.START_ROW}:B${lastSongRow})`];

  const gap = formulas.indexOf(null);
  if (gap !== -1) {
    throw new Error(`No category in CONFIG.CATEGORIES has archiveRow ${firstRow + gap}.`);
  }
  return formulas;
}


/**
 * @param {Array<number>} rows - Ascending row numbers.
 * @return {string} e.g. "=SUM(B2:B16)", or "=SUM(B336:B354,B600)" for a scattered category.
 */
function sumOfRows(rows) {
  if (!rows.length) return '=0';

  const parts = [];
  let start = rows[0];
  for (let i = 1; i <= rows.length; i++) {
    if (i < rows.length && rows[i] === rows[i - 1] + 1) continue;
    const end = rows[i - 1];
    parts.push(start === end ? `B${start}` : `B${start}:B${end}`);
    start = rows[i];
  }
  return `=SUM(${parts.join(',')})`;
}

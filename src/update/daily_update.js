function transferStats() {
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
  setAlbumFormulas(dailyArchiveSheet);

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

function setAlbumFormulas(dailyArchiveSheet) {
  // Define all formulas for the new column B
  var formulas = [
    ['=SUM(B2:B16)'],    // Row 550: Taylor Swift
    ['=SUM(B17:B35)'],   // Row 551: Fearless
    ['=SUM(B36:B56)'],   // Row 552: Speak Now
    ['=SUM(B57:B78)'],   // Row 553: Red
    ['=SUM(B79:B97)'],   // Row 554: 1989
    ['=SUM(B98:B112)'],  // Row 555: reputation
    ['=SUM(B113:B130)'], // Row 556: Lover
    ['=SUM(B131:B164)'], // Row 557: folklore
    ['=SUM(B165:B181)'], // Row 558: evermore
    ['=SUM(B238:B260)'], // Row 559: Midnights
    ['=SUM(B305:B335)'], // Row 560: The Tortured Poets Department
    ['=SUM(B182:B207)'], // Row 561: Fearless (Taylor's Version)
    ['=SUM(B208:B237)'], // Row 562: Red (Taylor's Version)
    ['=SUM(B261:B282)'], // Row 563: Speak Now (Taylor's Version)
    ['=SUM(B283:B304)'], // Row 564: 1989 (Taylor's Version)
    ['=SUM(B336:B354)'], // Row 565: The Life of a Showgirl
    ['=SUM(B429:B435)'], // Row 566: Droplets
    ['=SUM(B363:B368)'], // Row 567: The Taylor Swift Holiday Collection
    ['=SUM(B355:B362)'], // Row 568: Live Clear Channel Stripped 2008
    ['=SUM(B377:B392)'], // Row 569: Speak Now World Tour
    ['=SUM(B369:B376)'], // Row 570: Live From Paris
    ['=SUM(B405:B416)'], // Row 571: Soundtracks
    ['=SUM(B436:B518)'], // Row 572: Remixes and etc.
    ['=SUM(B417:B428)'], // Row 573: Features
    ['=SUM(B2:B549)']    // Row 574: Total Artist Streams
  ];

  // Write all formulas at once starting at CONFIG.ALBUMS.START_ROW
  dailyArchiveSheet.getRange(CONFIG.ALBUMS.START_ROW, 2, formulas.length, 1)
    .setFormulas(formulas)
    .setNumberFormat("#,##0")
    .setFontColor("black")
    .setFontWeight("normal");
}

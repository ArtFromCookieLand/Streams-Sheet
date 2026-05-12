/**
 * [HELPER] Checks all songs for new milestones and logs them.
 * 1. Calculates yesterday's total (Today's Total - Today's Daily).
 * 2. Finds all milestones passed between yesterday and today.
 * 3. Appends any new milestones to the "Milestone Log" sheet.
 */
function checkMilestones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var latestSheet = ss.getSheetByName(CONFIG.SHEETS.LATEST);
  var logSheet = ss.getSheetByName(CONFIG.SHEETS.MILESTONE_LOG);
  
  // --- 1. Create the Milestone Log sheet if it doesn't exist ---
  if (!logSheet) {
    logSheet = ss.insertSheet(CONFIG.SHEETS.MILESTONE_LOG);
    logSheet.getRange('A1:C1').setValues([['Date', 'Song Name', 'Milestone']])
      .setFontWeight('bold');
    logSheet.setColumnWidths(1, 2, 120); // Date, Song
    logSheet.setColumnWidth(3, 150);     // Milestone
  }

  // --- 2. Get all required data from 'Latest' ---
  var newDate = latestSheet.getRange(CONFIG.LATEST.DATE_CELL).getValue();
  var songNames = latestSheet.getRange(CONFIG.LATEST.SONG_NAMES).getValues();
  // F = Total Streams, G = Daily Streams
  var streamData = latestSheet.getRange(
    CONFIG.SONGS.START_ROW, 
    6, // Column F
    CONFIG.SONGS.COUNT, 
    2  // 2 columns (F and G)
  ).getValues();

  // --- 3. Define Milestones ---
  const staticMilestones = [1000000, 5000000, 10000000, 25000000];
  const dynamicMilestone = 50000000; // Every 50M
  
  var allMilestonesHit = []; // To store [Date, Name, Milestone]

  // --- 4. Loop through every song ---
  for (var i = 0; i < CONFIG.SONGS.COUNT; i++) {
    var songName = songNames[i][0];
    var todayTotal = Number(streamData[i][0]);
    var todayDaily = Number(streamData[i][1]);
    
    // Skip if data is bad or empty
    if (isNaN(todayTotal) || isNaN(todayDaily) || !songName) {
      continue;
    }

    var yesterdayTotal = todayTotal - todayDaily;
    var milestonesForThisSong = [];
    
    // a) Check static milestones
    for (var j = 0; j < staticMilestones.length; j++) {
      var m = staticMilestones[j];
      if (yesterdayTotal < m && todayTotal >= m) {
        milestonesForThisSong.push(m);
      }
    }
    
    // b) Check dynamic 50M milestones
    // Find the first 50M mark *above* yesterday
    var next50M = Math.ceil(yesterdayTotal / dynamicMilestone) * dynamicMilestone;
    if (next50M === 0) next50M = dynamicMilestone; // Fix for 0
    
    // Loop from that first mark until we pass today's total
    while (next50M <= todayTotal) {
      if (next50M > yesterdayTotal) { // Ensure it wasn't hit yesterday
        milestonesForThisSong.push(next50M);
      }
      next50M += dynamicMilestone;
    }

    // c) Add all found milestones to the main log
    for (var k = 0; k < milestonesForThisSong.length; k++) {
      allMilestonesHit.push([newDate, songName, milestonesForThisSong[k]]);
    }
  }

  // --- 5. Write all new milestones to the log sheet ---
  if (allMilestonesHit.length > 0) {
    var startRow = logSheet.getLastRow() + 1; // Start on the next blank row
    var numRows = allMilestonesHit.length;

    // Get the full range to write all data at once
    var dataRange = logSheet.getRange(
      startRow,
      1,       // Column A
      numRows,
      3        // 3 columns
    );
    
    // 1. Set all the values
    dataRange.setValues(allMilestonesHit);
    
    // 2. Format the Date column (Col A)
    logSheet.getRange(startRow, 1, numRows, 1).setNumberFormat('yyyy/mm/dd');
    
    // 3. Format the Milestone column (Col C)
    logSheet.getRange(startRow, 3, numRows, 1).setNumberFormat('#,##0');
  }
}


/**
 * [HELPER] Finds songs likely to hit a milestone "tomorrow".
 * 1. Reads data from 'Latest'.
 * 2. Clears the old watch list from the 'Tracks' sheet.
 * 3. Predicts tomorrow's total and finds upcoming milestones.
 * 4. Generates a text string for each.
 * 5. Writes the new watch list to 'Tracks!N45:Q...'.
 */
function updateUpcomingMilestones() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var latestSheet = ss.getSheetByName(CONFIG.SHEETS.LATEST);
  var tracksSheet = ss.getSheetByName(CONFIG.SHEETS.TRACKS);
  
  if (!tracksSheet) {
    Logger.log("Error: 'Tracks' sheet not found. Skipping upcoming milestones.");
    return;
  }

  // --- 1. Clear the old watch list from "Tracks" sheet ---
  tracksSheet.getRange(CONFIG.TRACKS.UPCOMING_MILESTONES_CLEAR).clearContent();
  
  // --- 2. Get all required data from 'Latest' ---
  var songNames = latestSheet.getRange(CONFIG.LATEST.SONG_NAMES).getValues();
  // F = Total Streams, G = Daily Streams
  var streamData = latestSheet.getRange(
    CONFIG.SONGS.START_ROW, 6, CONFIG.SONGS.COUNT, 2
  ).getValues();
  
  // Create a 1D array of just the total streams for rank calculation
  const allTotals = streamData.map(row => Number(row[0]) || 0);

  // --- 3. Define Milestones ---
  const staticMilestones = [1000000, 5000000, 10000000, 25000000];
  const dynamicMilestone = 50000000;

  var upcomingHits = []; // To store [Name, '', Milestone, Text]

  // --- 4. Loop through every song ---
  for (var i = 0; i < CONFIG.SONGS.COUNT; i++) {
    var songName = songNames[i][0];
    var todayTotal = Number(streamData[i][0]);
    var todayDaily = Number(streamData[i][1]);

    if (isNaN(todayTotal) || isNaN(todayDaily) || todayDaily <= 0 || !songName) {
      continue;
    }

    var predictedTotal = todayTotal + todayDaily;
    var milestonesForThisSong = [];

    // a) Check static milestones
    for (var j = 0; j < staticMilestones.length; j++) {
      var m = staticMilestones[j];
      if (todayTotal < m && predictedTotal >= m) {
        milestonesForThisSong.push(m);
      }
    }

    // b) Check dynamic 50M milestones
    var nextM = Math.ceil(todayTotal / dynamicMilestone) * dynamicMilestone;
    if (nextM <= todayTotal) nextM += dynamicMilestone;
    
    while (nextM <= predictedTotal) {
      milestonesForThisSong.push(nextM);
      nextM += dynamicMilestone;
    }

    // c) Add all found milestones to the main list
    for (var k = 0; k < milestonesForThisSong.length; k++) {
      var m = milestonesForThisSong[k];
      
      // --- NEW: Generate Text ---
      var rank = getMilestoneRank(m, allTotals);
      var ordinalRank = toOrdinal(rank);
      var streamsText = m.toLocaleString('en-US'); // Formats 50000000 -> "50,000,000"
      
      var text = `${songName} has surpassed ${streamsText} streams on Spotify. It is Taylor Swift's ${ordinalRank} song to reach the milestone.`;
      
      // Add all 4 pieces of data
      upcomingHits.push([songName, '', m, text]);
    }
  }

  // --- 5. Write all new milestones to the "Tracks" sheet ---
  
  // Clear any existing merges in the header row
  tracksSheet.getRange(CONFIG.TRACKS.UPCOMING_MILESTONES_DEST)
    .offset(0, 0, 1, 4).breakApart(); // Un-merge N45:Q45

  // Set Headers for N, O, P, Q
  tracksSheet.getRange(CONFIG.TRACKS.UPCOMING_MILESTONES_DEST)
    .offset(0, 0, 1, 4) // Get N45:Q45
    .setValues([['Upcoming Milestone', '', 'Milestone', 'Generated Text']])
    .setFontWeight('bold');
  
  // Merge N45:O45 for the main title
  tracksSheet.getRange(CONFIG.TRACKS.UPCOMING_MILESTONES_DEST)
    .offset(0, 0, 1, 2) // Get N45:O45
    .merge();
    
  // Write data (if any)
  if (upcomingHits.length > 0) {
    var dataRange = tracksSheet.getRange(
      CONFIG.TRACKS.UPCOMING_MILESTONES_DEST // 'N45'
    )
    // *** CHANGED: Start at row 46, write 4 columns wide (N, O, P, Q) ***
    .offset(1, 0, upcomingHits.length, 4); 

    dataRange.setValues(upcomingHits);
    
    // Format the Milestone column (Col P), which is the 3rd column
    dataRange.offset(0, 2, upcomingHits.length, 1).setNumberFormat('#,##0');
    
    // Format the Text column (Col Q) to wrap
    dataRange.offset(0, 3, upcomingHits.length, 1).setWrap(true);
  }
}
var ss = SpreadsheetApp.getActiveSpreadsheet();
var toolsSheet = ss.getSheetByName(CONFIG.SHEETS.TOOLS);;
var dailyArchiveSheet = ss.getSheetByName(CONFIG.SHEETS.ARCHIVE);
var latestSheet = ss.getSheetByName(CONFIG.SHEETS.LATEST);

/**
 * Creates a Custom Menu when the spreadsheet opens.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  // Labelled so you can never mistake the dev copy for the live sheet.
  const updateMenu = ui.createMenu(isDev() ? 'Update [DEV]' : 'Update')
      .addItem('⬇️ Import Data', 'importSpotifyData')
      .addItem('🔃 Update Daily Stats', 'main')
      .addItem('🔄 Switch Token', 'switchApifyToken')
      .addItem('🔗 Match Totals by ID', 'matchTotalsMenu')
      .addItem('➕ Add Pending Songs', 'addPendingSongsMenu')
      .addSeparator()
      .addItem('Update Auxiliary Stats', 'updateStats')
      .addItem('Update Best-Since-Days', 'transferBestSinceRows')
      .addItem("Check Today's Milestones", 'checkMilestones')
      .addItem("Check Upcoming Milestones", 'updateUpcomingMilestones')
      .addItem("Update Text Summaries", 'generateSummaries');

  // One-off setup (src/setup/setup_sheets.js), shown only while a sheet is missing.
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const setup = ui.createMenu('Setup');
  let setupNeeded = false;
  if (!ss.getSheetByName(CONFIG.SHEETS.CATEGORIES)) { setup.addItem('Create Categories sheet', 'createCategoriesSheet'); setupNeeded = true; }
  if (!ss.getSheetByName(CONFIG.SHEETS.PENDING)) { setup.addItem('Create Pending sheet', 'createPendingSheet'); setupNeeded = true; }
  if (setupNeeded) updateMenu.addSeparator().addSubMenu(setup);
  updateMenu.addToUi();

  // Read-only health checks (src/checks/checks.js)
  ui.createMenu(isDev() ? 'Checks [DEV]' : 'Checks')
      .addItem('🩺 Run All Checks', 'runAllChecksMenu')
      .addItem('↔️ Check Row Alignment', 'checkRowAlignmentMenu')
      .addToUi()
}

function main() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const checkSheet = ss.getSheetByName(CONFIG.SHEETS.TOOLS); 
  
  // --- 0a. TRACKLIST AND CATEGORIES CHECK ---
  // Everything below reads them, and transferStats() is destructive, so a broken Tracklist or
  // Categories sheet has to be caught first.
  try {
    buildAlbumFormulas();
  } catch (error) {
    ui.alert('Update Aborted', `The ${CONFIG.SHEETS.SONGS} or ${CONFIG.SHEETS.CATEGORIES} sheet has a problem:\n\n` + error.message, ui.ButtonSet.OK);
    return;
  }

  // --- 0b. MISSING TRACK IDS ---
  // matchTotalsById() marks songs whose ID was not in the import. They would otherwise surface
  // only as an unexplained error in C1.
  const totals = checkSheet.getRange(CONFIG.LAYOUT.FIRST_SONG_ROW, CONFIG.TOOLS.TOTALS_COLUMN, getSongRowCount(), 1).getValues();
  const missingRows = [];
  totals.forEach((r, i) => {
    if (r[0] === CONFIG.TOOLS.MISSING_MARKER) missingRows.push(CONFIG.LAYOUT.FIRST_SONG_ROW + i);
  });
  if (missingRows.length) {
    ui.alert(
      'Update Aborted',
      `${missingRows.length} song(s) have no total because their track ID was not in the import (rows ${missingRows.join(', ')}).\n\n` +
      `Correct their track IDs in the ${CONFIG.SHEETS.SONGS} sheet, then run "Match Totals by ID" and try again.`,
      ui.ButtonSet.OK
    );
    return;
  }

  // --- 0c. AUTOMATED SPOTIFY UPDATE CHECK ---
  const sumValue = checkSheet.getRange(CONFIG.TOOLS.SUM_OF_DAILYS).getValue();
  
  if (sumValue <= 0 || isNaN(sumValue)) {
    ui.alert(
      'Update Aborted',
      `The sum of daily streams in cell C1 is ${sumValue}.\n\nThis indicates that Spotify has not updated yet today, or there was an error importing the data. Please try again later.`,
      ui.ButtonSet.OK
    );
    return;
  }

  // --- 1. SAFETY CONFIRMATION ---
  // This pop-up prevents accidental clicks.
  const response = ui.alert(
    'Confirm Update' + envTag(),
    'Spotify data has been updated (C1 is positive). Are you sure you want to proceed with updating the sheets and generating summaries?',
    ui.ButtonSet.YES_NO
  );

  // If the user did NOT click "Yes", stop everything.
  if (response !== ui.Button.YES) {
    ss.toast('Update Cancelled.', 'Status');
    return;
  }

  // --- 2. EXECUTE UPDATES ---
  // Transfering stats
  ss.toast('Transferring stats...', 'Status');
  transferStats();

  // Updating data needed for further calculations in spreadsheets
  ss.toast('Updating data...', 'Status');
  const dailyArchiveSheet = ss.getSheetByName(CONFIG.SHEETS.ARCHIVE); // Ensure this is defined
  const latestSheet = ss.getSheetByName(CONFIG.SHEETS.LATEST);        // Ensure this is defined
  updateStats(dailyArchiveSheet, latestSheet);

  // Logging milestones
  ss.toast('Checking milestones...', 'Status');
  checkMilestones();
  
  ss.toast('Updating upcoming milestones...', 'Status');
  updateUpcomingMilestones();
  
  ss.toast('Generating summaries...', 'Status');
  generateSummaries();
  generateDiscographySummary();

  // --- 3. CHECKS ---
  // A dialog only if something is off; nothing to click through on a clean run.
  ss.toast('Running checks...', 'Status');
  runChecksAfterUpdate();

  // --- 4. COMPLETION ---
  ss.toast('All updates completed successfully!', 'Success', 5);
}

















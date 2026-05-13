var ss = SpreadsheetApp.getActiveSpreadsheet();
var toolsSheet = ss.getSheetByName(CONFIG.SHEETS.TOOLS);;
var dailyArchiveSheet = ss.getSheetByName(CONFIG.SHEETS.ARCHIVE);
var latestSheet = ss.getSheetByName(CONFIG.SHEETS.LATEST);

/**
 * Creates a Custom Menu when the spreadsheet opens.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Update')
      .addItem('⬇️ Import Data', 'importSpotifyData')
      .addItem('🔃 Update Daily Stats', 'main')
      .addItem('🔄 Switch Token', 'switchApifyToken')
      .addSeparator()
      .addItem('Update Auxiliary Stats', 'updateStats')
      .addItem('Update Best-Since-Days', 'transferBestSinceRows')
      .addItem("Check Today's Milestones", 'checkMilestones')
      .addItem("Check Upcoming Milestones", 'updateUpcomingMilestones')
      .addItem("Update Text Summaries", 'generateSummaries')
      .addToUi()
}

function main() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const checkSheet = ss.getSheetByName(CONFIG.SHEETS.TOOLS); 
  
  // --- 0. AUTOMATED SPOTIFY UPDATE CHECK ---
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
    'Confirm Update',
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

  // --- 3. COMPLETION ---
  ss.toast('All updates completed successfully!', 'Success', 5);
}

















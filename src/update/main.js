var ss = SpreadsheetApp.getActiveSpreadsheet();
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
      .addItem('🔗 Check Import', 'matchTotalsMenu')
      .addItem('🔍 Find New Tracks', 'findNewTracksMenu')
      .addItem('➕ Add Pending Songs', 'addPendingSongsMenu')
      .addItem('🖼️ Fill Missing Covers', 'fillMissingCoversMenu')
      .addItem('🗂️ Add New Categories', 'addCategoriesMenu')
      .addItem('↕️ Move a Song', 'moveSongMenu')
      .addItem('🧮 Rebuild Album History', 'rebuildAlbumHistoryMenu')
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
  if (!ss.getSheetByName(CONFIG.SHEETS.IGNORED)) { setup.addItem('Create Ignored sheet', 'createIgnoredSheet'); setupNeeded = true; }
  if (!ss.getSheetByName(CONFIG.SHEETS.SOURCES)) { setup.addItem('Create Sources sheet', 'createSourcesSheet'); setupNeeded = true; }
  if (!ss.getSheetByName(CONFIG.SHEETS.IMPORT)) { setup.addItem('Tidy up the Import sheet (2.1)', 'tidyImportSheet'); setupNeeded = true; }
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

  // --- 0a. TRACKLIST AND CATEGORIES CHECK ---
  // Everything below reads them, and transferStats() is destructive, so a broken Tracklist or
  // Categories sheet has to be caught first.
  try {
    buildAlbumFormulas();
  } catch (error) {
    ui.alert('Update Aborted', `The ${CONFIG.SHEETS.SONGS} or ${CONFIG.SHEETS.CATEGORIES} sheet has a problem:\n\n` + error.message, ui.ButtonSet.OK);
    return;
  }

  // --- 0b. TODAY'S FIGURES, FROM THE IMPORT ---
  // Each song's total by its track ID, and its daily against Latest!F. Worked out in full before
  // anything is written.
  let today;
  try {
    today = readTodayFromImport();
  } catch (error) {
    ui.alert('Update Aborted', error.message, ui.ButtonSet.OK);
    return;
  }

  // --- 0c. MISSING TRACK IDS ---
  // A song with no total would otherwise be archived as minus its whole total.
  if (today.missing.length) {
    ui.alert(
      'Update Aborted',
      `${today.missing.length} song(s) have no total because their track ID was not in the import:\n` +
      today.missing.map(m => `• row ${m.row}: ${m.title}`).join('\n') + '\n\n' +
      `Correct their track IDs in the ${CONFIG.SHEETS.SONGS} sheet, then run "Check Import" and try again.`,
      ui.ButtonSet.OK
    );
    return;
  }

  // --- 0d. HAS SPOTIFY REFRESHED? ---
  // The dailies add up to 0 or less when the import is no newer than Latest: Spotify hasn't
  // refreshed yet, or this import has already been used for an update. Running anyway would
  // archive a day of zeros, so this is also what stops a double run.
  const sum = today.sumOfDailies;
  writeSumOfDailies(sum);
  if (!(sum > 0)) {
    ui.alert(
      'Update Aborted',
      `The sum of today's daily streams is ${Math.round(sum).toLocaleString('en-US')}.\n\n` +
      'Spotify has not refreshed since the last update, or this import has already been used for one. Import again later.',
      ui.ButtonSet.OK
    );
    return;
  }

  // --- 1. SAFETY CONFIRMATION ---
  // This pop-up prevents accidental clicks.
  const response = ui.alert(
    'Confirm Update' + envTag(),
    `Today's dailies add up to ${Math.round(sum).toLocaleString('en-US')}. Are you sure you want to proceed with updating the sheets and generating summaries?`,
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
  transferStats(today);
  writeSumOfDailies(0);   // Latest now holds these totals, so nothing is left to add up

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

















/**
 * [BUTTON 1] IMPORTS DATA FROM SPOTIFY
 */
function importSpotifyData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 0. ENVIRONMENT GUARD ---
  // Apify credits are shared and limited (~17 runs per token), so the dev copy
  // does not spend them. The raw data that came across with the spreadsheet copy
  // in Tools!F2:G1000 is the fixture to develop against instead.
  if (isDev() && !CONFIG.ENV.ALLOW_APIFY_IN_DEV) {
    ui.alert(
      'Blocked in DEV',
      'This is a DEV copy, and an Apify run would spend credits from the shared token budget.\n\n' +
      'The raw data already sitting in Tools!F2:G1000 was copied from production — run "Update Daily Stats" against that instead.\n\n' +
      'To override, set CONFIG.ENV.ALLOW_APIFY_IN_DEV to true.',
      ui.ButtonSet.OK
    );
    return;
  }

  // --- 0b. SAFETY CONFIRMATION ---
  // This pop-up prevents accidental clicks.
  const response = ui.alert(
    'Confirm Import' + envTag(),
    'Are you sure you want to run the Spotify Scraper?',
    ui.ButtonSet.YES_NO
  );

  // If the user did NOT click "Yes", stop everything.
  if (response !== ui.Button.YES) {
    ss.toast('Import cancelled.', 'Status');
    return;
  }

  ss.toast('Starting Apify... Please wait.', 'Status');

  try {
    // --- 1. PREPARE INPUT ---
    
    // Extract just the URLs (Values) from the Config Dictionary
    const albumUrlList = Object.values(CONFIG.APIFY.ALBUMS);

    // Convert to Apify Object Format: { "url": "..." }
    const urlObjects = albumUrlList.map(link => {
      return { "url": link };
    });

    const actorInput = {
      "urls": urlObjects,
      "followAlbums": false,         
      "followSingles": false,        
      "followPopularReleases": false, 
      "proxy": { "useApifyProxy": true }
    };

    // --- 2. EXECUTE APIFY RUN ---

    const run = ApifyService.startActor(CONFIG.APIFY.ACTOR_ID, getActiveApifyToken(), actorInput);
    ApifyService.waitForCompletion(run.id, CONFIG.APIFY.ACTOR_ID, getActiveApifyToken());
    
    const rawData = ApifyService.fetchDataset(run.defaultDatasetId, getActiveApifyToken());
    
    if (!rawData || rawData.length === 0) {
      throw new Error("Apify returned no data.");
    }

    // --- 3. PARSE RESULTS ---
    
    let flatTrackList = [];

    rawData.forEach(album => {
      if (album.tracks && Array.isArray(album.tracks)) {
        album.tracks.forEach(track => {
          flatTrackList.push([track.name, Number(track.streamCount)]);
        });
      }
    });

    // --- 4. WRITE TO SHEET ---
    
    const toolsSheet = ss.getSheetByName(CONFIG.SHEETS.TOOLS);
    const range = toolsSheet.getRange(CONFIG.TOOLS.RAW_DATA);
    
    range.clearContent();
    
    if (flatTrackList.length > 0) {
      // Check for overflow
      if (flatTrackList.length > range.getNumRows()) {
        ui.alert("Warning: More songs returned than fit in your range!");
      }
      
      // Slice and Write
      const safeData = flatTrackList.slice(0, range.getNumRows());
      toolsSheet.getRange(range.getRow(), range.getColumn(), safeData.length, 2)
        .setValues(safeData);
    }
    
    // Success Message
    ui.alert('Success', `Imported ${flatTrackList.length} tracks successfully.`, ui.ButtonSet.OK);
    logRunAndCheckLimits();

  } catch (error) {
    console.error(error);
    const errorString = error.toString();
    
    // Аналізуємо текст помилки на наявність специфічного індикатора вичерпання кредитів
    if (errorString.includes("max-items-must-be-greater-than-zero")) {
      ui.alert(
        'Credits Exhausted', 
        'The current Apify token has run out of credits.\nPlease use the menu to switch the active token and try again.', 
        ui.ButtonSet.OK
      );
    } else {
      // Для всіх інших помилок виводимо стандартне повідомлення
      ui.alert('Error', errorString, ui.ButtonSet.OK);
    }
  }
}
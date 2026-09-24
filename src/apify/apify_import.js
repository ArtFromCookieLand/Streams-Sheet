/**
 * [BUTTON 1] IMPORTS DATA FROM SPOTIFY
 */
function importSpotifyData() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // --- 0. ENVIRONMENT GUARD ---
  // Apify credits are shared and limited (~17 runs per token), so the dev copy
  // does not spend them. The raw data that came across with the spreadsheet copy
  // in Import!E2:I1000 is the fixture to develop against instead.
  if (isDev() && !CONFIG.ENV.ALLOW_APIFY_IN_DEV) {
    ui.alert(
      'Blocked in DEV',
      'This is a DEV copy, and an Apify run would spend credits from the shared token budget.\n\n' +
      `The raw data already sitting in ${CONFIG.SHEETS.IMPORT}!${CONFIG.IMPORT.RAW_DATA} was copied from production — use "Check Import" and "Update Daily Stats" against that instead.\n\n` +
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
    
    // The albums to scrape: the Sources sheet, or CONFIG.APIFY.ALBUMS until it exists.
    const sources = readSources();
    if (sources.problems.length) throw new Error(`${CONFIG.SHEETS.SOURCES} sheet:\n• ` + sources.problems.join('\n• '));
    const albumUrlList = sources.urls;

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

    // [album, name, stream count, track ID, album cover] - the ID is what readTodayFromImport()
    // matches on; the album and name are for a human reading the sheet; the cover seeds new Covers rows.
    rawData.forEach(album => {
      if (album.tracks && Array.isArray(album.tracks)) {
        const cover = pickCoverArt(album.coverArt);
        album.tracks.forEach(track => {
          flatTrackList.push([album.name || '', track.name, Number(track.streamCount), track.id || '', cover]);
        });
      }
    });

    // --- 4. WRITE TO SHEET ---
    
    const importSheet = getImportSheet();
    const range = importSheet.getRange(CONFIG.IMPORT.RAW_DATA);
    
    range.clearContent();
    
    if (flatTrackList.length > 0) {
      // Check for overflow
      if (flatTrackList.length > range.getNumRows()) {
        ui.alert("Warning: More songs returned than fit in your range!");
      }
      
      // Slice and Write
      const safeData = flatTrackList.slice(0, range.getNumRows());
      importSheet.getRange(range.getRow(), range.getColumn(), safeData.length, safeData[0].length)
        .setValues(safeData);
    }

    // The run has already spent its credits, so count it before anything else can fail.
    logRunAndCheckLimits();

    // --- 5. MATCH BY TRACK ID: what's missing, and today's sum of dailies ---
    const matchResult = matchTotalsById();

    // --- 6. LOOK FOR NEW TRACKS ---
    // Links announced songs that are now out, and lists anything else new in Pending.
    let newTracksText;
    try {
      const found = findNewTracks();
      newTracksText = describeNewTracks(found);
      // A linked upcoming song has just got its ID, so it now counts towards the sum.
      if (found.linked.length) matchTotalsById();
    } catch (error) {
      newTracksText = 'Looking for new tracks was skipped: ' + error.message;
    }

    // Success Message
    ui.alert(
      'Import finished',
      `Imported ${flatTrackList.length} tracks from ${albumUrlList.length} albums` +
      (sources.from === 'config' ? ' (album list still from the code - run Update → Setup → Create Sources sheet).' : '.') +
      '\n\n' + describeMatchResult(matchResult) + '\n\n' + newTracksText,
      ui.ButtonSet.OK
    );

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

/**
 * @param {Array<{url, width, height}>} coverArt - As the actor returns it (usually 64, 300 and 640px).
 * @return {string} The URL at CONFIG.COVERS.SIZE if there is one, else the first, else ''.
 */
function pickCoverArt(coverArt) {
  if (!Array.isArray(coverArt) || !coverArt.length) return '';
  const sized = coverArt.filter(c => Number(c.width) === CONFIG.COVERS.SIZE)[0];
  return (sized || coverArt[0]).url || '';
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Taylor Swift Tracker')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getPublicData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.LATEST); 
  
  // --- 1. FETCH TRACK DATA (Unchanged) ---
  const startRow = 2;
  const numRows = 483; 
  
  const ranks = sheet.getRange(startRow, 2, numRows, 1).getValues().flat();       // B
  const rankChanges = sheet.getRange(startRow, 3, numRows, 1).getValues().flat(); // C
  const trackCovers = sheet.getRange(startRow, 16, numRows, 1).getValues().flat();// O (Track Covers)
  const names = sheet.getRange(startRow, 5, numRows, 1).getValues().flat();       // E
  const totalStreams = sheet.getRange(startRow, 6, numRows, 1).getValues().flat();// F
  const dailyStreams = sheet.getRange(startRow, 7, numRows, 1).getValues().flat();// G
  const dailyChanges = sheet.getRange(startRow, 8, numRows, 1).getValues().flat();// H
  const bestSinces = sheet.getRange(startRow, 12, numRows, 1).getValues().flat(); // L
  
  let tracks = [];
  for (let i = 0; i < numRows; i++) {
    if (names[i] && names[i] !== "") { 
      tracks.push({
        rank: Number(ranks[i]) || 0,
        rankChange: Number(rankChanges[i]) || 0,
        albumCover: String(trackCovers[i] || ""), 
        name: String(names[i]),
        totalStream: Number(totalStreams[i]) || 0,
        dailyStream: Number(dailyStreams[i]) || 0,
        dailyChange: Number(dailyChanges[i]) || 0,
        bestSince: formatDate(bestSinces[i], ss)
      });
    }
  }

  // --- 2. FETCH ALBUM DATA (Updated Mapping) ---
  // New Mapping:
  // T (20) - Titles
  // U (21) - Total
  // V (22) - Daily
  // W (23) - Daily %
  // Y (25) - Best Since
  // AB (28) - Covers
  
  // Range: T2 (20) to AB26 (28). Width = 28 - 20 + 1 = 9 columns.
  const albumRange = sheet.getRange(2, 20, 25, 9).getValues(); 
  
  let studio = [];
  let droplets = [];
  let total = null;

  for (let j = 0; j < albumRange.length; j++) {
    const row = albumRange[j];
    // Indices relative to T (0):
    // 0: T (Name)
    // 1: U (Total)
    // 2: V (Daily)
    // 3: W (%Daily)
    // 5: Y (Best Since) - skipping X(4)
    // 8: AB (Cover) - skipping Z(6), AA(7)
    
    const albumObj = {
      name: String(row[0]),
      total: Number(row[1]) || 0,
      daily: Number(row[2]) || 0,
      dailyPct: Number(row[3]) || 0,
      bestSince: formatDate(row[5], ss),
      cover: String(row[8] || "")
    };

    if (j < 16) {
      studio.push(albumObj);
    } else if (j < 24) {
      droplets.push(albumObj);
    } else {
      total = albumObj;
    }
  }
  
  return { 
    tracks: tracks, 
    albums: { studio: studio, droplets: droplets, total: total } 
  };
}

function formatDate(val, ss) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, ss.getSpreadsheetTimeZone(), "MMM d, yyyy");
  }
  return val ? String(val) : "-";
}
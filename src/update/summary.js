/** STRUCTURE  OF A SUMMARY
 * 1. Header
 *    Example:  
 *              "[album name] on Spotify yesterday ([date])"
 *              "evermore on Spotify yesterday (May 10th)"
 *    To be followed by an empty line afterwards.
 * 2. The biggest gainer
 *    Depends on whether there's in general any gainer among songs.
 *    Positive example: 
 *              "[song] was the biggest gainer, up [percent value]."
 *              "Enchanted was the biggest gainer, up 7.1%."
 *    Negative example: 
 *              "[song] was the most stabe song, down [percent value]."
 *              "Tim McGraw was the most stabe song, down 7.4%."
 * 3. The song with the best day value
 *    Generated only if there's a song with the best-since value being at least 7 days.
 *    Example:  
 *              "[song] scored its best day since [date]."
 *              "august scored its best day since June 18th"
 *    If multiple songs have same best-since value:
 *              "[list of songs separated by commas] scored their best update since [date]."
 *              "august, cardigan, hoax scored their best update since June 18th"
 * 3.1. 
 *    If the song for the biggest gainer and best-since value matches, the following should be displayed: 
 *              "[song] scored it's best day since [date] (up [percent value] from yesterday)."
 *              "Cruel Summer scored it's best day since March 18th (up 13% from yesterday)."
 *    If multiple songs have same best-since day, this message should not appear and the one form point 3 should be generated.
 * 4. Closing line
 *    The final line, preceeded by an empty line, should show the difference in streams compared to yesterday and the last week.
 *    Example:
 *              "[daily change in streams divided by 1000] from yesterday, [weekly change in streams divided by 1000] from the last week"
 *              "-121k from yesterday, +400k from the last week"
 *    Note: if change >1m streams, should be divided by 1000000 instead.
 * 4.1
 *    Similarly every album also has a best-since value. By the same idea as in songs, if it's bigger than 1 week ago, it is to mention.
 *    It can be in the next line after the point 4. Exmaple:
 *              "Best day since [date]"
 *              "Best day since August 7th"
 */

/**
 * GENERATES SUMMARIES FOR ALL ALBUMS
 */
function generateSummaries() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const latestSheet = ss.getSheetByName(CONFIG.SHEETS.LATEST);
  const albumsSheet = ss.getSheetByName(CONFIG.SHEETS.ALBUMS);

  // 1. Fetch Update Date
  const rawDate = latestSheet.getRange(CONFIG.LATEST.DATE_CELL).getValue();
  const formattedDate = formatDateString(rawDate);
  
  // Calculate the "7 days ago" threshold
  const sevenDaysAgo = new Date(rawDate);
  sevenDaysAgo.setDate(sevenDaysAgo.getDate() - 7);

  for (const [albumName, stats] of Object.entries(CONFIG.STATS)) {
    let summaryLines = [];

    // --- 1. HEADER ---
    summaryLines.push(`${albumName} on Spotify yesterday (${formattedDate})`);
    summaryLines.push("");

    // --- 2. EXTRACT SONG DATA ---
    const songDataRange = latestSheet.getRange(stats.songRow, 5, stats.count, 8);
    const songData = songDataRange.getValues();
    const songDisplayData = songDataRange.getDisplayValues(); 

    let maxPercent = -Infinity;
    let biggestGainer = null;
    let bestSinceSongs = [];

    // --- 3. ANALYZE SONG DATA ---
    for (let i = 0; i < songData.length; i++) {
      const title = songData[i][0];
      if (!title) continue; 

      const rawPercent = songData[i][3]; 
      const displayPercent = songDisplayData[i][3]; 
      const bestSinceDate = songData[i][7];

      // Logic for Biggest Gainer / Most Stable (smallest drop)
      if (typeof rawPercent === 'number' && rawPercent > maxPercent) {
        maxPercent = rawPercent;
        biggestGainer = {
          title: title,
          percentValue: rawPercent,
          displayPercent: displayPercent,
          isPositive: rawPercent > 0
        };
      }

      // Logic for Best Since (must be a date and at least 7 days ago)
      if (bestSinceDate instanceof Date && bestSinceDate <= sevenDaysAgo) {
        bestSinceSongs.push({
          title: title,
          dateStr: formatDateString(bestSinceDate)
        });
      }
    }

    // --- 4. BUILD POINTS 2, 3, & 3.1 ---
    let usePoint3_1 = false;

    // Check Point 3.1
    if (bestSinceSongs.length === 1 && biggestGainer && bestSinceSongs[0].title === biggestGainer.title && biggestGainer.isPositive) {
      usePoint3_1 = true;
      const cleanPercent = biggestGainer.displayPercent.replace('+', '');
      summaryLines.push(`• ${biggestGainer.title} was the biggest gainer and scored its best day since ${bestSinceSongs[0].dateStr} (up ${cleanPercent} from yesterday).`);
    }

    if (!usePoint3_1) {
      // Point 2: Standard Biggest Gainer / Most Stable
      if (biggestGainer) {
        if (biggestGainer.isPositive) {
          const cleanPercent = biggestGainer.displayPercent.replace('+', '');
          summaryLines.push(`• ${biggestGainer.title} was the biggest gainer, up ${cleanPercent}.`);
        } else {
          const cleanPercent = biggestGainer.displayPercent.replace('-', '');
          summaryLines.push(`• ${biggestGainer.title} was the most stable song, down ${cleanPercent}.`);
        }
      }

      // Point 3: Best Since lists
      if (bestSinceSongs.length === 1) {
        summaryLines.push(`• ${bestSinceSongs[0].title} scored its best day since ${bestSinceSongs[0].dateStr}.`);
      } else if (bestSinceSongs.length > 1) {
        const titles = bestSinceSongs.map(s => s.title).join(", ");
        summaryLines.push(`• ${titles} scored their best update since ${bestSinceSongs[0].dateStr}.`);
      }
    }

    // --- 5. EXTRACT ALBUM DATA ---
    const albumBestSinceRaw = latestSheet.getRange(stats.summaryRow, 25).getValue();
    const albumDailyRaw = latestSheet.getRange(stats.summaryRow, 26).getValue();
    const albumWeeklyRaw = latestSheet.getRange(stats.summaryRow, 27).getValue();

    // --- 6. BUILD POINT 4 (CLOSING LINE) ---
    summaryLines.push("");
    summaryLines.push(`${formatStreams(albumDailyRaw)} from yesterday, ${formatStreams(albumWeeklyRaw)} from the last week`);

    // --- 7. BUILD POINT 4.1 ---
    if (albumBestSinceRaw instanceof Date && albumBestSinceRaw <= sevenDaysAgo) {
      summaryLines.push(`Best day since ${formatDateString(albumBestSinceRaw)}`);
    }

    // --- 8. WRITE TO ALBUMS SHEET ---
    const destRange = albumsSheet.getRange(stats.destCell);
    const startRow = destRange.getRow();
    const startCol = destRange.getColumn();

    const outputData = summaryLines.map(line => [line]);
    
    albumsSheet.getRange(startRow, startCol, 10, 1).clearContent(); 
    albumsSheet.getRange(startRow, startCol, outputData.length, 1).setValues(outputData);
  }
}

// --- HELPER FUNCTIONS ---

/**
 * Formats a Date object into "Month Day+Ordinal" (e.g., "May 10th").
 * Appends the year if the date is not in the current year (e.g., "May 11th, 2025").
 */
function formatDateString(dateObj) {
  if (!(dateObj instanceof Date)) return "";
  const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  
  const month = months[dateObj.getMonth()];
  const day = dateObj.getDate();
  const dateYear = dateObj.getFullYear();
  const currentYear = new Date().getFullYear();
  
  // Використовуємо вашу існуючу функцію toOrdinal
  let formattedDate = `${month} ${toOrdinal(day)}`;
  
  // Додаємо рік, якщо він відрізняється від поточного
  if (dateYear !== currentYear) {
    formattedDate += `, ${dateYear}`;
  }
  
  return formattedDate;
}

/**
 * Divides stream numbers by 1k or 1m and attaches proper sign and letter.
 */
function formatStreams(num) {
  // Ensure it's parsed as a number just in case the sheet returns a string
  if (typeof num !== 'number') {
    num = parseFloat(num.toString().replace(/,/g, ''));
  }
  if (isNaN(num)) return "0";

  const sign = num > 0 ? "+" : (num < 0 ? "-" : "");
  const absNum = Math.abs(num);
  let formatted;

  if (absNum >= 1000000) {
    formatted = parseFloat((absNum / 1000000).toFixed(2)) + "m";
  } else {
    // Math.floor can also be used, but parseFloat + toFixed(1) handles .0 drop nicely
    formatted = parseFloat((absNum / 1000).toFixed(1)) + "k";
  }

  return sign + formatted;
}

/**
 * ===================================================================
 * SOURCES: THE ALBUM URLS TO SCRAPE
 * -------------------------------------------------------------------
 * The importer scrapes the Spotify albums listed in the `Sources` sheet
 * (name + url), so a new release needs no code change - add its URL and
 * import. Until the sheet exists, CONFIG.APIFY.ALBUMS is used instead,
 * and that map is also what Update → Setup → Create Sources sheet seeds
 * it from.
 * ===================================================================
 */

/**
 * @return {{urls: Array<string>, names: Array<string>, from: string, problems: Array<string>}}
 *   from is 'sheet' or 'config'.
 */
function readSources() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.SOURCES);
  if (!sheet) {
    const names = Object.keys(CONFIG.APIFY.ALBUMS);
    return { urls: names.map(n => CONFIG.APIFY.ALBUMS[n]), names: names, from: 'config', problems: [] };
  }

  const values = sheet.getDataRange().getValues();
  let col;
  try {
    col = findHeaderColumns(values[0], CONFIG.SOURCES_SHEET.HEADERS, CONFIG.SHEETS.SOURCES);
  } catch (error) {
    return { urls: [], names: [], from: 'sheet', problems: [error.message] };
  }

  const urls = [], names = [], problems = [], seen = {};
  for (let i = 1; i < values.length; i++) {
    const name = String(values[i][col.name]).trim();
    const url = String(values[i][col.url]).trim();
    if (!name && !url) continue;
    if (!url) { problems.push(`${CONFIG.SHEETS.SOURCES} row ${i + 1} ("${name}") has no URL.`); continue; }
    if (!/^https:\/\/open\.spotify\.com\/album\/[A-Za-z0-9]+/.test(url)) {
      problems.push(`${CONFIG.SHEETS.SOURCES} row ${i + 1} ("${name || url}") is not a Spotify album link.`);
      continue;
    }
    const id = url.split('/album/')[1].split('?')[0];
    if (seen[id]) { problems.push(`${CONFIG.SHEETS.SOURCES}: "${name}" is the same album as "${seen[id]}".`); continue; }
    seen[id] = name || url;
    names.push(name);
    urls.push(url);
  }
  if (!urls.length && !problems.length) problems.push(`The ${CONFIG.SHEETS.SOURCES} sheet has no album URLs.`);
  return { urls: urls, names: names, from: 'sheet', problems: problems };
}

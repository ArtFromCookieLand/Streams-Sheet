/**
 * Finds the first date where value > threshold.
 * @param {number} threshold - The value to beat (e.g., today's streams).
 * @param {Array<number>} valuesRange - 1D array of historical values.
 * @param {Array<Date>} datesRange - 1D array of historical dates.
 * @return {Date|String} The matching date, or "" if not found.
 */
function BESTSINCE(threshold, valuesRange, datesRange) {
  var values = to1D(valuesRange);
  var dates  = to1D(datesRange);

  if (!values.length || !dates.length) return "";
  var len = Math.min(values.length, dates.length);
  var t = Number(threshold);
  if (isNaN(t) || t === 0) return ""; // Also ignore if threshold is 0

  for (var i = 0; i < len; i++) {
    var v = Number(values[i]);
    if (!isNaN(v) && v > t) return dates[i];
  }
  return "";
}


/**
 * Normalizes an input range to a 1D array.
 * @param {any} a - A cell value, 1D array, or 2D array.
 * @return {Array} A 1D array.
 */
function to1D(a) {
  if (a == null) return [];
  if (!Array.isArray(a)) return [a];
  if (!a.length) return [];
  if (Array.isArray(a[0])) {
    if (a.length === 1) return a[0]; // single row
    if (a[0].length === 1) return a.map(r => r[0]); // single column
  }
  return a;
}


/**
 * [UTILITY] Calculates the rank for a new milestone.
 * Counts how many songs have already passed this milestone.
 * @param {number} milestone - The milestone to check (e.g., 50000000).
 * @param {Array<number>} allTotals - A 1D array of all song totals.
 * @return {number} The rank (e.g., 5 if 4 other songs have passed it).
 */
function getMilestoneRank(milestone, allTotals) {
  let count = 0;
  for (let i = 0; i < allTotals.length; i++) {
    if (allTotals[i] >= milestone) {
      count++;
    }
  }
  return count + 1; // The upcoming one will be the next in rank
}


/**
 * [UTILITY] Converts a number to its ordinal string (1 -> 1st, 2 -> 2nd).
 * @param {number} n - The number to convert.
 * @return {string} The ordinal string.
 */
function toOrdinal(n) {
  var s = ['th', 'st', 'nd', 'rd'];
  var v = n % 100;
  return n + (s[(v - 20) % 10] || s[v] || s[0]);
}
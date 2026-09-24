/**
 * A small in-memory stand-in for SpreadsheetApp, enough to run the project's code in node.
 *
 * What it models, because the code depends on it:
 *   - values, formulas and display text per cell
 *   - simple formulas are evaluated on read: numbers, cell and range references (also to other
 *     sheets), SUM, + - * / and brackets. Anything else (XLOOKUP, IMAGE, RANK...) keeps the value
 *     it was loaded with.
 *   - structural edits re-point formula references the way Sheets does: inserting rows or cells
 *     pushes references below them down, deleting pulls them up, and cut-and-paste (moveTo,
 *     moveRows) makes references follow the moved cells. A range's two ends are moved
 *     independently, which is what Sheets does for inserts and deletes.
 *   - R1C1 formulas, so copying a neighbour's formulas works as it does in Sheets
 *   - merged cells, for the "merges in the way" checks
 *
 * What it doesn't: formatting (accepted and ignored), data validation, charts, triggers.
 */

const blank = () => ({ v: '', f: '', d: undefined });

// ---------- A1 helpers ----------

const colToNum = s => s.split('').reduce((a, ch) => a * 26 + ch.charCodeAt(0) - 64, 0);
const numToCol = n => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

// A reference, optionally with a sheet, optionally a range: Sheet!A1, 'Daily Archive 2026'!B2:B9, $A$1.
const REF = /(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_]*))(!))?(\$?)([A-Z]{1,3})(\$?)(\d+)(?::(\$?)([A-Z]{1,3})(\$?)(\d+))?/g;

/**
 * Calls fn for every cell reference in a formula and splices in what it returns.
 * fn gets { sheet (name or null), c1, r1, c2, r2, abs } and returns replacement text, or null to keep it.
 */
function rewriteRefs(formula, fn) {
  // Skip string literals: references inside quotes are text.
  const parts = formula.split(/("(?:[^"]|"")*")/);
  return parts.map((part, i) => {
    if (i % 2) return part;
    return part.replace(REF, (m, qName, name, bang, ac1, c1, ar1, r1, ac2, c2, ar2, r2, offset, whole) => {
      const sheet = bang ? (qName !== undefined ? qName.replace(/''/g, "'") : name) : null;
      // Part of a longer word, or a function name like LOG10(.
      const prev = whole[offset - 1], next = whole[offset + m.length];
      if ((prev && /[A-Za-z0-9_$.]/.test(prev)) || next === '(') return m;
      const ref = { sheet, c1: colToNum(c1), r1: +r1, c2: c2 ? colToNum(c2) : null, r2: r2 ? +r2 : null,
        abs: { c1: !!ac1, r1: !!ar1, c2: !!ac2, r2: !!ar2 } };
      const res = fn(ref);
      return res === null || res === undefined ? m : res;
    });
  }).join('');
}

function refText(ref, withSheet) {
  const cell = (c, r, ac, ar) => (ac ? '$' : '') + numToCol(c) + (ar ? '$' : '') + r;
  let s = '';
  if (withSheet && ref.sheet !== null) s += (/^[A-Za-z_][A-Za-z0-9_]*$/.test(ref.sheet) ? ref.sheet : "'" + ref.sheet.replace(/'/g, "''") + "'") + '!';
  s += cell(ref.c1, ref.r1, ref.abs.c1, ref.abs.r1);
  if (ref.c2 !== null) s += ':' + cell(ref.c2, ref.r2, ref.abs.c2, ref.abs.r2);
  return s;
}


// ---------- Spreadsheet ----------

class Spreadsheet {
  constructor(opts) {
    this.sheets = [];
    this.gen = 0;
    this.opts = Object.assign({ id: 'dev', timeZone: 'Etc/UTC' }, opts || {});
  }
  getId() { return this.opts.id; }
  getSpreadsheetTimeZone() { return this.opts.timeZone; }
  getSheetByName(name) { return this.sheets.find(s => s.name === name) || null; }
  getSheets() { return this.sheets.slice(); }
  insertSheet(name) {
    if (this.getSheetByName(name)) throw new Error(`A sheet with the name "${name}" already exists.`);
    const s = new Sheet(this, name); this.sheets.push(s); return s;
  }
  toast() {}
  touch() { this.gen++; }

  /**
   * Applies a position mapping to every formula in the workbook that points at `sheet`.
   * map(r, c, end) -> [r, c] (new position), or null when the cell no longer exists.
   * `end` is 'start' / 'end' for the two ends of a range, 'cell' for a single reference.
   */
  remapRefs(sheet, map) {
    this.sheets.forEach(host => {
      host.g.forEach(row => row && row.forEach(x => {
        if (!x || !x.f) return;
        x.f = rewriteRefs(x.f, ref => {
          const target = ref.sheet === null ? host : this.getSheetByName(ref.sheet);
          if (target !== sheet) return null;
          if (ref.c2 === null) {
            const p = map(ref.r1, ref.c1, 'cell');
            if (!p) return '#REF!';
            return refText(Object.assign({}, ref, { r1: p[0], c1: p[1] }), true);
          }
          if (map.wholeRangesOnly) {
            // Cut and paste moves a range reference only when the whole range was moved.
            const a = map(ref.r1, ref.c1, 'cell'), b = map(ref.r2, ref.c2, 'cell');
            const movedA = a && (a[0] !== ref.r1 || a[1] !== ref.c1), movedB = b && (b[0] !== ref.r2 || b[1] !== ref.c2);
            if (!(movedA && movedB)) return null;
            return refText(Object.assign({}, ref, { r1: a[0], c1: a[1], r2: b[0], c2: b[1] }), true);
          }
          const a = map(ref.r1, ref.c1, 'start'), b = map(ref.r2, ref.c2, 'end');
          if (!a || !b || a[0] > b[0]) return '#REF!';
          return refText(Object.assign({}, ref, { r1: a[0], c1: a[1], r2: b[0], c2: b[1] }), true);
        });
      }));
    });
  }
}


// ---------- Sheet ----------

class Sheet {
  constructor(ss, name) { this.ss = ss; this.name = name; this.g = []; this.merges = []; this.maxRows = 1000; this.maxCols = 26; }
  getName() { return this.name; }
  setName(n) { this.name = n; return this; }
  getParent() { return this.ss; }
  cell(r, c) { const row = this.g[r - 1] || (this.g[r - 1] = []); return row[c - 1] || (row[c - 1] = blank()); }
  peek(r, c) { const row = this.g[r - 1]; return (row && row[c - 1]) || blank(); }
  value(r, c) { return evaluate(this, r, c); }

  getRange(r, c, nr, nc) {
    if (typeof r === 'string' && /^\d+:\d+$/.test(r)) {   // whole rows, e.g. "50:50"
      const [a, b] = r.split(':').map(Number);
      return new Range(this, a, 1, b - a + 1, this.getMaxColumns());
    }
    if (typeof r === 'string') {
      const m = r.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
      if (!m) throw new Error('fake getRange: unsupported A1 "' + r + '"');
      const r0 = +m[2], c0 = colToNum(m[1]);
      return m[3] ? new Range(this, r0, c0, +m[4] - r0 + 1, colToNum(m[3]) - c0 + 1) : new Range(this, r0, c0, 1, 1);
    }
    if (nr === 0 || nc === 0) throw new Error('The number of rows in the range must be at least 1.');
    return new Range(this, r, c, nr || 1, nc || 1);
  }
  getLastRow() { let n = 0; for (let i = 0; i < this.g.length; i++) if (this.g[i] && this.g[i].some(x => x && (x.v !== '' || x.f))) n = i + 1; return n; }
  getLastColumn() { let n = 0; this.g.forEach(row => row && row.forEach((x, j) => { if (x && (x.v !== '' || x.f)) n = Math.max(n, j + 1); })); return n; }
  getMaxRows() { return Math.max(this.g.length, this.maxRows); }
  getMaxColumns() { return Math.max(this.maxCols, ...this.g.map(r => (r ? r.length : 0))); }
  getDataRange() { return this.getRange(1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
  appendRow(values) { const r = this.getLastRow() + 1; this.getRange(r, 1, 1, values.length).setValues([values]); return this; }

  // --- rows ---
  insertRowsBefore(r, n) {
    this.ss.remapRefs(this, (row, col) => [row >= r ? row + n : row, col]);
    while (this.g.length < r - 1) this.g.push([]);
    const rows = []; for (let i = 0; i < n; i++) rows.push([]);
    this.g.splice(r - 1, 0, ...rows);
    this.merges.forEach(m => { if (m.row >= r) m.row += n; });
    this.maxRows += n; this.ss.touch();
    return this;
  }
  insertRowBefore(r) { return this.insertRowsBefore(r, 1); }
  insertRowAfter(r) { return this.insertRowsBefore(r + 1, 1); }
  insertRowsAfter(r, n) { return this.insertRowsBefore(r + 1, n); }
  deleteRows(r, n) {
    this.ss.remapRefs(this, (row, col, end) => {
      if (row < r) return [row, col];
      if (row >= r + n) return [row - n, col];
      if (end === 'start') return [r, col];
      if (end === 'end') return [r - 1, col];
      return null;
    });
    this.g.splice(r - 1, n);
    this.merges = this.merges.filter(m => m.row < r || m.row >= r + n);
    this.merges.forEach(m => { if (m.row >= r + n) m.row -= n; });
    this.maxRows -= n; this.ss.touch();
    return this;
  }
  deleteRow(r) { return this.deleteRows(r, 1); }

  /**
   * Sheet.moveRows(rowSpec, destinationIndex): moves whole rows so they end up just before what
   * was row destinationIndex. References follow the moved rows and the rows that shift.
   */
  moveRows(rowSpec, destinationIndex) {
    const from = rowSpec.getRow(), n = rowSpec.getNumRows();
    if (destinationIndex >= from && destinationIndex <= from + n) return; // no-op, as in Sheets
    const newRow = row => {
      if (row >= from && row < from + n) return destinationIndex > from ? destinationIndex - n + (row - from) : destinationIndex + (row - from);
      if (destinationIndex > from && row >= from + n && row < destinationIndex) return row - n;
      if (destinationIndex < from && row >= destinationIndex && row < from) return row + n;
      return row;
    };
    this.ss.remapRefs(this, (row, col) => [newRow(row), col]);
    while (this.g.length < Math.max(from + n, destinationIndex)) this.g.push([]);
    const moved = this.g.splice(from - 1, n);
    const at = destinationIndex > from ? destinationIndex - n - 1 : destinationIndex - 1;
    this.g.splice(at, 0, ...moved);
    this.merges.forEach(m => { m.row = newRow(m.row); });
    this.ss.touch();
  }

  // --- columns ---
  insertColumnAfter(c) { return this.insertColumnsBefore(c + 1, 1); }
  insertColumnBefore(c) { return this.insertColumnsBefore(c, 1); }
  insertColumnsBefore(c, n) {
    this.ss.remapRefs(this, (row, col) => [row, col >= c ? col + n : col]);
    this.g.forEach(row => { if (row && row.length >= c) { const add = []; for (let i = 0; i < n; i++) add.push(blank()); row.splice(c - 1, 0, ...add); } });
    this.merges.forEach(m => { if (m.col >= c) m.col += n; });
    this.maxCols += n; this.ss.touch();
    return this;
  }

  setColumnWidth() { return this; }
  setFrozenRows() { return this; }
  autoResizeColumns() { return this; }
  hideRows() { return this; }

  /** Test helper: a merged block of numCols columns on one row. */
  addMerge(row, col, numCols) { this.merges.push({ row, col, numCols }); return this; }
}


// ---------- Range ----------

class Range {
  constructor(s, r, c, nr, nc) { this.s = s; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  getSheet() { return this.s; }
  getRow() { return this.r; }
  getColumn() { return this.c; }
  getNumRows() { return this.nr; }
  getNumColumns() { return this.nc; }
  getLastRow() { return this.r + this.nr - 1; }
  getLastColumn() { return this.c + this.nc - 1; }
  getA1Notation() {
    const a = numToCol(this.c) + this.r;
    return this.nr === 1 && this.nc === 1 ? a : a + ':' + numToCol(this.c + this.nc - 1) + (this.r + this.nr - 1);
  }
  offset(dr, dc, nr, nc) { return new Range(this.s, this.r + dr, this.c + dc, nr || this.nr, nc || this.nc); }

  map(fn) {
    const o = [];
    for (let i = 0; i < this.nr; i++) { const row = []; for (let j = 0; j < this.nc; j++) row.push(fn(this.s.peek(this.r + i, this.c + j), this.r + i, this.c + j)); o.push(row); }
    return o;
  }
  getValues() { return this.map((x, r, c) => evaluate(this.s, r, c)); }
  getValue() { return evaluate(this.s, this.r, this.c); }
  getDisplayValues() {
    return this.map((x, r, c) => {
      if (x.d !== undefined) return String(x.d);
      const v = evaluate(this.s, r, c);
      return v instanceof Date ? v.toDateString() : String(v);
    });
  }
  getDisplayValue() { return this.getDisplayValues()[0][0]; }
  getFormulas() { return this.map(x => x.f); }
  getFormula() { return this.s.peek(this.r, this.c).f; }

  checkSize(vals, what) {
    if (vals.length !== this.nr || vals.some(row => row.length !== this.nc)) {
      throw new Error(`The number of rows or columns in the data does not match the range (${what} ${vals.length}x${vals[0] && vals[0].length} into ${this.nr}x${this.nc}).`);
    }
  }
  setValues(vals) {
    this.checkSize(vals, 'setValues');
    vals.forEach((row, i) => row.forEach((v, j) => {
      const x = this.s.cell(this.r + i, this.c + j);
      if (typeof v === 'string' && v[0] === '=') { x.f = v; x.v = ''; } else { x.v = v === null || v === undefined ? '' : v; x.f = ''; }
      x.d = undefined;
    }));
    this.s.ss.touch();
    return this;
  }
  setValue(v) { return this.setValues([[v]]); }
  setFormulas(f) {
    this.checkSize(f, 'setFormulas');
    f.forEach((row, i) => row.forEach((v, j) => { const x = this.s.cell(this.r + i, this.c + j); x.f = v || ''; x.v = ''; x.d = undefined; }));
    this.s.ss.touch();
    return this;
  }
  setFormula(f) { return this.setFormulas([[f]]); }

  getFormulasR1C1() {
    return this.map((x, r, c) => x.f ? rewriteRefs(x.f, ref => {
      const part = (col, row, ac, ar) => 'R' + (ar ? row : (row === r ? '' : '[' + (row - r) + ']')) + 'C' + (ac ? col : (col === c ? '' : '[' + (col - c) + ']'));
      let s = ref.sheet !== null ? refText(Object.assign({}, ref, { c2: null }), true).split('!')[0] + '!' : '';
      s += part(ref.c1, ref.r1, ref.abs.c1, ref.abs.r1);
      if (ref.c2 !== null) s += ':' + part(ref.c2, ref.r2, ref.abs.c2, ref.abs.r2);
      return s;
    }) : '');
  }
  setFormulasR1C1(f) {
    this.checkSize(f, 'setFormulasR1C1');
    f.forEach((row, i) => row.forEach((v, j) => {
      const x = this.s.cell(this.r + i, this.c + j);
      x.f = v ? fromR1C1(v, this.r + i, this.c + j) : '';
      x.v = ''; x.d = undefined;
    }));
    this.s.ss.touch();
    return this;
  }

  clearContent() { this.map((x, r, c) => { const y = this.s.cell(r, c); y.v = ''; y.f = ''; y.d = undefined; }); this.s.ss.touch(); return this; }
  clear() { return this.clearContent(); }

  /** Cut and paste: the cells move, and every reference to them follows. */
  moveTo(target) {
    const src = this, t = target;
    // Whatever was in the destination is overwritten; references to it break, as in Sheets.
    const inSrc = (r, c) => r >= src.r && r < src.r + src.nr && c >= src.c && c < src.c + src.nc;
    const inDst = (r, c) => r >= t.r && r < t.r + src.nr && c >= t.c && c < t.c + src.nc;
    if (src.s !== t.s) throw new Error('fake moveTo: only within one sheet');
    const map = (r, c) => {
      if (inSrc(r, c)) return [r - src.r + t.r, c - src.c + t.c];
      if (inDst(r, c)) return null;
      return [r, c];
    };
    map.wholeRangesOnly = true;
    src.s.ss.remapRefs(src.s, map);
    const cells = src.map(x => Object.assign({}, x));
    src.map((x, r, c) => Object.assign(src.s.cell(r, c), blank()));
    cells.forEach((row, i) => row.forEach((x, j) => Object.assign(t.s.cell(t.r + i, t.c + j), x)));
    src.s.ss.touch();
  }

  /** insertCells(SpreadsheetApp.Dimension.ROWS): shifts this block's columns down by nr rows. */
  insertCells(dimension) {
    if (dimension !== 'ROWS') throw new Error('fake insertCells: only ROWS');
    this.blockedByMerge();
    const { r, c, nr, nc } = this;
    this.s.ss.remapRefs(this.s, (row, col) => [row >= r && col >= c && col < c + nc ? row + nr : row, col]);
    for (let j = 0; j < nc; j++) {
      const col = c + j, last = this.s.g.length;
      for (let row = last; row >= r; row--) Object.assign(this.s.cell(row + nr, col), this.s.peek(row, col));
      for (let row = r; row < r + nr; row++) Object.assign(this.s.cell(row, col), blank());
    }
    this.s.merges.forEach(m => { if (m.row >= r && m.col >= c && m.col < c + nc) m.row += nr; });
    this.s.ss.touch();
    return this;
  }
  /** deleteCells(SpreadsheetApp.Dimension.ROWS): removes this block and pulls the cells below up. */
  deleteCells(dimension) {
    if (dimension !== 'ROWS') throw new Error('fake deleteCells: only ROWS');
    this.blockedByMerge();
    const { r, c, nr, nc } = this;
    this.s.ss.remapRefs(this.s, (row, col, end) => {
      if (col < c || col >= c + nc || row < r) return [row, col];
      if (row >= r + nr) return [row - nr, col];
      if (end === 'start') return [r, col];
      if (end === 'end') return [r - 1, col];
      return null;
    });
    for (let j = 0; j < nc; j++) {
      const col = c + j, last = this.s.g.length;
      for (let row = r; row <= last; row++) Object.assign(this.s.cell(row, col), this.s.peek(row + nr, col));
    }
    this.s.merges = this.s.merges.filter(m => !(m.row >= r && m.row < r + nr && m.col >= c && m.col < c + nc));
    this.s.merges.forEach(m => { if (m.row >= r + nr && m.col >= c && m.col < c + nc) m.row -= nr; });
    this.s.ss.touch();
    return this;
  }
  /** Sheets refuses to shift part of a merged cell. */
  blockedByMerge() {
    const last = this.c + this.nc - 1;
    this.s.merges.forEach(m => {
      const mLast = m.col + m.numCols - 1;
      if (m.row >= this.r && mLast >= this.c && m.col <= last && (m.col < this.c || mLast > last)) {
        throw new Error('You must select all cells in a merged range to merge or unmerge them.');
      }
    });
  }

  breakApart() {
    const last = this.c + this.nc - 1;
    this.s.merges = this.s.merges.filter(m => !(m.row >= this.r && m.row < this.r + this.nr && m.col + m.numCols - 1 >= this.c && m.col <= last));
    return this;
  }

  getMergedRanges() {
    const last = this.c + this.nc - 1;
    return this.s.merges
      .filter(m => m.row >= this.r && m.row < this.r + this.nr && m.col + m.numCols - 1 >= this.c && m.col <= last)
      .map(m => new Range(this.s, m.row, m.col, 1, m.numCols));
  }

  copyTo(target) {
    // Formatting only, as the code uses it; formulas are copied through R1C1 by the caller.
    return this;
  }
}
['setNumberFormat', 'setNumberFormats', 'setFontColor', 'setFontWeight', 'setFontStyle', 'setBackground', 'setHorizontalAlignment',
  'setVerticalAlignment', 'setWrap', 'setDataValidation', 'setNote', 'merge', 'setBorder', 'activate', 'setFontSize',
  'setFontFamily', 'clearFormat', 'clearDataValidations']
  .forEach(m => { Range.prototype[m] = function () { return this; }; });


function fromR1C1(f, r, c) {
  return f.replace(/(?:(?:'((?:[^']|'')+)'|([A-Za-z_][A-Za-z0-9_]*))(!))?R(\[-?\d+\]|\d*)C(\[-?\d+\]|\d*)(?::R(\[-?\d+\]|\d*)C(\[-?\d+\]|\d*))?/g,
    (m, qName, name, bang, R1, C1, R2, C2, offset, whole) => {
      const prev = whole[offset - 1], next = whole[offset + m.length];
      if ((prev && /[A-Za-z0-9_$.]/.test(prev)) || next === '(') return m;
      const part = (Rs, Cs) => {
        const row = Rs === '' ? r : Rs[0] === '[' ? r + +Rs.slice(1, -1) : +Rs;
        const col = Cs === '' ? c : Cs[0] === '[' ? c + +Cs.slice(1, -1) : +Cs;
        return (Cs !== '' && Cs[0] !== '[' ? '$' : '') + numToCol(col) + (Rs !== '' && Rs[0] !== '[' ? '$' : '') + row;
      };
      const sheet = bang ? (qName !== undefined ? "'" + qName + "'" : name) + '!' : '';
      return sheet + part(R1, C1) + (R2 !== undefined ? ':' + part(R2, C2) : '');
    });
}


// ---------- formula evaluation ----------

const cache = new WeakMap();   // spreadsheet -> { gen, values: Map }

function evaluate(sheet, r, c, depth) {
  const x = sheet.peek(r, c);
  if (!x.f) return x.v;
  const ss = sheet.ss;
  let memo = cache.get(ss);
  if (!memo || memo.gen !== ss.gen) { memo = { gen: ss.gen, values: new Map() }; cache.set(ss, memo); }
  const key = sheet.name + '|' + r + '|' + c;
  if (memo.values.has(key)) return memo.values.get(key);
  if ((depth || 0) > 50) return '#CYCLE';
  const v = evalFormula(sheet, x.f, depth || 0);
  const result = v === undefined ? x.v : v;
  memo.values.set(key, result);
  return result;
}

/** @return the value, or undefined when the formula uses something this fake can't compute. */
function evalFormula(sheet, f, depth) {
  const body = f.replace(/^=/, '');
  if (/"/.test(body)) return undefined;
  const fns = body.match(/[A-Z][A-Z0-9.]*\(/g) || [];
  if (fns.some(fn => fn !== 'SUM(')) return undefined;
  if (/#REF!/.test(body)) return '#REF!';

  let bad = false;
  const read = (s, r, c) => {
    const v = evaluate(s, r, c, depth + 1);
    if (typeof v === 'string' && v !== '') {
      const n = Number(v.replace(/,/g, ''));
      if (isNaN(n)) { bad = v[0] === '#' ? v : '#VALUE!'; return 0; }
      return n;
    }
    return typeof v === 'number' ? v : 0;
  };
  const js = rewriteRefs(body, ref => {
    const s = ref.sheet === null ? sheet : sheet.ss.getSheetByName(ref.sheet);
    if (!s) { bad = '#REF!'; return '0'; }
    if (ref.c2 === null) return '(' + read(s, ref.r1, ref.c1) + ')';
    let sum = 0;
    for (let r = ref.r1; r <= ref.r2; r++) for (let c = ref.c1; c <= ref.c2; c++) sum += read(s, r, c);
    return '(' + sum + ')';
  }).replace(/SUM\(/g, '_sum(').replace(/\$/g, '');
  if (bad) return bad;
  if (!/^[\d\s+\-*/().,_sume]*$/.test(js)) return undefined;
  try {
    // eslint-disable-next-line no-new-func
    const v = Function('_sum', 'return (' + js + ');')((...a) => a.reduce((s, n) => s + n, 0));
    return typeof v === 'number' && isFinite(v) ? v : '#DIV/0!';
  } catch (e) {
    return undefined;
  }
}


module.exports = { Spreadsheet, Sheet, Range, colToNum, numToCol, rewriteRefs, fromR1C1 };

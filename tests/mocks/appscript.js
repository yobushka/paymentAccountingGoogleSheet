/**
 * Minimal SpreadsheetApp mock for Node tests (vitest/jest).
 * Supports:
 *  - SpreadsheetApp.getActive()
 *  - Spreadsheet.getSheetByName(name)
 *  - Sheet.getRange(r,c,rows,cols) | getRange('A1:B2')
 *  - Range.getValues()/setValues(), getValue()/setValue()
 *  - Sheet.getLastRow()/getLastColumn()/getDataRange()
 *  - Named ranges: setNamedRange(name, range), getNamedRanges()
 */
export function createSpreadsheetMock() {
  const namedRanges = new Map();

  class MockRange {
    constructor(sheet, r, c, nr = 1, nc = 1) {
      this.sheet = sheet;
      this.r = r; this.c = c; this.nr = nr; this.nc = nc;
    }
    _ensureSize() {
      const needR = this.r + this.nr - 1;
      const needC = this.c + this.nc - 1;
      while (this.sheet._cells.length < needR) this.sheet._cells.push([]);
      for (let i = 0; i < this.sheet._cells.length; i++) {
        const row = this.sheet._cells[i];
        while (row.length < needC) row.push('');
      }
    }
    getValues() {
      this._ensureSize();
      const out = [];
      for (let i = 0; i < this.nr; i++) {
        const row = [];
        for (let j = 0; j < this.nc; j++) {
          row.push(this.sheet._cells[this.r - 1 + i][this.c - 1 + j] ?? '');
        }
        out.push(row);
      }
      return out;
    }
    setValues(values) {
      this._ensureSize();
      for (let i = 0; i < this.nr; i++) {
        for (let j = 0; j < this.nc; j++) {
          const v = values[i]?.[j];
          this.sheet._cells[this.r - 1 + i][this.c - 1 + j] = v;
        }
      }
      return this;
    }
    getValue() { return this.getValues()[0][0]; }
    setValue(v) { return this.setValues([[v]]); }
  }

  class MockSheet {
    constructor(name) { this._name = name; this._cells = []; }
    getName() { return this._name; }
    getLastRow() {
      for (let i = this._cells.length; i > 0; i--) {
        if ((this._cells[i-1] || []).some(v => v !== '' && v != null)) return i;
      }
      return 0;
    }
    getLastColumn() {
      let max = 0;
      for (const row of this._cells) max = Math.max(max, row.length);
      return max;
    }
    getDataRange() { return new MockRange(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); }
    getRange(a1OrR, c, nr, nc) {
      if (typeof a1OrR === 'string') {
        const m = a1OrR.match(/([A-Z]+)(\d+):([A-Z]+)(\d+)/i);
        if (!m) throw new Error('A1 range not supported: ' + a1OrR);
        const r1 = parseInt(m[2], 10);
        const r2 = parseInt(m[4], 10);
        const c1 = letterToCol(m[1]);
        const c2 = letterToCol(m[3]);
        return new MockRange(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
      }
      return new MockRange(this, a1OrR, c, nr, nc);
    }
  }

  class MockSpreadsheet {
    constructor() { this._sheets = new Map(); }
    getSheetByName(name) { return this._sheets.get(name) || null; }
    _addSheet(name) { const sh = new MockSheet(name); this._sheets.set(name, sh); return sh; }
    setNamedRange(name, range) { namedRanges.set(name, range); }
    getNamedRanges() {
      return Array.from(namedRanges.entries()).map(([name, range]) => ({
        getName: () => name,
        getRange: () => range
      }));
    }
  }

  function letterToCol(s) {
    let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n;
  }

  const spreadsheet = new MockSpreadsheet();
  const SpreadsheetApp = {
    getActive: () => spreadsheet
  };

  return { SpreadsheetApp, spreadsheet, MockRange, MockSheet };
}

/** Install mock as global SpreadsheetApp for tests */
export function installAppScriptMock() {
  const env = createSpreadsheetMock();
  globalThis.SpreadsheetApp = env.SpreadsheetApp;
  return env;
}

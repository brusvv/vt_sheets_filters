const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
const sourceFiles = [
  'config/config.gs',
  'core/utils.gs',
  'rankscroller/rankScrollerState.gs',
  'rankscroller/rankScrollerVisibility.gs',
  'rankscroller/rankScrollerCalculator.gs',
  'rankscroller/rankScrollerEngine.gs',
  'rankscroller/rankScroller.gs',
  'filtering/filters.gs',
  'filtering/filterOperations.gs'
];

function columnNumber(letters) {
  let result = 0;
  for (const char of letters) {
    result = result * 26 + char.charCodeAt(0) - 64;
  }
  return result;
}

function parseA1(address) {
  const match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i.exec(address);
  if (!match) throw new Error('Unsupported mock A1 range: ' + address);
  const row = Number(match[2]);
  const column = columnNumber(match[1].toUpperCase());
  const lastRow = match[4] ? Number(match[4]) : row;
  const lastColumn = match[3]
    ? columnNumber(match[3].toUpperCase())
    : column;
  return {
    row,
    column,
    numRows: lastRow - row + 1,
    numColumns: lastColumn - column + 1
  };
}

function makeSheet(name, lastRow = 2) {
  const grid = Array.from({ length: 40 }, () => Array(36).fill(''));
  const formulas = Array.from({ length: 40 }, () => Array(36).fill(''));
  const writes = [];
  const columnEvents = [];
  const hiddenColumns = new Set();

  if (['Novice', 'Intermediate', 'Advanced'].includes(name)) {
    grid[1][12] = 620;
    for (let index = 0; index < 16; index++) {
      grid[1][13 + index] = (index + 1) * 100;
    }
  }

  if (name === 'Options') {
    grid[4][0] = true;
    grid[5][0] = false;
    grid[6][0] = false;
    grid[9][0] = 2;
    grid[10][0] = 1;
  }

  if (name === 'RankReqs') {
    const columns = [8, 11, 14, 17, 20];
    for (const row of [3, 26]) {
      for (const column of columns) {
        formulas[row - 1][column - 1] =
          '=FORMULA_MARKER_' + String.fromCharCode(64 + column) + row + '()';
      }
    }

    for (let column = 1; column <= 30; column++) {
      if (![8, 11, 14, 17, 20].includes(column)) {
        hiddenColumns.add(column);
      }
    }
  }

  return {
    name,
    lastRow,
    grid,
    formulas,
    writes,
    columnEvents,
    hiddenColumns,

    getName() { return this.name; },
    getLastRow() { return this.lastRow; },

    getRange(...args) {
      const range = typeof args[0] === 'string'
        ? parseA1(args[0])
        : {
            row: args[0],
            column: args[1],
            numRows: args[2] === undefined ? 1 : args[2],
            numColumns: args[3] === undefined ? 1 : args[3]
          };
      return new MockRange(this, range.row, range.column, range.numRows, range.numColumns);
    },

    hideColumns(start, count = 1) {
      this.columnEvents.push({ method: 'hideColumns', start, count });
      for (let column = start; column < start + count; column++) {
        this.hiddenColumns.add(column);
      }
    },

    showColumns(start, count = 1) {
      this.columnEvents.push({ method: 'showColumns', start, count });
      for (let column = start; column < start + count; column++) {
        this.hiddenColumns.delete(column);
      }
    }
  };
}

class MockRange {
  constructor(sheet, row, column, numRows, numColumns) {
    this.sheet = sheet;
    this.row = row;
    this.column = column;
    this.numRows = numRows;
    this.numColumns = numColumns;
  }

  getValues() {
    return Array.from({ length: this.numRows }, (_, rowOffset) =>
      Array.from({ length: this.numColumns }, (_, columnOffset) => {
        const row = this.sheet.grid[this.row - 1 + rowOffset];
        return row && row[this.column - 1 + columnOffset] !== undefined
          ? row[this.column - 1 + columnOffset]
          : '';
      })
    );
  }

  getFormulas() {
    return Array.from({ length: this.numRows }, (_, rowOffset) =>
      Array.from({ length: this.numColumns }, (_, columnOffset) => {
        const row = this.sheet.formulas[this.row - 1 + rowOffset];
        return row && row[this.column - 1 + columnOffset] !== undefined
          ? row[this.column - 1 + columnOffset]
          : '';
      })
    );
  }

  getValue() {
    return this.getValues()[0][0];
  }

  setValue(value) {
    this.setValues([[value]]);
  }

  setValues(values) {
    for (let rowOffset = 0; rowOffset < this.numRows; rowOffset++) {
      for (let columnOffset = 0; columnOffset < this.numColumns; columnOffset++) {
        const row = this.row - 1 + rowOffset;
        const column = this.column - 1 + columnOffset;
        this.sheet.grid[row][column] = values[rowOffset][columnOffset];
        this.sheet.formulas[row][column] = '';
      }
    }
    this.sheet.writes.push({
      method: 'setValues',
      row: this.row,
      column: this.column,
      numRows: this.numRows,
      numColumns: this.numColumns
    });
  }

  setFormula(formula) {
    this.sheet.formulas[this.row - 1][this.column - 1] = formula;
    this.sheet.writes.push({
      method: 'setFormula',
      row: this.row,
      column: this.column
    });
  }

  setFormulas(formulas) {
    for (let rowOffset = 0; rowOffset < this.numRows; rowOffset++) {
      for (let columnOffset = 0; columnOffset < this.numColumns; columnOffset++) {
        this.sheet.formulas[this.row - 1 + rowOffset][this.column - 1 + columnOffset] =
          formulas[rowOffset][columnOffset];
      }
    }
    this.sheet.writes.push({
      method: 'setFormulas',
      row: this.row,
      column: this.column,
      numRows: this.numRows,
      numColumns: this.numColumns
    });
  }

  clearContent() {
    for (let rowOffset = 0; rowOffset < this.numRows; rowOffset++) {
      for (let columnOffset = 0; columnOffset < this.numColumns; columnOffset++) {
        const row = this.row - 1 + rowOffset;
        const column = this.column - 1 + columnOffset;
        this.sheet.grid[row][column] = '';
        this.sheet.formulas[row][column] = '';
      }
    }
    this.sheet.writes.push({
      method: 'clearContent',
      row: this.row,
      column: this.column,
      numRows: this.numRows,
      numColumns: this.numColumns
    });
  }
}

function createFixture() {
  const sheets = {
    Novice: makeSheet('Novice'),
    Intermediate: makeSheet('Intermediate'),
    Advanced: makeSheet('Advanced'),
    RankReqs: makeSheet('RankReqs', 26),
    Options: makeSheet('Options', 11)
  };

  const sheetLookups = [];
  const workbook = {
    getActiveSheet() {
      return sheets.RankReqs;
    },
    getSheetByName(name) {
      sheetLookups.push(name);
      return sheets[name] || null;
    }
  };

  const scriptProperties = {
    rankTiersValue: 'true',
    colBehindValue: '2',
    colAheadValue: '1',
    overallEnergyNovice: '580',
    overallEnergyIntermediate: '563',
    overallEnergyAdvanced: '580'
  };

  const propertiesApi = {
    getProperties() { return { ...scriptProperties }; },
    setProperties(values) { Object.assign(scriptProperties, values); },
    setProperty(key, value) { scriptProperties[key] = value; }
  };

  const noOp = () => {};
  const context = vm.createContext({
    console: { log: noOp, warn: noOp, error: noOp },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => workbook,
      flush: noOp
    },
    PropertiesService: {
      getScriptProperties: () => propertiesApi
    },
    withProjectLock: callback => callback(),
    triggerPerfStart_: () => ({}),
    triggerPerfStep_: noOp,
    triggerPerfEvent_: noOp,
    triggerPerfEnd_: noOp,
    operationPerfStep_: noOp,
    operationPerfEnd_: noOp,
    logFilterTiming: noOp
  });

  for (const relativePath of sourceFiles) {
    const source = fs.readFileSync(path.join(root, relativePath), 'utf8');
    vm.runInContext(source, context, { filename: relativePath });
  }

  return {
    context,
    workbook,
    sheets,
    sheetLookups,
    scriptProperties,
    optionsValues: [[true], [false], [false], [2], [1]]
  };
}

function formulaSnapshot(sheet) {
  return sheet.formulas.map(row => row.slice());
}

test('force-all Rank Scroller resolves only Novice, Intermediate and Advanced', () => {
  const fixture = createFixture();

  fixture.context.rankScroller(
    true,
    fixture.workbook,
    null,
    fixture.optionsValues
  );

  assert.deepEqual(
    fixture.sheetLookups.slice().sort(),
    ['Advanced', 'Intermediate', 'Novice'].sort()
  );

  for (const name of ['Novice', 'Intermediate', 'Advanced']) {
    assert.ok(fixture.sheets[name].columnEvents.length > 0);
    assert.ok(
      fixture.sheets[name].columnEvents.every(event =>
        ['hideColumns', 'showColumns'].includes(event.method)
      )
    );
  }

  assert.equal(fixture.sheets.RankReqs.columnEvents.length, 0);
  assert.equal(fixture.sheets.RankReqs.writes.length, 0);
});

test('RankReqs is rejected as an explicit Rank Scroller target before sheet lookup', () => {
  const fixture = createFixture();

  assert.throws(
    () => fixture.context.rankScroller(
      false,
      fixture.workbook,
      'RankReqs',
      fixture.optionsValues
    ),
    /Unknown Rank Scroller target sheet/
  );

  assert.deepEqual(fixture.sheetLookups, []);
  assert.equal(fixture.sheets.RankReqs.columnEvents.length, 0);
});

test('a RankReqs entry in CONFIG blocks the whole Rank Scroller before sheet operations', () => {
  const fixture = createFixture();

  vm.runInContext(
    'CONFIG.rankScroller.sheets.RankReqs = CONFIG.rankScroller.sheets.Intermediate;',
    fixture.context
  );

  assert.throws(
    () => fixture.context.rankScroller(
      true,
      fixture.workbook,
      null,
      fixture.optionsValues
    ),
    /must contain exactly the tracked gameplay sheets/
  );

  assert.deepEqual(fixture.sheetLookups, []);
  for (const sheet of Object.values(fixture.sheets)) {
    assert.equal(sheet.columnEvents.length, 0);
  }
});

test('engine rejects a non-game sheet before getSheetByName', () => {
  const fixture = createFixture();

  const result = fixture.context.readRankScrollerSheetData_(
    {
      sheetName: 'RankReqs',
      energyRow: 2,
      energyColumn: 13,
      thresholdColumns: 16,
      thresholdStartColumn: 14
    },
    2,
    fixture.workbook
  );

  assert.equal(result.success, false);
  assert.deepEqual(fixture.sheetLookups, []);
});

test('low-level visibility helpers fail closed for RankReqs', () => {
  const fixture = createFixture();
  const rankReqs = fixture.sheets.RankReqs;

  const columnsResult = fixture.context.showRankScrollerColumns_(
    rankReqs,
    14,
    16,
    19,
    4
  );

  const tiersResult = fixture.context.showRankScrollerTierColumns_(
    rankReqs,
    14,
    16,
    Array.from({ length: 16 }, (_, index) => (index + 1) * 100)
  );

  assert.equal(columnsResult.success, false);
  assert.equal(tiersResult.success, false);
  assert.equal(rankReqs.columnEvents.length, 0);
});

test('Intermediate refresh changes only Intermediate visibility and leaves RankReqs state untouched', () => {
  const fixture = createFixture();
  const rankReqs = fixture.sheets.RankReqs;
  const formulasBefore = formulaSnapshot(rankReqs);
  const hiddenBefore = new Set(rankReqs.hiddenColumns);

  fixture.context.rankScroller(
    false,
    fixture.workbook,
    'Intermediate',
    fixture.optionsValues
  );

  assert.deepEqual(fixture.sheetLookups, ['Intermediate']);
  assert.ok(fixture.sheets.Intermediate.columnEvents.length > 0);
  assert.ok(
    fixture.sheets.Intermediate.columnEvents.every(event =>
      ['hideColumns', 'showColumns'].includes(event.method)
    )
  );

  assert.equal(rankReqs.columnEvents.length, 0);
  assert.deepEqual(formulaSnapshot(rankReqs), formulasBefore);
  assert.deepEqual(Array.from(rankReqs.hiddenColumns), Array.from(hiddenBefore));
  assert.equal(rankReqs.writes.length, 0);
});

test('resetFilters on active RankReqs falls back to Intermediate and does not write RankReqs formulas', () => {
  const fixture = createFixture();
  const rankReqs = fixture.sheets.RankReqs;
  const formulasBefore = formulaSnapshot(rankReqs);
  const hiddenBefore = new Set(rankReqs.hiddenColumns);

  fixture.context.resetFilters();

  assert.ok(fixture.sheetLookups.includes('Options'));
  assert.ok(fixture.sheetLookups.includes('Intermediate'));
  assert.ok(!fixture.sheetLookups.includes('RankReqs'));

  assert.ok(
    fixture.sheets.Intermediate.writes.some(write =>
      write.method === 'clearContent' &&
      write.row === 1 &&
      write.column === 9 &&
      write.numRows === 1 &&
      write.numColumns === 2
    )
  );

  assert.equal(rankReqs.writes.length, 0);
  assert.equal(rankReqs.columnEvents.length, 0);
  assert.deepEqual(formulaSnapshot(rankReqs), formulasBefore);
  assert.deepEqual(Array.from(rankReqs.hiddenColumns), Array.from(hiddenBefore));
});

test('game-sheet helper rejects an explicit RankReqs target', () => {
  const fixture = createFixture();

  assert.throws(
    () => fixture.context.getTargetSheet('RankReqs', fixture.workbook),
    /Недопустимый целевой лист/
  );

  assert.equal(fixture.sheets.RankReqs.writes.length, 0);
  assert.equal(fixture.sheets.RankReqs.columnEvents.length, 0);
});

'use strict';

// Общий mock-окружение для интеграционных тестов Visible Sort и Energy mask.
// Вынесено без изменения поведения из visibleSort.integration.test.cjs.

const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..', '..');
const sourceFiles = [
  'filtering/manualState.gs',
  'filtering/filterData.gs',
  'sorting/sortVisibility.gs',
  'sorting/visibleSortPlan.gs',
  'sorting/sorting.gs',
  'filtering/filterEngine.gs',
  'filtering/filterState.gs',
  'filtering/filterVisibility.gs',
  'filtering/rowVisibility.gs',
  'filtering/filters.gs',
  'sorting/sortOperations.gs'
];

function makeRow(scenario, label, score, dots) {
  const result = Array(10).fill('');
  result[0] = scenario;
  result[1] = label;
  result[2] = score;
  result[6] = dots;
  return result;
}

function manualKey(row) {
  return JSON.stringify(row.map(value => {
    if (value === null || value === undefined) return ['null', ''];
    return [typeof value, value];
  }));
}

function columnNumber(letters) {
  let value = 0;
  for (const char of letters) value = value * 26 + char.charCodeAt(0) - 64;
  return value;
}

function parseA1(address) {
  const match = /^([A-Z]+)([0-9]+)(?::([A-Z]+)([0-9]+))?$/i.exec(address);
  if (!match) throw new Error('Unsupported mock A1 range: ' + address);
  const startRow = Number(match[2]);
  const startColumn = columnNumber(match[1].toUpperCase());
  const endRow = match[4] ? Number(match[4]) : startRow;
  const endColumn = match[3] ? columnNumber(match[3].toUpperCase()) : startColumn;
  return {
    row: startRow,
    column: startColumn,
    numRows: endRow - startRow + 1,
    numColumns: endColumn - startColumn + 1
  };
}

function createMockWorkbook(options = {}) {
  const rows = options.rows || [
    makeRow('A', 'duplicate', 10, 'A'),
    makeRow('A', 'duplicate', 10, 'A'),
    makeRow('A', 'other', 30, 'Z')
  ];
  const firstRow = 3;
  const lastRow = firstRow + rows.length - 1;
  const gameGrid = Array.from(
    { length: Math.max(lastRow + 2, 10) },
    () => Array(32).fill('')
  );
  rows.forEach((row, index) => {
    row.forEach((value, offset) => {
      gameGrid[firstRow - 1 + index][2 + offset] = value;
    });
  });

  // Filter controls are in I1:J1 (absolute columns 9 and 10).
  gameGrid[0][8] = '';
  gameGrid[0][9] = '';

  const stateGrid = Array.from(
    { length: Math.max(lastRow + 5, 20) },
    () => Array(24).fill('')
  );

  // Valid headers for all three FilterState blocks.
  stateGrid[0][0] = 'Novice Visible';
  stateGrid[0][1] = 'LastRow';
  stateGrid[0][3] = 'Intermediate Visible';
  stateGrid[0][4] = 'LastRow';
  stateGrid[0][6] = 'Advanced Visible';
  stateGrid[0][7] = 'LastRow';

  stateGrid[1][1] = lastRow;
  stateGrid[1][4] = options.savedLastRow === undefined ? lastRow : options.savedLastRow;
  stateGrid[1][7] = lastRow;

  const filterValues = options.filterVisibility || Array(rows.length).fill(true);
  filterValues.forEach((value, index) => {
    stateGrid[firstRow - 1 + index][3] = value;
  });

  stateGrid[0][12] = 'Intermediate Manual';
  stateGrid[0][13] = 'Count';
  stateGrid[1][13] = options.manualCount === undefined ? 1 : options.manualCount;

  const keyRows = options.manualKeys === undefined
    ? [manualKey(rows[1] || rows[0])]
    : options.manualKeys;
  keyRows.forEach((key, index) => {
    stateGrid[firstRow - 1 + index][12] = key;
  });

  const events = [];
  const sheetWrites = [];
  const stateWrites = [];
  const visibilityEvents = [];
  let rangeSortCalls = 0;

  class MockRange {
    constructor(grid, startRow, startColumn, numRows, numColumns, owner) {
      this.grid = grid;
      this.startRow = startRow;
      this.startColumn = startColumn;
      this.numRows = numRows;
      this.numColumns = numColumns;
      this.owner = owner;
    }
    getValues() {
      this.owner.reads.push({
        startRow: this.startRow,
        startColumn: this.startColumn,
        numRows: this.numRows,
        numColumns: this.numColumns
      });
      const result = [];
      for (let r = 0; r < this.numRows; r++) {
        const row = [];
        for (let c = 0; c < this.numColumns; c++) {
          row.push(
            this.grid[this.startRow - 1 + r] &&
            this.grid[this.startRow - 1 + r][this.startColumn - 1 + c] !== undefined
              ? this.grid[this.startRow - 1 + r][this.startColumn - 1 + c]
              : ''
          );
        }
        result.push(row);
      }
      return result;
    }
    getValue() {
      return this.getValues()[0][0];
    }
    setValues(values) {
      if (!Array.isArray(values) || values.length !== this.numRows) {
        throw new Error('Mock setValues row mismatch');
      }
      values.forEach((row, r) => {
        if (!Array.isArray(row) || row.length !== this.numColumns) {
          throw new Error('Mock setValues column mismatch');
        }
        row.forEach((value, c) => {
          this.grid[this.startRow - 1 + r][this.startColumn - 1 + c] = value;
        });
      });
      const record = {
        startRow: this.startRow,
        startColumn: this.startColumn,
        numRows: this.numRows,
        numColumns: this.numColumns,
        values: values.map(row => row.slice())
      };
      this.owner.writes.push(record);
      if (this.owner.name === '_FilterState') stateWrites.push(record);
      else sheetWrites.push(record);
      events.push(['setValues', this.owner.name, this.startRow, this.startColumn]);
    }
    clearContent() {
      for (let r = 0; r < this.numRows; r++) {
        for (let c = 0; c < this.numColumns; c++) {
          this.grid[this.startRow - 1 + r][this.startColumn - 1 + c] = '';
        }
      }
      this.owner.writes.push({
        method: 'clearContent',
        startRow: this.startRow,
        startColumn: this.startColumn,
        numRows: this.numRows,
        numColumns: this.numColumns
      });
      events.push(['clearContent', this.owner.name, this.startRow, this.startColumn]);
    }
    getA1Notation() {
      function col(n) {
        let result = '';
        while (n > 0) {
          const rem = (n - 1) % 26;
          result = String.fromCharCode(65 + rem) + result;
          n = Math.floor((n - 1) / 26);
        }
        return result;
      }
      const first = col(this.startColumn) + this.startRow;
      const last = col(this.startColumn + this.numColumns - 1) +
        (this.startRow + this.numRows - 1);
      return first === last ? first : first + ':' + last;
    }
    sort() {
      rangeSortCalls++;
      throw new Error('Legacy Range.sort must never run');
    }
  }

  class MockSheet {
    constructor(name, id, grid, parent = null) {
      this.name = name;
      this.id = id;
      this.grid = grid;
      this.parent = parent;
      this.reads = [];
      this.hiddenRows = new Set(options.hiddenRows || [4]);
      this.filterHiddenRows = new Set(options.filterHiddenRows || []);
      this.writes = [];
    }
    getName() { return this.name; }
    getSheetId() { return this.id; }
    getParent() { return this.parent; }
    getLastRow() { return this.lastRow; }
    setLastRow(value) { this.lastRow = value; }
    getRange(...args) {
      const parsed = typeof args[0] === 'string'
        ? parseA1(args[0])
        : {
            row: args[0],
            column: args[1],
            numRows: args[2] === undefined ? 1 : args[2],
            numColumns: args[3] === undefined ? 1 : args[3]
          };
      return new MockRange(
        this.grid,
        parsed.row,
        parsed.column,
        parsed.numRows,
        parsed.numColumns,
        this
      );
    }
    isRowHiddenByUser(row) { return this.hiddenRows.has(row); }
    isRowHiddenByFilter(row) { return this.filterHiddenRows.has(row); }
    showRows(start, count = 1) {
      visibilityEvents.push(['showRows', start, count]);
      for (let row = start; row < start + count; row++) this.hiddenRows.delete(row);
    }
    hideRows(start, count = 1) {
      visibilityEvents.push(['hideRows', start, count]);
      for (let row = start; row < start + count; row++) this.hiddenRows.add(row);
    }
  }

  const spreadsheet = {
    getSheetByName(name) {
      if (name === '_FilterState') return options.noStateSheet ? null : stateSheet;
      if (name === 'Intermediate') return gameSheet;
      return null;
    },
    getActiveSheet() { return gameSheet; },
    toast(message, title, seconds) {
      events.push(['toast', message, title, seconds]);
    }
  };

  const gameSheet = new MockSheet('Intermediate', 77, gameGrid, spreadsheet);
  gameSheet.setLastRow(options.gameLastRow === undefined ? lastRow : options.gameLastRow);

  const stateSheet = new MockSheet('_FilterState', 88, stateGrid, spreadsheet);
  stateSheet.setLastRow(lastRow);

  const sandbox = {
    CONFIG: {
      firstRow: 3,
      scenarioCol: 3,
      scoreCol: 5,
      dotsCol: 9,
      dataWidth: 7,
      sortStartCol: 3,
      sortWidth: 10
    },
    console: { log() {}, warn() {}, error() {} },
    SpreadsheetApp: {
      getActiveSpreadsheet() { return spreadsheet; },
      flush() {}
    },
    withProjectLock(fn) { return fn(); },
    getTargetSheet() { return gameSheet; },
    triggerPerfStart_(name, metadata) {
      events.push(['start', name, metadata || {}]);
      return { name: name };
    },
    triggerPerfEnd_(trace, status, metadata) {
      events.push(['end', trace && trace.name, status, metadata || {}]);
    },
    triggerPerfEvent_(trace, name, metadata) {
      events.push(['perfEvent', name, metadata || {}]);
    },
    filterPerfStep_() {},
    logFilterTiming() {},
    operationPerfStep_() {},
    operationPerfEnd_(trace, status) {
      events.push(['operationEnd', trace && trace.name, status]);
    }
  };

  vm.createContext(sandbox);
  sourceFiles.forEach(file => {
    vm.runInContext(
      fs.readFileSync(path.join(root, file), 'utf8'),
      sandbox,
      { filename: file }
    );
  });

  return {
    project: sandbox,
    spreadsheet: spreadsheet,
    gameSheet: gameSheet,
    stateSheet: stateSheet,
    events: events,
    sheetWrites: sheetWrites,
    stateWrites: stateWrites,
    visibilityEvents: visibilityEvents,
    rangeSortCalls: () => rangeSortCalls,
    rows: rows,
    firstRow: firstRow,
    lastRow: lastRow,
    manualKey: keyRows[0]
  };
}

function autoSortOptions(enabled = true) {
  return [
    [true],
    [enabled],
    [false],
    [0],
    [0]
  ];
}

module.exports = {
  sourceFiles,
  makeRow,
  manualKey,
  columnNumber,
  parseA1,
  createMockWorkbook,
  autoSortOptions
};

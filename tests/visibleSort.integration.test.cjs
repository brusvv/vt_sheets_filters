
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');
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

function readGameRows(fixture) {
  return fixture.gameSheet
    .getRange(
      fixture.firstRow,
      3,
      fixture.rows.length,
      10
    )
    .getValues();
}

function readManualState(fixture) {
  return fixture.stateSheet
    .getRange(3, 13, Math.max(1, fixture.stateGridRows || 1), 1)
    .getValues();
}

test('Resolver handles two duplicate ManualState keys mapped to two hidden physical rows', () => {
  const duplicate = makeRow('A', 'duplicate', 10, 'A');
  const fixture = createMockWorkbook({
    rows: [duplicate.slice(), duplicate.slice(), makeRow('A', 'other', 30, 'Z')],
    hiddenRows: [3, 4],
    manualCount: 2,
    manualKeys: [manualKey(duplicate), manualKey(duplicate)]
  });

  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );

  assert.equal(context.ok, true);
  assert.deepEqual(Array.from(context.manualHiddenRows), [3, 4]);
  assert.deepEqual(Array.from(context.visibleByPhysicalPosition), [false, false, true]);
});

test('Resolver rejects a ManualState key that overlaps a FilterState-hidden physical row', () => {
  const fixture = createMockWorkbook({
    filterVisibility: [true, false, true],
    hiddenRows: [4],
    manualCount: 1
  });

  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );

  assert.equal(context.ok, false);
  assert.equal(context.reason, 'MANUALSTATE_PHYSICAL_COUNT_MISMATCH');
});

test('Resolver rejects corrupted non-boolean FilterState values', () => {
  const fixture = createMockWorkbook();
  fixture.stateSheet.grid[2][3] = 'TRUE';

  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );

  assert.equal(context.ok, false);
  assert.equal(context.reason, 'VISIBILITY_SNAPSHOT_INVALID');
});

test('Resolver rejects an absent _FilterState sheet without creating it', () => {
  const fixture = createMockWorkbook({ noStateSheet: true });

  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );

  assert.equal(context.ok, false);
  assert.equal(context.reason, 'FILTERSTATE_SHEET_MISSING');
});

test('Resolver rejects invalid FilterState metadata', () => {
  const fixture = createMockWorkbook();
  fixture.stateSheet.grid[0][3] = 'wrong header';

  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );

  assert.equal(context.ok, false);
  assert.equal(context.reason, 'FILTERSTATE_METADATA_INVALID');
});

test('Resolver accepts a shortened data range when the saved FilterState is a longer valid prefix', () => {
  const fixture = createMockWorkbook({
    rows: [
      makeRow('A', 'duplicate', 10, 'A'),
      makeRow('A', 'duplicate', 10, 'A')
    ],
    hiddenRows: [4],
    savedLastRow: 5,
    gameLastRow: 5
  });

  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    4,
    readGameRows(fixture)
  );

  assert.equal(context.ok, true);
  assert.equal(context.lastRow, 4);
  assert.equal(context.savedLastRow, 5);
});

test('Resolver rejects an increased scenario range when FilterState has not caught up', () => {
  const fixture = createMockWorkbook({
    rows: [
      makeRow('A', 'duplicate', 10, 'A'),
      makeRow('A', 'duplicate', 10, 'A'),
      makeRow('A', 'other', 30, 'Z'),
      makeRow('A', 'new', 40, 'Y')
    ],
    gameLastRow: 6,
    savedLastRow: 5,
    manualCount: 1
  });

  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    6,
    readGameRows(fixture)
  );

  assert.equal(context.ok, false);
  assert.equal(context.reason, 'FILTERSTATE_TOO_SHORT');
});

test('Resolver rejects a physically hidden row with an empty scenario key', () => {
  const fixture = createMockWorkbook({
    rows: [
      makeRow('A', 'one', 10, 'A'),
      makeRow('', 'hidden empty', 99, 'Z'),
      makeRow('A', 'other', 30, 'B')
    ],
    hiddenRows: [4],
    manualCount: 0,
    manualKeys: []
  });

  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );

  assert.equal(context.ok, false);
  assert.equal(context.reason, 'UNTRACKED_HIDDEN_EMPTY_SCENARIO');
});

test('Resolver rejects a corrupted ManualState count', () => {
  const fixture = createMockWorkbook({ manualCount: 1 });
  fixture.stateSheet.grid[1][13] = 1.5;

  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );

  assert.equal(context.ok, false);
  assert.equal(context.reason, 'MANUALSTATE_COUNT_INVALID');
});

test('snapshot validator detects mutation of captured FilterState/ManualState fields', () => {
  const fixture = createMockWorkbook();
  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );
  assert.equal(context.ok, true);

  context.filterVisibility[0] = false;

  const result = fixture.project.validateSortVisibilityContext_(
    context,
    fixture.gameSheet,
    readGameRows(fixture),
    3,
    fixture.lastRow
  );

  assert.equal(result.ok, false);
  assert.equal(result.reason, 'VISIBILITY_CONTEXT_FINGERPRINT_STALE');
});

test('fresh validator accepts an unchanged sheet and state snapshot', () => {
  const fixture = createMockWorkbook();
  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );

  const result = fixture.project.validateSortVisibilityContextFresh_(
    fixture.gameSheet,
    context
  );

  assert.equal(result.ok, true);
});

test('fresh validator detects C:L changes since snapshot', () => {
  const fixture = createMockWorkbook();
  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );

  fixture.gameSheet.grid[2][4] = 777;

  const result = fixture.project.validateSortVisibilityContextFresh_(
    fixture.gameSheet,
    context
  );

  assert.equal(result.ok, false);
  assert.equal(result.reason, 'VISIBILITY_CONTEXT_DATA_STALE');
});

test('fresh validator detects FilterState and physical visibility changes since snapshot', () => {
  const fixture = createMockWorkbook();
  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );

  fixture.stateSheet.grid[2][3] = false;
  fixture.gameSheet.hideRows(3, 1);

  const result = fixture.project.validateSortVisibilityContextFresh_(
    fixture.gameSheet,
    context
  );

  assert.equal(result.ok, false);
  assert.equal(result.reason, 'VISIBILITY_CONTEXT_STATE_STALE');
});

test('fresh validator detects ManualState changes since snapshot', () => {
  const fixture = createMockWorkbook();
  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );

  fixture.stateSheet.grid[2][12] = 'different-manual-key';

  const result = fixture.project.validateSortVisibilityContextFresh_(
    fixture.gameSheet,
    context
  );

  assert.equal(result.ok, false);
  assert.equal(result.reason, 'VISIBILITY_CONTEXT_FRESH_READ_FAILED');
});

test('fresh validator detects a newly appended scenario row and range-boundary change', () => {
  const fixture = createMockWorkbook();
  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );

  const newRow = makeRow('B', 'new', 99, 'Z');
  newRow.forEach((value, offset) => {
    fixture.gameSheet.grid[5][2 + offset] = value;
  });
  fixture.gameSheet.setLastRow(6);

  const result = fixture.project.validateSortVisibilityContextFresh_(
    fixture.gameSheet,
    context
  );

  assert.equal(result.ok, false);
  assert.equal(result.reason, 'VISIBILITY_CONTEXT_BOUNDARY_STALE');
});

test('fresh validator rejects a different target sheet', () => {
  const fixture = createMockWorkbook();
  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );
  const otherSheet = {
    getName() { return 'Advanced'; },
    getSheetId() { return 999; }
  };

  const result = fixture.project.validateSortVisibilityContextFresh_(
    otherSheet,
    context
  );

  assert.equal(result.ok, false);
  assert.equal(result.reason, 'VISIBILITY_CONTEXT_SHEET_MISMATCH');
});

test('AutoSort integration: TRANSFER_NOT_VERIFIED is followed by filtering with resolved manual positions and no content/state loss', () => {
  const fixture = createMockWorkbook();
  const originalRows = JSON.stringify(readGameRows(fixture));
  const originalFilterState = JSON.stringify(
    fixture.stateSheet.getRange(3, 4, 3, 1).getValues()
  );
  const originalManualState = JSON.stringify(
    fixture.stateSheet.getRange(1, 13, 4, 2).getValues()
  );
  const originalHidden = Array.from(fixture.gameSheet.hiddenRows).sort((a, b) => a - b);

  const calls = [];
  const project = fixture.project;
  [
    'readSortVisibilityContext_',
    'sortScenarioBlocksDetailed_',
    'prepareFilterState_',
    'buildFilterVisibility_',
    'saveFilterVisibility_'
  ].forEach(name => {
    const original = project[name];
    project[name] = function(...args) {
      calls.push(name);
      return original.apply(this, args);
    };
  });

  project.applyFilters(false, fixture.gameSheet, false, autoSortOptions(true), null);

  assert.deepEqual(calls, [
    'readSortVisibilityContext_',
    'sortScenarioBlocksDetailed_',
    'prepareFilterState_',
    'buildFilterVisibility_',
    'saveFilterVisibility_'
  ]);

  const sortEvent = fixture.events.find(item =>
    item[0] === 'perfEvent' && item[1] === 'AUTO_SORT_PLAN'
  );
  assert.ok(sortEvent);
  assert.equal(sortEvent[2].status, 'TRANSFER_NOT_VERIFIED');
  assert.equal(sortEvent[2].applied, false);

  assert.equal(JSON.stringify(readGameRows(fixture)), originalRows);
  assert.equal(JSON.stringify(fixture.stateSheet.getRange(3, 4, 3, 1).getValues()), originalFilterState);
  assert.equal(JSON.stringify(fixture.stateSheet.getRange(1, 13, 4, 2).getValues()), originalManualState);
  assert.deepEqual(Array.from(fixture.gameSheet.hiddenRows).sort((a, b) => a - b), originalHidden);
  assert.equal(fixture.sheetWrites.length, 0);
  assert.equal(fixture.rangeSortCalls(), 0);
});

test('AutoSort resetVisibility integration: resolver runs before showRows and the same manual row is hidden again', () => {
  const fixture = createMockWorkbook();
  const originalRows = JSON.stringify(readGameRows(fixture));
  const originalManualState = JSON.stringify(
    fixture.stateSheet.getRange(1, 13, 4, 2).getValues()
  );

  fixture.project.applyFilters(
    false,
    fixture.gameSheet,
    true,
    autoSortOptions(true),
    null
  );

  assert.deepEqual(fixture.visibilityEvents, [
    ['showRows', 3, 3],
    ['hideRows', 4, 1]
  ]);
  assert.deepEqual(Array.from(fixture.gameSheet.hiddenRows), [4]);
  assert.equal(JSON.stringify(readGameRows(fixture)), originalRows);
  assert.equal(
    JSON.stringify(fixture.stateSheet.getRange(1, 13, 4, 2).getValues()),
    originalManualState
  );
  assert.equal(fixture.rangeSortCalls(), 0);
});

test('AutoSort integration: unresolved duplicate ManualState rows abort before visibility or state writes', () => {
  const fixture = createMockWorkbook({
    hiddenRows: [3, 4],
    manualCount: 1
  });
  const originalRows = JSON.stringify(readGameRows(fixture));
  const originalFilterState = JSON.stringify(
    fixture.stateSheet.getRange(3, 4, 3, 1).getValues()
  );
  const originalManualState = JSON.stringify(
    fixture.stateSheet.getRange(1, 13, 4, 2).getValues()
  );

  const result = fixture.project.applyFilters(
    false,
    fixture.gameSheet,
    false,
    autoSortOptions(true),
    null
  );

  assert.equal(result.status, 'SORT_VISIBILITY_BLOCKED');
  assert.equal(result.reason, 'UNRESOLVED_MANUALSTATE_VISIBILITY');
  assert.deepEqual(Array.from(fixture.gameSheet.hiddenRows).sort((a, b) => a - b), [3, 4]);
  assert.equal(fixture.visibilityEvents.length, 0);
  assert.equal(fixture.sheetWrites.length, 0);
  assert.equal(fixture.stateWrites.length, 0);
  assert.equal(JSON.stringify(readGameRows(fixture)), originalRows);
  assert.equal(JSON.stringify(fixture.stateSheet.getRange(3, 4, 3, 1).getValues()), originalFilterState);
  assert.equal(JSON.stringify(fixture.stateSheet.getRange(1, 13, 4, 2).getValues()), originalManualState);
  assert.equal(fixture.rangeSortCalls(), 0);
});

test('AutoSort resetVisibility integration: invalid resolver context aborts before showRows', () => {
  const fixture = createMockWorkbook({ noStateSheet: true });
  const originalRows = JSON.stringify(readGameRows(fixture));

  const result = fixture.project.applyFilters(
    false,
    fixture.gameSheet,
    true,
    autoSortOptions(true),
    null
  );

  assert.equal(result.status, 'SORT_VISIBILITY_BLOCKED');
  assert.deepEqual(fixture.visibilityEvents, []);
  assert.equal(JSON.stringify(readGameRows(fixture)), originalRows);
  assert.equal(fixture.rangeSortCalls(), 0);
});

test('AutoSort OFF integration: resolver and sorter are not called', () => {
  const fixture = createMockWorkbook({
    rows: [
      makeRow('A', 'one', 10, 'A'),
      makeRow('A', 'two', 30, 'B'),
      makeRow('A', 'three', 20, 'C')
    ],
    hiddenRows: [],
    manualCount: 0,
    manualKeys: []
  });
  let resolverCalls = 0;
  let sorterCalls = 0;
  const resolver = fixture.project.readSortVisibilityContext_;
  const sorter = fixture.project.sortScenarioBlocksDetailed_;
  fixture.project.readSortVisibilityContext_ = function(...args) {
    resolverCalls++;
    return resolver.apply(this, args);
  };
  fixture.project.sortScenarioBlocksDetailed_ = function(...args) {
    sorterCalls++;
    return sorter.apply(this, args);
  };

  fixture.project.applyFilters(false, fixture.gameSheet, false, autoSortOptions(false), null);

  assert.equal(resolverCalls, 0);
  assert.equal(sorterCalls, 0);
  assert.equal(fixture.rangeSortCalls(), 0);
});

test('ManualSort integration: actual entry point resolves and plans, but never mutates rows or FilterState/ManualState', () => {
  const fixture = createMockWorkbook();
  const originalRows = JSON.stringify(readGameRows(fixture));
  const originalFilterState = JSON.stringify(
    fixture.stateSheet.getRange(3, 4, 3, 1).getValues()
  );
  const originalManualState = JSON.stringify(
    fixture.stateSheet.getRange(1, 13, 4, 2).getValues()
  );
  let resolverCalls = 0;
  let plannerCalls = 0;

  const originalResolver = fixture.project.readSortVisibilityContext_;
  const originalPlanner = fixture.project.sortScenarioBlocksDetailed_;
  fixture.project.readSortVisibilityContext_ = function(...args) {
    resolverCalls++;
    return originalResolver.apply(this, args);
  };
  fixture.project.sortScenarioBlocksDetailed_ = function(...args) {
    plannerCalls++;
    return originalPlanner.apply(this, args);
  };

  const result = fixture.project.runManualSortOperation_(
    'manualSortByScoreDesc',
    'scoreDesc',
    'test'
  );

  assert.equal(result.status, 'TRANSFER_NOT_VERIFIED');
  assert.equal(result.applied, false);
  assert.equal(resolverCalls, 1);
  assert.equal(plannerCalls, 1);
  assert.equal(fixture.events.filter(item => item[0] === 'toast').length, 1);
  assert.equal(JSON.stringify(readGameRows(fixture)), originalRows);
  assert.equal(JSON.stringify(fixture.stateSheet.getRange(3, 4, 3, 1).getValues()), originalFilterState);
  assert.equal(JSON.stringify(fixture.stateSheet.getRange(1, 13, 4, 2).getValues()), originalManualState);
  assert.deepEqual(Array.from(fixture.gameSheet.hiddenRows), [4]);
  assert.equal(fixture.visibilityEvents.length, 0);
  assert.equal(fixture.sheetWrites.length, 0);
  assert.equal(fixture.stateWrites.length, 0);
  assert.equal(fixture.rangeSortCalls(), 0);
});


test('fresh validator detects a change specifically in scenario column C', () => {
  const fixture = createMockWorkbook();
  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );

  fixture.gameSheet.grid[2][2] = 'B';

  const result = fixture.project.validateSortVisibilityContextFresh_(
    fixture.gameSheet,
    context
  );

  assert.equal(result.ok, false);
  assert.equal(result.reason, 'VISIBILITY_CONTEXT_DATA_STALE');
});

test('fresh validator detects a physical visibility change even when FilterState is unchanged', () => {
  const fixture = createMockWorkbook();
  const context = fixture.project.readSortVisibilityContext_(
    fixture.gameSheet,
    fixture.lastRow,
    readGameRows(fixture)
  );

  fixture.gameSheet.showRows(4, 1);

  const result = fixture.project.validateSortVisibilityContextFresh_(
    fixture.gameSheet,
    context
  );

  assert.equal(result.ok, false);
  assert.equal(result.reason, 'VISIBILITY_CONTEXT_FRESH_READ_FAILED');
});

const {
  makeRow,
  manualKey,
  createMockWorkbook,
  autoSortOptions
} = require('./helpers/mockWorkbook.cjs');
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

test('AutoSort batches Resolver metadata reads while preserving independent state refresh', () => {
  const fixture = createMockWorkbook();
  fixture.stateSheet.reads.length = 0;

  let previousVisibilityReads = 0;
  const originalGetPreviousVisibility = fixture.project.getPreviousVisibility;
  fixture.project.getPreviousVisibility = function(...args) {
    previousVisibilityReads++;
    return originalGetPreviousVisibility.apply(this, args);
  };

  fixture.project.applyFilters(
    false,
    fixture.gameSheet,
    false,
    autoSortOptions(true),
    null
  );

  const reads = fixture.stateSheet.reads;
  const countReads = (row, column, rows, columns) => reads.filter(read =>
    read.startRow === row &&
    read.startColumn === column &&
    read.numRows === rows &&
    read.numColumns === columns
  ).length;

  assert.equal(countReads(1, 1, 2, 17), 2, 'Resolver and filter preparation should each read metadata once');
  assert.equal(countReads(3, 13, 1, 1), 2, 'Resolver and filter preparation should independently read ManualState keys');
  assert.equal(countReads(3, 4, 3, 1), 2, 'Resolver and previous-visibility validation should independently read FilterState visibility');
  assert.equal(previousVisibilityReads, 1, 'previous visibility should retain its independent validation path');
  assert.equal(fixture.project.FILTER_STATE_LAYOUT_CACHE, fixture.stateSheet);
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

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');

function loadProject() {
  const sandbox = {
    CONFIG: {
      firstRow: 3,
      scenarioCol: 3,
      scoreCol: 5,
      dotsCol: 9,
      sortStartCol: 3,
      sortWidth: 10
    },
    console: { log() {}, warn() {}, error() {} }
  };

  vm.createContext(sandbox);
  [
    'filtering/manualState.gs',
    'filtering/filterData.gs',
    'sorting/sortVisibility.gs',
    'sorting/visibleSortPlan.gs',
    'sorting/sorting.gs'
  ].forEach(file => {
    vm.runInContext(
      fs.readFileSync(path.join(root, file), 'utf8'),
      sandbox,
      { filename: file }
    );
  });
  return sandbox;
}

const project = loadProject();

function row(scenario, score, dots, identity) {
  const values = Array(10).fill('');
  values[0] = scenario;
  values[1] = identity || '';
  values[2] = score;
  values[6] = dots;
  return values;
}

function mask(firstRow, visible) {
  return {
    ok: true,
    firstRow,
    lastRow: firstRow + visible.length - 1,
    visibleByPhysicalPosition: visible
  };
}

function buildPlan(data, visible, sortType = 'scoreDesc', firstRow = 3) {
  return project.buildVisibleScenarioSortPlan_(
    data,
    firstRow,
    firstRow + data.length - 1,
    mask(firstRow, visible),
    sortType
  );
}

function resolve(data, args = {}) {
  const defaults = {
    firstRow: 3,
    lastRow: 3 + data.length - 1,
    filterVisibility: Array(data.length).fill(true),
    manualKeys: [],
    physicalHiddenByUser: Array(data.length).fill(false),
    physicalHiddenByFilter: Array(data.length).fill(false)
  };
  return project.buildSortVisibilityContextFromSnapshots_({
    data,
    ...defaults,
    ...args
  });
}

test('all-visible records sort by Score descending', () => {
  const data = [
    row('A', 10, 'B', 'r3'),
    row('A', 30, 'C', 'r4'),
    row('A', 20, 'A', 'r5')
  ];
  const plan = buildPlan(data, [true, true, true]);

  assert.equal(plan.changed, true);
  assert.deepEqual(Array.from(plan.blocks[0].sourceRows), [4, 5, 3]);
  assert.deepEqual(Array.from(plan.blocks[0].visibleRows), [3, 4, 5]);
});

test('one hidden row is neither a source nor a destination', () => {
  const data = [
    row('A', 20, 'B', 'r3'),
    row('A', 999, 'Z', 'hidden-r4'),
    row('A', 10, 'A', 'r5'),
    row('A', 30, 'C', 'r6')
  ];
  const block = buildPlan(data, [true, false, true, true]).blocks[0];

  assert.deepEqual(Array.from(block.sourceRows), [6, 3, 5]);
  assert.deepEqual(Array.from(block.visibleRows), [3, 5, 6]);
  assert.equal(block.sourceRows.includes(4), false);
  assert.equal(block.visibleRows.includes(4), false);
});

test('consecutive hidden rows remain pinned', () => {
  const data = [
    row('A', 10, 'A', 'r3'),
    row('A', 900, 'Z', 'h4'),
    row('A', 800, 'Y', 'h5'),
    row('A', 30, 'C', 'r6'),
    row('A', 20, 'B', 'r7')
  ];
  const block = buildPlan(data, [true, false, false, true, true]).blocks[0];

  assert.deepEqual(Array.from(block.sourceRows), [6, 7, 3]);
  assert.deepEqual(Array.from(block.visibleRows), [3, 6, 7]);
  assert.equal(block.sourceRows.some(r => r === 4 || r === 5), false);
});

test('several disjoint visible islands are represented as physical slots', () => {
  const data = [
    row('A', 10, 'A', 'r3'),
    row('A', 100, 'X', 'h4'),
    row('A', 40, 'D', 'r5'),
    row('A', 90, 'W', 'h6'),
    row('A', 20, 'B', 'r7'),
    row('A', 30, 'C', 'r8')
  ];
  const block = buildPlan(data, [true, false, true, false, true, true]).blocks[0];

  assert.deepEqual(Array.from(block.sourceRows), [5, 8, 7, 3]);
  assert.deepEqual(Array.from(block.visibleRows), [3, 5, 7, 8]);
  assert.equal(block.sourceRows.includes(4), false);
  assert.equal(block.sourceRows.includes(6), false);
  assert.equal(block.visibleRows.includes(4), false);
  assert.equal(block.visibleRows.includes(6), false);
});

test('hidden first and last positions are excluded', () => {
  const data = [
    row('A', 999, 'Z', 'h3'),
    row('A', 10, 'A', 'r4'),
    row('A', 20, 'B', 'r5'),
    row('A', 888, 'Y', 'h6')
  ];
  const block = buildPlan(data, [false, true, true, false]).blocks[0];

  assert.deepEqual(Array.from(block.sourceRows), [5, 4]);
  assert.deepEqual(Array.from(block.visibleRows), [4, 5]);
});

test('equal Scores use Dots ascending and exact ties remain stable', () => {
  const data = [
    row('A', 50, 'B', 'r3'),
    row('A', 50, 'A', 'r4-first'),
    row('A', 50, 'A', 'r5-second')
  ];
  const plan = buildPlan(data, [true, true, true]);

  assert.deepEqual(Array.from(plan.blocks[0].sourceRows), [4, 5, 3]);
});

test('blank Scores go last and preserve their original mutual order', () => {
  const data = [
    row('A', 40, 'Z', 'numeric-40'),
    row('A', '', 'A', 'blank-first'),
    row('A', 10, 'B', 'numeric-10'),
    row('A', null, 'C', 'blank-second')
  ];
  const plan = buildPlan(data, [true, true, true, true], 'scoreDesc');

  assert.deepEqual(Array.from(plan.blocks[0].sourceRows), [3, 5, 4, 6]);
});

test('non-empty non-numeric Score fails closed', () => {
  const data = [
    row('A', 10, 'A', 'r3'),
    row('A', 'not-a-score', 'B', 'r4')
  ];
  assert.throws(
    () => buildPlan(data, [true, true]),
    /INVALID_SCORE at row 4/
  );
});

test('empty and non-ASCII Dots fail closed', () => {
  const emptyDots = [
    row('A', 10, 'A', 'r3'),
    row('A', 20, '', 'r4')
  ];
  const unicodeDots = [
    row('A', 10, 'A', 'r3'),
    row('A', 20, 'Ж', 'r4')
  ];

  assert.throws(() => buildPlan(emptyDots, [true, true]), /INVALID_DOTS at row 4/);
  assert.throws(() => buildPlan(unicodeDots, [true, true]), /NON_ASCII_DOTS at row 4/);
});

test('non-contiguous occurrences of the same scenario create separate blocks', () => {
  const data = [
    row('A', 10, 'A', 'a3'),
    row('A', 20, 'B', 'a4'),
    row('B', 10, 'A', 'b5'),
    row('B', 20, 'B', 'b6'),
    row('A', 10, 'A', 'a7'),
    row('A', 20, 'B', 'a8')
  ];
  const plan = buildPlan(data, [true, true, true, true, true, true]);

  assert.equal(plan.blocks.length, 3);
  assert.deepEqual(
    Array.from(plan.blocks, block => [block.startRow, block.endRow]),
    [[3, 4], [5, 6], [7, 8]]
  );
});

test('Dots-only mode ignores Score values', () => {
  const data = [
    row('A', 'not-numeric', 'Z', 'r3'),
    row('A', '', 'A', 'r4')
  ];
  const plan = buildPlan(data, [true, true], 'dots');

  assert.deepEqual(Array.from(plan.blocks[0].sourceRows), [4, 3]);
});

test('planner does not mutate the source snapshot', () => {
  const data = [
    row('A', 10, 'A', 'r3'),
    row('A', 20, 'B', 'r4')
  ];
  const before = JSON.stringify(data);

  buildPlan(data, [true, true]);

  assert.equal(JSON.stringify(data), before);
});

test('resolver maps a duplicate ManualState key to its exact physically hidden row', () => {
  const duplicate = row('A', 10, 'A', 'same');
  const data = [duplicate.slice(), duplicate.slice()];
  const key = project.getManualRecordKey_(duplicate);
  const result = resolve(data, {
    manualKeys: [key],
    physicalHiddenByUser: [false, true]
  });

  assert.equal(result.ok, true);
  assert.deepEqual(Array.from(result.visibleByPhysicalPosition), [true, false]);
  assert.deepEqual(Array.from(result.manualHiddenRows), [4]);
});

test('resolver rejects duplicate physical candidates when ManualState count differs', () => {
  const duplicate = row('A', 10, 'A', 'same');
  const data = [duplicate.slice(), duplicate.slice()];
  const key = project.getManualRecordKey_(duplicate);
  const result = resolve(data, {
    manualKeys: [key],
    physicalHiddenByUser: [true, true]
  });

  assert.equal(result.ok, false);
  assert.equal(result.reason, 'MANUALSTATE_PHYSICAL_COUNT_MISMATCH');
});

test('resolver rejects hidden FilterState=true rows without a ManualState key', () => {
  const result = resolve([row('A', 10, 'A', 'r3')], {
    physicalHiddenByUser: [true]
  });

  assert.equal(result.ok, false);
  assert.equal(result.reason, 'PHYSICAL_HIDDEN_ROW_MISSING_MANUALSTATE_KEY');
});

test('resolver rejects stale ManualState keys with no physical candidate', () => {
  const result = resolve([row('A', 10, 'A', 'r3')], {
    manualKeys: ['stale-key']
  });

  assert.equal(result.ok, false);
  assert.equal(result.reason, 'MANUALSTATE_PHYSICAL_COUNT_MISMATCH');
});

test('resolver rejects FilterState-hidden rows that are physically visible', () => {
  const result = resolve([row('A', 10, 'A', 'r3')], {
    filterVisibility: [false]
  });

  assert.equal(result.ok, false);
  assert.equal(result.reason, 'FILTERSTATE_PHYSICAL_VISIBILITY_MISMATCH');
});

test('resolver rejects external basic-filter-hidden rows', () => {
  const result = resolve([row('A', 10, 'A', 'r3')], {
    physicalHiddenByFilter: [true]
  });

  assert.equal(result.ok, false);
  assert.equal(result.reason, 'EXTERNAL_FILTER_UNSUPPORTED');
});

test('resolver rejects invalid snapshot dimensions and types', () => {
  const result = project.buildSortVisibilityContextFromSnapshots_({
    data: [row('A', 10, 'A', 'r3')],
    firstRow: 3,
    lastRow: 4,
    filterVisibility: [true],
    manualKeys: [],
    physicalHiddenByUser: [false],
    physicalHiddenByFilter: [false]
  });

  assert.equal(result.ok, false);
  assert.equal(result.reason, 'VISIBILITY_SNAPSHOT_INVALID');
});

test('resolver accepts FilterState-hidden rows hidden by the filter engine', () => {
  const result = resolve([row('A', 10, 'A', 'r3')], {
    filterVisibility: [false],
    physicalHiddenByUser: [true]
  });

  assert.equal(result.ok, true);
  assert.deepEqual(Array.from(result.visibleByPhysicalPosition), [false]);
});

test('sorter rejects a stale sixth-argument visibility snapshot', () => {
  const data = [
    row('A', 10, 'A', 'r3'),
    row('A', 20, 'B', 'r4')
  ];
  const staleMask = resolve(data);
  staleMask.sheetId = 77;
  staleMask.sheetName = 'Intermediate';
  staleMask.dataFingerprint = 'wrong';

  const result = project.sortScenarioBlocksDetailed_(
    {
      getLastRow: () => 4,
      getName: () => 'Intermediate',
      getSheetId: () => 77
    },
    'scoreDesc',
    4,
    null,
    data,
    staleMask
  );

  assert.equal(result.status, 'BLOCKED_VISIBILITY');
  assert.equal(result.reason, 'VISIBILITY_CONTEXT_DATA_STALE');
  assert.equal(result.applied, false);
});

test('sorter rejects a missing mask without resolving visibility itself', () => {
  const data = [
    row('A', 10, 'A', 'r3'),
    row('A', 20, 'B', 'r4')
  ];
  let rangeReads = 0;

  const result = project.sortScenarioBlocksDetailed_(
    {
      getLastRow: () => 4,
      getName: () => 'Intermediate',
      getSheetId: () => 77,
      getRange: () => {
        rangeReads++;
        throw new Error('unexpected data read');
      }
    },
    'scoreDesc',
    4,
    null,
    data,
    null
  );

  assert.equal(result.status, 'BLOCKED_VISIBILITY');
  assert.equal(result.applied, false);
  assert.equal(rangeReads, 0);
});

test('changed permutation is planned but no source, destination, or state write occurs', () => {
  const data = [
    row('A', 10, 'A', 'r3'),
    row('A', 20, 'B', 'r4')
  ];
  const context = resolve(data);
  context.sheetId = 77;
  context.sheetName = 'Intermediate';

  let writes = 0;
  const sheet = {
    getLastRow: () => 4,
    getName: () => 'Intermediate',
    getSheetId: () => 77,
    getRange: () => {
      writes++;
      throw new Error('unexpected range access after complete snapshot');
    },
    sort: () => {
      writes++;
      throw new Error('unexpected Range.sort write');
    },
    showRows: () => { writes++; },
    hideRows: () => { writes++; }
  };

  const detailed = project.sortScenarioBlocksDetailed_(
    sheet, 'scoreDesc', 4, null, data, context
  );

  assert.equal(detailed.status, 'TRANSFER_NOT_VERIFIED');
  assert.equal(detailed.applied, false);
  assert.equal(detailed.plan.changed, true);
  assert.equal(writes, 0);

  assert.throws(
    () => project.sortScenarioBlocks(
      sheet, 'scoreDesc', 4, null, data, context
    ),
    error => error && error.code === 'FULL_CELL_STAGING_ROUND_TRIP_NOT_VERIFIED'
  );
  assert.equal(writes, 0);
});

test('already sorted visible slots produce a no-op', () => {
  const data = [
    row('A', 30, 'A', 'r3'),
    row('A', 20, 'B', 'r4'),
    row('A', 10, 'C', 'r5')
  ];
  const plan = buildPlan(data, [true, false, true]);

  assert.equal(plan.changed, false);
  assert.equal(plan.blocks.length, 0);
});

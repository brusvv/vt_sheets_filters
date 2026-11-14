const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');

const FORMAT_METHODS = [
  'getNumberFormats',
  'getBackgrounds',
  'getFontColors',
  'getFontFamilies',
  'getFontSizes',
  'getFontWeights',
  'getFontStyles',
  'getFontLines',
  'getHorizontalAlignments',
  'getVerticalAlignments',
  'getWraps',
  'getTextRotations'
];

function clone(value) {
  return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
}

function makeRichText(descriptor) {
  if (!descriptor) return null;
  const style = descriptor.style || {};
  const run = {
    getText: () => descriptor.text || '',
    getLinkUrl: () => descriptor.linkUrl || null,
    getTextStyle: () => ({
      isBold: () => Boolean(style.bold),
      isItalic: () => Boolean(style.italic),
      isUnderline: () => Boolean(style.underline),
      isStrikethrough: () => Boolean(style.strikethrough),
      getFontFamily: () => style.fontFamily || null,
      getFontSize: () => style.fontSize || null,
      getForegroundColor: () => style.foregroundColor || null
    })
  };
  return {
    getText: () => descriptor.text || '',
    getLinkUrl: () => descriptor.linkUrl || null,
    getRuns: () => [run]
  };
}

function makeValidation(descriptor) {
  if (!descriptor) return null;
  return {
    getCriteriaType: () => descriptor.criteriaType,
    getCriteriaValues: () => descriptor.criteriaValues || [],
    getAllowInvalid: () => descriptor.allowInvalid,
    getHelpText: () => descriptor.helpText || null
  };
}

function defaultCell(value = '') {
  const format = {};
  FORMAT_METHODS.forEach(method => {
    format[method] = ({
      getNumberFormats: '0.00',
      getBackgrounds: '#ffffff',
      getFontColors: '#000000',
      getFontFamilies: 'Arial',
      getFontSizes: 10,
      getFontWeights: 'normal',
      getFontStyles: 'normal',
      getFontLines: 'none',
      getHorizontalAlignments: 'left',
      getVerticalAlignments: 'middle',
      getWraps: false,
      getTextRotations: 'ROTATION_0'
    })[method];
  });
  return {
    value,
    formulaR1C1: '',
    note: '',
    validation: null,
    richText: null,
    format
  };
}

function makeRange(sheet, row, column, numRows, numColumns) {
  const matrix = getter => Array.from({length: numRows}, (_, r) =>
    Array.from({length: numColumns}, (_, c) => getter(
      sheet.getCell(row + r, column + c)
    ))
  );
  const range = {
    getValues: () => matrix(cell => cell.value),
    getFormulasR1C1: () => matrix(cell => cell.formulaR1C1),
    getNotes: () => matrix(cell => cell.note),
    getDataValidations: () => matrix(cell => makeValidation(cell.validation)),
    getRichTextValues: () => matrix(cell => makeRichText(cell.richText)),
    getMergedRanges: () => sheet.mergedRanges || []
  };
  FORMAT_METHODS.forEach(method => {
    range[method] = () => matrix(cell => cell.format[method]);
  });
  return range;
}

function makeWorkbook(data, visible, options = {}) {
  const sheets = new Map();
  const properties = new Map();
  let nextSheetId = 100;
  let uuidCounter = 0;
  const events = [];
  const targetGrid = new Map();

  function gridKey(row, column) {
    return row + ':' + column;
  }

  const workbook = {
    getId: () => 'spreadsheet-test-id',
    getSheetByName: name => Array.from(sheets.values()).find(sheet => sheet.name === name) || null,
    insertSheet(name) {
      const sheet = createSheet(nextSheetId++, name, new Map());
      sheets.set(sheet.id, sheet);
      events.push(['insertSheet', name]);
      if (options.failInsertStage) throw new Error('injected staging insert failure');
      return sheet;
    },
    deleteSheet(sheet) {
      sheets.delete(sheet.id);
      events.push(['deleteSheet', sheet.name]);
    }
  };

  function createSheet(id, name, grid) {
    const sheet = {
      id,
      name,
      grid,
      hidden: false,
      hiddenRows: new Set(),
      mergedRanges: [],
      getName: () => name,
      getSheetId: () => id,
      getParent: () => workbook,
      getMaxRows: () => 1000,
      getMaxColumns: () => 26,
      getRange: (r, c, nr, nc) => makeRange(sheet, r, c, nr, nc),
      getMergedRanges: () => sheet.mergedRanges.slice(),
      hideSheet() {
        sheet.hidden = true;
        events.push(['hideSheet', name]);
      },
      setHiddenRow(rowNumber, hidden) {
        if (hidden) sheet.hiddenRows.add(rowNumber);
        else sheet.hiddenRows.delete(rowNumber);
      },
      isRowHiddenByUser(rowNumber) {
        return sheet.hiddenRows.has(rowNumber);
      },
      setCell(rowNumber, column, patch) {
        const key = gridKey(rowNumber, column);
        const current = sheet.grid.get(key) || defaultCell('');
        sheet.grid.set(key, Object.assign(current, clone(patch)));
      },
      getCell(rowNumber, column) {
        const key = gridKey(rowNumber, column);
        if (!sheet.grid.has(key)) sheet.grid.set(key, defaultCell(''));
        return sheet.grid.get(key);
      }
    };
    return sheet;
  }

  const target = createSheet(10, 'Intermediate', targetGrid);
  sheets.set(target.id, target);

  data.forEach((row, index) => {
    const rowNumber = 3 + index;
    row.forEach((value, offset) => {
      target.setCell(rowNumber, 3 + offset, {value: value});
    });
    // Data outside C:L must not be touched by the transfer.
    target.setCell(rowNumber, 1, {value: 'outside-left-' + rowNumber});
    target.setCell(rowNumber, 13, {value: 'outside-right-' + rowNumber});
    if (visible[index] === false) target.setHiddenRow(rowNumber, true);
  });

  function getSheetById(id) {
    return sheets.get(id);
  }

  const propertiesService = {
    getDocumentProperties() {
      return {
        getProperty(key) {
          return properties.has(key) ? properties.get(key) : null;
        },
        setProperty(key, value) {
          properties.set(key, value);
          return this;
        }
      };
    }
  };

  const sheetsService = {
    Spreadsheets: {
      batchUpdate(resource, spreadsheetId) {
        assert.equal(spreadsheetId, 'spreadsheet-test-id');
        const requests = resource.requests || [];
        const isTargetWrite = requests.some(request =>
          request.copyPaste &&
          request.copyPaste.destination.sheetId === target.id
        );
        const isStageSeed = requests.some(request =>
          request.copyPaste &&
          request.copyPaste.source.sheetId === target.id &&
          request.copyPaste.destination.sheetId !== target.id
        );
        if (isStageSeed && options.failStageCopy) {
          throw new Error('injected staging copy failure');
        }
        if (isTargetWrite && options.failTargetWrite) {
          throw new Error('injected destination write failure');
        }
        requests.forEach(request => {
          const copy = request.copyPaste;
          assert.ok(copy, 'only CopyPasteRequest is expected');
          const sourceSheet = getSheetById(copy.source.sheetId);
          const destinationSheet = getSheetById(copy.destination.sheetId);
          assert.ok(sourceSheet, 'source sheet must exist');
          assert.ok(destinationSheet, 'destination sheet must exist');

          const rowCount = copy.source.endRowIndex - copy.source.startRowIndex;
          const columnCount = copy.source.endColumnIndex - copy.source.startColumnIndex;
          assert.equal(
            rowCount,
            copy.destination.endRowIndex - copy.destination.startRowIndex
          );
          assert.equal(
            columnCount,
            copy.destination.endColumnIndex - copy.destination.startColumnIndex
          );

          for (let r = 0; r < rowCount; r++) {
            for (let c = 0; c < columnCount; c++) {
              const sourceCell = sourceSheet.getCell(
                copy.source.startRowIndex + r + 1,
                copy.source.startColumnIndex + c + 1
              );
              const destinationRow = copy.destination.startRowIndex + r + 1;
              const destinationColumn = copy.destination.startColumnIndex + c + 1;
              const destinationCell = destinationSheet.getCell(
                destinationRow,
                destinationColumn
              );
              if (copy.pasteType === 'PASTE_DATA_VALIDATION') {
                destinationCell.validation = clone(sourceCell.validation);
              } else if (copy.pasteType === 'PASTE_NORMAL') {
                const copied = clone(sourceCell);
                destinationSheet.grid.set(
                  gridKey(destinationRow, destinationColumn),
                  copied
                );
              } else {
                throw new Error('unsupported paste type: ' + copy.pasteType);
              }
            }
          }
          events.push([
            'batchCopy',
            copy.source.sheetId,
            copy.destination.sheetId,
            copy.source.startRowIndex + 1,
            copy.destination.startRowIndex + 1,
            rowCount,
            copy.pasteType
          ]);
        });

        if (isTargetWrite && options.corruptAfterTargetWrite) {
          target.getCell(3, 4).note = 'post-write corruption';
        }
        return {replies: requests.map(() => ({}))};
      }
    }
  };

  const sandbox = {
    CONFIG: {
      firstRow: 3,
      scoreCol: 5,
      dotsCol: 9,
      sortStartCol: 3,
      sortWidth: 10,
      visibleSortTransferEnabled: true
    },
    console: {log() {}, warn() {}, error() {}},
    PropertiesService: propertiesService,
    Sheets: sheetsService,
    Utilities: {
      getUuid() {
        uuidCounter++;
        return 'uuid-' + uuidCounter;
      }
    }
  };
  vm.createContext(sandbox);
  ['sorting/visibleSortPlan.gs', 'sorting/visibleSortTransfer.gs'].forEach(file => {
    vm.runInContext(fs.readFileSync(path.join(root, file), 'utf8'), sandbox, {
      filename: file
    });
  });

  sandbox.validateSortVisibilityContextFresh_ = function() {
    if (typeof options.onFreshValidation === 'function') {
      return options.onFreshValidation(target);
    }
    return {ok: options.freshOk !== false, reason: options.freshReason || 'STALE'};
  };

  return {
    project: sandbox,
    workbook,
    target,
    events,
    sheets,
    properties,
    getCell: (row, column) => target.getCell(row, column)
  };
}

function makeRow(scenario, score, dots, label) {
  const row = Array(10).fill('');
  row[0] = scenario;
  row[1] = label || '';
  row[2] = score;
  row[6] = dots;
  return row;
}

function makePlan(fixture, data, visible, sortType = 'scoreDesc') {
  const firstRow = 3;
  const lastRow = firstRow + data.length - 1;
  const mask = {
    ok: true,
    firstRow,
    lastRow,
    visibleByPhysicalPosition: visible.slice()
  };
  const context = {
    ok: true,
    firstRow,
    lastRow,
    sheetId: fixture.target.getSheetId(),
    sheetName: fixture.target.getName(),
    visibleByPhysicalPosition: visible.slice(),
    physicalHiddenByFilter: Array(data.length).fill(false),
    dataFingerprint: 'test-data-fingerprint',
    stateFingerprint: 'test-state-fingerprint',
    fingerprint: 'test-state-fingerprint'
  };
  const plan = fixture.project.buildVisibleScenarioSortPlan_(
    data,
    firstRow,
    lastRow,
    mask,
    sortType
  );
  return {plan, context, mask};
}

function currentRows(fixture, firstRow, count) {
  return Array.from({length: count}, (_, r) =>
    Array.from({length: 10}, (_, c) => fixture.getCell(firstRow + r, 3 + c).value)
  );
}

function runTransfer(fixture, built) {
  return fixture.project.transferVisibleSortPlan_(
    fixture.target,
    built.plan,
    built.context
  );
}

test('transfer module parses and is integrated after the Planner', () => {
  const source = fs.readFileSync(path.join(root, 'sorting/visibleSortTransfer.gs'), 'utf8');
  assert.doesNotThrow(() => new vm.Script(source));
  const sorting = fs.readFileSync(path.join(root, 'sorting/sorting.gs'), 'utf8');
  assert.match(sorting, /transferVisibleSortPlan_\(/);
  assert.doesNotMatch(sorting, /FULL_CELL_STAGING_ROUND_TRIP_NOT_VERIFIED/);
  const config = fs.readFileSync(path.join(root, 'config/config.gs'), 'utf8');
  assert.match(config, /visibleSortTransferEnabled:\s*false/);
});

test('ordinary visible rows are moved according to the plan', () => {
  const data = [
    makeRow('A', 10, 'A', 'low'),
    makeRow('A', 30, 'C', 'high'),
    makeRow('A', 20, 'B', 'middle')
  ];
  const fixture = makeWorkbook(data, [true, true, true]);
  const built = makePlan(fixture, data, [true, true, true]);

  const result = runTransfer(fixture, built);

  assert.equal(result.status, 'SUCCESS');
  assert.equal(result.applied, true);
  assert.deepEqual(
    currentRows(fixture, 3, 3).map(row => row[1]),
    ['high', 'middle', 'low']
  );
  assert.equal(fixture.events.filter(event => event[0] === 'deleteSheet').length, 2);
});

test('manually hidden rows remain pinned and are never source or destination rows', () => {
  const data = [
    makeRow('A', 10, 'A', 'visible-low'),
    makeRow('A', 999, 'Z', 'manual-hidden'),
    makeRow('A', 30, 'C', 'visible-high'),
    makeRow('A', 20, 'B', 'visible-middle')
  ];
  const visible = [true, false, true, true];
  const fixture = makeWorkbook(data, visible);
  const hiddenBefore = currentRows(fixture, 4, 1);
  const built = makePlan(fixture, data, visible);

  const result = runTransfer(fixture, built);

  assert.equal(result.status, 'SUCCESS');
  assert.deepEqual(
    currentRows(fixture, 3, 4).map(row => row[1]),
    ['visible-high', 'manual-hidden', 'visible-middle', 'visible-low']
  );
  assert.deepEqual(currentRows(fixture, 4, 1), hiddenBefore);
  assert.equal(fixture.target.isRowHiddenByUser(4), true);
});

test('filter-hidden rows fail closed instead of being considered transferable', () => {
  const data = [
    makeRow('A', 10, 'A', 'visible-low'),
    makeRow('A', 999, 'Z', 'filter-hidden'),
    makeRow('A', 30, 'C', 'visible-high')
  ];
  const visible = [true, false, true];
  const fixture = makeWorkbook(data, visible);
  const built = makePlan(fixture, data, visible);
  built.context.physicalHiddenByFilter[1] = true;

  const result = runTransfer(fixture, built);

  assert.equal(result.status, 'TRANSFER_NOT_VERIFIED');
  assert.equal(result.reason, 'FILTER_HIDDEN_ROWS_UNSUPPORTED_BY_RESOLVER');
  assert.equal(fixture.events.some(event => event[0] === 'insertSheet'), false);
  assert.deepEqual(currentRows(fixture, 3, 3), data.map(row => row));
});

test('alternating visible and hidden rows transfer only visible positions', () => {
  const data = [
    makeRow('A', 10, 'A', 'v1'),
    makeRow('A', 999, 'Z', 'h1'),
    makeRow('A', 40, 'D', 'v2'),
    makeRow('A', 888, 'Y', 'h2'),
    makeRow('A', 20, 'B', 'v3'),
    makeRow('A', 30, 'C', 'v4')
  ];
  const visible = [true, false, true, false, true, true];
  const fixture = makeWorkbook(data, visible);
  const hiddenBefore = [currentRows(fixture, 4, 1), currentRows(fixture, 6, 1)];
  const built = makePlan(fixture, data, visible);

  const result = runTransfer(fixture, built);

  assert.equal(result.status, 'SUCCESS');
  assert.deepEqual(
    currentRows(fixture, 3, 6).map(row => row[1]),
    ['v2', 'h1', 'v4', 'h2', 'v3', 'v1']
  );
  assert.deepEqual(currentRows(fixture, 4, 1), hiddenBefore[0]);
  assert.deepEqual(currentRows(fixture, 6, 1), hiddenBefore[1]);
});

test('multiple consecutive hidden groups stay pinned during transfer', () => {
  const data = [
    makeRow('A', 10, 'A', 'v1'),
    makeRow('A', 900, 'Z', 'h1'),
    makeRow('A', 800, 'Y', 'h2'),
    makeRow('A', 40, 'D', 'v2'),
    makeRow('A', 700, 'X', 'h3'),
    makeRow('A', 600, 'W', 'h4'),
    makeRow('A', 20, 'B', 'v3'),
    makeRow('A', 30, 'C', 'v4')
  ];
  const visible = [true, false, false, true, false, false, true, true];
  const fixture = makeWorkbook(data, visible);
  const hiddenBefore = [4, 5, 7, 8].map(row => currentRows(fixture, row, 1));
  const built = makePlan(fixture, data, visible);

  const result = runTransfer(fixture, built);

  assert.equal(result.status, 'SUCCESS');
  assert.deepEqual(
    currentRows(fixture, 3, 8).map(row => row[1]),
    ['v2', 'h1', 'h2', 'v4', 'h3', 'h4', 'v3', 'v1']
  );
  [4, 5, 7, 8].forEach((row, index) => {
    assert.deepEqual(currentRows(fixture, row, 1), hiddenBefore[index]);
  });
});

test('no visible rows is a no-op and creates no staging sheets', () => {
  const data = [makeRow('A', 10, 'A', 'h1'), makeRow('A', 20, 'B', 'h2')];
  const visible = [false, false];
  const fixture = makeWorkbook(data, visible);
  const built = makePlan(fixture, data, visible);

  const result = runTransfer(fixture, built);

  assert.equal(result.status, 'NO_OP');
  assert.equal(result.applied, false);
  assert.equal(fixture.events.some(event => event[0] === 'insertSheet'), false);
});

test('one visible row is a no-op and does not create staging sheets', () => {
  const data = [makeRow('A', 10, 'A', 'only'), makeRow('A', 20, 'B', 'hidden')];
  const visible = [true, false];
  const fixture = makeWorkbook(data, visible);
  const built = makePlan(fixture, data, visible);

  const result = runTransfer(fixture, built);

  assert.equal(result.status, 'NO_OP');
  assert.equal(fixture.events.some(event => event[0] === 'insertSheet'), false);
});

test('stale context blocks before destination write and cleans staging', () => {
  const data = [
    makeRow('A', 10, 'A', 'low'),
    makeRow('A', 30, 'C', 'high'),
    makeRow('A', 20, 'B', 'middle')
  ];
  const fixture = makeWorkbook(data, [true, true, true], {freshOk: false});
  const built = makePlan(fixture, data, [true, true, true]);
  const before = currentRows(fixture, 3, 3);

  const result = runTransfer(fixture, built);

  assert.equal(result.status, 'BLOCKED_STALE_CONTEXT');
  assert.equal(result.writeAttempted, false);
  assert.deepEqual(currentRows(fixture, 3, 3), before);
  assert.equal(fixture.events.filter(event => event[0] === 'deleteSheet').length, 2);
});

test('a visibility change after planning is caught by the fresh gate', () => {
  const data = [
    makeRow('A', 10, 'A', 'low'),
    makeRow('A', 30, 'C', 'high'),
    makeRow('A', 20, 'B', 'middle')
  ];
  const fixture = makeWorkbook(data, [true, true, true], {
    onFreshValidation(target) {
      target.setHiddenRow(4, true);
      return {ok: false, reason: 'PHYSICAL_VISIBILITY_CHANGED'};
    }
  });
  const built = makePlan(fixture, data, [true, true, true]);
  const before = currentRows(fixture, 3, 3);

  const result = runTransfer(fixture, built);

  assert.equal(result.status, 'BLOCKED_STALE_CONTEXT');
  assert.equal(result.reason, 'PHYSICAL_VISIBILITY_CHANGED');
  assert.equal(result.writeAttempted, false);
  assert.deepEqual(currentRows(fixture, 3, 3), before);
  assert.equal(fixture.target.isRowHiddenByUser(4), true);
});

test('source and destination count mismatch fails closed', () => {
  const data = [
    makeRow('A', 10, 'A', 'low'),
    makeRow('A', 30, 'C', 'high'),
    makeRow('A', 20, 'B', 'middle')
  ];
  const fixture = makeWorkbook(data, [true, true, true]);
  const built = makePlan(fixture, data, [true, true, true]);
  built.plan.blocks[0].destinations.pop();

  const result = runTransfer(fixture, built);

  assert.equal(result.status, 'TRANSFER_NOT_VERIFIED');
  assert.equal(result.reason.startsWith('SORT_PLAN_INVALID'), true);
  assert.equal(fixture.events.some(event => event[0] === 'insertSheet'), false);
});

test('staging error blocks without changing the target sheet', () => {
  const data = [
    makeRow('A', 10, 'A', 'low'),
    makeRow('A', 30, 'C', 'high'),
    makeRow('A', 20, 'B', 'middle')
  ];
  const fixture = makeWorkbook(data, [true, true, true], {failStageCopy: true});
  const built = makePlan(fixture, data, [true, true, true]);
  const before = currentRows(fixture, 3, 3);

  const result = runTransfer(fixture, built);

  assert.equal(result.status, 'BLOCKED_STAGING');
  assert.equal(result.writeAttempted, false);
  assert.deepEqual(currentRows(fixture, 3, 3), before);
  assert.equal(fixture.events.filter(event => event[0] === 'deleteSheet').length, 2);
});

test('destination write error is reported as unknown and staging is retained', () => {
  const data = [
    makeRow('A', 10, 'A', 'low'),
    makeRow('A', 30, 'C', 'high'),
    makeRow('A', 20, 'B', 'middle')
  ];
  const fixture = makeWorkbook(data, [true, true, true], {failTargetWrite: true});
  const built = makePlan(fixture, data, [true, true, true]);
  const before = currentRows(fixture, 3, 3);

  const result = runTransfer(fixture, built);

  assert.equal(result.status, 'TRANSFER_WRITE_UNKNOWN');
  assert.equal(result.writeAttempted, true);
  assert.deepEqual(currentRows(fixture, 3, 3), before);
  assert.equal(fixture.events.filter(event => event[0] === 'deleteSheet').length, 0);
  assert.equal(fixture.events.filter(event => event[0] === 'insertSheet').length, 2);
});

test('post-write verification failure does not perform an unsafe rollback', () => {
  const data = [
    makeRow('A', 10, 'A', 'low'),
    makeRow('A', 30, 'C', 'high'),
    makeRow('A', 20, 'B', 'middle')
  ];
  const fixture = makeWorkbook(data, [true, true, true], {corruptAfterTargetWrite: true});
  const built = makePlan(fixture, data, [true, true, true]);

  const result = runTransfer(fixture, built);

  assert.equal(result.status, 'TRANSFER_VERIFY_FAILED');
  assert.equal(result.writeAttempted, true);
  assert.equal(fixture.getCell(3, 4).note, 'post-write corruption');
  assert.equal(fixture.events.filter(event => event[0] === 'deleteSheet').length, 0);
});

test('values, R1C1 formulas, formatting, notes, rich text, and validation are preserved', () => {
  const data = [
    makeRow('A', 10, 'A', 'low'),
    makeRow('A', 30, 'C', 'high'),
    makeRow('A', 20, 'B', 'middle')
  ];
  const fixture = makeWorkbook(data, [true, true, true]);
  fixture.target.setCell(3, 4, {
    value: 'linked label',
    note: 'row note',
    richText: {
      text: 'linked label',
      linkUrl: 'https://example.com',
      style: {bold: true, foregroundColor: '#ff0000'}
    },
    format: {
      getBackgrounds: '#ffeecc',
      getFontColors: '#ff0000',
      getNumberFormats: '@'
    }
  });
  fixture.target.setCell(4, 5, {
    value: 30,
    formulaR1C1: '=RC[-2]*2',
    note: 'formula note'
  });
  fixture.target.setCell(5, 6, {
    value: 'validated',
    validation: {
      criteriaType: 'TEXT_CONTAINS',
      criteriaValues: ['valid'],
      allowInvalid: false,
      helpText: 'must contain valid'
    }
  });
  const built = makePlan(fixture, data, [true, true, true]);

  const result = runTransfer(fixture, built);

  assert.equal(result.status, 'SUCCESS');
  const highRow = 3;
  assert.equal(fixture.getCell(5, 4).value, 'linked label');
  assert.equal(fixture.getCell(5, 4).note, 'row note');
  assert.equal(fixture.getCell(5, 4).richText.linkUrl, 'https://example.com');
  assert.equal(fixture.getCell(5, 4).format.getBackgrounds, '#ffeecc');
  assert.equal(fixture.getCell(3, 5).formulaR1C1, '=RC[-2]*2');
  assert.equal(fixture.getCell(4, 6).validation.criteriaType, 'TEXT_CONTAINS');
});

test('cells outside C:L and hidden-row positions remain unchanged', () => {
  const data = [
    makeRow('A', 10, 'A', 'low'),
    makeRow('A', 999, 'Z', 'hidden'),
    makeRow('A', 30, 'C', 'high')
  ];
  const visible = [true, false, true];
  const fixture = makeWorkbook(data, visible);
  const leftBefore = [3, 4, 5].map(row => fixture.getCell(row, 1).value);
  const rightBefore = [3, 4, 5].map(row => fixture.getCell(row, 13).value);
  const hiddenBefore = currentRows(fixture, 4, 1);
  const built = makePlan(fixture, data, visible);

  const result = runTransfer(fixture, built);

  assert.equal(result.status, 'SUCCESS');
  assert.deepEqual([3, 4, 5].map(row => fixture.getCell(row, 1).value), leftBefore);
  assert.deepEqual([3, 4, 5].map(row => fixture.getCell(row, 13).value), rightBefore);
  assert.deepEqual(currentRows(fixture, 4, 1), hiddenBefore);
  assert.equal(fixture.target.isRowHiddenByUser(4), true);
});

test('the same operation cannot be replayed after successful transfer', () => {
  const data = [
    makeRow('A', 10, 'A', 'low'),
    makeRow('A', 30, 'C', 'high'),
    makeRow('A', 20, 'B', 'middle')
  ];
  const fixture = makeWorkbook(data, [true, true, true]);
  const built = makePlan(fixture, data, [true, true, true]);

  const first = runTransfer(fixture, built);
  const afterFirst = currentRows(fixture, 3, 3);
  const second = runTransfer(fixture, built);

  assert.equal(first.status, 'SUCCESS');
  assert.equal(second.status, 'TRANSFER_NOT_VERIFIED');
  assert.equal(second.reason, 'DUPLICATE_TRANSFER_OPERATION_BLOCKED');
  assert.deepEqual(currentRows(fixture, 3, 3), afterFirst);
  assert.equal(fixture.events.filter(event => event[0] === 'insertSheet').length, 2);
});

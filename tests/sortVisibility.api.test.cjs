const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');

function loadReader(apiResponseFactory) {
  const calls = [];
  const sandbox = {
    Date,
    Sheets: {
      Spreadsheets: {
        get(spreadsheetId, options) {
          calls.push({ spreadsheetId, options });
          return apiResponseFactory();
        }
      }
    }
  };
  vm.createContext(sandbox);
  const source = fs.readFileSync(
    path.join(root, 'sorting/sortVisibility.gs'),
    'utf8'
  );
  vm.runInContext(source, sandbox, { filename: 'sorting/sortVisibility.gs' });
  const sheet = {
    getParent: () => ({ getId: () => 'spreadsheet-id' }),
    getName: () => "O'Brien",
    getSheetId: () => 42
  };
  return {
    read: (firstRow, rowCount) =>
      sandbox.readSortPhysicalVisibilityFromSheetsApi_(sheet, firstRow, rowCount),
    calls
  };
}

function response(rowMetadata, startRow = 2, sheetId = 42) {
  return {
    sheets: [{
      properties: { sheetId },
      data: [{ startRow, rowMetadata }]
    }]
  };
}

test('batch reader maps user and filter visibility from one GridData request', () => {
  const reader = loadReader(() => response([
    { hiddenByUser: true },
    { hiddenByFilter: true },
    {},
    { hiddenByUser: false, hiddenByFilter: false }
  ]));

  const result = reader.read(3, 4);

  assert.deepEqual(Array.from(result.physicalHiddenByUser), [true, false, false, false]);
  assert.deepEqual(Array.from(result.physicalHiddenByFilter), [false, true, false, false]);
  assert.equal(result.rowsRead, 4);
  assert.equal(result.requestCount, 1);
  assert.equal(result.range, "'O''Brien'!A3:A6");
  assert.equal(reader.calls.length, 1);
  assert.equal(reader.calls[0].spreadsheetId, 'spreadsheet-id');
  assert.equal(reader.calls[0].options.ranges[0], "'O''Brien'!A3:A6");
});

test('batch reader fails closed when row metadata is missing or incomplete', () => {
  const missingArray = loadReader(() => ({
    sheets: [{ properties: { sheetId: 42 }, data: [{ startRow: 2 }] }]
  }));
  assert.throws(() => missingArray.read(3, 2), /rowMetadata is missing/);

  const shortArray = loadReader(() => response([{}]));
  assert.throws(() => shortArray.read(3, 2), /rowMetadata is missing/);

  const invalidEntry = loadReader(() => response([{}, null]));
  assert.throws(() => invalidEntry.read(3, 2), /Invalid rowMetadata object/);
});

test('batch reader fails closed on wrong sheet identity and invalid flag types', () => {
  const wrongSheet = loadReader(() => response([{}, {}], 2, 99));
  assert.throws(() => wrongSheet.read(3, 2), /matching sheetId/);

  const invalidFlag = loadReader(() => response([
    { hiddenByUser: 'true' },
    {}
  ]));
  assert.throws(() => invalidFlag.read(3, 2), /Invalid hiddenByUser type/);
});

test('batch reader rejects misaligned GridData instead of guessing row positions', () => {
  const reader = loadReader(() => response([{}, {}], 1));
  assert.throws(() => reader.read(3, 2), /missing, incomplete, or misaligned/);
});

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const source = fs.readFileSync(path.join(__dirname, '..', 'experiments', 'visibility-probe', 'visibilityProbe.gs'), 'utf8');

function loadProbe(extra = {}) {
  const context = {
    Date,
    Math,
    Number,
    String,
    Object,
    Array,
    JSON,
    ...extra
  };
  vm.createContext(context);
  vm.runInContext(source, context, { filename: 'visibilityProbe.gs' });
  return context;
}

function gridResponse(sheetId, data) {
  return {
    sheets: [{
      properties: { sheetId, title: 'Target' },
      data
    }]
  };
}

test('field mask requests only sheet identity and row visibility metadata', () => {
  const probe = loadProbe();
  assert.equal(
    probe.VISIBILITY_PROBE_FIELDS,
    'sheets(properties(sheetId,title),data(startRow,startColumn,rowMetadata(hiddenByUser,hiddenByFilter)))'
  );
});

test('maps multiple GridData segments by sheetId and zero-based startRow, independent of response order', () => {
  const probe = loadProbe();
  const response = {
    sheets: [
      { properties: { sheetId: 999, title: 'Not target' }, data: [
        { startRow: 0, startColumn: 0, rowMetadata: [{ hiddenByUser: true, hiddenByFilter: true }] }
      ] },
      { properties: { sheetId: 7, title: 'Target' }, data: [
        { startRow: 19, startColumn: 0, rowMetadata: [{ hiddenByUser: true, hiddenByFilter: false }] },
        { startRow: 9, startColumn: 0, rowMetadata: [
          { hiddenByUser: false, hiddenByFilter: false },
          { hiddenByUser: true, hiddenByFilter: false }
        ] }
      ] }
    ]
  };
  const result = probe.visibilityProbeMapGridData_(response, 7, [
    { startRow: 10, numRows: 2, endRow: 11 },
    { startRow: 20, numRows: 1, endRow: 20 }
  ]);

  assert.equal(result.sheetMatched, true);
  assert.equal(result.gridDataCount, 2);
  assert.equal(result.rowMetadataCount, 3);
  assert.equal(result.rowsByNumber[10][0].gridDataStartRow, 9);
  assert.equal(result.rowsByNumber[10][0].rowMetadataIndex, 0);
  assert.equal(result.rowsByNumber[11][0].rowMetadataIndex, 1);
  assert.equal(result.rowsByNumber[20][0].gridDataStartRow, 19);
  assert.equal(result.rowsByNumber[20][0].metadata.hiddenByUser, true);
});

test('does not use sheet response position or sheet title instead of sheetId', () => {
  const probe = loadProbe();
  const response = {
    sheets: [
      { properties: { sheetId: 12, title: 'Target' }, data: [{ startRow: 0, rowMetadata: [{ hiddenByUser: false, hiddenByFilter: false }] }] }
    ]
  };
  const result = probe.visibilityProbeMapGridData_(response, 7, [
    { startRow: 1, numRows: 1, endRow: 1 }
  ]);
  assert.equal(result.sheetMatched, false);
  assert.equal(result.errors[0].code, 'SHEET_ID_NOT_FOUND');
});

test('missing rowMetadata array is reported, not interpreted as visible rows', () => {
  const probe = loadProbe();
  const response = gridResponse(7, [{ startRow: 0, startColumn: 0 }]);
  const mapped = probe.visibilityProbeMapGridData_(response, 7, [
    { startRow: 1, numRows: 2, endRow: 2 }
  ]);
  const status = probe.visibilityProbeDetermineApiStatus_(
    mapped.rowsByNumber,
    [{ startRow: 1, numRows: 2, endRow: 2 }],
    mapped.sheetMatched,
    mapped.errors
  );
  assert.equal(status, 'INCOMPLETE');
  assert.equal(mapped.rowMetadataCount, 0);
  assert.equal(mapped.errors[0].code, 'ROW_METADATA_ARRAY_MISSING');
});

test('missing Boolean fields remain null with MISSING status, never false', () => {
  const probe = loadProbe();
  const user = probe.visibilityProbeReadApiBoolean_({ hiddenByFilter: false }, 'hiddenByUser');
  const filter = probe.visibilityProbeReadApiBoolean_({ hiddenByFilter: false }, 'hiddenByFilter');
  assert.equal(user.value, null);
  assert.equal(user.present, false);
  assert.equal(user.status, 'MISSING');
  assert.equal(filter.value, false);
  assert.equal(filter.present, true);
  assert.equal(filter.status, 'PRESENT_FALSE');
});

test('invalid Boolean types and null metadata are not coerced to false', () => {
  const probe = loadProbe();
  assert.equal(probe.visibilityProbeReadApiBoolean_({ hiddenByUser: null }, 'hiddenByUser').status, 'INVALID_TYPE');
  assert.equal(probe.visibilityProbeReadApiBoolean_(null, 'hiddenByUser').status, 'ROW_METADATA_INVALID');
});

test('partial API response maps present rows and leaves missing rows incomplete', () => {
  const probe = loadProbe();
  const response = gridResponse(7, [{
    startRow: 4,
    startColumn: 0,
    rowMetadata: [{ hiddenByUser: false, hiddenByFilter: false }]
  }]);
  const mapped = probe.visibilityProbeMapGridData_(response, 7, [
    { startRow: 5, numRows: 2, endRow: 6 }
  ]);
  const app = {
    rowsByNumber: {
      5: { hiddenByUser: false, hiddenByFilter: false, status: { hiddenByUser: 'OK', hiddenByFilter: 'OK' } },
      6: { hiddenByUser: false, hiddenByFilter: false, status: { hiddenByUser: 'OK', hiddenByFilter: 'OK' } }
    }
  };
  const api = { rowsByNumber: mapped.rowsByNumber };
  const joined = probe.visibilityProbeJoinRows_([
    { startRow: 5, numRows: 2, endRow: 6 }
  ], app, api);

  assert.equal(joined[0].status, 'COMPLETE');
  assert.equal(joined[1].status, 'INCOMPLETE');
  assert.equal(joined[1].sheetsApi.hiddenByUser, null);
  assert.equal(joined[1].comparison.hiddenByUser, 'UNAVAILABLE_API');
});

test('duplicate row metadata is ambiguous and is not silently selected', () => {
  const probe = loadProbe();
  const response = gridResponse(7, [
    { startRow: 0, rowMetadata: [{ hiddenByUser: false, hiddenByFilter: false }] },
    { startRow: 0, rowMetadata: [{ hiddenByUser: true, hiddenByFilter: false }] }
  ]);
  const mapped = probe.visibilityProbeMapGridData_(response, 7, [
    { startRow: 1, numRows: 1, endRow: 1 }
  ]);
  const app = {
    rowsByNumber: {
      1: { hiddenByUser: false, hiddenByFilter: false, status: { hiddenByUser: 'OK', hiddenByFilter: 'OK' } }
    }
  };
  const rows = probe.visibilityProbeJoinRows_([
    { startRow: 1, numRows: 1, endRow: 1 }
  ], app, { rowsByNumber: mapped.rowsByNumber });
  assert.equal(rows[0].rowMetadataStatus, 'DUPLICATE_ROW_METADATA');
  assert.equal(rows[0].sheetsApi.hiddenByUser, null);
  assert.equal(rows[0].status, 'INCOMPLETE');
});

test('unknown sheetId yields ERROR status and no mapped values', () => {
  const probe = loadProbe();
  const mapped = probe.visibilityProbeMapGridData_({ sheets: [] }, 7, [
    { startRow: 1, numRows: 1, endRow: 1 }
  ]);
  const status = probe.visibilityProbeDetermineApiStatus_(
    mapped.rowsByNumber,
    [{ startRow: 1, numRows: 1, endRow: 1 }],
    mapped.sheetMatched,
    mapped.errors
  );
  assert.equal(status, 'ERROR');
  assert.equal(mapped.errors[0].code, 'SHEET_ID_NOT_FOUND');
});

test('API exceptions are returned explicitly and do not fabricate false values', () => {
  const probe = loadProbe({
    Sheets: { Spreadsheets: { get() { throw new Error('simulated API failure'); } } }
  });
  const sheet = { getName: () => 'Target', getSheetId: () => 7 };
  const result = probe.visibilityProbeReadSheetsApi_('id', sheet, [
    { startRow: 1, numRows: 1, endRow: 1 }
  ]);
  assert.equal(result.status, 'ERROR');
  assert.equal(result.requestAttempted, true);
  assert.equal(result.requestOk, false);
  assert.equal(result.rowsByNumber[1], undefined);
  assert.equal(result.errors[0].message, 'simulated API failure');
});

test('API request uses multiple A1 ranges and exact minimal field mask in one call', () => {
  let calls = 0;
  let capturedId;
  let capturedOptions;
  const probe = loadProbe({
    Sheets: { Spreadsheets: { get(id, options) {
      calls++;
      capturedId = id;
      capturedOptions = options;
      return gridResponse(7, [
        { startRow: 0, startColumn: 0, rowMetadata: [{ hiddenByUser: false, hiddenByFilter: false }] },
        { startRow: 9, startColumn: 0, rowMetadata: [{ hiddenByUser: true, hiddenByFilter: false }] }
      ]);
    } } },
    Utilities: { newBlob(text) { return { getBytes: () => Array.from(Buffer.from(text, 'utf8')) }; } }
  });
  const sheet = { getName: () => "Target's Sheet", getSheetId: () => 7 };
  const result = probe.visibilityProbeReadSheetsApi_('spreadsheet-id', sheet, [
    { startRow: 1, numRows: 1, endRow: 1 },
    { startRow: 10, numRows: 1, endRow: 10 }
  ]);

  assert.equal(calls, 1);
  assert.equal(capturedId, 'spreadsheet-id');
  assert.deepEqual(Array.from(capturedOptions.ranges), ["'Target''s Sheet'!A1:A1", "'Target''s Sheet'!A10:A10"]);
  assert.equal(capturedOptions.fields, probe.VISIBILITY_PROBE_FIELDS);
  assert.equal(capturedOptions.includeGridData, true);
  assert.equal(result.status, 'SUCCESS');
  assert.equal(result.gridDataCount, 2);
  assert.equal(result.rowMetadataCount, 2);
  assert.ok(result.responseJsonUtf8Bytes > 0);
});

test('comparison summary counts matches, mismatches, missing values and incomplete rows separately', () => {
  const probe = loadProbe();
  const rows = [
    {
      comparison: { hiddenByUser: 'MATCH', hiddenByFilter: 'MISMATCH' },
      sheetsApiFieldStatus: { hiddenByUser: 'PRESENT_FALSE', hiddenByFilter: 'PRESENT_TRUE' },
      appsScriptFieldStatus: { hiddenByUser: 'OK', hiddenByFilter: 'OK' },
      status: 'COMPLETE'
    },
    {
      comparison: { hiddenByUser: 'UNAVAILABLE_API', hiddenByFilter: 'MATCH' },
      sheetsApiFieldStatus: { hiddenByUser: 'MISSING', hiddenByFilter: 'PRESENT_FALSE' },
      appsScriptFieldStatus: { hiddenByUser: 'OK', hiddenByFilter: 'OK' },
      status: 'INCOMPLETE'
    },
    {
      comparison: { hiddenByUser: 'UNAVAILABLE_APPS_SCRIPT', hiddenByFilter: 'UNAVAILABLE_API' },
      sheetsApiFieldStatus: { hiddenByUser: 'PRESENT_FALSE', hiddenByFilter: 'INVALID_TYPE' },
      appsScriptFieldStatus: { hiddenByUser: 'ERROR', hiddenByFilter: 'OK' },
      status: 'INCOMPLETE'
    }
  ];
  const summary = probe.visibilityProbeBuildSummary_(rows);
  assert.equal(summary.hiddenByUserMatches, 1);
  assert.equal(summary.hiddenByUserMismatches, 0);
  assert.equal(summary.hiddenByFilterMatches, 1);
  assert.equal(summary.hiddenByFilterMismatches, 1);
  assert.equal(summary.missingApiFieldCount, 2);
  assert.equal(summary.appsScriptReadErrorCount, 1);
  assert.equal(summary.missingValueCount, 3);
  assert.equal(summary.incompleteRowCount, 2);
});

test('p50 and p95 use linear interpolation on sorted samples', () => {
  const probe = loadProbe();
  assert.equal(probe.visibilityProbePercentile_([4, 1, 3, 2], 0.50), 2.5);
  assert.ok(Math.abs(probe.visibilityProbePercentile_([4, 1, 3, 2], 0.95) - 3.85) < 1e-12);
  assert.equal(probe.visibilityProbePercentile_([], 0.50), null);
  assert.equal(probe.visibilityProbePercentile_([7], 0.95), 7);
});

test('range validation supports multiple disjoint ranges and rejects overlap / out-of-sheet ranges', () => {
  const probe = loadProbe();
  const good = probe.visibilityProbeValidateRanges_([
    { startRow: 20, numRows: 2 },
    { startRow: 1, numRows: 3 }
  ], 100);
  assert.deepEqual(Array.from(good, x => [x.startRow, x.endRow]), [[1, 3], [20, 21]]);
  assert.throws(() => probe.visibilityProbeValidateRanges_([
    { startRow: 1, numRows: 5 }, { startRow: 5, numRows: 2 }
  ], 100), /must not overlap/);
  assert.throws(() => probe.visibilityProbeValidateRanges_([
    { startRow: 99, numRows: 3 }
  ], 100), /exceeds sheet maxRows/);
});

test('Apps Script errors are preserved separately while the other visibility method still runs', () => {
  const probe = loadProbe();
  const calls = [];
  const sheet = {
    isRowHiddenByUser() { calls.push('user'); throw new Error('user method error'); },
    isRowHiddenByFilter() { calls.push('filter'); return false; }
  };
  const result = probe.visibilityProbeReadAppsScript_(sheet, [
    { startRow: 1, numRows: 1, endRow: 1 }
  ]);
  assert.deepEqual(calls, ['user', 'filter']);
  assert.equal(result.rowsByNumber[1].hiddenByUser, null);
  assert.equal(result.rowsByNumber[1].hiddenByFilter, false);
  assert.equal(result.success, false);
  assert.equal(result.errors[0].message, 'user method error');
});

test('source contains no spreadsheet mutation, visibility mutation, sort, or flush operations', () => {
  assert.doesNotMatch(source, /\.setValue\s*\(/);
  assert.doesNotMatch(source, /\.setValues\s*\(/);
  assert.doesNotMatch(source, /\.hideRows\s*\(/);
  assert.doesNotMatch(source, /\.showRows\s*\(/);
  assert.doesNotMatch(source, /\.createFilter\s*\(/);
  assert.doesNotMatch(source, /\.removeFilter\s*\(/);
  assert.doesNotMatch(source, /\.sort\s*\(\s*\[/); // No Range.sort([...]) fallback.
  assert.doesNotMatch(source, /SpreadsheetApp\.flush\s*\(/);
  assert.doesNotMatch(source, /\.getRange\([^\n]*\)\.getValues\s*\(/);
});

test('benchmark alternates methods, performs warmups, and enforces at least 10 samples per method', () => {
  const events = [];
  let apiCalls = 0;
  const sheet = {
    getName: () => 'Target',
    getSheetId: () => 7,
    getMaxRows: () => 100,
    isRowHiddenByUser() { events.push('APPS_SCRIPT'); return false; },
    isRowHiddenByFilter() { return false; }
  };
  const spreadsheet = {
    getId: () => 'spreadsheet-id',
    getActiveSheet: () => sheet
  };
  const probe = loadProbe({
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    Sheets: { Spreadsheets: { get() {
      apiCalls++;
      events.push('SHEETS_API');
      return gridResponse(7, [{ startRow: 2, startColumn: 0, rowMetadata: [
        { hiddenByUser: false, hiddenByFilter: false }
      ] }]);
    } } },
    Utilities: { newBlob(text) { return { getBytes: () => Array.from(Buffer.from(text, 'utf8')) }; } },
    Logger: { log() {} }
  });
  probe.VISIBILITY_PROBE_CONFIG.ranges = [{ startRow: 3, numRows: 1 }];
  probe.VISIBILITY_PROBE_CONFIG.benchmark = { warmupRuns: 1, sampleRuns: 10 };

  const result = probe.runVisibilityProbeBenchmark();
  assert.equal(result.status, 'SUCCESS');
  assert.equal(result.performance.appsScript.successfulSamples, 10);
  assert.equal(result.performance.sheetsApi.successfulRequests, 10);
  assert.equal(result.performance.appsScript.methodCallsAttemptedIncludingWarmups, 22);
  assert.equal(result.performance.sheetsApi.requestsAttemptedIncludingWarmups, 11);
  assert.equal(apiCalls, 11);
  assert.equal(result.performance.order.length, 11);
  assert.equal(result.performance.coldStart.appsScript.phase, 'FIRST_METHOD_INVOCATION_IN_THIS_SCRIPT_RUN');
  assert.equal(result.performance.coldStart.appsScript.methodCallsAttempted, 2);
  assert.equal(result.performance.coldStart.sheetsApi.phase, 'FIRST_METHOD_INVOCATION_IN_THIS_SCRIPT_RUN');
  assert.equal(result.performance.coldStart.sheetsApi.requestAttempted, true);
  assert.equal(result.performance.coldStart.sheetsApi.status, 'SUCCESS');
  assert.ok(result.performance.coldStart.appsScript.elapsedMs >= 0);
  assert.ok(result.performance.coldStart.sheetsApi.totalMs >= 0);
  assert.equal(result.performance.appsScript.methodCallsPerRun, 2);
  assert.equal(result.performance.sheetsApi.requestsPerRun, 1);
  assert.equal(result.performance.order[0].phase, 'WARMUP');
  assert.deepEqual(Array.from(result.performance.order[0].order), ['APPS_SCRIPT', 'SHEETS_API']);
  assert.deepEqual(Array.from(result.performance.order[1].order), ['SHEETS_API', 'APPS_SCRIPT']);
  assert.ok(result.performance.appsScript.p50Ms !== null);
  assert.ok(result.performance.appsScript.p95Ms !== null);
  assert.ok(result.performance.sheetsApi.p50TotalMs !== null);
  assert.ok(events.length > 0);
});

test('benchmark rejects fewer than 10 measured samples', () => {
  const sheet = { getName: () => 'Target', getSheetId: () => 7, getMaxRows: () => 100 };
  const spreadsheet = { getId: () => 'id', getActiveSheet: () => sheet };
  const probe = loadProbe({
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    Logger: { log() {} }
  });
  probe.VISIBILITY_PROBE_CONFIG.ranges = [{ startRow: 1, numRows: 1 }];
  probe.VISIBILITY_PROBE_CONFIG.benchmark = { warmupRuns: 1, sampleRuns: 9 };
  const result = probe.runVisibilityProbeBenchmark();
  assert.equal(result.status, 'ERROR');
  assert.match(result.errors[0].message, /sampleRuns must be an integer >= 10/);
});

test('main probe preserves Apps Script values when the Sheets API fails', () => {
  const logs = [];
  const sheet = {
    getName: () => 'Test sheet',
    getSheetId: () => 77,
    getMaxRows: () => 50,
    isRowHiddenByUser: () => true,
    isRowHiddenByFilter: () => false
  };
  const spreadsheet = {
    getId: () => 'test-spreadsheet-id',
    getActiveSheet: () => sheet
  };
  const probe = loadProbe({
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    Sheets: { Spreadsheets: { get() { throw new Error('API disabled for test'); } } },
    Logger: { log(value) { logs.push(value); } }
  });
  probe.VISIBILITY_PROBE_CONFIG.ranges = [{ startRow: 3, numRows: 1 }];
  probe.VISIBILITY_PROBE_CONFIG.scenarioLabel = 'api-error-preserves-appscript';

  const report = probe.runVisibilityProbe();
  assert.equal(report.status, 'ERROR');
  assert.equal(report.apiStatus, 'ERROR');
  assert.equal(report.spreadsheetId, 'test-spreadsheet-id');
  assert.equal(report.sheetId, 77);
  assert.equal(report.rows.length, 1);
  assert.equal(report.performance.totalElapsedMs >= 0, true);
  assert.equal(report.comparisonSummary.totalRowCount, 1);
  assert.equal(report.comparisonSummary.incompleteRowCount, 1);
  assert.equal(report.rows[0].appsScript.hiddenByUser, true);
  assert.equal(report.rows[0].appsScript.hiddenByFilter, false);
  assert.equal(report.rows[0].sheetsApi.hiddenByUser, null);
  assert.equal(report.rows[0].sheetsApi.hiddenByFilter, null);
  assert.equal(report.rows[0].comparison.hiddenByUser, 'UNAVAILABLE_API');
  assert.equal(report.rows[0].comparison.hiddenByFilter, 'UNAVAILABLE_API');
  assert.match(report.errors[0].message, /API disabled for test/);
  assert.equal(logs.length, 1);
  assert.equal(JSON.parse(logs[0]).scenarioLabel, 'api-error-preserves-appscript');
});

test('benchmark excludes incomplete API responses from successful timing percentiles', () => {
  let apiCall = 0;
  const sheet = {
    getName: () => 'Target',
    getSheetId: () => 7,
    getMaxRows: () => 100,
    isRowHiddenByUser: () => false,
    isRowHiddenByFilter: () => false
  };
  const spreadsheet = { getId: () => 'id', getActiveSheet: () => sheet };
  const probe = loadProbe({
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    Sheets: { Spreadsheets: { get() {
      apiCall++;
      if (apiCall === 2) {
        return gridResponse(7, [{ startRow: 2, startColumn: 0, rowMetadata: [{}] }]);
      }
      return gridResponse(7, [{ startRow: 2, startColumn: 0, rowMetadata: [
        { hiddenByUser: false, hiddenByFilter: false }
      ] }]);
    } } },
    Utilities: { newBlob(text) { return { getBytes: () => Array.from(Buffer.from(text, 'utf8')) }; } },
    Logger: { log() {} }
  });
  probe.VISIBILITY_PROBE_CONFIG.ranges = [{ startRow: 3, numRows: 1 }];
  probe.VISIBILITY_PROBE_CONFIG.benchmark = { warmupRuns: 1, sampleRuns: 10 };
  const result = probe.runVisibilityProbeBenchmark();
  assert.equal(result.performance.sheetsApi.incompleteResponses, 1);
  assert.equal(result.performance.sheetsApi.successfulRequests, 9);
  assert.equal(result.performance.sheetsApi.p50TotalMs !== null, true);
  assert.equal(result.performance.coldStart.appsScript.phase, 'FIRST_METHOD_INVOCATION_IN_THIS_SCRIPT_RUN');
  assert.equal(result.performance.coldStart.sheetsApi.phase, 'FIRST_METHOD_INVOCATION_IN_THIS_SCRIPT_RUN');
  assert.equal(result.performance.pairedProbe.measuredRounds, 10);
  assert.equal(result.performance.pairedProbe.completePairedRounds, 9);
  assert.equal(result.performance.pairedProbe.comparisonTotals.completeRowCount, 9);
  assert.equal(result.performance.pairedProbe.comparisonTotals.matchedRowsBothFields, 9);
  assert.equal(result.performance.pairedProbe.equivalentPairedRounds, 9);
  assert.equal(result.performance.pairedProbe.comparisonTotals.incompleteRowCount, 1);
  assert.equal(result.performance.pairedProbe.rounds[0].incompleteRows, 1);
  assert.ok(result.performance.pairedProbe.p50TotalMs >= 0);
  assert.ok(result.performance.totalElapsedMs >= 0);
});


test('benchmark records per-round row matches and mismatches in JSON output', () => {
  let apiCall = 0;
  const sheet = {
    getName: () => 'Target',
    getSheetId: () => 7,
    getMaxRows: () => 100,
    isRowHiddenByUser: () => false,
    isRowHiddenByFilter: () => false
  };
  const spreadsheet = { getId: () => 'id', getActiveSheet: () => sheet };
  const probe = loadProbe({
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    Sheets: { Spreadsheets: { get() {
      apiCall++;
      return gridResponse(7, [{ startRow: 2, startColumn: 0, rowMetadata: [
        { hiddenByUser: apiCall === 2, hiddenByFilter: false }
      ] }]);
    } } },
    Utilities: { newBlob(text) { return { getBytes: () => Array.from(Buffer.from(text, 'utf8')) }; } },
    Logger: { log() {} }
  });
  probe.VISIBILITY_PROBE_CONFIG.ranges = [{ startRow: 3, numRows: 1 }];
  probe.VISIBILITY_PROBE_CONFIG.benchmark = { warmupRuns: 1, sampleRuns: 10 };
  probe.VISIBILITY_PROBE_CONFIG.scenarioLabel = 'benchmark-mismatch';

  const result = probe.runVisibilityProbeBenchmark();
  const serialized = JSON.stringify(result);
  const parsed = JSON.parse(serialized);

  assert.equal(parsed.scenarioLabel, 'benchmark-mismatch');
  assert.equal(parsed.performance.pairedProbe.measuredRounds, 10);
  assert.equal(parsed.performance.pairedProbe.comparisonTotals.mismatchedRows, 1);
  assert.equal(parsed.performance.pairedProbe.comparisonTotals.hiddenByUserMismatches, 1);
  assert.equal(parsed.performance.pairedProbe.comparisonTotals.hiddenByFilterMismatches, 0);
  assert.equal(parsed.performance.pairedProbe.comparisonTotals.matchedRowsBothFields, 9);
  assert.equal(parsed.performance.pairedProbe.rounds[0].mismatchRows, 1);
  assert.equal(typeof parsed.generatedAt, 'string');
});

test('benchmark requires at least one warmup so the first invocation stays outside percentiles', () => {
  const sheet = { getName: () => 'Target', getSheetId: () => 7, getMaxRows: () => 100 };
  const spreadsheet = { getId: () => 'id', getActiveSheet: () => sheet };
  const probe = loadProbe({
    SpreadsheetApp: { getActiveSpreadsheet: () => spreadsheet },
    Logger: { log() {} }
  });
  probe.VISIBILITY_PROBE_CONFIG.ranges = [{ startRow: 1, numRows: 1 }];
  probe.VISIBILITY_PROBE_CONFIG.benchmark = { warmupRuns: 0, sampleRuns: 10 };
  const result = probe.runVisibilityProbeBenchmark();
  assert.equal(result.status, 'ERROR');
  assert.match(result.errors[0].message, /warmupRuns must be an integer >= 1/);
});


test('invalid or missing GridData.startRow is reported without guessing physical row numbers', () => {
  const probe = loadProbe();
  const mapped = probe.visibilityProbeMapGridData_(gridResponse(7, [
    { rowMetadata: [{ hiddenByUser: false, hiddenByFilter: false }] },
    { startRow: -1, rowMetadata: [{ hiddenByUser: false, hiddenByFilter: false }] },
    { startRow: '2', rowMetadata: [{ hiddenByUser: false, hiddenByFilter: false }] }
  ]), 7, [
    { startRow: 1, numRows: 1, endRow: 1 }
  ]);
  assert.equal(mapped.rowMetadataCount, 0);
  assert.equal(mapped.errors.length, 3);
  assert.ok(mapped.errors.every(error => error.code === 'GRIDDATA_STARTROW_INVALID'));
  assert.equal(mapped.rowsByNumber[1], undefined);
});

test('missing hiddenByFilter is incomplete independently from hiddenByUser', () => {
  const probe = loadProbe();
  const user = probe.visibilityProbeReadApiBoolean_({ hiddenByUser: true }, 'hiddenByUser');
  const filter = probe.visibilityProbeReadApiBoolean_({ hiddenByUser: true }, 'hiddenByFilter');
  assert.equal(user.status, 'PRESENT_TRUE');
  assert.equal(user.value, true);
  assert.equal(filter.status, 'MISSING');
  assert.equal(filter.value, null);
});

test('metadata returned outside requested row intervals is flagged as incomplete', () => {
  const probe = loadProbe();
  const mapped = probe.visibilityProbeMapGridData_(gridResponse(7, [{
    startRow: 0,
    rowMetadata: [
      { hiddenByUser: false, hiddenByFilter: false },
      { hiddenByUser: false, hiddenByFilter: false }
    ]
  }]), 7, [{ startRow: 2, numRows: 1, endRow: 2 }]);
  assert.equal(mapped.errors.length, 1);
  assert.equal(mapped.errors[0].code, 'METADATA_OUTSIDE_REQUESTED_RANGE');
  assert.equal(mapped.rowsByNumber[2][0].rowNumber, 2);
});
/**
 * READ-ONLY GOOGLE SHEETS ROW VISIBILITY PROBE
 *
 * Standalone Apps Script experiment. It reads row visibility through:
 *   A) Sheet.isRowHiddenByUser() / Sheet.isRowHiddenByFilter()
 *   B) Sheets.Spreadsheets.get() GridData.rowMetadata
 *
 * It never changes cells, row visibility, filters, sheet state, or project state.
 * The only output is structured JSON in the Apps Script execution log.
 *
 * IMPORTANT:
 * - Use only in a disposable test spreadsheet / separate Apps Script project.
 * - SpreadsheetApp.openById() requires the broad spreadsheets OAuth scope.
 *   The probe code itself performs no writes, but that scope technically permits
 *   writes. See README for the lower-scope container-bound alternative.
 */

var VISIBILITY_PROBE_CONFIG = {
  // Required for a standalone script. Leave blank only when this project is
  // container-bound to the disposable test spreadsheet.
  spreadsheetId: '',

  // Select by sheetId when supplied. Otherwise use sheetName. When both are
  // blank, the active sheet is used (container-bound project only).
  sheetId: null,
  sheetName: '',

  // Change scenarioLabel before every manual scenario and save each JSON log.
  scenarioLabel: '01-visible-baseline',

  // One or more NON-OVERLAPPING row ranges. Default is intentionally small.
  // For the current Manual Sort baseline, set this to [{ startRow: 3, numRows: 129 }]
  // (or the exact physical row interval chosen in the test sheet).
  ranges: [
    { startRow: 3, numRows: 20 }
  ],

  benchmark: {
    warmupRuns: 2,
    sampleRuns: 12 // Minimum enforced by the probe: 10.
  }
};

var VISIBILITY_PROBE_VERSION = '1.5.0';
// Keep the field mask minimal and request only the row metadata needed by the probe.
// A fields mask already enables GridData in spreadsheets.get; includeGridData is redundant.
var VISIBILITY_PROBE_FIELDS =
  'sheets(properties(sheetId),data(startRow,rowMetadata(hiddenByUser,hiddenByFilter)))';

/**
 * Single-request API-only probe for the 129-row VisibilityTest fixture.
 *
 * Reads rows 2–130 (129 rows) using one Sheets API request. It deliberately
 * does not call SpreadsheetApp row-visibility methods, so it can measure the
 * API path without the slow per-row Apps Script calls.
 *
 * Requires the container-bound test project to target the VisibilityTest sheet.
 * Read-only: no sheet or spreadsheet state is changed.
 */
function runVisibilityProbeApi129() {
  var started = Date.now();
  var report = {
    probeVersion: VISIBILITY_PROBE_VERSION,
    generatedAt: new Date().toISOString(),
    mode: 'API_ONLY_129_ROWS',
    status: 'NOT_STARTED',
    spreadsheetId: null,
    sheetId: null,
    sheetName: null,
    scenarioLabel: String(VISIBILITY_PROBE_CONFIG.scenarioLabel || ''),
    requestedRange: null,
    requestedRanges: [],
    apiStatus: 'NOT_STARTED',
    apiGridDataCount: 0,
    apiRowMetadataCount: 0,
    expectedRowCount: 129,
    actualMappedRowCount: 0,
    hiddenByUserTrueCount: 0,
    hiddenByFilterTrueCount: 0,
    falseOrDefaultFalseCount: 0,
    performance: null,
    errors: []
  };

  try {
    var target = visibilityProbeResolveTarget_();
    if (target.sheet.getName() !== 'VisibilityTest') {
      throw new Error(
        'API-129 probe is restricted to the VisibilityTest sheet; actual sheet: ' +
        target.sheet.getName()
      );
    }

    // Fixture layout: header in row 1, data in rows 2–130.
    var ranges = visibilityProbeValidateRanges_(
      [{ startRow: 2, numRows: 129 }],
      target.sheet.getMaxRows()
    );
    visibilityProbeAddTarget_(report, target, ranges);

    var apiRead = visibilityProbeReadSheetsApi_(
      target.spreadsheet.getId(),
      target.sheet,
      ranges
    );

    report.apiStatus = apiRead.status;
    report.apiGridDataCount = apiRead.gridDataCount;
    report.apiRowMetadataCount = apiRead.rowMetadataCount;
    report.actualMappedRowCount = Object.keys(apiRead.rowsByNumber).length;

    if (apiRead.status === 'SUCCESS') {
      var rowNumbers = visibilityProbeExpandRows_(ranges);
      rowNumbers.forEach(function(rowNumber) {
        var candidates = apiRead.rowsByNumber[rowNumber] || [];
        if (candidates.length !== 1) {
          throw new Error('Expected exactly one rowMetadata record for row ' + rowNumber + '.');
        }
        var metadata = candidates[0].metadata;
        var user = visibilityProbeReadApiBoolean_(metadata, 'hiddenByUser');
        var filter = visibilityProbeReadApiBoolean_(metadata, 'hiddenByFilter');
        if (!visibilityProbeIsResolvedApiBoolean_(user) ||
            !visibilityProbeIsResolvedApiBoolean_(filter)) {
          throw new Error('Unresolved visibility metadata for row ' + rowNumber + '.');
        }
        if (user.value) report.hiddenByUserTrueCount++;
        if (filter.value) report.hiddenByFilterTrueCount++;
        if (!user.value && !filter.value) report.falseOrDefaultFalseCount++;
      });
    }

    report.errors = apiRead.errors.slice();
    report.performance = {
      mode: 'SINGLE_API_REQUEST_NO_APPS_SCRIPT_VISIBILITY_CALLS',
      totalElapsedMs: Date.now() - started,
      requestMs: apiRead.requestMs,
      parseMs: apiRead.parseMs,
      totalApiMs: apiRead.requestMs === null || apiRead.parseMs === null
        ? null
        : apiRead.requestMs + apiRead.parseMs,
      requestsAttempted: apiRead.requestAttempted ? 1 : 0,
      responseJsonUtf8Bytes: apiRead.responseJsonUtf8Bytes
    };

    report.status = apiRead.status === 'SUCCESS' &&
      report.actualMappedRowCount === 129 &&
      report.apiRowMetadataCount === 129
        ? 'SUCCESS'
        : (apiRead.status === 'ERROR' ? 'ERROR' : 'INCOMPLETE');
  } catch (error) {
    report.status = 'ERROR';
    report.apiStatus = report.apiStatus === 'NOT_STARTED' ? 'ERROR' : report.apiStatus;
    report.errors.push(visibilityProbeError_('API_ONLY_129', error));
    report.performance = {
      mode: 'SINGLE_API_REQUEST_NO_APPS_SCRIPT_VISIBILITY_CALLS',
      totalElapsedMs: Date.now() - started,
      requestsAttempted: report.performance ? report.performance.requestsAttempted : 0
    };
  }

  return visibilityProbeLogReport_(report);
}

/**
 * Small API-only diagnostic. Run this before the full probe when the Sheets API
 * request fails or returns no GridData. Its single log line stays compact and
 * exposes the actual API exception or the returned GridData structure.
 */
function runVisibilityProbeApiDiagnostics() {
  var report = {
    probeVersion: VISIBILITY_PROBE_VERSION,
    generatedAt: new Date().toISOString(),
    mode: 'API_DIAGNOSTICS',
    status: 'NOT_STARTED',
    spreadsheetId: null,
    sheetId: null,
    sheetName: null,
    requestedRanges: [],
    requestAttempted: false,
    requestOk: false,
    requestMs: null,
    apiStatus: 'NOT_STARTED',
    gridDataCount: 0,
    rowMetadataCount: 0,
    responseSummary: null,
    errors: []
  };

  try {
    var target = visibilityProbeResolveTarget_();
    var ranges = visibilityProbeValidateRanges_(
      VISIBILITY_PROBE_CONFIG.ranges,
      target.sheet.getMaxRows()
    );
    report.spreadsheetId = target.spreadsheet.getId();
    report.sheetId = target.sheet.getSheetId();
    report.sheetName = target.sheet.getName();
    report.requestedRanges = visibilityProbeBuildA1Ranges_(report.sheetName, ranges);

    var apiRead = visibilityProbeReadSheetsApi_(
      report.spreadsheetId,
      target.sheet,
      ranges
    );
    report.requestAttempted = apiRead.requestAttempted;
    report.requestOk = apiRead.requestOk;
    report.requestMs = apiRead.requestMs;
    report.apiStatus = apiRead.status;
    report.gridDataCount = apiRead.gridDataCount;
    report.rowMetadataCount = apiRead.rowMetadataCount;
    report.responseSummary = apiRead.responseSummary;
    report.errors = apiRead.errors;
    report.status = apiRead.status;
  } catch (error) {
    report.status = 'ERROR';
    report.apiStatus = 'ERROR';
    report.errors.push(visibilityProbeError_('API_DIAGNOSTICS', error));
  }

  return visibilityProbeLogDiagnostic_(report);
}

function visibilityProbeLogDiagnostic_(report) {
  var json = JSON.stringify(report);
  if (typeof Logger !== 'undefined' && Logger.log) {
    Logger.log('VISIBILITY_PROBE_API_DIAGNOSTICS ' + json);
  }
  return report;
}

/** Main one-shot comparison. Logs and returns one JSON report. */
function runVisibilityProbe() {
  var runStarted = Date.now();
  var report = visibilityProbeCreateBaseReport_('SINGLE_RUN');
  var target;
  var normalizedRanges;

  try {
    target = visibilityProbeResolveTarget_();
    normalizedRanges = visibilityProbeValidateRanges_(
      VISIBILITY_PROBE_CONFIG.ranges,
      target.sheet.getMaxRows()
    );
  } catch (error) {
    report.status = 'ERROR';
    report.apiStatus = 'ERROR';
    report.performance = {
      mode: 'SINGLE_RUN',
      totalElapsedMs: Date.now() - runStarted
    };
    report.errors.push(visibilityProbeError_('SETUP', error));
    return visibilityProbeLogReport_(report);
  }

  visibilityProbeAddTarget_(report, target, normalizedRanges);

  // One-shot diagnostic deliberately reads Apps Script first, then Sheets API.
  // Use the alternating benchmark for timing comparisons.
  var appScriptRead = visibilityProbeReadAppsScript_(
    target.sheet,
    normalizedRanges
  );
  var apiRead = visibilityProbeReadSheetsApi_(
    target.spreadsheet.getId(),
    target.sheet,
    normalizedRanges
  );

  report.apiStatus = apiRead.status;
  report.apiGridDataCount = apiRead.gridDataCount;
  report.apiRowMetadataCount = apiRead.rowMetadataCount;
  report.rows = visibilityProbeJoinRows_(
    normalizedRanges,
    appScriptRead,
    apiRead
  );
  report.comparisonSummary = visibilityProbeBuildSummary_(report.rows);
  report.performance = {
    mode: 'SINGLE_RUN_NOT_A_BENCHMARK',
    totalElapsedMs: Date.now() - runStarted,
    appsScriptReadMs: appScriptRead.elapsedMs,
    sheetsApiRequestMs: apiRead.requestMs,
    sheetsApiParseMs: apiRead.parseMs,
    sheetsApiTotalMs: apiRead.requestMs === null || apiRead.parseMs === null
      ? null
      : apiRead.requestMs + apiRead.parseMs,
    appsScriptMethodCallsAttempted: appScriptRead.methodCallsAttempted,
    sheetsApiRequestsAttempted: apiRead.requestAttempted ? 1 : 0,
    responseJsonUtf8Bytes: apiRead.responseJsonUtf8Bytes
  };

  report.errors = report.errors
    .concat(appScriptRead.errors)
    .concat(apiRead.errors);
  report.status = report.apiStatus === 'ERROR'
    ? 'ERROR'
    : (report.apiStatus === 'INCOMPLETE' ||
       !appScriptRead.success ||
       report.comparisonSummary.incompleteRowCount > 0
      ? 'INCOMPLETE'
      : 'SUCCESS');

  return visibilityProbeLogReport_(report);
}

/**
 * Repeated performance comparison. Executes at least 10 measured rounds per
 * method, after warm-up, alternating method order on every round.
 * This mode performs read requests only; it does not create test scenarios.
 */
function runVisibilityProbeBenchmark() {
  var benchmarkStarted = Date.now();
  var report = visibilityProbeCreateBaseReport_('BENCHMARK');
  var target;
  var normalizedRanges;
  var warmupRuns;
  var sampleRuns;

  try {
    target = visibilityProbeResolveTarget_();
    normalizedRanges = visibilityProbeValidateRanges_(
      VISIBILITY_PROBE_CONFIG.ranges,
      target.sheet.getMaxRows()
    );
    warmupRuns = Number(VISIBILITY_PROBE_CONFIG.benchmark.warmupRuns);
    sampleRuns = Number(VISIBILITY_PROBE_CONFIG.benchmark.sampleRuns);
    // At least one warmup is required so the first invocation of each method
    // is reported as coldStart and is not mixed into measured percentiles.
    if (!Number.isInteger(warmupRuns) || warmupRuns < 1) {
      throw new Error('benchmark.warmupRuns must be an integer >= 1.');
    }
    if (!Number.isInteger(sampleRuns) || sampleRuns < 10) {
      throw new Error('benchmark.sampleRuns must be an integer >= 10.');
    }
  } catch (error) {
    report.status = 'ERROR';
    report.apiStatus = 'ERROR';
    report.performance = {
      mode: 'ALTERNATING_REPEATED_BENCHMARK',
      totalElapsedMs: Date.now() - benchmarkStarted
    };
    report.errors.push(visibilityProbeError_('SETUP', error));
    return visibilityProbeLogReport_(report);
  }

  visibilityProbeAddTarget_(report, target, normalizedRanges);

  var appScriptSamples = [];
  var apiRequestSamples = [];
  var apiParseSamples = [];
  var apiTotalSamples = [];
  var pairTotalSamples = [];
  var responseSizes = [];
  var coldStart = { appsScript: null, sheetsApi: null };
  var lastApiGridDataCount = 0;
  var lastApiRowMetadataCount = 0;
  var appScriptFailedRuns = 0;
  var apiFailedRequests = 0;
  var apiIncompleteResponses = 0;
  var apiErrorResponses = 0;
  var appScriptFailedWarmups = 0;
  var apiFailedWarmups = 0;
  var apiIncompleteWarmups = 0;
  var apiErrorWarmups = 0;
  var appsScriptMethodCallsAttempted = 0;
  var sheetsApiRequestsAttempted = 0;
  var order = [];
  var benchmarkErrors = [];
  var measuredRoundDetails = [];
  var comparisonTotals = visibilityProbeCreateEmptySummary_();
  var completePairedRounds = 0;
  var equivalentPairedRounds = 0;
  var rowCount = visibilityProbeExpandRows_(normalizedRanges).length;
  var totalRounds = warmupRuns + sampleRuns;

  for (var round = 0; round < totalRounds; round++) {
    var isWarmup = round < warmupRuns;
    var firstMethod = round % 2 === 0 ? 'APPS_SCRIPT' : 'SHEETS_API';
    var secondMethod = firstMethod === 'APPS_SCRIPT' ? 'SHEETS_API' : 'APPS_SCRIPT';
    var methods = [firstMethod, secondMethod];
    var pairStarted = Date.now();
    var appResultForRound = null;
    var apiResultForRound = null;

    order.push({
      round: round + 1,
      phase: isWarmup ? 'WARMUP' : 'MEASURED',
      order: methods.slice()
    });

    for (var methodIndex = 0; methodIndex < methods.length; methodIndex++) {
      var method = methods[methodIndex];

      if (method === 'APPS_SCRIPT') {
        var appResult = visibilityProbeReadAppsScript_(target.sheet, normalizedRanges);
        appResultForRound = appResult;
        appsScriptMethodCallsAttempted += appResult.methodCallsAttempted;

        if (coldStart.appsScript === null) {
          coldStart.appsScript = {
            phase: 'FIRST_METHOD_INVOCATION_IN_THIS_SCRIPT_RUN',
            elapsedMs: appResult.elapsedMs,
            methodCallsAttempted: appResult.methodCallsAttempted,
            status: appResult.success ? 'SUCCESS' : 'ERROR',
            errors: appResult.errors.slice(0, 5)
          };
        }

        if (!appResult.success) {
          if (isWarmup) appScriptFailedWarmups++;
          else appScriptFailedRuns++;
          if (benchmarkErrors.length < 20) {
            benchmarkErrors.push({
              stage: 'APPS_SCRIPT_BENCHMARK',
              phase: isWarmup ? 'WARMUP' : 'MEASURED',
              round: round + 1,
              errors: appResult.errors.slice(0, 5)
            });
          }
        } else if (!isWarmup) {
          appScriptSamples.push(appResult.elapsedMs);
        }
      } else {
        var apiResult = visibilityProbeReadSheetsApi_(
          target.spreadsheet.getId(),
          target.sheet,
          normalizedRanges
        );
        apiResultForRound = apiResult;
        if (apiResult.requestAttempted) sheetsApiRequestsAttempted++;
        if (apiResult.requestOk) {
          lastApiGridDataCount = apiResult.gridDataCount;
          lastApiRowMetadataCount = apiResult.rowMetadataCount;
        }

        if (coldStart.sheetsApi === null) {
          coldStart.sheetsApi = {
            phase: 'FIRST_METHOD_INVOCATION_IN_THIS_SCRIPT_RUN',
            requestAttempted: apiResult.requestAttempted,
            requestOk: apiResult.requestOk,
            requestMs: apiResult.requestMs,
            parseMs: apiResult.parseMs,
            totalMs: apiResult.requestMs === null || apiResult.parseMs === null
              ? null
              : apiResult.requestMs + apiResult.parseMs,
            status: apiResult.status,
            gridDataCount: apiResult.gridDataCount,
            rowMetadataCount: apiResult.rowMetadataCount,
            responseJsonUtf8Bytes: apiResult.responseJsonUtf8Bytes,
            errors: apiResult.errors.slice(0, 5)
          };
        }

        if (!apiResult.requestOk) {
          if (isWarmup) apiFailedWarmups++;
          else apiFailedRequests++;
          if (benchmarkErrors.length < 20) {
            benchmarkErrors.push({
              stage: 'SHEETS_API_BENCHMARK',
              phase: isWarmup ? 'WARMUP' : 'MEASURED',
              round: round + 1,
              errors: apiResult.errors.slice(0, 5)
            });
          }
        } else {
          if (apiResult.status === 'ERROR') {
            if (isWarmup) apiErrorWarmups++;
            else apiErrorResponses++;
          } else if (apiResult.status === 'INCOMPLETE') {
            if (isWarmup) apiIncompleteWarmups++;
            else apiIncompleteResponses++;
          }
          if (!isWarmup && apiResult.status === 'SUCCESS') {
            // Only complete, type-valid metadata responses are timing samples.
            apiRequestSamples.push(apiResult.requestMs);
            apiParseSamples.push(apiResult.parseMs);
            apiTotalSamples.push(apiResult.requestMs + apiResult.parseMs);
            if (apiResult.responseJsonUtf8Bytes !== null) {
              responseSizes.push(apiResult.responseJsonUtf8Bytes);
            }
          }
        }
      }
    }

    var pairTotalMs = Date.now() - pairStarted;
    var pairedRows = visibilityProbeJoinRows_(
      normalizedRanges,
      appResultForRound || visibilityProbeEmptyAppsScriptRead_(),
      apiResultForRound || visibilityProbeEmptySheetsApiRead_()
    );
    var pairedSummary = visibilityProbeBuildSummary_(pairedRows);

    if (!isWarmup) {
      pairTotalSamples.push(pairTotalMs);
      visibilityProbeAccumulateSummary_(comparisonTotals, pairedSummary);
      if (
        appResultForRound && appResultForRound.success &&
        apiResultForRound && apiResultForRound.status === 'SUCCESS' &&
        pairedSummary.completeRowCount === rowCount
      ) {
        completePairedRounds++;
        if (pairedSummary.matchedRowsBothFields === rowCount && pairedSummary.mismatchedRows === 0) {
          equivalentPairedRounds++;
        }
      }
      measuredRoundDetails.push({
        round: round + 1,
        order: methods.slice(),
        appsScript: {
          elapsedMs: appResultForRound ? appResultForRound.elapsedMs : null,
          status: appResultForRound && appResultForRound.success ? 'SUCCESS' : 'ERROR'
        },
        sheetsApi: {
          requestMs: apiResultForRound ? apiResultForRound.requestMs : null,
          parseMs: apiResultForRound ? apiResultForRound.parseMs : null,
          status: apiResultForRound ? apiResultForRound.status : 'NOT_RUN'
        },
        pairedTotalMs: pairTotalMs,
        comparedRows: pairedSummary.completeRowCount,
        matchedRowsBothFields: pairedSummary.matchedRowsBothFields,
        mismatchRows: pairedSummary.mismatchedRows,
        hiddenByUserMismatches: pairedSummary.hiddenByUserMismatches,
        hiddenByFilterMismatches: pairedSummary.hiddenByFilterMismatches,
        incompleteRows: pairedSummary.incompleteRowCount
      });
    }
  }

  report.performance = {
    mode: 'ALTERNATING_REPEATED_BENCHMARK',
    totalElapsedMs: Date.now() - benchmarkStarted,
    coldStart: coldStart,
    warmupRunsPerMethod: warmupRuns,
    requestedMeasuredRunsPerMethod: sampleRuns,
    failedWarmupRuns: {
      appsScript: appScriptFailedWarmups,
      sheetsApiRequests: apiFailedWarmups,
      sheetsApiIncompleteResponses: apiIncompleteWarmups,
      sheetsApiErrorResponses: apiErrorWarmups
    },
    orderPolicy: 'Alternating; APPS_SCRIPT first on even rounds, SHEETS_API first on odd rounds.',
    samplePolicy: 'First invocation of each method is reported in coldStart; warmups are excluded from all percentiles; only complete, successful Sheets API responses enter API percentiles.',
    percentileMethod: 'Linear interpolation over sorted successful samples (Type 7 / PERCENTILE.INC).',
    appsScript: {
      method: 'Two Sheet row-visibility methods per row',
      successfulSamples: appScriptSamples.length,
      failedSamples: appScriptFailedRuns,
      p50Ms: visibilityProbePercentile_(appScriptSamples, 0.50),
      p95Ms: visibilityProbePercentile_(appScriptSamples, 0.95),
      methodCallsAttemptedIncludingWarmups: appsScriptMethodCallsAttempted,
      methodCallsPerRun: rowCount * 2
    },
    sheetsApi: {
      method: 'One Sheets.Spreadsheets.get() per run',
      successfulRequests: apiRequestSamples.length,
      failedRequests: apiFailedRequests,
      incompleteResponses: apiIncompleteResponses,
      errorResponses: apiErrorResponses,
      p50RequestMs: visibilityProbePercentile_(apiRequestSamples, 0.50),
      p95RequestMs: visibilityProbePercentile_(apiRequestSamples, 0.95),
      p50ParseMs: visibilityProbePercentile_(apiParseSamples, 0.50),
      p95ParseMs: visibilityProbePercentile_(apiParseSamples, 0.95),
      p50TotalMs: visibilityProbePercentile_(apiTotalSamples, 0.50),
      p95TotalMs: visibilityProbePercentile_(apiTotalSamples, 0.95),
      requestsAttemptedIncludingWarmups: sheetsApiRequestsAttempted,
      requestsPerRun: 1,
      responseJsonUtf8BytesMedian: visibilityProbePercentile_(responseSizes, 0.50)
    },
    pairedProbe: {
      measuredRounds: measuredRoundDetails.length,
      completePairedRounds: completePairedRounds,
      equivalentPairedRounds: equivalentPairedRounds,
      p50TotalMs: visibilityProbePercentile_(pairTotalSamples, 0.50),
      p95TotalMs: visibilityProbePercentile_(pairTotalSamples, 0.95),
      comparisonTotals: comparisonTotals,
      rounds: measuredRoundDetails
    },
    rowCount: rowCount,
    order: order,
    errors: benchmarkErrors
  };

  report.apiGridDataCount = lastApiGridDataCount;
  report.apiRowMetadataCount = lastApiRowMetadataCount;
  report.apiStatus = apiFailedRequests > 0 || apiErrorResponses > 0
    ? 'ERROR'
    : (apiIncompleteResponses > 0 ? 'INCOMPLETE' : 'SUCCESS');
  report.status = report.apiStatus === 'ERROR'
    ? 'ERROR'
    : (report.apiStatus === 'INCOMPLETE' ||
       appScriptFailedRuns > 0 ||
       comparisonTotals.incompleteRowCount > 0
      ? 'INCOMPLETE'
      : 'SUCCESS');
  report.errors = benchmarkErrors;

  return visibilityProbeLogReport_(report);
}

function visibilityProbeEmptyAppsScriptRead_() {
  return {
    rowsByNumber: Object.create(null),
    errors: [],
    success: false,
    methodCallsAttempted: 0,
    elapsedMs: null
  };
}

function visibilityProbeEmptySheetsApiRead_() {
  return {
    rowsByNumber: Object.create(null),
    errors: [],
    status: 'ERROR',
    requestAttempted: false,
    requestOk: false,
    requestMs: null,
    parseMs: null,
    responseJsonUtf8Bytes: null,
    responseSummary: null,
    gridDataCount: 0,
    rowMetadataCount: 0
  };
}

function visibilityProbeCreateBaseReport_(mode) {
  return {
    probeVersion: VISIBILITY_PROBE_VERSION,
    generatedAt: new Date().toISOString(),
    mode: mode,
    status: 'NOT_STARTED',
    spreadsheetId: null,
    sheetId: null,
    sheetName: null,
    scenarioLabel: String(VISIBILITY_PROBE_CONFIG.scenarioLabel || ''),
    requestedRange: null,
    requestedRanges: [],
    apiStatus: 'NOT_STARTED',
    apiGridDataCount: 0,
    apiRowMetadataCount: 0,
    rows: [],
    comparisonSummary: null,
    performance: null,
    errors: []
  };
}

function visibilityProbeResolveTarget_() {
  if (typeof SpreadsheetApp === 'undefined') {
    throw new Error('SpreadsheetApp is unavailable in this execution context.');
  }

  var spreadsheetId = String(VISIBILITY_PROBE_CONFIG.spreadsheetId || '').trim();
  var spreadsheet = spreadsheetId
    ? SpreadsheetApp.openById(spreadsheetId)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      'No active spreadsheet. In a standalone project set spreadsheetId in VISIBILITY_PROBE_CONFIG.'
    );
  }

  var configuredSheetId = VISIBILITY_PROBE_CONFIG.sheetId;
  var configuredSheetName = String(VISIBILITY_PROBE_CONFIG.sheetName || '').trim();
  var sheet = null;

  if (configuredSheetId !== null && configuredSheetId !== undefined && configuredSheetId !== '') {
    var wantedId = Number(configuredSheetId);
    if (!Number.isInteger(wantedId)) {
      throw new Error('sheetId must be an integer when supplied.');
    }
    var sheets = spreadsheet.getSheets();
    for (var i = 0; i < sheets.length; i++) {
      if (sheets[i].getSheetId() === wantedId) {
        sheet = sheets[i];
        break;
      }
    }
    if (!sheet) throw new Error('Configured sheetId was not found: ' + wantedId);
  } else if (configuredSheetName) {
    sheet = spreadsheet.getSheetByName(configuredSheetName);
    if (!sheet) throw new Error('Configured sheetName was not found: ' + configuredSheetName);
  } else {
    if (spreadsheetId) {
      throw new Error('For a standalone project set sheetId or sheetName in VISIBILITY_PROBE_CONFIG.');
    }
    sheet = spreadsheet.getActiveSheet();
  }

  if (!sheet) throw new Error('Target sheet could not be resolved.');
  return { spreadsheet: spreadsheet, sheet: sheet };
}

function visibilityProbeValidateRanges_(ranges, maxRows) {
  if (!Array.isArray(ranges) || ranges.length === 0) {
    throw new Error('ranges must contain at least one { startRow, numRows } item.');
  }
  if (!Number.isInteger(maxRows) || maxRows < 1) {
    throw new Error('The target sheet returned an invalid maxRows value.');
  }

  var normalized = ranges.map(function(range, index) {
    if (!range || !Number.isInteger(range.startRow) || !Number.isInteger(range.numRows) ||
        range.startRow < 1 || range.numRows < 1) {
      throw new Error('Invalid range at index ' + index + ': startRow and numRows must be positive integers.');
    }
    var endRow = range.startRow + range.numRows - 1;
    if (endRow > maxRows) {
      throw new Error('Requested row ' + endRow + ' exceeds sheet maxRows=' + maxRows + '.');
    }
    return { startRow: range.startRow, numRows: range.numRows, endRow: endRow };
  }).sort(function(a, b) { return a.startRow - b.startRow; });

  for (var i = 1; i < normalized.length; i++) {
    if (normalized[i].startRow <= normalized[i - 1].endRow) {
      throw new Error('Configured ranges must not overlap.');
    }
  }
  return normalized;
}

function visibilityProbeAddTarget_(report, target, ranges) {
  report.spreadsheetId = target.spreadsheet.getId();
  report.sheetId = target.sheet.getSheetId();
  report.sheetName = target.sheet.getName();
  report.requestedRanges = visibilityProbeBuildA1Ranges_(report.sheetName, ranges);
  report.requestedRange = report.requestedRanges.length === 1
    ? report.requestedRanges[0]
    : null;
}

function visibilityProbeBuildA1Ranges_(sheetName, ranges) {
  var quotedSheetName = "'" + String(sheetName).replace(/'/g, "''") + "'";
  return ranges.map(function(range) {
    return quotedSheetName + '!A' + range.startRow + ':A' + range.endRow;
  });
}

function visibilityProbeExpandRows_(ranges) {
  var rows = [];
  ranges.forEach(function(range) {
    for (var row = range.startRow; row <= range.endRow; row++) rows.push(row);
  });
  return rows;
}

function visibilityProbeReadAppsScript_(sheet, ranges) {
  var started = Date.now();
  var rowsByNumber = Object.create(null);
  var errors = [];
  var methodCallsAttempted = 0;
  var rowNumbers = visibilityProbeExpandRows_(ranges);

  rowNumbers.forEach(function(rowNumber) {
    var record = {
      rowNumber: rowNumber,
      hiddenByUser: null,
      hiddenByFilter: null,
      status: { hiddenByUser: 'NOT_READ', hiddenByFilter: 'NOT_READ' }
    };

    methodCallsAttempted++;
    try {
      var hiddenByUserResult = sheet.isRowHiddenByUser(rowNumber);
      record.status.hiddenByUser = typeof hiddenByUserResult === 'boolean' ? 'OK' : 'INVALID_TYPE';
      record.hiddenByUser = record.status.hiddenByUser === 'OK' ? hiddenByUserResult : null;
      if (record.status.hiddenByUser !== 'OK') {
        errors.push({ stage: 'APPS_SCRIPT', rowNumber: rowNumber, field: 'hiddenByUser', message: 'Method did not return a boolean.' });
      }
    } catch (error) {
      record.status.hiddenByUser = 'ERROR';
      errors.push(visibilityProbeError_('APPS_SCRIPT_HIDDEN_BY_USER', error, { rowNumber: rowNumber }));
    }

    methodCallsAttempted++;
    try {
      var hiddenByFilterResult = sheet.isRowHiddenByFilter(rowNumber);
      record.status.hiddenByFilter = typeof hiddenByFilterResult === 'boolean' ? 'OK' : 'INVALID_TYPE';
      record.hiddenByFilter = record.status.hiddenByFilter === 'OK' ? hiddenByFilterResult : null;
      if (record.status.hiddenByFilter !== 'OK') {
        errors.push({ stage: 'APPS_SCRIPT', rowNumber: rowNumber, field: 'hiddenByFilter', message: 'Method did not return a boolean.' });
      }
    } catch (error) {
      record.status.hiddenByFilter = 'ERROR';
      errors.push(visibilityProbeError_('APPS_SCRIPT_HIDDEN_BY_FILTER', error, { rowNumber: rowNumber }));
    }

    rowsByNumber[rowNumber] = record;
  });

  var allSuccess = rowNumbers.every(function(rowNumber) {
    var record = rowsByNumber[rowNumber];
    return record && record.status.hiddenByUser === 'OK' && record.status.hiddenByFilter === 'OK';
  });

  return {
    rowsByNumber: rowsByNumber,
    elapsedMs: Date.now() - started,
    methodCallsAttempted: methodCallsAttempted,
    errors: errors,
    success: allSuccess
  };
}

function visibilityProbeReadSheetsApi_(spreadsheetId, sheet, ranges) {
  var result = {
    status: 'ERROR',
    requestAttempted: false,
    requestOk: false,
    requestMs: null,
    parseMs: null,
    responseJsonUtf8Bytes: null,
    responseSummary: null,
    gridDataCount: 0,
    rowMetadataCount: 0,
    rowsByNumber: Object.create(null),
    errors: []
  };

  try {
    if (typeof Sheets === 'undefined' || !Sheets.Spreadsheets ||
        typeof Sheets.Spreadsheets.get !== 'function') {
      throw new Error('Advanced Sheets Service is unavailable. Enable the Sheets API service in this test project.');
    }

    var a1Ranges = visibilityProbeBuildA1Ranges_(sheet.getName(), ranges);
    var requestStarted = Date.now();
    result.requestAttempted = true;
    var response = Sheets.Spreadsheets.get(spreadsheetId, {
      ranges: a1Ranges,
      fields: VISIBILITY_PROBE_FIELDS
    });
    result.requestMs = Date.now() - requestStarted;
    result.requestOk = true;
    result.responseSummary = visibilityProbeSummarizeApiResponse_(
      response,
      sheet.getSheetId()
    );

    var parseStarted = Date.now();
    var mapped = visibilityProbeMapGridData_(response, sheet.getSheetId(), ranges);
    result.parseMs = Date.now() - parseStarted;
    result.gridDataCount = mapped.gridDataCount;
    result.rowMetadataCount = mapped.rowMetadataCount;
    result.rowsByNumber = mapped.rowsByNumber;
    result.errors = result.errors.concat(mapped.errors);

    try {
      if (typeof Utilities !== 'undefined' && Utilities.newBlob) {
        result.responseJsonUtf8Bytes = Utilities.newBlob(
          JSON.stringify(response),
          'application/json'
        ).getBytes().length;
      }
    } catch (sizeError) {
      result.errors.push(visibilityProbeError_('RESPONSE_SIZE', sizeError));
    }

    result.status = visibilityProbeDetermineApiStatus_(
      result.rowsByNumber,
      ranges,
      mapped.sheetMatched,
      mapped.errors
    );
  } catch (error) {
    result.errors.push(visibilityProbeError_('SHEETS_API', error));
    result.status = 'ERROR';
    if (result.requestAttempted && result.requestMs === null) {
      // The request failed before a response was returned. Timing is retained
      // as null rather than pretending the API returned in zero milliseconds.
      result.requestMs = null;
    }
  }

  return result;
}

/** Returns only structural details; never includes cell values. */
function visibilityProbeSummarizeApiResponse_(response, wantedSheetId) {
  var summary = {
    responseKeys: response && typeof response === 'object'
      ? Object.keys(response)
      : [],
    sheetsIsArray: !!(response && Array.isArray(response.sheets)),
    sheetsCount: response && Array.isArray(response.sheets)
      ? response.sheets.length
      : null,
    matchingSheetCount: 0,
    sheets: []
  };

  if (!response || !Array.isArray(response.sheets)) return summary;

  summary.sheets = response.sheets.map(function(item) {
    var properties = item && item.properties ? item.properties : null;
    var data = item && Array.isArray(item.data) ? item.data : null;
    var sheetId = properties && properties.sheetId !== undefined
      ? Number(properties.sheetId)
      : null;
    if (sheetId === Number(wantedSheetId)) summary.matchingSheetCount++;

    return {
      sheetId: sheetId,
      title: properties && properties.title ? properties.title : null,
      dataIsArray: !!data,
      dataCount: data ? data.length : null,
      gridData: data ? data.map(function(grid) {
        var metadata = grid && Array.isArray(grid.rowMetadata)
          ? grid.rowMetadata
          : null;
        var first = metadata && metadata.length ? metadata[0] : null;
        return {
          startRow: grid && grid.startRow !== undefined ? grid.startRow : null,
          startColumn: grid && grid.startColumn !== undefined ? grid.startColumn : null,
          rowMetadataIsArray: !!metadata,
          rowMetadataCount: metadata ? metadata.length : null,
          firstMetadata: first ? {
            hiddenByUser: typeof first.hiddenByUser === 'boolean'
              ? first.hiddenByUser
              : null,
            hiddenByUserType: first.hiddenByUser === null
              ? 'null'
              : typeof first.hiddenByUser,
            hiddenByFilter: typeof first.hiddenByFilter === 'boolean'
              ? first.hiddenByFilter
              : null,
            hiddenByFilterType: first.hiddenByFilter === null
              ? 'null'
              : typeof first.hiddenByFilter
          } : null
        };
      }) : []
    };
  });

  return summary;
}

/** Maps GridData by sheetId and startRow, never by response array position. */
function visibilityProbeMapGridData_(response, wantedSheetId, requestedRanges) {
  var result = {
    sheetMatched: false,
    gridDataCount: 0,
    rowMetadataCount: 0,
    rowsByNumber: Object.create(null),
    errors: []
  };

  if (!response || !Array.isArray(response.sheets)) {
    result.errors.push({ stage: 'API_RESPONSE', code: 'SHEETS_ARRAY_MISSING', message: 'Response has no sheets array.' });
    return result;
  }

  var matchingSheets = response.sheets.filter(function(item) {
    return item && item.properties && Number(item.properties.sheetId) === Number(wantedSheetId);
  });
  if (matchingSheets.length !== 1) {
    result.errors.push({
      stage: 'API_RESPONSE',
      code: matchingSheets.length === 0 ? 'SHEET_ID_NOT_FOUND' : 'DUPLICATE_SHEET_ID',
      message: matchingSheets.length === 0
        ? 'The API response did not contain the requested sheetId.'
        : 'The API response contained more than one matching sheetId.'
    });
    return result;
  }
  result.sheetMatched = true;

  var sheetData = matchingSheets[0].data;
  if (!Array.isArray(sheetData)) {
    result.errors.push({ stage: 'API_RESPONSE', code: 'GRIDDATA_ARRAY_MISSING', message: 'Matching sheet has no data array.' });
    return result;
  }

  var requestedRowSet = Object.create(null);
  visibilityProbeExpandRows_(requestedRanges).forEach(function(rowNumber) {
    requestedRowSet[rowNumber] = true;
  });

  result.gridDataCount = sheetData.length;
  sheetData.forEach(function(gridData, gridDataIndex) {
    if (!gridData || !Number.isInteger(gridData.startRow) || gridData.startRow < 0) {
      result.errors.push({
        stage: 'API_RESPONSE',
        code: 'GRIDDATA_STARTROW_INVALID',
        gridDataIndex: gridDataIndex,
        message: 'GridData.startRow is missing or invalid; rows cannot be mapped safely.'
      });
      return;
    }

    if (!Array.isArray(gridData.rowMetadata)) {
      result.errors.push({
        stage: 'API_RESPONSE',
        code: 'ROW_METADATA_ARRAY_MISSING',
        gridDataIndex: gridDataIndex,
        startRow: gridData.startRow,
        message: 'GridData.rowMetadata is absent or is not an array.'
      });
      return;
    }

    result.rowMetadataCount += gridData.rowMetadata.length;
    gridData.rowMetadata.forEach(function(metadata, metadataIndex) {
      var rowNumber = gridData.startRow + metadataIndex + 1;
      if (!requestedRowSet[rowNumber]) {
        result.errors.push({
          stage: 'API_RESPONSE',
          code: 'METADATA_OUTSIDE_REQUESTED_RANGE',
          rowNumber: rowNumber,
          gridDataIndex: gridDataIndex,
          rowMetadataIndex: metadataIndex,
          message: 'Returned row metadata does not belong to any requested row range.'
        });
        return;
      }

      var candidate = {
        rowNumber: rowNumber,
        gridDataIndex: gridDataIndex,
        gridDataStartRow: gridData.startRow,
        gridDataStartColumn: Number.isInteger(gridData.startColumn) ? gridData.startColumn : null,
        rowMetadataIndex: metadataIndex,
        rowMetadataPresent: true,
        metadata: metadata
      };
      if (!result.rowsByNumber[rowNumber]) result.rowsByNumber[rowNumber] = [];
      result.rowsByNumber[rowNumber].push(candidate);
    });
  });

  return result;
}

function visibilityProbeReadApiBoolean_(metadata, field) {
  if (!metadata || typeof metadata !== 'object' || Array.isArray(metadata)) {
    return { value: null, present: false, status: 'ROW_METADATA_INVALID' };
  }
  if (!Object.prototype.hasOwnProperty.call(metadata, field)) {
    // Sheets API uses protobuf JSON serialization: default false Boolean
    // values may be omitted. A present rowMetadata object plus an explicitly
    // requested Boolean field therefore represents false when the property
    // is absent. Missing rowMetadata objects are still rejected separately.
    return { value: false, present: false, status: 'OMITTED_DEFAULT_FALSE' };
  }
  if (typeof metadata[field] !== 'boolean') {
    return {
      value: null,
      present: true,
      status: 'INVALID_TYPE',
      rawType: metadata[field] === null ? 'null' : typeof metadata[field]
    };
  }
  return {
    value: metadata[field],
    present: true,
    status: metadata[field] ? 'PRESENT_TRUE' : 'PRESENT_FALSE'
  };
}

function visibilityProbeIsResolvedApiBoolean_(field) {
  return field && (
    field.status === 'PRESENT_TRUE' ||
    field.status === 'PRESENT_FALSE' ||
    field.status === 'OMITTED_DEFAULT_FALSE'
  );
}

function visibilityProbeDetermineApiStatus_(rowsByNumber, ranges, sheetMatched, errors) {
  if (!sheetMatched) return 'ERROR';
  if (errors.some(function(error) {
    return error.code === 'GRIDDATA_STARTROW_INVALID' ||
      error.code === 'ROW_METADATA_ARRAY_MISSING' ||
      error.code === 'METADATA_OUTSIDE_REQUESTED_RANGE' ||
      error.code === 'GRIDDATA_ARRAY_MISSING';
  })) return 'INCOMPLETE';

  var rows = visibilityProbeExpandRows_(ranges);
  for (var i = 0; i < rows.length; i++) {
    var candidates = rowsByNumber[rows[i]] || [];
    if (candidates.length !== 1) return 'INCOMPLETE';
    var user = visibilityProbeReadApiBoolean_(candidates[0].metadata, 'hiddenByUser');
    var filter = visibilityProbeReadApiBoolean_(candidates[0].metadata, 'hiddenByFilter');
    if (!visibilityProbeIsResolvedApiBoolean_(user)) return 'INCOMPLETE';
    if (!visibilityProbeIsResolvedApiBoolean_(filter)) return 'INCOMPLETE';
  }
  return 'SUCCESS';
}

function visibilityProbeJoinRows_(ranges, appScriptRead, apiRead) {
  return visibilityProbeExpandRows_(ranges).map(function(rowNumber) {
    var app = appScriptRead.rowsByNumber[rowNumber] || null;
    var candidates = apiRead.rowsByNumber[rowNumber] || [];
    var uniqueCandidate = candidates.length === 1 ? candidates[0] : null;
    var apiMetadata = uniqueCandidate ? uniqueCandidate.metadata : null;
    var userField = visibilityProbeReadApiBoolean_(apiMetadata, 'hiddenByUser');
    var filterField = visibilityProbeReadApiBoolean_(apiMetadata, 'hiddenByFilter');
    var apiRowStatus = candidates.length === 0
      ? 'MISSING_ROW_METADATA'
      : (candidates.length > 1 ? 'DUPLICATE_ROW_METADATA' :
        (!apiMetadata || typeof apiMetadata !== 'object' ? 'ROW_METADATA_INVALID' : 'PRESENT'));

    var userComparison = visibilityProbeCompare_(
      app ? app.hiddenByUser : null,
      app ? app.status.hiddenByUser : 'MISSING',
      userField
    );
    var filterComparison = visibilityProbeCompare_(
      app ? app.hiddenByFilter : null,
      app ? app.status.hiddenByFilter : 'MISSING',
      filterField
    );

    var complete = apiRowStatus === 'PRESENT' &&
      visibilityProbeIsResolvedApiBoolean_(userField) &&
      visibilityProbeIsResolvedApiBoolean_(filterField) &&
      app && app.status.hiddenByUser === 'OK' &&
      app.status.hiddenByFilter === 'OK';

    return {
      rowNumber: rowNumber,
      rowMetadataIndex: uniqueCandidate ? uniqueCandidate.rowMetadataIndex : null,
      gridDataIndex: uniqueCandidate ? uniqueCandidate.gridDataIndex : null,
      gridDataStartRow: uniqueCandidate ? uniqueCandidate.gridDataStartRow : null,
      appsScript: {
        hiddenByUser: app ? app.hiddenByUser : null,
        hiddenByFilter: app ? app.hiddenByFilter : null
      },
      appsScriptFieldStatus: app ? app.status : {
        hiddenByUser: 'MISSING', hiddenByFilter: 'MISSING'
      },
      sheetsApi: {
        hiddenByUser: userField.value,
        hiddenByFilter: filterField.value
      },
      sheetsApiFieldStatus: {
        hiddenByUser: userField.status,
        hiddenByFilter: filterField.status
      },
      comparison: {
        hiddenByUser: userComparison,
        hiddenByFilter: filterComparison
      },
      rowMetadataStatus: apiRowStatus,
      status: complete ? 'COMPLETE' : 'INCOMPLETE'
    };
  });
}

function visibilityProbeCompare_(appsValue, appsStatus, apiField) {
  if (appsStatus !== 'OK' || typeof appsValue !== 'boolean') return 'UNAVAILABLE_APPS_SCRIPT';
  if (!visibilityProbeIsResolvedApiBoolean_(apiField)) return 'UNAVAILABLE_API';
  return appsValue === apiField.value ? 'MATCH' : 'MISMATCH';
}

function visibilityProbeCreateEmptySummary_() {
  return {
    totalRowCount: 0,
    completeRowCount: 0,
    matchedRowsBothFields: 0,
    mismatchedRows: 0,
    hiddenByUserMatches: 0,
    hiddenByUserMismatches: 0,
    hiddenByFilterMatches: 0,
    hiddenByFilterMismatches: 0,
    missingValueCount: 0,
    missingApiFieldCount: 0,
    appsScriptReadErrorCount: 0,
    incompleteRowCount: 0
  };
}

function visibilityProbeBuildSummary_(rows) {
  var summary = visibilityProbeCreateEmptySummary_();
  summary.totalRowCount = rows.length;

  rows.forEach(function(row) {
    var userComparison = row.comparison.hiddenByUser;
    var filterComparison = row.comparison.hiddenByFilter;
    var complete = row.status === 'COMPLETE';

    if (userComparison === 'MATCH') summary.hiddenByUserMatches++;
    if (userComparison === 'MISMATCH') summary.hiddenByUserMismatches++;
    if (filterComparison === 'MATCH') summary.hiddenByFilterMatches++;
    if (filterComparison === 'MISMATCH') summary.hiddenByFilterMismatches++;

    if (complete) {
      summary.completeRowCount++;
      if (userComparison === 'MATCH' && filterComparison === 'MATCH') {
        summary.matchedRowsBothFields++;
      }
      if (userComparison === 'MISMATCH' || filterComparison === 'MISMATCH') {
        summary.mismatchedRows++;
      }
    }

    ['hiddenByUser', 'hiddenByFilter'].forEach(function(field) {
      var apiStatus = row.sheetsApiFieldStatus[field];
      var appStatus = row.appsScriptFieldStatus[field];
      if (
        apiStatus !== 'PRESENT_TRUE' &&
        apiStatus !== 'PRESENT_FALSE' &&
        apiStatus !== 'OMITTED_DEFAULT_FALSE'
      ) {
        summary.missingApiFieldCount++;
        summary.missingValueCount++;
      }
      if (appStatus !== 'OK') {
        summary.appsScriptReadErrorCount++;
        summary.missingValueCount++;
      }
    });
    if (!complete) summary.incompleteRowCount++;
  });
  return summary;
}

function visibilityProbeAccumulateSummary_(total, sample) {
  [
    'totalRowCount',
    'completeRowCount',
    'matchedRowsBothFields',
    'mismatchedRows',
    'hiddenByUserMatches',
    'hiddenByUserMismatches',
    'hiddenByFilterMatches',
    'hiddenByFilterMismatches',
    'missingValueCount',
    'missingApiFieldCount',
    'appsScriptReadErrorCount',
    'incompleteRowCount'
  ].forEach(function(key) {
    total[key] += Number(sample[key] || 0);
  });
}

function visibilityProbePercentile_(values, percentile) {
  if (!Array.isArray(values) || values.length === 0) return null;
  var sorted = values.slice().sort(function(a, b) { return a - b; });
  if (sorted.length === 1) return sorted[0];
  var position = (sorted.length - 1) * percentile;
  var lowerIndex = Math.floor(position);
  var upperIndex = Math.ceil(position);
  if (lowerIndex === upperIndex) return sorted[lowerIndex];
  var weight = position - lowerIndex;
  return sorted[lowerIndex] * (1 - weight) + sorted[upperIndex] * weight;
}

function visibilityProbeError_(stage, error, extra) {
  var result = {
    stage: stage,
    message: error && error.message ? error.message : String(error)
  };
  if (error && error.name) result.name = error.name;
  if (extra && typeof extra === 'object') {
    Object.keys(extra).forEach(function(key) { result[key] = extra[key]; });
  }
  return result;
}

function visibilityProbeLogReport_(report) {
  var json = JSON.stringify(report);
  if (typeof Logger !== 'undefined' && Logger.log) Logger.log(json);
  return report;
}

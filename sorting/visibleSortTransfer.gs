/**
 * SAFE VISIBLE SORT TRANSFER
 *
 * Transfers only the visible physical destination slots described by the
 * existing Visible Sort Planner. The original C:L range is staged before any
 * destination write. A second hidden staging sheet simulates the final
 * placement so formulas can be checked in R1C1 notation before commit.
 *
 * Requires the Apps Script Advanced Sheets service (Sheets) to be enabled.
 * CONFIG.visibleSortTransferEnabled deliberately defaults to false until the
 * live disposable-workbook validation described in docs/visible-sort-transfer.md
 * has passed.
 */

const VISIBLE_SORT_TRANSFER_LEDGER_KEY_ =
  'VISIBLE_SORT_TRANSFER_LEDGER_V1';
const VISIBLE_SORT_TRANSFER_LEDGER_LIMIT_ = 40;
const VISIBLE_SORT_TRANSFER_FORMAT_METHODS_ = [
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

/**
 * Transfers a validated planner permutation without touching hidden rows.
 *
 * Return statuses:
 * - SUCCESS: transfer and full C:L cell snapshot verification succeeded.
 * - NO_OP: plan has no changed destinations.
 * - TRANSFER_NOT_VERIFIED: feature flag/service/preconditions are unavailable.
 * - BLOCKED_STALE_CONTEXT / BLOCKED_STAGING: no target cells were written.
 * - TRANSFER_WRITE_UNKNOWN / TRANSFER_VERIFY_FAILED: write was attempted;
 *   staging sheets are intentionally retained for manual recovery.
 */
function transferVisibleSortPlan_(
  sheet,
  plan,
  visibilityContext,
  perfTrace = null
) {
  const startedAt = Date.now();
  let spreadsheet = null;
  let stageSource = null;
  let stageOutput = null;
  let operationId = null;
  let writeAttempted = false;
  let ledgerProperties = null;

  function finish(status, applied, reason, extra = {}) {
    const result = Object.assign({
      status: status,
      applied: applied,
      reason: reason || null,
      writeAttempted: writeAttempted
    }, extra);

    sortPerfStep_(perfTrace, 'Visible transfer: Total', startedAt, {
      status: status,
      applied: applied,
      reason: reason || null,
      rows: extra.rowsMoved || 0,
      writeAttempted: writeAttempted
    });
    return result;
  }

  function cleanupStages() {
    if (!spreadsheet || writeAttempted) return;
    [stageOutput, stageSource].forEach(stage => {
      if (!stage) return;
      try {
        spreadsheet.deleteSheet(stage);
      } catch (cleanupError) {
        console.warn('[VISIBLE SORT TRANSFER] staging cleanup failed', JSON.stringify({
          sheetId: stage.getSheetId ? stage.getSheetId() : null,
          message: cleanupError && cleanupError.message
            ? cleanupError.message
            : String(cleanupError)
        }));
      }
    });
  }

  try {
    if (!sheet || !plan || !visibilityContext) {
      return finish('TRANSFER_NOT_VERIFIED', false, 'TRANSFER_INPUT_MISSING');
    }

    if (CONFIG.visibleSortTransferEnabled !== true) {
      return finish(
        'TRANSFER_NOT_VERIFIED',
        false,
        'TRANSFER_FEATURE_DISABLED_UNTIL_LIVE_VALIDATION'
      );
    }

    if (
      typeof Sheets === 'undefined' ||
      !Sheets ||
      !Sheets.Spreadsheets ||
      typeof Sheets.Spreadsheets.batchUpdate !== 'function'
    ) {
      return finish(
        'TRANSFER_NOT_VERIFIED',
        false,
        'ADVANCED_SHEETS_SERVICE_NOT_ENABLED'
      );
    }

    const planValidationStarted = Date.now();
    try {
      assertVisibleSortPlan_(plan, visibilityContext);
    } catch (planError) {
      return finish(
        'TRANSFER_NOT_VERIFIED',
        false,
        'SORT_PLAN_INVALID: ' + visibleSortTransferErrorMessage_(planError)
      );
    }
    if (
      plan.firstRow !== visibilityContext.firstRow ||
      plan.lastRow !== visibilityContext.lastRow
    ) {
      return finish('TRANSFER_NOT_VERIFIED', false, 'SORT_PLAN_CONTEXT_RANGE_MISMATCH');
    }
    if (
      Array.isArray(visibilityContext.physicalHiddenByFilter) &&
      visibilityContext.physicalHiddenByFilter.some(Boolean)
    ) {
      return finish(
        'TRANSFER_NOT_VERIFIED',
        false,
        'FILTER_HIDDEN_ROWS_UNSUPPORTED_BY_RESOLVER'
      );
    }

    const moves = flattenVisibleSortMoves_(plan);
    if (moves.length === 0) {
      return finish('NO_OP', false, 'NO_VISIBLE_ROWS_TO_MOVE');
    }

    const sourceRows = new Set();
    const destinationRows = new Set();
    for (const move of moves) {
      if (
        !Number.isInteger(move.sourceRow) ||
        !Number.isInteger(move.destinationRow) ||
        move.sourceRow < plan.firstRow ||
        move.sourceRow > plan.lastRow ||
        move.destinationRow < plan.firstRow ||
        move.destinationRow > plan.lastRow ||
        !visibilityContext.visibleByPhysicalPosition[move.sourceRow - plan.firstRow] ||
        !visibilityContext.visibleByPhysicalPosition[move.destinationRow - plan.firstRow] ||
        sourceRows.has(move.sourceRow) ||
        destinationRows.has(move.destinationRow)
      ) {
        return finish('TRANSFER_NOT_VERIFIED', false, 'SORT_PLAN_ROW_MAPPING_INVALID');
      }
      sourceRows.add(move.sourceRow);
      destinationRows.add(move.destinationRow);
    }
    if (sourceRows.size !== destinationRows.size) {
      return finish('TRANSFER_NOT_VERIFIED', false, 'SOURCE_DESTINATION_COUNT_MISMATCH');
    }
    for (const row of sourceRows) {
      if (!destinationRows.has(row)) {
        return finish('TRANSFER_NOT_VERIFIED', false, 'SOURCE_DESTINATION_SET_MISMATCH');
      }
    }
    sortPerfStep_(perfTrace, 'Visible transfer: Validate plan', planValidationStarted, {
      rows: moves.length,
      blocks: plan.blocks.length
    });

    spreadsheet = sheet.getParent();
    if (!spreadsheet || typeof spreadsheet.getId !== 'function') {
      return finish('TRANSFER_NOT_VERIFIED', false, 'SPREADSHEET_ID_UNAVAILABLE');
    }
    const spreadsheetId = spreadsheet.getId();
    if (!spreadsheetId) {
      return finish('TRANSFER_NOT_VERIFIED', false, 'SPREADSHEET_ID_UNAVAILABLE');
    }

    const firstRow = plan.firstRow;
    const lastRow = plan.lastRow;
    const rowCount = lastRow - firstRow + 1;
    const startColumn = CONFIG.sortStartCol;
    const width = CONFIG.sortWidth;
    const endColumn = startColumn + width - 1;
    const targetRange = sheet.getRange(firstRow, startColumn, rowCount, width);

    // Moving only cell ranges is not safe for merges that cross row slots.
    // Reject all merged cells intersecting C:L rather than attempting to
    // recreate merge geometry over pinned hidden positions.
    const mergedRanges = targetRange.getMergedRanges();
    if (mergedRanges && mergedRanges.length > 0) {
      return finish('TRANSFER_NOT_VERIFIED', false, 'MERGED_CELLS_IN_SORT_RANGE_UNSUPPORTED');
    }

    ledgerProperties = PropertiesService.getDocumentProperties();
    if (
      !ledgerProperties ||
      typeof ledgerProperties.getProperty !== 'function' ||
      typeof ledgerProperties.setProperty !== 'function'
    ) {
      return finish('TRANSFER_NOT_VERIFIED', false, 'TRANSFER_LEDGER_UNAVAILABLE');
    }

    operationId = buildVisibleSortOperationId_(
      sheet,
      visibilityContext,
      plan,
      moves
    );
    const ledger = readVisibleSortTransferLedger_(ledgerProperties);
    const previous = ledger.find(entry => entry.id === operationId);
    if (previous && (
      previous.status === 'WRITE_ATTEMPTED' ||
      previous.status === 'COMPLETE' ||
      previous.status === 'VERIFY_FAILED' ||
      previous.status === 'WRITE_UNKNOWN'
    )) {
      return finish(
        'TRANSFER_NOT_VERIFIED',
        false,
        'DUPLICATE_TRANSFER_OPERATION_BLOCKED',
        { operationId: operationId, previousStatus: previous.status }
      );
    }
    writeVisibleSortTransferLedger_(
      ledgerProperties,
      ledger,
      operationId,
      'PREPARING',
      sheet.getSheetId(),
      null
    );

    const sourceSnapshotStarted = Date.now();
    const originalSnapshot = captureVisibleSortRangeSnapshot_(targetRange);
    sortPerfStep_(perfTrace, 'Visible transfer: Capture original C:L', sourceSnapshotStarted, {
      rows: rowCount,
      columns: width
    });

    const stageToken = Utilities.getUuid().replace(/-/g, '').slice(0, 12);
    const stageSourceName = '__VS_SRC_' + stageToken;
    const stageOutputName = '__VS_OUT_' + stageToken;

    stageSource = spreadsheet.insertSheet(stageSourceName);
    ensureVisibleSortStageGrid_(stageSource, lastRow, endColumn);
    stageSource.hideSheet();

    stageOutput = spreadsheet.insertSheet(stageOutputName);
    ensureVisibleSortStageGrid_(stageOutput, lastRow, endColumn);
    stageOutput.hideSheet();

    const stageCopyStarted = Date.now();
    Sheets.Spreadsheets.batchUpdate({
      requests: buildVisibleSortCopyPasteRequests_(
        sheet.getSheetId(),
        firstRow,
        stageSource.getSheetId(),
        firstRow,
        rowCount,
        startColumn,
        width,
        ['PASTE_NORMAL', 'PASTE_DATA_VALIDATION']
      )
    }, spreadsheetId);

    const sourceStageSnapshot = captureVisibleSortRangeSnapshot_(
      stageSource.getRange(firstRow, startColumn, rowCount, width)
    );
    const sourceStageMismatch = findVisibleSortSnapshotMismatch_(
      originalSnapshot,
      sourceStageSnapshot,
      firstRow
    );
    if (sourceStageMismatch) {
      throw new Error(
        'STAGING_SOURCE_FIDELITY_MISMATCH: ' +
        JSON.stringify(sourceStageMismatch)
      );
    }
    sortPerfStep_(perfTrace, 'Visible transfer: Stage original C:L', stageCopyStarted, {
      rows: rowCount,
      columns: width
    });

    const runs = buildVisibleSortCopyRuns_(moves);
    const prepareOutputStarted = Date.now();
    const prepareOutputRequests = [];
    runs.forEach(run => {
      prepareOutputRequests.push.apply(
        prepareOutputRequests,
        buildVisibleSortCopyPasteRequests_(
          stageSource.getSheetId(),
          run.sourceStartRow,
          stageOutput.getSheetId(),
          run.destinationStartRow,
          run.rowCount,
          startColumn,
          width,
          ['PASTE_NORMAL', 'PASTE_DATA_VALIDATION']
        )
      );
    });
    Sheets.Spreadsheets.batchUpdate({
      requests: prepareOutputRequests
    }, spreadsheetId);

    const outputSnapshot = captureVisibleSortRangeSnapshot_(
      stageOutput.getRange(firstRow, startColumn, rowCount, width)
    );
    for (const move of moves) {
      const expectedSource = sourceStageSnapshot.cells[move.sourceRow - firstRow];
      const actualDestination = outputSnapshot.cells[move.destinationRow - firstRow];
      const mismatch = findVisibleSortCellMismatch_(
        expectedSource,
        actualDestination,
        move.destinationRow
      );
      if (mismatch) {
        throw new Error(
          'STAGING_OUTPUT_FIDELITY_MISMATCH: ' +
          JSON.stringify(Object.assign({ sourceRow: move.sourceRow }, mismatch))
        );
      }
    }
    sortPerfStep_(perfTrace, 'Visible transfer: Prepare destination staging', prepareOutputStarted, {
      rows: moves.length,
      runs: runs.length,
      copyPasteRequests: prepareOutputRequests.length
    });

    // Detect edits made while staging was prepared. This also catches formula,
    // note, validation, rich-text, and formatting changes that are not part
    // of the Resolver's values-only data fingerprint.
    const preCommitSnapshot = captureVisibleSortRangeSnapshot_(targetRange);
    const preCommitMismatch = findVisibleSortSnapshotMismatch_(
      originalSnapshot,
      preCommitSnapshot,
      firstRow
    );
    if (preCommitMismatch) {
      writeVisibleSortTransferLedger_(
        ledgerProperties,
        readVisibleSortTransferLedger_(ledgerProperties),
        operationId,
        'ABORTED_BEFORE_WRITE',
        sheet.getSheetId(),
        { reason: 'SOURCE_CHANGED_DURING_STAGING', mismatch: preCommitMismatch }
      );
      cleanupStages();
      return finish(
        'BLOCKED_STALE_CONTEXT',
        false,
        'SOURCE_CHANGED_DURING_STAGING',
        { mismatch: preCommitMismatch, operationId: operationId }
      );
    }

    // This is deliberately the last spreadsheet-state validation before the
    // first write to the target sheet.
    const freshValidationStarted = Date.now();
    const freshValidation = validateSortVisibilityContextFresh_(
      sheet,
      visibilityContext
    );
    if (!freshValidation || freshValidation.ok !== true) {
      writeVisibleSortTransferLedger_(
        ledgerProperties,
        readVisibleSortTransferLedger_(ledgerProperties),
        operationId,
        'ABORTED_BEFORE_WRITE',
        sheet.getSheetId(),
        { reason: freshValidation && freshValidation.reason
            ? freshValidation.reason
            : 'FRESH_CONTEXT_VALIDATION_FAILED' }
      );
      cleanupStages();
      return finish(
        'BLOCKED_STALE_CONTEXT',
        false,
        freshValidation && freshValidation.reason
          ? freshValidation.reason
          : 'FRESH_CONTEXT_VALIDATION_FAILED',
        { operationId: operationId }
      );
    }
    sortPerfStep_(perfTrace, 'Visible transfer: Fresh context gate', freshValidationStarted, {
      ok: true,
      rows: rowCount
    });

    const finalWriteRequests = [];
    runs.forEach(run => {
      finalWriteRequests.push.apply(
        finalWriteRequests,
        buildVisibleSortCopyPasteRequests_(
          stageOutput.getSheetId(),
          run.destinationStartRow,
          sheet.getSheetId(),
          run.destinationStartRow,
          run.rowCount,
          startColumn,
          width,
          ['PASTE_NORMAL', 'PASTE_DATA_VALIDATION']
        )
      );
    });
    if (finalWriteRequests.length === 0) {
      writeVisibleSortTransferLedger_(
        ledgerProperties,
        readVisibleSortTransferLedger_(ledgerProperties),
        operationId,
        'ABORTED_BEFORE_WRITE',
        sheet.getSheetId(),
        { reason: 'NO_TRANSFER_REQUESTS' }
      );
      cleanupStages();
      return finish('TRANSFER_NOT_VERIFIED', false, 'NO_TRANSFER_REQUESTS');
    }

    // Persist the replay guard before sending any destination write. A crash
    // or transport timeout after this point is treated as outcome-unknown;
    // the same operation cannot be replayed automatically.
    writeVisibleSortTransferLedger_(
      ledgerProperties,
      readVisibleSortTransferLedger_(ledgerProperties),
      operationId,
      'WRITE_ATTEMPTED',
      sheet.getSheetId(),
      {
        rows: moves.length,
        runs: runs.length,
        sourceStageSheetId: stageSource.getSheetId(),
        outputStageSheetId: stageOutput.getSheetId()
      }
    );
    writeAttempted = true;

    const writeStarted = Date.now();
    try {
      Sheets.Spreadsheets.batchUpdate({
        requests: finalWriteRequests
      }, spreadsheetId);
    } catch (writeError) {
      writeVisibleSortTransferLedger_(
        ledgerProperties,
        readVisibleSortTransferLedger_(ledgerProperties),
        operationId,
        'WRITE_UNKNOWN',
        sheet.getSheetId(),
        { message: visibleSortTransferErrorMessage_(writeError) }
      );
      console.error('[VISIBLE SORT TRANSFER] destination write outcome unknown', JSON.stringify({
        operationId: operationId,
        sheetId: sheet.getSheetId(),
        stageSource: stageSourceName,
        stageOutput: stageOutputName,
        rows: moves.length,
        message: visibleSortTransferErrorMessage_(writeError)
      }));
      return finish(
        'TRANSFER_WRITE_UNKNOWN',
        false,
        'DESTINATION_WRITE_OUTCOME_UNKNOWN',
        {
          operationId: operationId,
          rowsMoved: moves.length,
          stagingSheets: [stageSourceName, stageOutputName],
          message: visibleSortTransferErrorMessage_(writeError)
        }
      );
    }
    sortPerfStep_(perfTrace, 'Visible transfer: Atomic destination batch', writeStarted, {
      rows: moves.length,
      runs: runs.length,
      copyPasteRequests: finalWriteRequests.length
    });

    const verificationStarted = Date.now();
    let actualSnapshot;
    try {
      actualSnapshot = captureVisibleSortRangeSnapshot_(targetRange);
    } catch (verificationReadError) {
      writeVisibleSortTransferLedger_(
        ledgerProperties,
        readVisibleSortTransferLedger_(ledgerProperties),
        operationId,
        'VERIFY_FAILED',
        sheet.getSheetId(),
        { message: visibleSortTransferErrorMessage_(verificationReadError) }
      );
      return finish(
        'TRANSFER_VERIFY_FAILED',
        false,
        'POST_WRITE_READ_FAILED',
        {
          operationId: operationId,
          rowsMoved: moves.length,
          stagingSheets: [stageSourceName, stageOutputName],
          message: visibleSortTransferErrorMessage_(verificationReadError)
        }
      );
    }

    const expectedSnapshot = buildVisibleSortExpectedSnapshot_(
      originalSnapshot,
      moves,
      firstRow
    );
    const finalMismatch = findVisibleSortSnapshotMismatch_(
      expectedSnapshot,
      actualSnapshot,
      firstRow
    );
    if (finalMismatch) {
      writeVisibleSortTransferLedger_(
        ledgerProperties,
        readVisibleSortTransferLedger_(ledgerProperties),
        operationId,
        'VERIFY_FAILED',
        sheet.getSheetId(),
        { mismatch: finalMismatch }
      );
      console.error('[VISIBLE SORT TRANSFER] post-write verification failed', JSON.stringify({
        operationId: operationId,
        sheetId: sheet.getSheetId(),
        stageSource: stageSourceName,
        stageOutput: stageOutputName,
        mismatch: finalMismatch
      }));
      return finish(
        'TRANSFER_VERIFY_FAILED',
        false,
        'POST_WRITE_SNAPSHOT_MISMATCH',
        {
          operationId: operationId,
          rowsMoved: moves.length,
          stagingSheets: [stageSourceName, stageOutputName],
          mismatch: finalMismatch
        }
      );
    }
    sortPerfStep_(perfTrace, 'Visible transfer: Verify full C:L', verificationStarted, {
      rows: rowCount,
      movedRows: moves.length,
      hiddenRowsPinned: rowCount - destinationRows.size
    });

    writeVisibleSortTransferLedger_(
      ledgerProperties,
      readVisibleSortTransferLedger_(ledgerProperties),
      operationId,
      'COMPLETE',
      sheet.getSheetId(),
      {
        rows: moves.length,
        stageSourceSheetId: stageSource.getSheetId(),
        stageOutputSheetId: stageOutput.getSheetId()
      }
    );

    const cleanupWarnings = [];
    [stageOutput, stageSource].forEach(stage => {
      try {
        spreadsheet.deleteSheet(stage);
      } catch (cleanupError) {
        cleanupWarnings.push({
          sheetId: stage.getSheetId(),
          message: visibleSortTransferErrorMessage_(cleanupError)
        });
      }
    });

    return finish('SUCCESS', true, null, {
      operationId: operationId,
      rowsMoved: moves.length,
      runs: runs.length,
      cleanupWarnings: cleanupWarnings
    });

  } catch (error) {
    const message = visibleSortTransferErrorMessage_(error);
    if (ledgerProperties && operationId) {
      try {
        writeVisibleSortTransferLedger_(
          ledgerProperties,
          readVisibleSortTransferLedger_(ledgerProperties),
          operationId,
          writeAttempted ? 'WRITE_UNKNOWN' : 'ABORTED_BEFORE_WRITE',
          sheet && sheet.getSheetId ? sheet.getSheetId() : null,
          { message: message }
        );
      } catch (ledgerError) {
        console.error('[VISIBLE SORT TRANSFER] could not persist failure state', JSON.stringify({
          operationId: operationId,
          message: visibleSortTransferErrorMessage_(ledgerError)
        }));
      }
    }
    if (!writeAttempted) {
      cleanupStages();
      return finish('BLOCKED_STAGING', false, message, {
        operationId: operationId
      });
    }
    console.error('[VISIBLE SORT TRANSFER] failure after destination write began', JSON.stringify({
      operationId: operationId,
      stageSourceSheetId: stageSource && stageSource.getSheetId
        ? stageSource.getSheetId()
        : null,
      stageOutputSheetId: stageOutput && stageOutput.getSheetId
        ? stageOutput.getSheetId()
        : null,
      message: message
    }));
    return finish('TRANSFER_WRITE_UNKNOWN', false, 'TRANSFER_FAILED_AFTER_WRITE_ATTEMPT', {
      operationId: operationId,
      stagingSheets: [
        stageSource && stageSource.getName ? stageSource.getName() : null,
        stageOutput && stageOutput.getName ? stageOutput.getName() : null
      ].filter(Boolean),
      message: message
    });
  }
}

function visibleSortTransferErrorMessage_(error) {
  return error && error.message ? error.message : String(error);
}

function flattenVisibleSortMoves_(plan) {
  const moves = [];
  (plan.blocks || []).forEach(block => {
    (block.destinations || []).forEach(move => {
      moves.push({
        sourceRow: move.sourceRow,
        destinationRow: move.destinationRow
      });
    });
  });
  return moves;
}

function buildVisibleSortCopyRuns_(moves) {
  const ordered = moves.slice().sort((left, right) =>
    left.destinationRow - right.destinationRow
  );
  const runs = [];
  ordered.forEach(move => {
    const previous = runs[runs.length - 1];
    if (
      previous &&
      move.destinationRow === previous.destinationStartRow + previous.rowCount &&
      move.sourceRow === previous.sourceStartRow + previous.rowCount
    ) {
      previous.rowCount++;
    } else {
      runs.push({
        sourceStartRow: move.sourceRow,
        destinationStartRow: move.destinationRow,
        rowCount: 1
      });
    }
  });
  return runs;
}

function buildVisibleSortCopyPasteRequests_(
  sourceSheetId,
  sourceStartRow,
  destinationSheetId,
  destinationStartRow,
  rowCount,
  startColumn,
  width,
  pasteTypes
) {
  const source = {
    sheetId: sourceSheetId,
    startRowIndex: sourceStartRow - 1,
    endRowIndex: sourceStartRow - 1 + rowCount,
    startColumnIndex: startColumn - 1,
    endColumnIndex: startColumn - 1 + width
  };
  const destination = {
    sheetId: destinationSheetId,
    startRowIndex: destinationStartRow - 1,
    endRowIndex: destinationStartRow - 1 + rowCount,
    startColumnIndex: startColumn - 1,
    endColumnIndex: startColumn - 1 + width
  };
  return pasteTypes.map(pasteType => ({
    copyPaste: {
      source: Object.assign({}, source),
      destination: Object.assign({}, destination),
      pasteType: pasteType
    }
  }));
}

function ensureVisibleSortStageGrid_(sheet, lastRow, lastColumn) {
  const currentRows = sheet.getMaxRows();
  if (currentRows < lastRow) {
    sheet.insertRowsAfter(currentRows, lastRow - currentRows);
  }
  const currentColumns = sheet.getMaxColumns();
  if (currentColumns < lastColumn) {
    sheet.insertColumnsAfter(currentColumns, lastColumn - currentColumns);
  }
}

function buildVisibleSortOperationId_(sheet, context, plan, moves) {
  const payload = JSON.stringify({
    version: 1,
    sheetId: sheet.getSheetId(),
    firstRow: plan.firstRow,
    lastRow: plan.lastRow,
    sortType: plan.sortType,
    dataFingerprint: context.dataFingerprint,
    stateFingerprint: context.stateFingerprint,
    moves: moves
  });
  let hash = 2166136261;
  for (let index = 0; index < payload.length; index++) {
    hash ^= payload.charCodeAt(index);
    hash = Math.imul(hash, 16777619);
  }
  return 'visible-sort-v1-' + sheet.getSheetId() + '-' +
    (hash >>> 0).toString(16).padStart(8, '0');
}

function readVisibleSortTransferLedger_(properties) {
  const raw = properties.getProperty(VISIBLE_SORT_TRANSFER_LEDGER_KEY_);
  if (!raw) return [];
  try {
    const parsed = JSON.parse(raw);
    return Array.isArray(parsed) ? parsed : [];
  } catch (error) {
    throw new Error('TRANSFER_LEDGER_CORRUPT');
  }
}

function writeVisibleSortTransferLedger_(
  properties,
  ledger,
  operationId,
  status,
  sheetId,
  details
) {
  const next = (ledger || []).filter(entry => entry.id !== operationId);
  next.push({
    id: operationId,
    status: status,
    sheetId: sheetId,
    updatedAt: Date.now(),
    details: details || null
  });
  properties.setProperty(
    VISIBLE_SORT_TRANSFER_LEDGER_KEY_,
    JSON.stringify(next.slice(-VISIBLE_SORT_TRANSFER_LEDGER_LIMIT_))
  );
}

function captureVisibleSortRangeSnapshot_(range) {
  const values = range.getValues();
  const formulasR1C1 = range.getFormulasR1C1();
  const notes = range.getNotes();
  const validations = range.getDataValidations();
  const richText = range.getRichTextValues();
  const formats = {};

  VISIBLE_SORT_TRANSFER_FORMAT_METHODS_.forEach(method => {
    if (typeof range[method] !== 'function') {
      throw new Error('SNAPSHOT_METHOD_UNAVAILABLE: ' + method);
    }
    formats[method] = range[method]().map(row =>
      row.map(value => normalizeVisibleSortTransferValue_(value))
    );
  });

  const cells = values.map((row, rowIndex) =>
    row.map((value, columnIndex) => {
      const formula = formulasR1C1[rowIndex][columnIndex] || '';
      return {
        formulaR1C1: formula,
        value: formula ? null : normalizeVisibleSortTransferValue_(value),
        note: notes[rowIndex][columnIndex] || '',
        validation: describeVisibleSortValidation_(validations[rowIndex][columnIndex]),
        richText: describeVisibleSortRichText_(richText[rowIndex][columnIndex]),
        format: VISIBLE_SORT_TRANSFER_FORMAT_METHODS_.reduce((result, method) => {
          result[method] = formats[method][rowIndex][columnIndex];
          return result;
        }, {})
      };
    })
  );
  return { cells: cells };
}

function normalizeVisibleSortTransferValue_(value) {
  if (value === null || value === undefined) return value === undefined ? null : null;
  if (value instanceof Date) return { __type: 'date', milliseconds: value.getTime() };
  if (Array.isArray(value)) {
    return value.map(item => normalizeVisibleSortTransferValue_(item));
  }
  if (typeof value === 'object') {
    if (typeof value.getA1Notation === 'function') {
      return {
        __type: 'range',
        a1: value.getA1Notation(),
        sheetId: typeof value.getSheet === 'function' &&
          value.getSheet() &&
          typeof value.getSheet().getSheetId === 'function'
          ? value.getSheet().getSheetId()
          : null
      };
    }
    if (typeof value.getAngle === 'function' || typeof value.isVertical === 'function') {
      return {
        angle: typeof value.getAngle === 'function' ? value.getAngle() : null,
        vertical: typeof value.isVertical === 'function' ? value.isVertical() : null
      };
    }
    return String(value);
  }
  return value;
}

function describeVisibleSortValidation_(rule) {
  if (!rule) return null;
  return {
    criteriaType: String(rule.getCriteriaType()),
    criteriaValues: normalizeVisibleSortTransferValue_(rule.getCriteriaValues()),
    allowInvalid: rule.getAllowInvalid(),
    helpText: rule.getHelpText()
  };
}

function describeVisibleSortTextStyle_(style) {
  if (!style) return null;
  const methods = [
    'isBold',
    'isItalic',
    'isUnderline',
    'isStrikethrough',
    'getFontFamily',
    'getFontSize',
    'getForegroundColor'
  ];
  const result = {};
  methods.forEach(method => {
    if (typeof style[method] === 'function') {
      result[method] = style[method]();
    }
  });
  return result;
}

function describeVisibleSortRichText_(value) {
  if (!value) return null;
  const runs = typeof value.getRuns === 'function' ? value.getRuns() : [];
  return {
    text: value.getText(),
    linkUrl: typeof value.getLinkUrl === 'function' ? value.getLinkUrl() : null,
    runs: runs.map(run => ({
      text: run.getText(),
      linkUrl: typeof run.getLinkUrl === 'function' ? run.getLinkUrl() : null,
      style: describeVisibleSortTextStyle_(
        typeof run.getTextStyle === 'function' ? run.getTextStyle() : null
      )
    }))
  };
}

function findVisibleSortCellMismatch_(expected, actual, rowNumber) {
  if (!expected || !actual) {
    return { row: rowNumber, reason: 'CELL_SNAPSHOT_MISSING' };
  }
  const fields = [
    'formulaR1C1',
    'value',
    'note',
    'validation',
    'richText',
    'format'
  ];
  for (const field of fields) {
    if (JSON.stringify(expected[field]) !== JSON.stringify(actual[field])) {
      return {
        row: rowNumber,
        field: field,
        expected: expected[field],
        actual: actual[field]
      };
    }
  }
  return null;
}

function findVisibleSortSnapshotMismatch_(expected, actual, firstRow) {
  if (
    !expected ||
    !actual ||
    !Array.isArray(expected.cells) ||
    !Array.isArray(actual.cells) ||
    expected.cells.length !== actual.cells.length
  ) {
    return { reason: 'SNAPSHOT_DIMENSION_MISMATCH' };
  }
  for (let rowIndex = 0; rowIndex < expected.cells.length; rowIndex++) {
    if (
      !Array.isArray(expected.cells[rowIndex]) ||
      !Array.isArray(actual.cells[rowIndex]) ||
      expected.cells[rowIndex].length !== actual.cells[rowIndex].length
    ) {
      return { row: firstRow + rowIndex, reason: 'SNAPSHOT_WIDTH_MISMATCH' };
    }
    for (let columnIndex = 0; columnIndex < expected.cells[rowIndex].length; columnIndex++) {
      const mismatch = findVisibleSortCellMismatch_(
        expected.cells[rowIndex][columnIndex],
        actual.cells[rowIndex][columnIndex],
        firstRow + rowIndex
      );
      if (mismatch) {
        mismatch.columnOffset = columnIndex;
        return mismatch;
      }
    }
  }
  return null;
}

function buildVisibleSortExpectedSnapshot_(originalSnapshot, moves, firstRow) {
  const expected = JSON.parse(JSON.stringify(originalSnapshot));
  moves.forEach(move => {
    expected.cells[move.destinationRow - firstRow] = JSON.parse(
      JSON.stringify(originalSnapshot.cells[move.sourceRow - firstRow])
    );
  });
  return expected;
}

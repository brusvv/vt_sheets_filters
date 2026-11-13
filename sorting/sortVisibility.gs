/**
 * READ-ONLY VISIBILITY RESOLVER FOR PINNED SORTING.
 *
 * Never creates or migrates FilterState, writes fallback state, or mutates
 * ManualState. Any inconsistent or unsupported state prevents sorting.
 */
function normalizeSortSnapshotValue_(value) {
  if (value instanceof Date) return ['date', value.toISOString()];
  if (value === null || value === undefined) return ['null', ''];
  return [typeof value, value];
}

/**
 * Rebuilds the complete resolver fingerprint from its captured fields.
 * This detects accidental mutation of the snapshot object itself.
 * It does not replace a fresh read from Sheets when a commit is introduced.
 */
function buildSortVisibilityStateFingerprint_(context) {
  return JSON.stringify({
    firstRow: context.firstRow,
    lastRow: context.lastRow,
    dataFingerprint: context.dataFingerprint,
    filterVisibility: context.filterVisibility,
    manualKeys: context.manualKeys,
    physicalHiddenByUser: context.physicalHiddenByUser,
    physicalHiddenByFilter: context.physicalHiddenByFilter,
    filterMetadata: context.filterMetadata || null,
    manualMetadata: context.manualMetadata || null,
    savedLastRow: context.savedLastRow === undefined
      ? null
      : context.savedLastRow
  });
}

function buildSortVisibilityContextFromSnapshots_(args, perfTrace = null) {
  const data = args.data;
  const firstRow = args.firstRow;
  const lastRow = args.lastRow;
  const filterVisibility = args.filterVisibility;
  const manualKeys = args.manualKeys;
  const physicalHiddenByUser = args.physicalHiddenByUser;
  const physicalHiddenByFilter = args.physicalHiddenByFilter;

  const validationStarted = perfTrace ? Date.now() : 0;
  if (!Number.isInteger(firstRow) || !Number.isInteger(lastRow) || lastRow < firstRow) {
    return { ok: false, reason: 'INVALID_ROW_BOUNDARY' };
  }

  const expectedLength = lastRow - firstRow + 1;
  if (
    !Array.isArray(data) ||
    data.length !== expectedLength ||
    !data.every(row => Array.isArray(row) && row.length >= CONFIG.sortWidth) ||
    !Array.isArray(filterVisibility) ||
    filterVisibility.length !== expectedLength ||
    !filterVisibility.every(value => typeof value === 'boolean') ||
    !Array.isArray(physicalHiddenByUser) ||
    physicalHiddenByUser.length !== expectedLength ||
    !physicalHiddenByUser.every(value => typeof value === 'boolean') ||
    !Array.isArray(physicalHiddenByFilter) ||
    physicalHiddenByFilter.length !== expectedLength ||
    !physicalHiddenByFilter.every(value => typeof value === 'boolean') ||
    !Array.isArray(manualKeys) ||
    !manualKeys.every(key => typeof key === 'string' && key.length > 0)
  ) {
    return { ok: false, reason: 'VISIBILITY_SNAPSHOT_INVALID' };
  }

  if (physicalHiddenByFilter.some(Boolean)) {
    return { ok: false, reason: 'EXTERNAL_FILTER_UNSUPPORTED' };
  }

  if (perfTrace) {
    sortPerfStep_(perfTrace, 'Resolver: Validate snapshots', validationStarted, {
      rows: expectedLength,
      valid: true
    });
  }

  const contextBuildStarted = perfTrace ? Date.now() : 0;
  const manualKeyCounts = new Map();
  manualKeys.forEach(key => {
    manualKeyCounts.set(key, (manualKeyCounts.get(key) || 0) + 1);
  });

  const physicalManualCandidates = new Map();

  for (let index = 0; index < expectedLength; index++) {
    const row = firstRow + index;
    const filterVisible = filterVisibility[index];
    const hiddenByUser = physicalHiddenByUser[index];

    if (!filterVisible && !hiddenByUser) {
      return {
        ok: false,
        reason: 'FILTERSTATE_PHYSICAL_VISIBILITY_MISMATCH',
        row: row
      };
    }

    if (filterVisible && hiddenByUser) {
      if (!String(data[index][0] || '').trim()) {
        return {
          ok: false,
          reason: 'UNTRACKED_HIDDEN_EMPTY_SCENARIO',
          row: row
        };
      }

      const key = getManualRecordKey_(data[index]);
      if (!physicalManualCandidates.has(key)) {
        physicalManualCandidates.set(key, []);
      }
      physicalManualCandidates.get(key).push(row);
    }
  }

  const manualHiddenRows = new Set();

  for (const [key, expectedCount] of manualKeyCounts.entries()) {
    const candidates = physicalManualCandidates.get(key) || [];

    if (candidates.length !== expectedCount) {
      return {
        ok: false,
        reason: 'MANUALSTATE_PHYSICAL_COUNT_MISMATCH',
        key: key,
        expectedCount: expectedCount,
        physicalCount: candidates.length
      };
    }

    candidates.forEach(row => manualHiddenRows.add(row));
  }

  for (const [key, candidates] of physicalManualCandidates.entries()) {
    if (!manualKeyCounts.has(key)) {
      return {
        ok: false,
        reason: 'PHYSICAL_HIDDEN_ROW_MISSING_MANUALSTATE_KEY',
        key: key,
        rows: candidates.slice()
      };
    }
  }

  const visibleByPhysicalPosition = filterVisibility.map(
    (filterVisible, index) =>
      filterVisible &&
      !physicalHiddenByUser[index] &&
      !manualHiddenRows.has(firstRow + index)
  );

  if (perfTrace) {
    sortPerfStep_(perfTrace, 'Resolver: Build context', contextBuildStarted, {
      rows: expectedLength,
      manualKeys: manualKeys.length,
      manualHiddenRows: manualHiddenRows.size
    });
  }

  const dataFingerprintStarted = perfTrace ? Date.now() : 0;
  const dataFingerprint = JSON.stringify(
    data.map(row => row.map(normalizeSortSnapshotValue_))
  );
  if (perfTrace) {
    sortPerfStep_(perfTrace, 'Resolver: Data fingerprint', dataFingerprintStarted, {
      rows: expectedLength
    });
  }

  const context = {
    ok: true,
    firstRow: firstRow,
    lastRow: lastRow,
    visibleByPhysicalPosition: visibleByPhysicalPosition,
    filterVisibility: filterVisibility.slice(),
    manualHiddenRows: manualHiddenRows,
    manualKeyCounts: manualKeyCounts,
    manualKeys: manualKeys.slice(),
    physicalHiddenByUser: physicalHiddenByUser.slice(),
    physicalHiddenByFilter: physicalHiddenByFilter.slice(),
    dataFingerprint: dataFingerprint,
    filterMetadata: null,
    manualMetadata: null,
    savedLastRow: null,
    source: 'validated-filterstate-manualstate-physical'
  };

  const stateFingerprintStarted = perfTrace ? Date.now() : 0;
  context.stateFingerprint = buildSortVisibilityStateFingerprint_(context);
  context.fingerprint = context.stateFingerprint;
  if (perfTrace) {
    sortPerfStep_(perfTrace, 'Resolver: State fingerprint', stateFingerprintStarted, {
      rows: expectedLength,
      manualKeys: manualKeys.length
    });
  }

  return context;
}

function sortManualCountsEqual_(left, right) {
  if (!(left instanceof Map) || !(right instanceof Map)) return false;
  if (left.size !== right.size) return false;

  for (const [key, count] of left.entries()) {
    if (right.get(key) !== count) return false;
  }

  return true;
}

/**
 * Reads current state without ensuring, creating, or migrating its layout.
 */
function readSortVisibilityContext_(sheet, lastRow, knownSortData, perfTrace = null) {
  try {
    if (!sheet || !Number.isInteger(lastRow) || lastRow < CONFIG.firstRow) {
      return { ok: false, reason: 'INVALID_SORT_SHEET_OR_BOUNDARY' };
    }

    const snapshotStarted = perfTrace ? Date.now() : 0;
    const upperRowCount = lastRow - CONFIG.firstRow + 1;
    let data;

    if (
      Array.isArray(knownSortData) &&
      knownSortData.length === upperRowCount &&
      knownSortData.every(row => Array.isArray(row) && row.length >= CONFIG.sortWidth)
    ) {
      data = knownSortData.map(row => row.slice(0, CONFIG.sortWidth));
    } else {
      data = sheet.getRange(
        CONFIG.firstRow,
        CONFIG.sortStartCol,
        upperRowCount,
        CONFIG.sortWidth
      ).getValues();
    }

    const exactLastRow = getLastScenarioRowFromData(data);
    if (exactLastRow < CONFIG.firstRow) {
      return { ok: false, reason: 'NO_SCENARIO_DATA' };
    }

    const exactCount = exactLastRow - CONFIG.firstRow + 1;
    data = data.slice(0, exactCount);
    if (perfTrace) {
      sortPerfStep_(perfTrace, 'Resolver: Prepare C:L snapshot', snapshotStarted, {
        rows: exactCount,
        readFromKnownSortData: Array.isArray(knownSortData)
      });
    }

    const stateSheet = sheet.getParent().getSheetByName('_FilterState');
    if (!stateSheet) {
      return { ok: false, reason: 'FILTERSTATE_SHEET_MISSING' };
    }

    const filterColumns = getFilterStateColumns(sheet);
    const manualColumns = getManualStateColumns(sheet);

    const filterMetadataStarted = perfTrace ? Date.now() : 0;
    const filterMetadata = stateSheet.getRange(
      1,
      filterColumns.visibleCol,
      2,
      2
    ).getValues();
    if (perfTrace) sortPerfStep_(perfTrace, 'Resolver: Read FilterState metadata', filterMetadataStarted, {
      rows: 2,
      columns: 2
    });

    if (
      filterMetadata[0][0] !== sheet.getName() + ' Visible' ||
      filterMetadata[0][1] !== 'LastRow'
    ) {
      return { ok: false, reason: 'FILTERSTATE_METADATA_INVALID' };
    }

    const savedLastRow = Number(filterMetadata[1][1]);
    if (!Number.isInteger(savedLastRow) || savedLastRow < exactLastRow) {
      return { ok: false, reason: 'FILTERSTATE_TOO_SHORT' };
    }

    const filterVisibilityStarted = perfTrace ? Date.now() : 0;
    const filterVisibility = stateSheet.getRange(
      CONFIG.firstRow,
      filterColumns.visibleCol,
      exactCount,
      1
    ).getValues().map(row => row[0]);
    if (perfTrace) sortPerfStep_(perfTrace, 'Resolver: Read FilterState visibility', filterVisibilityStarted, {
      rows: exactCount
    });

    const manualMetadataStarted = perfTrace ? Date.now() : 0;
    const manualMetadata = stateSheet.getRange(
      1,
      manualColumns.keyCol,
      2,
      2
    ).getValues();
    if (perfTrace) sortPerfStep_(perfTrace, 'Resolver: Read ManualState metadata', manualMetadataStarted, {
      rows: 2,
      columns: 2
    });

    if (
      manualMetadata[0][0] !== sheet.getName() + ' Manual' ||
      manualMetadata[0][1] !== 'Count'
    ) {
      return { ok: false, reason: 'MANUALSTATE_METADATA_INVALID' };
    }

    const manualCount = Number(manualMetadata[1][1]);
    if (!Number.isInteger(manualCount) || manualCount < 0) {
      return { ok: false, reason: 'MANUALSTATE_COUNT_INVALID' };
    }

    const manualKeysStarted = perfTrace ? Date.now() : 0;
    const manualKeys = manualCount > 0
      ? stateSheet.getRange(
          CONFIG.firstRow,
          manualColumns.keyCol,
          manualCount,
          1
        ).getValues().map(row => row[0])
      : [];
    if (perfTrace) sortPerfStep_(perfTrace, 'Resolver: Read ManualState keys', manualKeysStarted, {
      keys: manualCount
    });

    const physicalVisibilityStarted = perfTrace ? Date.now() : 0;
    const physicalHiddenByUser = [];
    const physicalHiddenByFilter = [];

    for (let offset = 0; offset < exactCount; offset++) {
      const row = CONFIG.firstRow + offset;
      physicalHiddenByUser.push(sheet.isRowHiddenByUser(row));
      physicalHiddenByFilter.push(sheet.isRowHiddenByFilter(row));
    }

    if (perfTrace) {
      sortPerfStep_(perfTrace, 'Resolver: Read physical visibility', physicalVisibilityStarted, {
        rowsChecked: exactCount,
        isRowHiddenByUserCalls: exactCount,
        isRowHiddenByFilterCalls: exactCount,
        spreadsheetApiCalls: exactCount * 2
      });
    }

    const context = buildSortVisibilityContextFromSnapshots_({
      data: data,
      firstRow: CONFIG.firstRow,
      lastRow: exactLastRow,
      filterVisibility: filterVisibility,
      manualKeys: manualKeys,
      physicalHiddenByUser: physicalHiddenByUser,
      physicalHiddenByFilter: physicalHiddenByFilter
    }, perfTrace);

    if (!context.ok) return context;

    context.sheetId = sheet.getSheetId();
    context.sheetName = sheet.getName();
    context.filterMetadata = filterMetadata.map(row => row.slice());
    context.manualMetadata = manualMetadata.map(row => row.slice());
    context.savedLastRow = savedLastRow;
    const finalStateFingerprintStarted = perfTrace ? Date.now() : 0;
    context.stateFingerprint = buildSortVisibilityStateFingerprint_(context);
    context.fingerprint = context.stateFingerprint;
    if (perfTrace) sortPerfStep_(perfTrace, 'Resolver: Final state fingerprint', finalStateFingerprintStarted, {
      rows: exactCount,
      manualKeys: manualKeys.length
    });
    return context;

  } catch (error) {
    return {
      ok: false,
      reason: 'VISIBILITY_READ_ERROR',
      message: error && error.message ? error.message : String(error)
    };
  }
}

/**
 * Validates a supplied visibility context against the exact C:L snapshot.
 * This is a pure check: no spreadsheet reads or writes.
 */
function validateSortVisibilityContext_(context, sheet, data, firstRow, lastRow) {
  if (!context || context.ok !== true) {
    return { ok: false, reason: context && context.reason
      ? context.reason
      : 'VISIBILITY_CONTEXT_MISSING' };
  }

  if (
    context.sheetId !== sheet.getSheetId() ||
    context.sheetName !== sheet.getName()
  ) {
    return { ok: false, reason: 'VISIBILITY_CONTEXT_SHEET_MISMATCH' };
  }

  const rowCount = lastRow - firstRow + 1;

  if (
    context.firstRow !== firstRow ||
    context.lastRow !== lastRow ||
    !Array.isArray(data) ||
    data.length !== rowCount ||
    !Array.isArray(context.visibleByPhysicalPosition) ||
    context.visibleByPhysicalPosition.length !== rowCount ||
    !context.visibleByPhysicalPosition.every(value => typeof value === 'boolean') ||
    !Array.isArray(context.filterVisibility) ||
    context.filterVisibility.length !== rowCount ||
    !context.filterVisibility.every(value => typeof value === 'boolean') ||
    !Array.isArray(context.physicalHiddenByUser) ||
    context.physicalHiddenByUser.length !== rowCount ||
    !context.physicalHiddenByUser.every(value => typeof value === 'boolean') ||
    !Array.isArray(context.physicalHiddenByFilter) ||
    context.physicalHiddenByFilter.length !== rowCount ||
    !context.physicalHiddenByFilter.every(value => typeof value === 'boolean') ||
    context.physicalHiddenByFilter.some(Boolean) ||
    !(context.manualHiddenRows instanceof Set) ||
    !(context.manualKeyCounts instanceof Map) ||
    typeof context.dataFingerprint !== 'string' ||
    typeof context.stateFingerprint !== 'string' ||
    typeof context.fingerprint !== 'string' ||
    !context.dataFingerprint ||
    !context.stateFingerprint ||
    !Array.isArray(context.manualKeys) ||
    !context.manualKeys.every(key => typeof key === 'string' && key.length > 0)
  ) {
    return { ok: false, reason: 'VISIBILITY_CONTEXT_INVALID' };
  }

  const expectedDataFingerprint = JSON.stringify(
    data.map(row => row.map(normalizeSortSnapshotValue_))
  );

  if (context.dataFingerprint !== expectedDataFingerprint) {
    return { ok: false, reason: 'VISIBILITY_CONTEXT_DATA_STALE' };
  }

  const expectedStateFingerprint =
    buildSortVisibilityStateFingerprint_(context);

  if (
    context.stateFingerprint !== expectedStateFingerprint ||
    context.fingerprint !== context.stateFingerprint
  ) {
    return {
      ok: false,
      reason: 'VISIBILITY_CONTEXT_FINGERPRINT_STALE'
    };
  }

  const candidateRows = new Set();
  const candidateCounts = new Map();

  for (let index = 0; index < rowCount; index++) {
    const rowNumber = firstRow + index;
    const filterVisible = context.filterVisibility[index];
    const physicallyHidden = context.physicalHiddenByUser[index];
    const isManualCandidate = filterVisible && physicallyHidden;

    if (!filterVisible && !physicallyHidden) {
      return {
        ok: false,
        reason: 'FILTERSTATE_PHYSICAL_VISIBILITY_MISMATCH',
        row: rowNumber
      };
    }

    if (isManualCandidate) {
      if (!String(data[index][0] || '').trim()) {
        return {
          ok: false,
          reason: 'UNTRACKED_HIDDEN_EMPTY_SCENARIO',
          row: rowNumber
        };
      }

      candidateRows.add(rowNumber);
      const key = getManualRecordKey_(data[index]);
      candidateCounts.set(key, (candidateCounts.get(key) || 0) + 1);
    }
  }

  if (candidateRows.size !== context.manualHiddenRows.size) {
    return { ok: false, reason: 'MANUAL_HIDDEN_ROWS_MISMATCH' };
  }

  for (const rowNumber of candidateRows) {
    if (!context.manualHiddenRows.has(rowNumber)) {
      return {
        ok: false,
        reason: 'MANUAL_HIDDEN_ROWS_MISMATCH',
        row: rowNumber
      };
    }
  }

  if (candidateCounts.size !== context.manualKeyCounts.size) {
    return { ok: false, reason: 'MANUAL_KEY_COUNTS_MISMATCH' };
  }

  for (const [key, count] of candidateCounts.entries()) {
    if (context.manualKeyCounts.get(key) !== count) {
      return {
        ok: false,
        reason: 'MANUAL_KEY_COUNTS_MISMATCH',
        key: key
      };
    }
  }

  const expectedVisible = context.filterVisibility.map(
    (filterVisible, index) =>
      filterVisible &&
      !context.physicalHiddenByUser[index]
  );

  for (let index = 0; index < rowCount; index++) {
    if (context.visibleByPhysicalPosition[index] !== expectedVisible[index]) {
      return {
        ok: false,
        reason: 'VISIBLE_MASK_PHYSICAL_STATE_MISMATCH',
        row: firstRow + index
      };
    }
  }

  return { ok: true };
}


/**
 * Re-reads the current sheet, FilterState, ManualState and physical row state,
 * then compares them with a previously resolved snapshot.
 *
 * This is the required optimistic-concurrency gate for a future transfer
 * commit. The current sorter deliberately has no write phase and therefore
 * does not call this helper as part of TRANSFER_NOT_VERIFIED.
 */
function validateSortVisibilityContextFresh_(sheet, context) {
  if (!sheet || !context || context.ok !== true) {
    return {
      ok: false,
      reason: 'VISIBILITY_CONTEXT_MISSING_OR_INVALID'
    };
  }

  if (
    context.sheetId !== sheet.getSheetId() ||
    context.sheetName !== sheet.getName()
  ) {
    return {
      ok: false,
      reason: 'VISIBILITY_CONTEXT_SHEET_MISMATCH'
    };
  }

  if (
    !Number.isInteger(context.firstRow) ||
    !Number.isInteger(context.lastRow) ||
    context.firstRow !== CONFIG.firstRow
  ) {
    return {
      ok: false,
      reason: 'VISIBILITY_CONTEXT_RANGE_INVALID'
    };
  }

  try {
    const sheetLastRow = sheet.getLastRow();

    if (sheetLastRow < CONFIG.firstRow) {
      return {
        ok: false,
        reason: 'VISIBILITY_CONTEXT_BOUNDARY_STALE'
      };
    }

    const currentUpperData = sheet.getRange(
      CONFIG.firstRow,
      CONFIG.sortStartCol,
      sheetLastRow - CONFIG.firstRow + 1,
      CONFIG.sortWidth
    ).getValues();

    const currentLastRow = getLastScenarioRowFromData(currentUpperData);

    if (currentLastRow !== context.lastRow) {
      return {
        ok: false,
        reason: 'VISIBILITY_CONTEXT_BOUNDARY_STALE',
        expectedLastRow: context.lastRow,
        currentLastRow: currentLastRow
      };
    }

    const currentData = currentUpperData
      .slice(0, currentLastRow - CONFIG.firstRow + 1)
      .map(row => row.slice(0, CONFIG.sortWidth));

    const freshContext = readSortVisibilityContext_(
      sheet,
      currentLastRow,
      currentData
    );

    if (!freshContext.ok) {
      return {
        ok: false,
        reason: 'VISIBILITY_CONTEXT_FRESH_READ_FAILED',
        detail: freshContext.reason || null
      };
    }

    if (freshContext.dataFingerprint !== context.dataFingerprint) {
      return {
        ok: false,
        reason: 'VISIBILITY_CONTEXT_DATA_STALE'
      };
    }

    if (freshContext.stateFingerprint !== context.stateFingerprint) {
      return {
        ok: false,
        reason: 'VISIBILITY_CONTEXT_STATE_STALE'
      };
    }

    return {
      ok: true,
      context: freshContext
    };

  } catch (error) {
    return {
      ok: false,
      reason: 'VISIBILITY_CONTEXT_FRESH_READ_ERROR',
      message: error && error.message ? error.message : String(error)
    };
  }
}

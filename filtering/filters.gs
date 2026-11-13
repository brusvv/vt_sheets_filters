/**
 * FILTERS
 *
 * Основная логика фильтрации Dots и TOP.
 * Подготовка FilterState и ManualState.
 * Построение и сохранение видимости строк.
 *
 * Основная последовательность:
 * 1. Получение размера листа и настроек.
 * 2. Чтение игровых данных.
 * 3. Подготовка FilterState и ManualState.
 * 4. Построение фильтрационной видимости.
 * 5. Определение предыдущей физической видимости.
 * 6. Обновление видимости и сохранение FilterState.
 */


/**
 * Основная функция применения фильтров.
 *
 * Сохраняет прежнюю сигнатуру для совместимости.
 */
function applyFilters(
  skipAutoSort = false,
  targetSheet = null,
  resetVisibility = false,
  optionsValues = null,
  parentTrace = null,
  knownLastRow = null,
  knownFilterState = null
) {
  const perfTrace = triggerPerfStart_(
    'applyFilters',
    {
      sheet: targetSheet ? targetSheet.getName() : 'active',
      skipAutoSort: skipAutoSort,
      resetVisibility: resetVisibility
    },
    parentTrace
  );

  const totalStarted = Date.now();

  let status = 'SUCCESS';
  let errorMetadata = {};

  try {
    // --------------------------------------------------
    // 1. Получение листа и его размера
    // --------------------------------------------------

    let stepStarted = Date.now();

    const sheet = targetSheet || getTargetSheet();

    if (
      !sheet ||
      !RANK_TRACKED_SHEETS.includes(sheet.getName())
    ) {
      throw new Error(
        'Filters accept only tracked gameplay sheets: ' +
        RANK_TRACKED_SHEETS.join(', ')
      );
    }

    const sheetLastRow = knownLastRow !== null
      ? knownLastRow
      : sheet.getLastRow();

    filterPerfStep_(
      perfTrace,
      'Get sheet dimensions',
      stepStarted,
      { sheet: sheet.getName() }
    );

    if (sheetLastRow < CONFIG.firstRow) {
      status = 'EMPTY_SHEET';

      triggerPerfEnd_(perfTrace, status, {
        sheet: sheet.getName(),
        lastRow: CONFIG.firstRow - 1
      });

      logFilterTiming('APPLY FILTERS TOTAL', totalStarted);

      return;
    }

    // --------------------------------------------------
    // 2. Автоматическая сортировка и настройки
    // --------------------------------------------------

    stepStarted = Date.now();

    const autoSortEnabled = !skipAutoSort && (
      optionsValues
        ? optionsValues[1][0] === true
        : isAutoSortEnabled()
    );

    const hideEmptyScores = optionsValues
      ? optionsValues[2][0] === true
      : getOptionBoolean_('hideEmptyScores');

    // --------------------------------------------------
    // 3. Чтение игровых данных и настроек фильтров
    // --------------------------------------------------

    let inputData = null;
    let sortChanged = false;
    let sortVisibilityContext = null;

    if (autoSortEnabled) {
      inputData = readFilterInputData_(
        sheet,
        sheetLastRow,
        perfTrace
      );

      const sortLastRow = getLastScenarioRowFromData(inputData.data);

      if (sortLastRow >= CONFIG.firstRow) {
        const sortRowCount = sortLastRow - CONFIG.firstRow + 1;
        const sortData = inputData.data
          .slice(0, sortRowCount)
          .map(row => row.slice(0, CONFIG.sortWidth));

        stepStarted = Date.now();

        // Resolve the physical visibility before resetVisibility/showRows.
        sortVisibilityContext = readSortVisibilityContext_(
          sheet,
          sortLastRow,
          sortData
        );

        const sortResult = sortScenarioBlocksDetailed_(
          sheet,
          'scoreDesc',
          sortLastRow,
          null,
          sortData,
          sortVisibilityContext
        );

        sortChanged = sortResult.applied === true;

        triggerPerfEvent_(
          perfTrace,
          'AUTO_SORT_PLAN',
          {
            sheet: sheet.getName(),
            status: sortResult.status,
            reason: sortResult.reason || null,
            changed: Boolean(sortResult.plan && sortResult.plan.changed),
            applied: sortChanged
          }
        );

      } else {
        triggerPerfEvent_(
          perfTrace,
          'AUTO_SORT_PLAN',
          {
            sheet: sheet.getName(),
            status: 'NO_OP',
            reason: 'NO_SCENARIO_DATA',
            changed: false,
            applied: false
          }
        );
      }
    }

    filterPerfStep_(
      perfTrace,
      'Automatic sorting',
      stepStarted,
      {
        sheet: sheet.getName(),
        enabled: !skipAutoSort && autoSortEnabled,
        hideEmptyScores: hideEmptyScores,
        sortChanged: sortChanged
      }
    );

    // После Range.sort() исходный snapshot устаревает.
    // Перечитываем только если сортировка действительно запускалась.
    if (sortChanged) {
      inputData = readFilterInputData_(
        sheet,
        sheetLastRow,
        perfTrace
      );
    } else if (!inputData) {
      inputData = readFilterInputData_(
        sheet,
        sheetLastRow,
        perfTrace
      );
    }

    // --------------------------------------------------
    // 4. Определение последней игровой строки
    // --------------------------------------------------

    stepStarted = Date.now();

    const lastRow = getLastScenarioRowFromData(
      inputData.data
    );

    if (lastRow < CONFIG.firstRow) {
      status = 'EMPTY_SHEET';

      triggerPerfEnd_(perfTrace, status, {
        sheet: sheet.getName(),
        lastRow: lastRow
      });

      logFilterTiming(
        'Поиск последней игровой строки',
        stepStarted
      );

      logFilterTiming('APPLY FILTERS TOTAL', totalStarted);

      return;
    }

    const numRows = lastRow - CONFIG.firstRow + 1;

    inputData.data.length = numRows;

    filterPerfStep_(
      perfTrace,
      'Find last scenario row',
      stepStarted,
      {
        sheet: sheet.getName(),
        lastRow: lastRow,
        rows: numRows
      }
    );

    // A reset shows every row before rebuilding visibility. If the
    // pre-reset resolver could not classify the current state, do not run
    // showRows() and risk exposing an unmapped manual exclusion.
    if (
      autoSortEnabled &&
      resetVisibility &&
      (!sortVisibilityContext || !sortVisibilityContext.ok)
    ) {
      const reason = sortVisibilityContext
        ? sortVisibilityContext.reason
        : 'VISIBILITY_CONTEXT_MISSING';

      status = 'SORT_VISIBILITY_BLOCKED';

      triggerPerfEvent_(
        perfTrace,
        'FILTERS_ABORTED_UNRESOLVED_SORT_VISIBILITY',
        {
          sheet: sheet.getName(),
          reason: reason,
          resetVisibility: true
        }
      );

      triggerPerfEnd_(perfTrace, status, {
        sheet: sheet.getName(),
        reason: reason
      });

      return {
        status: status,
        applied: false,
        reason: reason
      };
    }

    // --------------------------------------------------
    // 5. Подготовка FilterState и ManualState
    // --------------------------------------------------

    const filterState = prepareFilterState_(
      sheet,
      numRows,
      knownFilterState,
      perfTrace
    );

    // --------------------------------------------------
    // 6. Построение новой фильтрационной видимости
    // --------------------------------------------------

    const manualVisibilityResolved = Boolean(
      autoSortEnabled &&
      sortVisibilityContext &&
      sortVisibilityContext.ok &&
      sortManualCountsEqual_(
        sortVisibilityContext.manualKeyCounts,
        filterState.manualHiddenCounts
      )
    );

    const hasManualStateEvidence =
      filterState.manualHiddenCounts.size > 0 ||
      Boolean(
        sortVisibilityContext &&
        sortVisibilityContext.ok &&
        (
          sortVisibilityContext.manualKeyCounts.size > 0 ||
          sortVisibilityContext.manualHiddenRows.size > 0
        )
      );

    const unresolvedManualState =
      autoSortEnabled &&
      !manualVisibilityResolved &&
      hasManualStateEvidence;

    if (unresolvedManualState) {
      const reason = sortVisibilityContext
        ? (sortVisibilityContext.reason || 'SORT_VISIBILITY_CONTEXT_STALE')
        : 'VISIBILITY_CONTEXT_MISSING';

      status = 'SORT_VISIBILITY_BLOCKED';

      triggerPerfEvent_(
        perfTrace,
        'FILTERS_ABORTED_UNRESOLVED_MANUALSTATE',
        {
          sheet: sheet.getName(),
          reason: reason,
          manualKeys: filterState.manualHiddenCounts.size,
          resetVisibility: resetVisibility
        }
      );

      triggerPerfEnd_(perfTrace, status, {
        sheet: sheet.getName(),
        reason: reason
      });

      return {
        status: status,
        applied: false,
        reason: 'UNRESOLVED_MANUALSTATE_VISIBILITY',
        detail: reason
      };
    }

    let resolvedManualHiddenRows = null;

    if (manualVisibilityResolved) {
      resolvedManualHiddenRows = sortVisibilityContext.manualHiddenRows;
    } else if (sortVisibilityContext && sortVisibilityContext.ok) {
      triggerPerfEvent_(
        perfTrace,
        'SORT_VISIBILITY_CONTEXT_STALE',
        { sheet: sheet.getName() }
      );
    }

    const visibilityData = buildFilterVisibility_(
      inputData.data,
      inputData.selectedDots,
      inputData.topN,
      hideEmptyScores,
      sheet,
      filterState.manualHiddenCounts,
      perfTrace,
      resolvedManualHiddenRows
    );

    // --------------------------------------------------
    // 7. Получение предыдущей фильтрационной видимости
    // --------------------------------------------------

    stepStarted = Date.now();

    let previousFilterVisibility;

    if (resetVisibility) {
      sheet.showRows(CONFIG.firstRow, numRows);

      previousFilterVisibility = Array(numRows).fill(true);

    } else if (filterState.hasKnownFilterState) {
      previousFilterVisibility =
        knownFilterState.previousVisibility.slice(0, numRows);

    } else {
      previousFilterVisibility = getPreviousVisibility(
        sheet,
        numRows,
        lastRow,
        filterState.filterStateSheet,
        filterState.filterStateMetadata
      );
    }

    filterPerfStep_(
      perfTrace,
      'Read previous filter visibility',
      stepStarted,
      {
        resetVisibility: resetVisibility,
        reused: filterState.hasKnownFilterState
      }
    );

    // --------------------------------------------------
    // 8. Формирование итоговой видимости
    // --------------------------------------------------

    stepStarted = Date.now();

    const newFilterVisibility = visibilityData.filterVisibility;

    const newVisibility = newFilterVisibility.map(
      (filterVisible, index) => {
        const sheetRow = CONFIG.firstRow + index;

        return (
          filterVisible &&
          !visibilityData.manualRowsToHide.has(sheetRow)
        );
      }
    );

    filterPerfStep_(
      perfTrace,
      'Build final visibility',
      stepStarted,
      {
        filterVisible: newFilterVisibility.filter(Boolean).length,
        manualHidden: visibilityData.manualRowsToHide.size,
        finalVisible: newVisibility.filter(Boolean).length
      }
    );

    // --------------------------------------------------
    // 9. Определение предыдущей физической видимости
    // --------------------------------------------------

    const previousPhysicalVisibility =
      resolvePreviousPhysicalVisibility_(
        sheet,
        previousFilterVisibility,
        visibilityData.manuallyHiddenRows,
        numRows,
        resetVisibility,
        perfTrace
      );

    // --------------------------------------------------
    // 10. Обновление физической видимости и FilterState
    // --------------------------------------------------

    saveFilterVisibility_(
      sheet,
      previousPhysicalVisibility,
      newVisibility,
      newFilterVisibility,
      lastRow,
      resetVisibility,
      previousFilterVisibility,
      filterState,
      perfTrace
    );

    // --------------------------------------------------
    // 11. Итоговая информация
    // --------------------------------------------------

    const totalHidden =
      numRows - newVisibility.filter(Boolean).length;

    logFilterTiming(
      'APPLY FILTERS TOTAL',
      totalStarted
    );

    console.log(
      '[FILTER INFO] Sheet: ' + sheet.getName() +
      ', Rows: ' + numRows +
      ', Filter hidden: ' + visibilityData.filterRowsToHide.size +
      ', Manual hidden: ' + visibilityData.manualRowsToHide.size +
      ', Total hidden: ' + totalHidden +
      ', Visible: ' + (numRows - totalHidden)
    );

    triggerPerfEnd_(perfTrace, 'SUCCESS', {
      sheet: sheet.getName(),
      lastRow: lastRow,
      rows: numRows,
      filterHidden: visibilityData.filterRowsToHide.size,
      manualHidden: visibilityData.manualRowsToHide.size,
      hidden: totalHidden,
      visible: numRows - totalHidden
    });

  } catch (error) {
    status = 'ERROR';

    errorMetadata = {
      message: error.message,
      stack: error.stack
    };

    throw error;

  } finally {
    if (status === 'ERROR') {
      triggerPerfEnd_(
        perfTrace,
        status,
        errorMetadata
      );
    }
  }
}


/**
 * Подготавливает FilterState и ManualState.
 *
 * Использует единое чтение A1:Q2.
 * Поддерживает ранее подготовленный knownFilterState.
 */
function prepareFilterState_(
  sheet,
  numRows,
  knownFilterState,
  perfTrace
) {
  const hasKnownFilterState = Boolean(
    knownFilterState &&
    knownFilterState.stateSheet &&
    knownFilterState.columns &&
    Array.isArray(knownFilterState.metadata) &&
    Array.isArray(knownFilterState.previousVisibility) &&
    knownFilterState.previousVisibility.length >= numRows
  );

  let stepStarted = Date.now();

  const filterStateSheet = hasKnownFilterState
    ? knownFilterState.stateSheet
    : (
        FILTER_STATE_LAYOUT_CACHE ||
        getFilterStateSheet()
      );

  filterPerfStep_(
    perfTrace,
    'Prepare FilterState',
    stepStarted,
    {
      sheet: sheet.getName(),
      reused: hasKnownFilterState
    }
  );

  // Единое чтение метаданных A1:Q2.

  stepStarted = Date.now();

  const stateMetadataValues = filterStateSheet
    .getRange('A1:Q2')
    .getValues();

  if (!hasKnownFilterState) {
    ensureFilterStateLayout(
      filterStateSheet,
      stateMetadataValues
    );
  }

  const filterStateColumns = hasKnownFilterState
    ? knownFilterState.columns
    : getFilterStateColumns(sheet);

  const manualStateColumns = getManualStateColumns(sheet);

  const getMetadataBlock = startCol => [
    [
      stateMetadataValues[0][startCol - 1],
      stateMetadataValues[0][startCol]
    ],
    [
      stateMetadataValues[1][startCol - 1],
      stateMetadataValues[1][startCol]
    ]
  ];

  const manualStateMetadata = getMetadataBlock(
    manualStateColumns.keyCol
  );

  const filterStateMetadata = hasKnownFilterState
    ? knownFilterState.metadata
    : getMetadataBlock(
        filterStateColumns.visibleCol
      );

  filterPerfStep_(
    perfTrace,
    'Read combined state metadata',
    stepStarted,
    {
      range: 'A1:Q2',
      reusedFilterState: hasKnownFilterState
    }
  );

  // Загрузка ручных исключений.

  stepStarted = Date.now();

  const manualHiddenCounts = loadManualHiddenKeyCounts_(
    sheet,
    filterStateSheet,
    manualStateMetadata
  );

  filterPerfStep_(
    perfTrace,
    'Load ManualState',
    stepStarted,
    {
      manualKeys: manualHiddenCounts.size
    }
  );

  return {
    hasKnownFilterState: hasKnownFilterState,
    filterStateSheet: filterStateSheet,
    filterStateColumns: filterStateColumns,
    manualStateColumns: manualStateColumns,
    filterStateMetadata: filterStateMetadata,
    manualStateMetadata: manualStateMetadata,
    manualHiddenCounts: manualHiddenCounts
  };
}


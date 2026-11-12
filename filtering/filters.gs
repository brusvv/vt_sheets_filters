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

    if (autoSortEnabled) {
      // Используем уже прочитанные C:L для проверки порядка.
      inputData = readFilterInputData_(
        sheet,
        sheetLastRow,
        perfTrace
      );

      stepStarted = Date.now();

      sortChanged = sortScenarioBlocks(
        sheet,
        'scoreDesc',
        sheetLastRow,
        null,
        inputData.data
      );
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

    const visibilityData = buildFilterVisibility_(
      inputData.data,
      inputData.selectedDots,
      inputData.topN,
      hideEmptyScores,
      sheet,
      filterState.manualHiddenCounts,
      perfTrace
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


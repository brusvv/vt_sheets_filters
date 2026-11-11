
/**
 * FILTER VISIBILITY
 *
 * Управление фильтрационной и физической видимостью строк.
 *
 * Содержит:
 * - показ строк с данными;
 * - скрытие пустых строк;
 * - сравнение предыдущей и новой видимости;
 * - определение предыдущей физической видимости;
 * - применение изменений;
 * - сохранение FilterState.
 */


/**
 * Показывает строки с данными и скрывает пустые строки.
 * Синхронизирует фактическую видимость с _FilterState.
 *
 * Используется комплексными операциями из operations.gs.
 */
function showAllDataRows(
  sheet,
  knownLastRow = null,
  stateSnapshot = null
) {
  const lastRow = knownLastRow !== null
    ? knownLastRow
    : sheet.getLastRow();

  if (lastRow < CONFIG.firstRow) return [];

  const numRows = lastRow - CONFIG.firstRow + 1;

  const scenarios = sheet
    .getRange(
      CONFIG.firstRow,
      CONFIG.scenarioCol,
      numRows,
      1
    )
    .getValues();

  const visibility = scenarios.map(row => {
    return String(row[0] || '').trim() !== '';
  });

  const rowsToShow = [];
  const rowsToHide = [];

  visibility.forEach((isVisible, index) => {
    const sheetRow = CONFIG.firstRow + index;

    if (isVisible) {
      rowsToShow.push(sheetRow);
    } else {
      rowsToHide.push(sheetRow);
    }
  });

  showRowsInGroups(sheet, rowsToShow);
  hideRowsInGroups(sheet, new Set(rowsToHide));

  const stateSheet = stateSnapshot
    ? ensureFilterStateLayout()
    : null;

  const columns = stateSnapshot
    ? getFilterStateColumns(sheet)
    : null;

  updateVisibleRowsState(
    visibility,
    lastRow,
    sheet,
    null,
    stateSheet
  );

  if (stateSnapshot) {
    stateSnapshot.stateSheet = stateSheet;
    stateSnapshot.columns = columns;
    stateSnapshot.metadata = [
      [sheet.getName() + ' Visible', 'LastRow'],
      ['', lastRow]
    ];
    stateSnapshot.previousVisibility = visibility.slice();
  }

  return scenarios;
}


/**
 * Изменяет видимость только тех строк,
 * у которых новое состояние отличается от текущего.
 */
function updateRowsVisibility(
  sheet,
  previousVisibility,
  newVisibility
) {
  const rowsToShow = [];
  const rowsToHide = [];

  for (let i = 0; i < newVisibility.length; i++) {
    if (previousVisibility[i] === newVisibility[i]) continue;

    const sheetRow = CONFIG.firstRow + i;

    if (newVisibility[i]) {
      rowsToShow.push(sheetRow);
    } else {
      rowsToHide.push(sheetRow);
    }
  }

  showRowsInGroups(sheet, rowsToShow);
  hideRowsInGroups(sheet, new Set(rowsToHide));
}


/**
 * Определяет предыдущую физическую видимость.
 *
 * Проверяет только строки ManualState.
 * Сохраняет поведение ручного показа и скрытия строк.
 */
function resolvePreviousPhysicalVisibility_(
  sheet,
  previousFilterVisibility,
  manuallyHiddenRows,
  numRows,
  resetVisibility,
  perfTrace
) {
  const stepStarted = Date.now();

  let previousPhysicalVisibility;

  if (resetVisibility) {
    previousPhysicalVisibility = Array(numRows).fill(true);

  } else {
    previousPhysicalVisibility = previousFilterVisibility.map(
      (filterVisible, index) => {
        const sheetRow = CONFIG.firstRow + index;

        return (
          filterVisible &&
          !manuallyHiddenRows.has(sheetRow)
        );
      }
    );

    manuallyHiddenRows.forEach(sheetRow => {
      const index = sheetRow - CONFIG.firstRow;

      if (
        index >= 0 &&
        index < numRows &&
        sheet.isRowHiddenByUser(sheetRow)
      ) {
        previousPhysicalVisibility[index] = false;

      } else if (
        index >= 0 &&
        index < numRows
      ) {
        previousPhysicalVisibility[index] = true;
      }
    });
  }

  filterPerfStep_(
    perfTrace,
    'Resolve previous physical visibility',
    stepStarted,
    {
      manualRowsChecked: manuallyHiddenRows.size
    }
  );

  return previousPhysicalVisibility;
}


/**
 * Обновляет физическую видимость строк
 * и сохраняет FilterState.
 *
 * ManualState здесь не изменяется.
 */
function saveFilterVisibility_(
  sheet,
  previousPhysicalVisibility,
  newVisibility,
  newFilterVisibility,
  lastRow,
  resetVisibility,
  previousFilterVisibility,
  filterState,
  perfTrace
) {
  // --------------------------------------------------
  // Изменение физической видимости
  // --------------------------------------------------

  let stepStarted = Date.now();

  updateRowsVisibility(
    sheet,
    previousPhysicalVisibility,
    newVisibility
  );

  filterPerfStep_(
    perfTrace,
    'Update physical row visibility',
    stepStarted
  );

  // --------------------------------------------------
  // Сохранение только FilterState
  // --------------------------------------------------

  stepStarted = Date.now();

  updateVisibleRowsState(
    newFilterVisibility,
    lastRow,
    sheet,
    resetVisibility ? null : previousFilterVisibility,
    filterState.filterStateSheet,
    filterState.filterStateMetadata
  );

  filterPerfStep_(
    perfTrace,
    'Save FilterState',
    stepStarted,
    {
      sheet: sheet.getName()
    }
  );
}

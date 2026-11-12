/**
 * SORT OPERATIONS
 *
 * Комплексные операции ручной сортировки.
 */


/**
 * Ручная сортировка High Score:
 * от большего к меньшему.
 */
function manualSortByScoreDesc() {
  return runManualSortOperation_(
    'manualSortByScoreDesc',
    'scoreDesc',
    'High Score: highest to lowest'
  );
}


/**
 * Ручная сортировка High Score:
 * от меньшего к большему.
 */
function manualSortByScoreAsc() {
  return runManualSortOperation_(
    'manualSortByScoreAsc',
    'scoreAsc',
    'High Score: lowest to highest'
  );
}


/**
 * Ручная сортировка по Dots.
 */
function manualSortByDots() {
  return runManualSortOperation_(
    'manualSortByDots',
    'dots',
    'Dots: A-Z'
  );
}



/**
 * Создаёт ключ строки по всему содержимому C:L.
 *
 * Даты нормализуются отдельно, чтобы корректно
 * сопоставлять строки до и после сортировки.
 */
function getSortRowKey_(row) {
  return JSON.stringify(
    row.map(value => {
      if (value instanceof Date) {
        return ['date', value.getTime()];
      }

      return [typeof value, value];
    })
  );
}


/**
 * Сохраняет видимость строк в очередях,
 * сгруппированных по содержимому C:L.
 *
 * Очереди нужны для обработки полностью одинаковых строк.
 */
function buildSortVisibilityQueues_(
  data,
  visibility
) {
  const queues = new Map();

  data.forEach((row, index) => {
    const key = getSortRowKey_(row);

    if (!queues.has(key)) {
      queues.set(key, []);
    }

    queues.get(key).push(visibility[index]);
  });

  return queues;
}


/**
 * Восстанавливает видимость после сортировки.
 *
 * Сопоставляет отсортированные строки с исходными
 * по содержимому C:L.
 *
 * После восстановления:
 * - обновляет физическую видимость строк;
 * - синхронизирует _FilterState.
 */
function restoreSortedRowVisibility_(
  sheet,
  sortedData,
  visibilityQueues,
  originalVisibility,
  lastRow
) {
  const newVisibility = sortedData.map(row => {
    const key = getSortRowKey_(row);
    const queue = visibilityQueues.get(key);

    if (!queue || queue.length === 0) {
      throw new Error(
        'Не удалось сопоставить строку после сортировки: ' +
        key
      );
    }

    return queue.shift();
  });

  // Проверяем, что все исходные строки сопоставлены.
  for (const queue of visibilityQueues.values()) {
    if (queue.length > 0) {
      throw new Error(
        'После сортировки остались несопоставленные строки.'
      );
    }
  }

  const rowsToShow = [];
  const rowsToHide = [];

  newVisibility.forEach((isVisible, index) => {
    const sheetRow = CONFIG.firstRow + index;

    if (isVisible) {
      rowsToShow.push(sheetRow);
    } else {
      rowsToHide.push(sheetRow);
    }
  });

  // Восстанавливаем физическую видимость.
  showRowsInGroups(sheet, rowsToShow);
  hideRowsInGroups(sheet, new Set(rowsToHide));

  // Синхронизируем сохранённую видимость.
  updateVisibleRowsState(
    newVisibility,
    lastRow,
    sheet,
    originalVisibility
  );

  return newVisibility;
}



/**
 * Общая реализация ручной сортировки.
 *
 * Сохраняет видимость записей до сортировки
 * и восстанавливает её после перестановки строк.
 *
 * applyFilters() не запускается.
 */
function runManualSortOperation_(
  operation,
  sortType,
  toastMessage
) {
  const trace = triggerPerfStart_(
    operation,
    {
      sortType: sortType
    }
  );

  let status = 'SUCCESS';
  let error = null;

  try {
    return withProjectLock(function() {
      const spreadsheet =
        SpreadsheetApp.getActiveSpreadsheet();

      const sheet = getTargetSheet(
        null,
        spreadsheet
      );

      const sheetName = sheet.getName();

      // --------------------------------------------------
      // 1. Получение последней строки
      // --------------------------------------------------

      let stepStarted = Date.now();

      const lastRow = sheet.getLastRow();

      operationPerfStep_(
        trace,
        'Get last row',
        stepStarted,
        {
          sheet: sheetName,
          lastRow: lastRow
        }
      );

      if (lastRow < CONFIG.firstRow) {
        spreadsheet.toast(
          'No data to sort',
          'Sorting',
          3
        );

        return;
      }

      const numRows =
        lastRow - CONFIG.firstRow + 1;

      // --------------------------------------------------
      // 2. Сохранение исходных данных и видимости
      // --------------------------------------------------

      stepStarted = Date.now();

      const originalData = sheet
        .getRange(
          CONFIG.firstRow,
          CONFIG.sortStartCol,
          numRows,
          CONFIG.sortWidth
        )
        .getValues();

      const originalVisibility =
        getPreviousVisibility(
          sheet,
          numRows,
          lastRow
        );

      const visibilityQueues =
        buildSortVisibilityQueues_(
          originalData,
          originalVisibility
        );

      operationPerfStep_(
        trace,
        'Save original data and visibility',
        stepStarted,
        {
          sheet: sheetName,
          rows: numRows
        }
      );

      // --------------------------------------------------
      // 3. Сортировка сценариев
      // --------------------------------------------------

      stepStarted = Date.now();

      const sortChanged = sortScenarioBlocks(
        sheet,
        sortType,
        lastRow,
        null,
        originalData
      );

      SpreadsheetApp.flush();

      operationPerfStep_(
        trace,
        'Sort scenario blocks',
        stepStarted,
        {
          sheet: sheetName,
          sortType: sortType,
          lastRow: lastRow
        }
      );

      // --------------------------------------------------
      // 4. Чтение отсортированных данных
      // --------------------------------------------------

      stepStarted = Date.now();

      const sortedData = sortChanged
        ? sheet
            .getRange(
              CONFIG.firstRow,
              CONFIG.sortStartCol,
              numRows,
              CONFIG.sortWidth
            )
            .getValues()
        : originalData;

      operationPerfStep_(
        trace,
        'Read sorted data',
        stepStarted,
        {
          sheet: sheetName,
          rows: numRows,
          reusedOriginalData: !sortChanged
        }
      );

      // --------------------------------------------------
      // 5. Восстановление видимости
      // --------------------------------------------------

      stepStarted = Date.now();

      const newVisibility =
        restoreSortedRowVisibility_(
          sheet,
          sortedData,
          visibilityQueues,
          originalVisibility,
          lastRow
        );

      const visibleCount = newVisibility.filter(
        value => value === true
      ).length;

      const hiddenCount =
        newVisibility.length - visibleCount;

      operationPerfStep_(
        trace,
        'Restore row visibility',
        stepStarted,
        {
          sheet: sheetName,
          visibleRows: visibleCount,
          hiddenRows: hiddenCount
        }
      );

      // --------------------------------------------------
      // 6. Синхронизация таблицы
      // --------------------------------------------------

      stepStarted = Date.now();

      SpreadsheetApp.flush();

      operationPerfStep_(
        trace,
        'Spreadsheet flush',
        stepStarted
      );

      spreadsheet.toast(
        toastMessage,
        'Sorting',
        3
      );

    }, trace);

  } catch (caughtError) {
    status = 'ERROR';
    error = caughtError;

    throw caughtError;

  } finally {
    operationPerfEnd_(
      trace,
      status,
      error
    );
  }
}



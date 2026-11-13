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
/**
 * Manual sort entry point.
 *
 * It builds a validated visible-row plan but does not write any cell until
 * full C:L source -> staging -> destination fidelity is verified.
 */
/**
 * Manual sort entry point. The resolver produces the visibility snapshot;
 * the sorter only consumes it and never reinterprets physical visibility.
 *
 * No transfer is attempted until full C:L staging round-trip fidelity is
 * verified on a disposable Google workbook.
 */
function runManualSortOperation_(
  operation,
  sortType,
  toastMessage
) {
  const trace = triggerPerfStart_(
    operation,
    { sortType: sortType }
  );

  let status = 'SUCCESS';
  let error = null;

  try {
    const result = withProjectLock(function() {
      const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = getTargetSheet(null, spreadsheet);
      const sheetLastRow = sheet.getLastRow();

      operationPerfStep_(
        trace,
        'Get last row',
        Date.now(),
        { sheet: sheet.getName(), lastRow: sheetLastRow }
      );

      if (sheetLastRow < CONFIG.firstRow) {
        spreadsheet.toast('No data to sort', 'Sorting', 3);
        return { status: 'NO_OP', applied: false, reason: 'NO_DATA' };
      }

      const upperRowCount = sheetLastRow - CONFIG.firstRow + 1;
      const allData = sheet.getRange(
        CONFIG.firstRow,
        CONFIG.sortStartCol,
        upperRowCount,
        CONFIG.sortWidth
      ).getValues();

      const lastScenarioRow = getLastScenarioRowFromData(allData);
      if (lastScenarioRow < CONFIG.firstRow) {
        spreadsheet.toast('No data to sort', 'Sorting', 3);
        return {
          status: 'NO_OP',
          applied: false,
          reason: 'NO_SCENARIO_DATA'
        };
      }

      const rowCount = lastScenarioRow - CONFIG.firstRow + 1;
      const sortData = allData
        .slice(0, rowCount)
        .map(row => row.slice(0, CONFIG.sortWidth));

      const visibilityContext = readSortVisibilityContext_(
        sheet,
        lastScenarioRow,
        sortData
      );

      const sortResult = sortScenarioBlocksDetailed_(
        sheet,
        sortType,
        lastScenarioRow,
        null,
        sortData,
        visibilityContext
      );

      triggerPerfEvent_(trace, 'MANUAL_SORT_PLAN', {
        sheet: sheet.getName(),
        status: sortResult.status,
        reason: sortResult.reason || null,
        changed: Boolean(sortResult.plan && sortResult.plan.changed),
        applied: sortResult.applied === true
      });

      if (sortResult.status === 'TRANSFER_NOT_VERIFIED') {
        spreadsheet.toast(
          'Sorting not applied: full C:L staging transfer has not passed real Google Sheets round-trip verification.',
          'Sorting',
          6
        );
        status = 'TRANSFER_NOT_VERIFIED';

      } else if (sortResult.status === 'NO_OP') {
        spreadsheet.toast('Rows are already sorted.', 'Sorting', 3);

      } else if (sortResult.status !== 'SUCCESS') {
        spreadsheet.toast(
          'Sorting skipped: ' + (sortResult.reason || sortResult.status),
          'Sorting',
          5
        );
        status = sortResult.status;

      } else {
        spreadsheet.toast(toastMessage, 'Sorting', 3);
      }

      return sortResult;
    }, trace);

    return result;

  } catch (caughtError) {
    status = 'ERROR';
    error = caughtError;
    throw caughtError;

  } finally {
    operationPerfEnd_(trace, status, error);
  }
}

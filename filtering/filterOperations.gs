
/**
 * FILTER OPERATIONS
 *
 * Комплексные операции фильтрации.
 */


/**
 * Сбрасывает фильтры активного игрового листа.
 */
function resetFilters() {
  const trace = triggerPerfStart_('resetFilters');

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

      let stepStarted = Date.now();

      const optionsSheet =
        spreadsheet.getSheetByName(
          CONFIG.optionsSheetName
        );

      if (!optionsSheet) {
        throw new Error(
          `Лист "${CONFIG.optionsSheetName}" не найден.`
        );
      }

      const optionsValues =
        getRankScrollerOptions_(spreadsheet);

      operationPerfStep_(
        trace,
        'Read Options A5:A9',
        stepStarted,
        {
          sheet: sheet.getName()
        }
      );

      stepStarted = Date.now();

      sheet.getRange('I1:J1').clearContent();

      operationPerfStep_(
        trace,
        'Clear filters',
        stepStarted,
        {
          sheet: sheet.getName()
        }
      );

      stepStarted = Date.now();

      applyFilters(
        false,
        sheet,
        true,
        optionsValues,
        trace
      );

      SpreadsheetApp.flush();

      operationPerfStep_(
        trace,
        'Apply filters and flush',
        stepStarted,
        {
          sheet: sheet.getName()
        }
      );

      stepStarted = Date.now();

      rankScroller(
        false,
        spreadsheet,
        sheet.getName(),
        optionsValues,
        trace
      );

      operationPerfStep_(
        trace,
        'Rank Scroller',
        stepStarted,
        {
          sheet: sheet.getName()
        }
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



/**
 * Обновляет фильтры и, при необходимости,
 * Rank Scroller для одного листа.
 */
function refreshFiltersAndRanks(
  skipAutoSort = false,
  targetSheet = null,
  parentTrace = null,
  updateRankScroller = true
) {
  const trace = triggerPerfStart_(
    'refreshFiltersAndRanks',
    {
      skipAutoSort: skipAutoSort,
      updateRankScroller: updateRankScroller,
      sheet: targetSheet
        ? targetSheet.getName()
        : 'active'
    },
    parentTrace
  );

  let status = 'SUCCESS';
  let error = null;

  try {
    return withProjectLock(() => {
      const spreadsheet =
        SpreadsheetApp.getActiveSpreadsheet();

      const sheet = targetSheet ||
        spreadsheet.getActiveSheet();

      if (
        !RANK_TRACKED_SHEETS.includes(
          sheet.getName()
        )
      ) {
        triggerPerfEvent_(
          trace,
          'REFRESH_SKIPPED_INVALID_SHEET',
          {
            sheet: sheet.getName()
          }
        );

        return;
      }

      // Читаем Options один раз.
      // Полученные значения используются и фильтрами,
      // и Rank Scroller.
      let stepStarted = Date.now();

      const optionsValues =
        getRankScrollerOptions_(spreadsheet);

      operationPerfStep_(
        trace,
        'Read Options',
        stepStarted,
        {
          sheet: sheet.getName()
        }
      );

      // Применяем фильтры.
      stepStarted = Date.now();

      applyFilters(
        skipAutoSort,
        sheet,
        false,
        optionsValues,
        trace
      );

      operationPerfStep_(
        trace,
        'Apply filters',
        stepStarted,
        {
          sheet: sheet.getName()
        }
      );

      // Обновляем Rank Scroller, если требуется.
      if (updateRankScroller) {
        stepStarted = Date.now();

        SpreadsheetApp.flush();

        operationPerfStep_(
          trace,
          'Spreadsheet flush',
          stepStarted
        );

        stepStarted = Date.now();

        rankScroller(
          false,
          spreadsheet,
          sheet.getName(),
          optionsValues,
          trace
        );

        operationPerfStep_(
          trace,
          'Rank Scroller',
          stepStarted,
          {
            sheet: sheet.getName(),
            optionsProvided: true
          }
        );

      } else {
        triggerPerfEvent_(
          trace,
          'RANK_SCROLLER_SKIPPED',
          {
            reason: 'Not required for this edit',
            sheet: sheet.getName()
          }
        );
      }

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


/**
 * Обновляет фильтры на всех трёх листах
 * без автоматической сортировки
 * и без запуска Rank Scroller.
 */

function refreshAllFiltersOnly(parentTrace = null) {
  const trace = triggerPerfStart_(
    'refreshAllFiltersOnly',
    {},
    parentTrace
  );

  let status = 'SUCCESS';
  let error = null;

  try {
    return withProjectLock(() => {
      const spreadsheet =
        SpreadsheetApp.getActiveSpreadsheet();

      // Читаем настройки один раз для всех листов.
      const optionsValues =
        getRankScrollerOptions_(spreadsheet);

      RANK_TRACKED_SHEETS.forEach(sheetName => {
        const sheet =
          spreadsheet.getSheetByName(sheetName);

        if (!sheet) {
          throw new Error(
            `Лист "${sheetName}" не найден.`
          );
        }

        const stepStarted = Date.now();

        applyFilters(
          true,
          sheet,
          false,
          optionsValues,
          trace
        );

        operationPerfStep_(
          trace,
          `Apply filters: ${sheetName}`,
          stepStarted,
          {
            sheet: sheetName,
            skipAutoSort: true
          }
        );
      });

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



/**
 * Обновляет фильтры на всех трёх листах
 * и последовательно синхронизирует Rank Scroller.
 */

function refreshAllFiltersAndRanks(parentTrace = null) {
  const trace = triggerPerfStart_(
    'refreshAllFiltersAndRanks',
    {},
    parentTrace
  );

  let status = 'SUCCESS';
  let error = null;

  try {
    return withProjectLock(() => {
      const spreadsheet =
        SpreadsheetApp.getActiveSpreadsheet();

      // Читаем настройки один раз для всех листов
      // и последующего Rank Scroller.
      const optionsValues =
        getRankScrollerOptions_(spreadsheet);

      RANK_TRACKED_SHEETS.forEach(sheetName => {
        const sheet =
          spreadsheet.getSheetByName(sheetName);

        if (!sheet) {
          triggerPerfEvent_(
            trace,
            'FILTER_SHEET_NOT_FOUND',
            {
              sheet: sheetName
            }
          );

          return;
        }

        const stepStarted = Date.now();

        applyFilters(
          false,
          sheet,
          false,
          optionsValues,
          trace
        );

        operationPerfStep_(
          trace,
          `Apply filters: ${sheetName}`,
          stepStarted,
          {
            sheet: sheetName
          }
        );
      });

      let stepStarted = Date.now();

      SpreadsheetApp.flush();

      operationPerfStep_(
        trace,
        'Spreadsheet flush',
        stepStarted
      );

      stepStarted = Date.now();

      rankScroller(
        false,
        spreadsheet,
        null,
        optionsValues,
        trace
      );

      operationPerfStep_(
        trace,
        'Rank Scroller',
        stepStarted
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


/**
 * RANK OPERATIONS
 *
 * Комплексные операции обновления Rank Scroller.
 */


/**
 * Запуск Rank Scroller.
 *
 * forceAll = true:
 * принудительная проверка всех листов.
 *
 * forceAll = false:
 * проверка только изменённых листов.
 */
function runRankScroller(
  forceAll = false,
  parentTrace = null
) {
  const trace = triggerPerfStart_(
    'runRankScroller',
    {
      forceAll: forceAll
    },
    parentTrace
  );

  let status = 'SUCCESS';
  let error = null;

  try {
    return withProjectLock(() => {
      let stepStarted = Date.now();

      SpreadsheetApp.flush();

      operationPerfStep_(
        trace,
        'Spreadsheet flush before Rank Scroller',
        stepStarted
      );

      stepStarted = Date.now();

      const spreadsheet =
        SpreadsheetApp.getActiveSpreadsheet();

      operationPerfStep_(
        trace,
        'Get active spreadsheet',
        stepStarted
      );

      stepStarted = Date.now();

      rankScroller(
        forceAll,
        spreadsheet,
        null,
        null,
        trace
      );

      operationPerfStep_(
        trace,
        'Rank Scroller',
        stepStarted,
        {
          forceAll: forceAll
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
 * Принудительная полная синхронизация
 * всех трёх листов.
 */
function forceRankScroller() {
  const trace = triggerPerfStart_(
    'forceRankScroller'
  );

  let status = 'SUCCESS';
  let error = null;

  try {
    markAllRankSheetsDirty();

    triggerPerfEvent_(
      trace,
      'ALL_RANK_SHEETS_MARKED_DIRTY'
    );

    return runRankScroller(
      true,
      trace
    );

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

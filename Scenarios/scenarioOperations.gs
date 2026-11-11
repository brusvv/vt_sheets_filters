
/**
 * SCENARIO OPERATIONS
 *
 * Комплексные операции со сценариями.
 *
 * Ответственность:
 * - дублирование выбранной строки сценария;
 * - копирование данных, форматирования и проверок;
 * - управление объединениями категорий.
 */


/**
 * Дублирует выбранную строку сценария ниже неё.
 *
 * Валидация, копирование и обновление выполняются
 * под одной блокировкой.
 *
 * Использует единый Performance Tracer.
 *
 * Логика копирования:
 * - C:D — содержимое и формат;
 * - E — только формат;
 * - F:H — содержимое и формат;
 * - I:K — формат и проверки данных;
 * - L — содержимое и формат;
 * - высота строки;
 * - объединения категорий;
 * - перемещение курсора;
 * - flush.
 */
function duplicateScenario() {
  const trace = triggerPerfStart_(
    'duplicateScenario'
  );

  let status = 'SUCCESS';
  let error = null;

  try {
    return withProjectLock(function() {

      // -----------------------------------------
      // GET TARGET SHEET
      // -----------------------------------------

      let stepStarted = Date.now();

      const spreadsheet =
        SpreadsheetApp.getActiveSpreadsheet();

      const sheet = getTargetSheet(
        null,
        spreadsheet
      );

      operationPerfStep_(
        trace,
        'Get target sheet',
        stepStarted,
        {
          sheet: sheet.getName()
        }
      );


      // -----------------------------------------
      // VALIDATION
      // -----------------------------------------

      stepStarted = Date.now();

      const activeRange = sheet.getActiveRange();

      operationPerfStep_(
        trace,
        'Get active range',
        stepStarted
      );

      if (!activeRange) {
        triggerPerfEvent_(
          trace,
          'DUPLICATE_ABORTED',
          {
            reason: 'No active range'
          }
        );

        SpreadsheetApp.getUi().alert(
          'Сначала выбери строку сценария.'
        );

        return;
      }

      const sourceRow = activeRange.getRow();

      if (sourceRow < CONFIG.firstRow) {
        triggerPerfEvent_(
          trace,
          'DUPLICATE_ABORTED',
          {
            reason: 'Selected row is above scenario data',
            sourceRow: sourceRow
          }
        );

        SpreadsheetApp.getUi().alert(
          'Выбери строку сценария ниже заголовков.'
        );

        return;
      }

      stepStarted = Date.now();

      const scenario = String(
        sheet.getRange(
          sourceRow,
          CONFIG.scenarioCol
        ).getValue() || ''
      ).trim();

      operationPerfStep_(
        trace,
        'Read scenario',
        stepStarted,
        {
          sourceRow: sourceRow
        }
      );

      if (!scenario) {
        triggerPerfEvent_(
          trace,
          'DUPLICATE_ABORTED',
          {
            reason: 'Selected row has no scenario',
            sourceRow: sourceRow
          }
        );

        SpreadsheetApp.getUi().alert(
          'В выбранной строке нет сценария.'
        );

        return;
      }

      const targetRow = sourceRow + 1;


      // -----------------------------------------
      // SOURCE ROW SETTINGS
      // -----------------------------------------

      stepStarted = Date.now();

      const sourceHeight =
        sheet.getRowHeight(sourceRow);

      operationPerfStep_(
        trace,
        'Get row height',
        stepStarted
      );


      // -----------------------------------------
      // INSERT ROW
      // -----------------------------------------

      stepStarted = Date.now();

      sheet.insertRowAfter(sourceRow);

      operationPerfStep_(
        trace,
        'Insert row',
        stepStarted,
        {
          sourceRow: sourceRow,
          targetRow: targetRow
        }
      );


      // -----------------------------------------
      // COPY C:D
      // -----------------------------------------

      stepStarted = Date.now();

      sheet.getRange(
        sourceRow,
        3,
        1,
        2
      ).copyTo(
        sheet.getRange(
          targetRow,
          3,
          1,
          2
        ),
        {
          contentsOnly: false
        }
      );

      operationPerfStep_(
        trace,
        'Copy C:D',
        stepStarted
      );


      // -----------------------------------------
      // COPY E FORMAT ONLY
      // -----------------------------------------

      stepStarted = Date.now();

      sheet.getRange(
        sourceRow,
        CONFIG.scoreCol
      ).copyTo(
        sheet.getRange(
          targetRow,
          CONFIG.scoreCol
        ),
        {
          formatOnly: true
        }
      );

      operationPerfStep_(
        trace,
        'Copy E format',
        stepStarted
      );


      // -----------------------------------------
      // COPY F:H
      // -----------------------------------------

      stepStarted = Date.now();

      sheet.getRange(
        sourceRow,
        6,
        1,
        3
      ).copyTo(
        sheet.getRange(
          targetRow,
          6,
          1,
          3
        ),
        {
          contentsOnly: false
        }
      );

      operationPerfStep_(
        trace,
        'Copy F:H',
        stepStarted
      );


      // -----------------------------------------
      // COPY I:K FORMAT ONLY
      // -----------------------------------------

      stepStarted = Date.now();

      const sourceDotsRange = sheet.getRange(
        sourceRow,
        CONFIG.clearDotsStartCol,
        1,
        CONFIG.clearDotsEndCol -
          CONFIG.clearDotsStartCol + 1
      );

      const targetDotsRange = sheet.getRange(
        targetRow,
        CONFIG.clearDotsStartCol,
        1,
        CONFIG.clearDotsEndCol -
          CONFIG.clearDotsStartCol + 1
      );

      sourceDotsRange.copyTo(
        targetDotsRange,
        {
          formatOnly: true
        }
      );

      operationPerfStep_(
        trace,
        'Copy I:K format',
        stepStarted
      );


      // -----------------------------------------
      // COPY I:K DATA VALIDATIONS
      // -----------------------------------------

      stepStarted = Date.now();

      sourceDotsRange.copyTo(
        targetDotsRange,
        SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
        false
      );

      operationPerfStep_(
        trace,
        'Copy I:K validations',
        stepStarted
      );


      // -----------------------------------------
      // COPY L
      // -----------------------------------------

      stepStarted = Date.now();

      sheet.getRange(
        sourceRow,
        12
      ).copyTo(
        sheet.getRange(
          targetRow,
          12
        ),
        {
          contentsOnly: false
        }
      );

      operationPerfStep_(
        trace,
        'Copy L',
        stepStarted
      );


      // -----------------------------------------
      // COPY ROW HEIGHT
      // -----------------------------------------

      stepStarted = Date.now();

      sheet.setRowHeight(
        targetRow,
        sourceHeight
      );

      operationPerfStep_(
        trace,
        'Set row height',
        stepStarted
      );


      // -----------------------------------------
      // EXTEND CATEGORY MERGES
      // -----------------------------------------

      stepStarted = Date.now();

      extendCategoryMerges(
        sheet,
        sourceRow,
        targetRow
      );

      operationPerfStep_(
        trace,
        'Extend category merges',
        stepStarted
      );


      // -----------------------------------------
      // MOVE CURSOR
      // -----------------------------------------

      stepStarted = Date.now();

      sheet.setActiveRange(
        sheet.getRange(
          targetRow,
          CONFIG.scoreCol
        )
      );

      operationPerfStep_(
        trace,
        'Set active range',
        stepStarted
      );


      // -----------------------------------------
      // FLUSH
      // -----------------------------------------

      stepStarted = Date.now();

      SpreadsheetApp.flush();

      operationPerfStep_(
        trace,
        'Spreadsheet flush',
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


/**
 * @OnlyCurrentDoc
 */

/**
 * RANK SCROLLER 3.3
 *
 * Интеграция с единым Performance Tracer.
 *
 * Ответственность файла:
 * - оркестрация Rank Scroller;
 * - определение листов для обработки;
 * - получение настроек;
 * - определение необходимости обновления;
 * - вызов обработчика листа;
 * - подготовка и сохранение изменений состояния.
 *
 * Управление Script Properties и dirty-флагами
 * находится в RankScrollerState.gs.
 *
 * Обработка отдельных листов находится
 * в RankScrollerEngine.gs.
 *
 * rankScroller() — принудительная проверка всех листов.
 * rankScroller(false) — проверка только изменённых листов.
 *
 * Пятый аргумент parentTrace необязательный.
 */
function rankScroller(
  forceAll = true,
  spreadsheet = null,
  targetSheetName = null,
  providedOptionsValues = null,
  parentTrace = null
) {
  const trace = triggerPerfStart_(
    'rankScroller',
    {
      forceAll: forceAll,
      targetSheet: targetSheetName || 'none',
      spreadsheetProvided: Boolean(spreadsheet),
      optionsProvided: Boolean(providedOptionsValues)
    },
    parentTrace
  );

  let status = 'SUCCESS';
  let errorMetadata = {};

  const recordStep = function(name, startedAt, metadata = {}) {
    triggerPerfStep_(
      trace,
      name,
      startedAt,
      metadata
    );
  };

  const recordEvent = function(name, metadata = {}) {
    triggerPerfEvent_(
      trace,
      name,
      metadata
    );
  };

  try {
    const totalStartedAt = Date.now();

    // Fail closed before resolving or mutating any target sheet.
    const allowedSheetNames =
      validateRankScrollerConfiguration_();

    if (
      targetSheetName !== null &&
      !isRankScrollerSheetAllowed_(targetSheetName)
    ) {
      throw new Error(
        'Unknown Rank Scroller target sheet: ' +
        targetSheetName
      );
    }

    // Получение Spreadsheet.
    if (!spreadsheet) {
      const stepStartedAt = Date.now();

      spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

      recordStep(
        'Get active spreadsheet',
        stepStartedAt,
        {
          provided: false
        }
      );
    }

    // Получение Script Properties.
    // Все свойства читаются одним запросом
    // через RankScrollerState.gs.
    let stepStartedAt = Date.now();

    const properties =
      getRankScrollerProperties_();

    recordStep(
      'Read Script Properties',
      stepStartedAt,
      {
        propertiesCount: Object.keys(properties).length
      }
    );

    // Конфигурация листов и расположения диапазонов.
    const energyRanges = allowedSheetNames.map(
      sheetName => ({
        sheetName: sheetName,
        ...CONFIG.rankScroller.sheets[sheetName]
      })
    );

    // Проверка имени принудительно выбранного листа.
    stepStartedAt = Date.now();

    if (
      targetSheetName !== null &&
      !energyRanges.some(
        config => config.sheetName === targetSheetName
      )
    ) {
      throw new Error(
        'Unknown Rank Scroller target sheet: ' +
        targetSheetName
      );
    }

    recordStep(
      'Validate target sheet',
      stepStartedAt
    );

    // Получение Options.
    // Если настройки переданы вызывающей функцией,
    // повторно обращаться к листу Options не нужно.
    stepStartedAt = Date.now();

    let optionsValues = providedOptionsValues;

    if (!optionsValues) {
      optionsValues =
        getRankScrollerOptions_(spreadsheet);
    }

    recordStep(
      'Read Options',
      stepStartedAt,
      {
        provided: Boolean(providedOptionsValues)
      }
    );

    // Диагностика исходных значений Options.
    const rawOptions = {
      rankTiers: optionsValues?.[0]?.[0],
      autoSort: optionsValues?.[1]?.[0],
      hideEmptyScores: optionsValues?.[2]?.[0],
      columnsBehind: optionsValues?.[3]?.[0],
      columnsAhead: optionsValues?.[4]?.[0]
    };

    console.log(
      '[RANK SCROLLER] Raw Options: ' +
      JSON.stringify(rawOptions)
    );

    recordEvent(
      'RANK_SCROLLER_OPTIONS_RAW',
      rawOptions
    );

    const currentRankTiersValue =
      optionsValues[0][0];

    const colBehind =
      parseFloat(optionsValues[3][0]);

    const colAhead =
      parseFloat(optionsValues[4][0]);

    const visibleColumns =
      colBehind + colAhead + 1;

    // Диагностика разобранных значений.
    const parsedOptions = {
      rankTiers: currentRankTiersValue,
      rankTiersType: typeof currentRankTiersValue,
      colBehind: colBehind,
      colAhead: colAhead,
      visibleColumns: visibleColumns,
      validColBehind:
        Number.isFinite(colBehind) &&
        Number.isInteger(colBehind) &&
        colBehind >= 0,
      validColAhead:
        Number.isFinite(colAhead) &&
        Number.isInteger(colAhead) &&
        colAhead >= 0,
      validVisibleColumns: Number.isFinite(visibleColumns),
      integerVisibleColumns: Number.isInteger(visibleColumns),
      visibleColumnsAtLeastOne: visibleColumns >= 1
    };

    console.log(
      '[RANK SCROLLER] Parsed Options: ' +
      JSON.stringify(parsedOptions)
    );

    recordEvent(
      'RANK_SCROLLER_OPTIONS_PARSED',
      parsedOptions
    );

    // Проверка корректности настроек.
    if (
      !Number.isFinite(colBehind) ||
      !Number.isInteger(colBehind) ||
      colBehind < 0 ||
      !Number.isFinite(colAhead) ||
      !Number.isInteger(colAhead) ||
      colAhead < 0 ||
      !Number.isFinite(visibleColumns) ||
      !Number.isInteger(visibleColumns) ||
      visibleColumns < 1
    ) {
      const invalidFields = [];

      if (
        !Number.isFinite(colBehind) ||
        !Number.isInteger(colBehind) ||
        colBehind < 0
      ) {
        invalidFields.push(
          'Columns Behind (Options!A10) must be a non-negative integer: ' +
          String(optionsValues[3][0])
        );
      }

      if (
        !Number.isFinite(colAhead) ||
        !Number.isInteger(colAhead) ||
        colAhead < 0
      ) {
        invalidFields.push(
          'Columns Ahead (Options!A11) must be a non-negative integer: ' +
          String(optionsValues[4][0])
        );
      }

      if (!Number.isInteger(visibleColumns)) {
        invalidFields.push(
          'Visible Columns is not an integer: ' +
          visibleColumns
        );
      }

      if (visibleColumns < 1) {
        invalidFields.push(
          'Visible Columns must be at least 1: ' +
          visibleColumns
        );
      }

      throw new Error(
        'Invalid Rank Scroller Options. ' +
        invalidFields.join('; ')
      );
    }

    // Получение предыдущих настроек.
    const previousRankTiersValue =
      properties.rankTiersValue === 'true';

    const previousColBehindValue =
      parseFloat(properties.colBehindValue);

    const previousColAheadValue =
      parseFloat(properties.colAheadValue);

    const optionsChanged =
      currentRankTiersValue !== previousRankTiersValue ||
      colBehind !== previousColBehindValue ||
      colAhead !== previousColAheadValue;

    // Определяем необходимость полной проверки.
    const checkAll =
      forceAll || optionsChanged;

    recordEvent(
      'RANK_SCROLLER_CONFIGURATION',
      {
        forceAll: forceAll,
        optionsChanged: optionsChanged,
        checkAll: checkAll,
        targetSheet: targetSheetName || null,
        colBehind: colBehind,
        colAhead: colAhead,
        visibleColumns: visibleColumns,
        rankTiers: currentRankTiersValue
      }
    );

    const propertiesToSave = {};

    let processedSheets = 0;
    let skippedSheets = 0;

    // Обработка листов.
    const processStartedAt = Date.now();

    energyRanges.forEach(function(config) {
      // Проверяем dirty-флаг через State.
      const isDirty =
        isRankSheetDirty_(
          properties,
          config.sheetName
        );

      // Проверяем, является ли лист принудительно выбранным.
      const isTargetSheet =
        targetSheetName === config.sheetName;

      // Пропускаем неизменённые листы,
      // кроме принудительно выбранного.
      if (
        !checkAll &&
        !isDirty &&
        !isTargetSheet
      ) {
        skippedSheets++;

        recordEvent(
          'RANK_SCROLLER_SHEET_SKIPPED',
          {
            sheet: config.sheetName,
            reason: 'NOT_DIRTY'
          }
        );

        return;
      }

      const sheetStartedAt = Date.now();

      // Принудительно обновляем видимость, если:
      // - запущена полная проверка;
      // - лист помечен dirty;
      // - лист выбран как target.
      const forceUpdate =
        checkAll ||
        isDirty ||
        isTargetSheet;

      recordEvent(
        'RANK_SCROLLER_SHEET_STARTED',
        {
          sheet: config.sheetName,
          dirty: isDirty,
          forced: isTargetSheet,
          forceUpdate: forceUpdate
        }
      );

      const result =
        hideUnhideColumnsForNamedRange(
          config,
          colBehind,
          colAhead,
          visibleColumns,
          currentRankTiersValue,
          optionsChanged,
          spreadsheet,
          properties,
          trace,
          forceUpdate
        );

      if (!result.success) {
        throw new Error(
          'Rank Scroller failed for ' +
          config.sheetName
        );
      }

      // Сохраняем Energy только при фактическом изменении.
      if (result.energyValue !== undefined) {
        const previousEnergy =
          parseFloat(properties[config.propertyKey]);

        if (previousEnergy !== result.energyValue) {
          propertiesToSave[config.propertyKey] =
            String(result.energyValue);

          recordEvent(
            'RANK_SCROLLER_ENERGY_CHANGED',
            {
              sheet: config.sheetName,
              previousValue: previousEnergy,
              newValue: result.energyValue
            }
          );
        }
      }

      // Сбрасываем dirty только после успешной обработки.
      if (isDirty) {
        resetRankSheetDirty_(
          propertiesToSave,
          config.sheetName
        );

        recordEvent(
          'RANK_SCROLLER_DIRTY_RESET',
          {
            sheet: config.sheetName
          }
        );
      }

      processedSheets++;

      recordStep(
        'Process sheet: ' + config.sheetName,
        sheetStartedAt,
        {
          sheet: config.sheetName,
          energyValue: result.energyValue,
          dirty: isDirty,
          forced: isTargetSheet,
          forceUpdate: forceUpdate
        }
      );

      recordEvent(
        'RANK_SCROLLER_SHEET_COMPLETED',
        {
          sheet: config.sheetName,
          durationMs: Date.now() - sheetStartedAt,
          energyValue: result.energyValue
        }
      );
    });

    recordStep(
      'Process all sheets',
      processStartedAt,
      {
        processedSheets: processedSheets,
        skippedSheets: skippedSheets
      }
    );

    // Сохраняем настройки только при их изменении.
    stepStartedAt = Date.now();

    if (optionsChanged) {
      propertiesToSave.rankTiersValue =
        String(currentRankTiersValue);

      propertiesToSave.colBehindValue =
        String(colBehind);

      propertiesToSave.colAheadValue =
        String(colAhead);
    }

    recordStep(
      'Prepare properties',
      stepStartedAt,
      {
        optionsChanged: optionsChanged
      }
    );

    // Записываем свойства только при наличии изменений.
    // Фактическая запись выполняется через State.
    stepStartedAt = Date.now();

    const saveResult =
      saveRankScrollerProperties_(
        propertiesToSave
      );

    const keysToSave = saveResult.keys;

    if (saveResult.saved) {
      recordEvent(
        'RANK_SCROLLER_PROPERTIES_SAVED',
        {
          keys: keysToSave,
          count: keysToSave.length
        }
      );

    } else {
      recordEvent(
        'RANK_SCROLLER_PROPERTIES_SAVE_SKIPPED',
        {
          reason: 'NO_CHANGES'
        }
      );
    }

    recordStep(
      'Save Script Properties',
      stepStartedAt,
      {
        saved: saveResult.saved,
        keysCount: keysToSave.length
      }
    );

    recordEvent(
      'RANK_SCROLLER_COMPLETED',
      {
        durationMs: Date.now() - totalStartedAt,
        processedSheets: processedSheets,
        skippedSheets: skippedSheets,
        savedProperties: keysToSave.length
      }
    );

  } catch (error) {
    status = 'ERROR';

    errorMetadata = {
      message: error.message,
      stack: error.stack
    };

    triggerPerfEvent_(
      trace,
      'RANK_SCROLLER_ERROR',
      errorMetadata
    );

    throw error;

  } finally {
    triggerPerfEnd_(
      trace,
      status,
      errorMetadata
    );
  }
}


/**
 * Тест повторного запуска без изменений.
 */
function testRankScrollerNoChanges() {
  rankScroller(false);
}


/**
 * Тест расчёта целевой колонки.
 */
function testRankScrollerCalculator() {
  const result =
    calculateRankScrollerTargetColumn_(
      564,
      [
        100, 200, 300, 400,
        500, 600, 700, 800,
        900, 1000, 1100, 1200,
        1300, 1400, 1500, 1600
      ],
      14,
      16,
      4
    );

  console.log(
    '[RANK SCROLLER CALCULATOR TEST] ' +
    JSON.stringify(result)
  );
}


/**
 * Тест обработки листа Intermediate.
 */
function testRankScrollerIntermediate() {
  rankScroller(
    false,
    null,
    'Intermediate'
  );
}


/**
 * @OnlyCurrentDoc
 */

/**
 * RANK SCROLLER ENGINE
 *
 * Ответственность:
 * - обработка отдельного листа;
 * - чтение Energy и EnergyThreshold;
 * - проверка изменения Energy;
 * - определение необходимости обновления;
 * - вызов Calculator;
 * - вызов Visibility.
 *
 * Не управляет Script Properties.
 * Не изменяет dirty-флаги.
 * Не сохраняет состояние.
 *
 * Состояние обрабатывается в RankScrollerState.gs
 * и RankScroller.gs.
 *
 * Performance Tracer использует родительскую
 * трассировку rankScroller().
 */


/**
 * Главная функция обработки одного листа.
 *
 * Координирует чтение данных, проверку Energy
 * и обновление видимости колонок.
 *
 * Внешняя сигнатура сохранена для совместимости
 * с RankScroller.gs.
 *
 * @return {{
 *   success: boolean,
 *   energyValue?: number
 * }}
 */
function hideUnhideColumnsForNamedRange(
  config,
  colBehind,
  colAhead,
  visibleColumns,
  rankTiersValue,
  optionsChanged,
  spreadsheet,
  properties,
  trace = null,
  forceUpdate = false
) {
  // Читаем данные рабочего листа.
  const sheetData =
    readRankScrollerSheetData_(
      config,
      colBehind,
      spreadsheet,
      trace
    );

  if (!sheetData.success) {
    return { success: false };
  }

  // Проверяем изменение Energy и необходимость обновления.
  const energyCheck =
    checkRankScrollerEnergy_(
      config,
      sheetData.energyValue,
      optionsChanged,
      properties,
      forceUpdate,
      trace
    );

  if (!energyCheck.success) {
    return { success: false };
  }

  // Если обновление не требуется, возвращаем текущую Energy.
  if (!energyCheck.shouldUpdate) {
    return {
      success: true,
      energyValue: sheetData.energyValue
    };
  }

  // Обновляем видимость колонок.
  const visibilityResult =
    updateRankScrollerVisibility_(
      config,
      sheetData,
      visibleColumns,
      rankTiersValue,
      trace
    );

  if (!visibilityResult.success) {
    return { success: false };
  }

  return {
    success: true,
    energyValue: sheetData.energyValue
  };
}


/**
 * Получает рабочий лист и читает Energy с EnergyThreshold.
 *
 * Energy и EnergyThreshold читаются одним запросом
 * из общей строки M2:AB2 или M2:AC2.
 *
 * Корректировка Energy:
 * energyValue -= 20 * colBehind
 *
 * @return {{
 *   success: boolean,
 *   sheet?: Sheet,
 *   energyValue?: number,
 *   energyValues?: Array,
 *   startColumn?: number,
 *   totalColumns?: number
 * }}
 */
function readRankScrollerSheetData_(
  config,
  colBehind,
  spreadsheet,
  trace = null
) {
  const recordStep = function(name, startedAt, metadata = {}) {
    if (!trace) return;

    triggerPerfStep_(
      trace,
      name,
      startedAt,
      metadata
    );
  };

  const recordEvent = function(name, metadata = {}) {
    if (!trace) return;

    triggerPerfEvent_(
      trace,
      name,
      metadata
    );
  };

  // Reject unsupported names before asking Spreadsheet for a Sheet.
  if (
    !config ||
    !isRankScrollerSheetAllowed_(config.sheetName)
  ) {
    recordEvent(
      'RANK_SCROLLER_UNSUPPORTED_SHEET',
      {
        sheet: config && config.sheetName
          ? config.sheetName
          : null
      }
    );

    return { success: false };
  }

  // Получаем только канонический игровой лист.
  let stepStartedAt = Date.now();

  const sheet =
    spreadsheet.getSheetByName(config.sheetName);

  recordStep(
    'Get sheet: ' + config.sheetName,
    stepStartedAt,
    {
      sheet: config.sheetName
    }
  );

  if (
    !sheet ||
    sheet.getName() !== config.sheetName ||
    !isRankScrollerSheetAllowed_(sheet)
  ) {
    recordEvent(
      'RANK_SCROLLER_SHEET_NOT_FOUND',
      {
        sheet: config.sheetName
      }
    );

    return { success: false };
  }

  // Читаем Energy и EnergyThreshold одним запросом.
  stepStartedAt = Date.now();

  const rowValues = sheet.getRange(
    config.energyRow,
    config.energyColumn,
    1,
    config.thresholdColumns + 1
  ).getValues()[0];

  const rawEnergyValue = rowValues[0];

  const energyValues = rowValues.slice(
    1,
    config.thresholdColumns + 1
  );

  let energyValue =
    parseFloat(rawEnergyValue);

  if (!Number.isFinite(energyValue)) {
    recordEvent(
      'RANK_SCROLLER_INVALID_ENERGY',
      {
        sheet: config.sheetName,
        rawValue: rawEnergyValue
      }
    );

    recordStep(
      'Read Energy and thresholds: ' + config.sheetName,
      stepStartedAt,
      {
        sheet: config.sheetName,
        status: 'INVALID_ENERGY'
      }
    );

    return { success: false };
  }

  energyValue -= 20 * colBehind;

  recordStep(
    'Read Energy and thresholds: ' + config.sheetName,
    stepStartedAt,
    {
      sheet: config.sheetName,
      energyValue: energyValue,
      thresholdsCount: energyValues.length
    }
  );

  return {
    success: true,
    sheet: sheet,
    energyValue: energyValue,
    energyValues: energyValues,
    startColumn: config.thresholdStartColumn,
    totalColumns: config.thresholdColumns
  };
}


/**
 * Проверяет изменение Energy и определяет,
 * требуется ли обновление видимости колонок.
 *
 * Ранний выход разрешён только если:
 * - Energy не изменилась;
 * - настройки не изменились;
 * - принудительное обновление не запрошено.
 *
 * @return {{
 *   success: boolean,
 *   shouldUpdate?: boolean,
 *   energyChanged?: boolean,
 *   previousValue?: number
 * }}
 */
function checkRankScrollerEnergy_(
  config,
  energyValue,
  optionsChanged,
  properties,
  forceUpdate,
  trace = null
) {
  const recordStep = function(name, startedAt, metadata = {}) {
    if (!trace) return;

    triggerPerfStep_(
      trace,
      name,
      startedAt,
      metadata
    );
  };

  const recordEvent = function(name, metadata = {}) {
    if (!trace) return;

    triggerPerfEvent_(
      trace,
      name,
      metadata
    );
  };

  // Проверяем изменение Energy.
  const stepStartedAt = Date.now();

  const previousValue =
    parseFloat(properties[config.propertyKey]);

  const energyChanged =
    energyValue !== previousValue;

  recordStep(
    'Check Energy changes: ' + config.sheetName,
    stepStartedAt,
    {
      sheet: config.sheetName,
      energyValue: energyValue,
      previousValue: previousValue,
      changed: energyChanged,
      optionsChanged: optionsChanged,
      forceUpdate: forceUpdate
    }
  );

  // Ранний выход.
  if (
    !energyChanged &&
    !optionsChanged &&
    !forceUpdate
  ) {
    recordEvent(
      'RANK_SCROLLER_EARLY_EXIT',
      {
        sheet: config.sheetName,
        reason: 'ENERGY_UNCHANGED',
        energyValue: energyValue
      }
    );

    return {
      success: true,
      shouldUpdate: false,
      energyChanged: false,
      previousValue: previousValue
    };
  }

  return {
    success: true,
    shouldUpdate: true,
    energyChanged: energyChanged,
    previousValue: previousValue
  };
}



/**
 * Обновляет видимость колонок.
 *
 * В зависимости от rankTiersValue:
 * - false: показывает группы колонок, кратных 100;
 * - true: рассчитывает целевую колонку через Calculator
 *   и показывает заданное количество колонок.
 *
 * @return {{
 *   success: boolean,
 *   currentColumn?: number,
 *   groupsCount?: number
 * }}
 */
function updateRankScrollerVisibility_(
  config,
  sheetData,
  visibleColumns,
  rankTiersValue,
  trace = null
) {
  const recordStep = function(name, startedAt, metadata = {}) {
    if (!trace) return;

    triggerPerfStep_(
      trace,
      name,
      startedAt,
      metadata
    );
  };

  const recordEvent = function(name, metadata = {}) {
    if (!trace) return;

    triggerPerfEvent_(
      trace,
      name,
      metadata
    );
  };

  const sheet = sheetData.sheet;
  const energyValue = sheetData.energyValue;
  const energyValues = sheetData.energyValues;
  const startColumn = sheetData.startColumn;
  const totalColumns = sheetData.totalColumns;

  // Режим rankTiers = false.
  // visibleColumns здесь не используется.
  if (rankTiersValue === false) {
    const stepStartedAt = Date.now();

    const visibilityResult =
      showRankScrollerTierColumns_(
        sheet,
        startColumn,
        totalColumns,
        energyValues
      );

    if (!visibilityResult.success) {
      recordEvent(
        'RANK_SCROLLER_VISIBILITY_ERROR',
        {
          sheet: config.sheetName,
          mode: 'RANK_TIERS_DISABLED'
        }
      );

      return { success: false };
    }

    recordStep(
      'Update column visibility: ' + config.sheetName,
      stepStartedAt,
      {
        sheet: config.sheetName,
        mode: 'RANK_TIERS_DISABLED',
        groupsCount: visibilityResult.groupsCount
      }
    );

    recordEvent(
      'RANK_SCROLLER_VISIBILITY_UPDATED',
      {
        sheet: config.sheetName,
        mode: 'RANK_TIERS_DISABLED',
        groupsCount: visibilityResult.groupsCount
      }
    );

    return {
      success: true,
      groupsCount: visibilityResult.groupsCount
    };
  }

  // Проверяем количество отображаемых колонок
  // только в режиме Rank Tiers = true.
  if (
    visibleColumns > totalColumns ||
    visibleColumns < 1
  ) {
    recordEvent(
      'RANK_SCROLLER_INVALID_VISIBLE_COLUMNS',
      {
        sheet: config.sheetName,
        visibleColumns: visibleColumns,
        totalColumns: totalColumns
      }
    );

    return { success: false };
  }

  // Определяем текущую колонку.
  const calculationStartedAt = Date.now();

  const calculation =
    calculateRankScrollerTargetColumn_(
      energyValue,
      energyValues,
      startColumn,
      totalColumns,
      visibleColumns
    );

  if (!calculation.success) {
    recordEvent(
      'RANK_SCROLLER_TARGET_COLUMN_NOT_FOUND',
      {
        sheet: config.sheetName,
        energyValue: energyValue
      }
    );

    recordStep(
      'Calculate target column: ' + config.sheetName,
      calculationStartedAt,
      {
        sheet: config.sheetName,
        status: 'NOT_FOUND'
      }
    );

    return { success: false };
  }

  const currentColumn =
    calculation.currentColumn;

  recordStep(
    'Calculate target column: ' + config.sheetName,
    calculationStartedAt,
    {
      sheet: config.sheetName,
      currentColumn: currentColumn,
      visibleColumns: visibleColumns
    }
  );

  // Обновляем видимость колонок.
  const visibilityStartedAt = Date.now();

  const visibilityResult =
    showRankScrollerColumns_(
      sheet,
      startColumn,
      totalColumns,
      currentColumn,
      visibleColumns
    );

  if (!visibilityResult.success) {
    recordEvent(
      'RANK_SCROLLER_VISIBILITY_ERROR',
      {
        sheet: config.sheetName,
        mode: 'RANK_TIERS_ENABLED',
        currentColumn: currentColumn,
        visibleColumns: visibleColumns
      }
    );

    return { success: false };
  }

  recordStep(
    'Update column visibility: ' + config.sheetName,
    visibilityStartedAt,
    {
      sheet: config.sheetName,
      currentColumn: currentColumn,
      visibleColumns: visibleColumns
    }
  );

  recordEvent(
    'RANK_SCROLLER_VISIBILITY_UPDATED',
    {
      sheet: config.sheetName,
      currentColumn: currentColumn,
      visibleColumns: visibleColumns
    }
  );

  return {
    success: true,
    currentColumn: currentColumn
  };
}


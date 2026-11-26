/**
 * Energy mask - производная колонка в _FilterState.
 *
 * Формулы Energy читают маску, чтобы узнать, какие строки учитывать:
 *
 *   mask[i] = filterVisible[i] И строка i не сохранена в ManualState.
 *
 * Маска вычисляется из FilterState и ManualState и не хранится ни в одном из
 * них: столбец D и ключи ManualState не меняются. Ручные строки определяются
 * по физической позиции, а не по содержимому.
 *
 * Статус маски (ячейка statusRow) защищает от устаревших данных: OK - маска
 * соответствует последнему успешному применению фильтров или Save, STALE -
 * применение прервано, формулы не должны считать маску актуальной.
 *
 * Раскладка блока для листа (колонка из CONFIG.energyMask.columns):
 *   строка 1: заголовок "<лист> Energy Mask"
 *   строка 2: статус (OK / STALE)
 *   строка 3+: булевы значения по строкам данных
 * Колонка versionCol (R) в строке 1 хранит метку версии схемы.
 *
 * Модуль только создаёт и пишет блок. Подключение к applyFilters и
 * saveManualHiddenRows выполняется на следующих этапах.
 */

/**
 * Включена ли запись маски для листа.
 *
 * @param {Sheet|string} sheet
 * @return {boolean}
 */
function isEnergyMaskEnabledForSheet_(sheet) {
  const settings = CONFIG.energyMask;

  if (!settings || settings.enabled !== true) {
    return false;
  }

  const sheetName = getEnergyMaskSheetName_(sheet);

  return Array.isArray(settings.sheets) &&
    settings.sheets.indexOf(sheetName) !== -1 &&
    Boolean(settings.columns[sheetName]);
}

/**
 * Номер колонки маски для листа.
 *
 * @param {Sheet|string} sheet
 * @return {number}
 */
function getEnergyMaskColumn_(sheet) {
  const sheetName = getEnergyMaskSheetName_(sheet);
  const column = CONFIG.energyMask.columns[sheetName];

  if (!column) {
    throw new Error(
      `Нет блока Energy mask для листа "${sheetName}".`
    );
  }

  return column;
}

/**
 * Вычисляет маску Energy для строк данных.
 *
 * @param {boolean[]} filterVisibility Результат фильтров по строкам данных.
 * @param {Set<number>|number[]|null} manualRowsToHide Физические номера строк,
 *   сохранённых в ManualState.
 * @param {number} firstRow Номер первой строки данных.
 * @return {boolean[]}
 */
function computeEnergyMask_(filterVisibility, manualRowsToHide, firstRow) {
  if (!Array.isArray(filterVisibility)) {
    throw new TypeError('filterVisibility должен быть массивом.');
  }

  if (!Number.isInteger(firstRow) || firstRow < 1) {
    throw new TypeError('firstRow должен быть положительным целым числом.');
  }

  const manualRows = manualRowsToHide instanceof Set
    ? manualRowsToHide
    : new Set(manualRowsToHide || []);

  return filterVisibility.map(function(visible, index) {
    return visible === true && !manualRows.has(firstRow + index);
  });
}

/**
 * Проверяет блок маски и создаёт его, если он ещё не создан.
 * При неизвестной версии схемы или чужом заголовке ничего не записывает.
 *
 * @param {Sheet} stateSheet Лист _FilterState.
 * @param {Sheet|string} sheet Игровой лист.
 * @return {{ok: boolean, created: boolean, reason: (string|null)}}
 */
function ensureEnergyMaskLayout_(stateSheet, sheet) {
  const settings = CONFIG.energyMask;
  const sheetName = getEnergyMaskSheetName_(sheet);
  const maskCol = getEnergyMaskColumn_(sheetName);
  const versionLabel = getEnergyMaskVersionLabel_();
  const headerLabel = `${sheetName} Energy Mask`;

  const lastCol = getEnergyMaskLastColumn_();

  if (stateSheet.getMaxColumns() < lastCol) {
    stateSheet.insertColumnsAfter(
      stateSheet.getMaxColumns(),
      lastCol - stateSheet.getMaxColumns()
    );
  }

  // Одним чтением: версия (R1) и заголовки колонок маски (S1:U1).
  const headerRow = stateSheet
    .getRange(
      settings.headerRow,
      settings.versionCol,
      1,
      lastCol - settings.versionCol + 1
    )
    .getValues()[0];

  const foundVersion = headerRow[0];
  const foundHeader = headerRow[maskCol - settings.versionCol];

  if (!isEmptyCellValue_(foundVersion) && foundVersion !== versionLabel) {
    return {
      ok: false,
      created: false,
      reason: 'SCHEMA_VERSION_MISMATCH'
    };
  }

  if (!isEmptyCellValue_(foundHeader) && foundHeader !== headerLabel) {
    return {
      ok: false,
      created: false,
      reason: 'HEADER_MISMATCH'
    };
  }

  let created = false;

  if (isEmptyCellValue_(foundVersion)) {
    stateSheet
      .getRange(settings.headerRow, settings.versionCol)
      .setValue(versionLabel);

    created = true;
  }

  if (isEmptyCellValue_(foundHeader)) {
    stateSheet
      .getRange(settings.headerRow, maskCol)
      .setValue(headerLabel);

    created = true;
  }

  return {
    ok: true,
    created: created,
    reason: null
  };
}

/**
 * Читает статус маски листа.
 *
 * @param {Sheet} stateSheet
 * @param {Sheet|string} sheet
 * @return {{status: *, ok: boolean, stale: boolean}}
 */
function readEnergyMaskStatus_(stateSheet, sheet) {
  const settings = CONFIG.energyMask;
  const status = stateSheet
    .getRange(settings.statusRow, getEnergyMaskColumn_(sheet))
    .getValue();

  return {
    status: status,
    ok: status === settings.statusOk,
    stale: status === settings.statusStale
  };
}

/**
 * Читает значения маски для строк данных (для диагностики и тестов).
 *
 * @param {Sheet} stateSheet
 * @param {Sheet|string} sheet
 * @param {number} lastRow Последняя строка данных.
 * @return {{valid: boolean, values: Array}}
 */
function readEnergyMask_(stateSheet, sheet, lastRow) {
  const rowCount = lastRow - CONFIG.firstRow + 1;

  if (rowCount <= 0) {
    return {valid: true, values: []};
  }

  const values = stateSheet
    .getRange(CONFIG.firstRow, getEnergyMaskColumn_(sheet), rowCount, 1)
    .getValues()
    .map(function(row) {
      return row[0];
    });

  return {
    valid: values.every(function(value) {
      return typeof value === 'boolean';
    }),
    values: values
  };
}

/**
 * Записывает маску и статус OK одной операцией setValues, чтобы формула не
 * увидела новую маску со старым статусом. Значения ниже lastRow очищаются.
 *
 * При неизвестной схеме или чужом заголовке ничего не записывает.
 *
 * @param {Sheet} stateSheet
 * @param {Sheet|string} sheet
 * @param {boolean[]} mask Значения для строк firstRow..lastRow.
 * @param {number} lastRow Последняя строка данных.
 * @return {{ok: boolean, reason: (string|null), rowsWritten: number,
 *   rowsCleared: number, layoutCreated: boolean}}
 */
function writeEnergyMask_(stateSheet, sheet, mask, lastRow) {
  const settings = CONFIG.energyMask;
  const firstRow = CONFIG.firstRow;

  if (settings.statusRow + 1 !== firstRow) {
    throw new Error(
      'Строка статуса Energy mask должна идти непосредственно перед ' +
      'первой строкой данных.'
    );
  }

  const rowCount = lastRow - firstRow + 1;

  if (!Array.isArray(mask) || mask.length !== Math.max(rowCount, 0)) {
    throw new TypeError(
      'Длина маски должна совпадать с числом строк данных.'
    );
  }

  const nonBoolean = mask.some(function(value) {
    return typeof value !== 'boolean';
  });

  if (nonBoolean) {
    throw new TypeError('Маска должна содержать только boolean.');
  }

  const layout = ensureEnergyMaskLayout_(stateSheet, sheet);

  if (!layout.ok) {
    return {
      ok: false,
      reason: layout.reason,
      rowsWritten: 0,
      rowsCleared: 0,
      layoutCreated: false
    };
  }

  const maskCol = getEnergyMaskColumn_(sheet);

  const values = [[settings.statusOk]].concat(
    mask.map(function(value) {
      return [value];
    })
  );

  stateSheet
    .getRange(settings.statusRow, maskCol, values.length, 1)
    .setValues(values);

  const tailStart = Math.max(lastRow + 1, firstRow);
  const maxRows = stateSheet.getMaxRows();
  let rowsCleared = 0;

  if (maxRows >= tailStart) {
    rowsCleared = maxRows - tailStart + 1;

    stateSheet
      .getRange(tailStart, maskCol, rowsCleared, 1)
      .clearContent();
  }

  return {
    ok: true,
    reason: null,
    rowsWritten: mask.length,
    rowsCleared: rowsCleared,
    layoutCreated: layout.created
  };
}

/**
 * Помечает маску листа как устаревшую (STALE). Меняет только ячейку статуса.
 * Не бросает исключений: используется на путях сбоя.
 *
 * @param {Sheet} stateSheet
 * @param {Sheet|string} sheet
 * @return {{ok: boolean, reason: (string|null)}}
 */
function markEnergyMaskStale_(stateSheet, sheet) {
  try {
    const layout = ensureEnergyMaskLayout_(stateSheet, sheet);

    if (!layout.ok) {
      return {ok: false, reason: layout.reason};
    }

    stateSheet
      .getRange(
        CONFIG.energyMask.statusRow,
        getEnergyMaskColumn_(sheet)
      )
      .setValue(CONFIG.energyMask.statusStale);

    return {ok: true, reason: null};
  } catch (error) {
    return {
      ok: false,
      reason: 'ERROR',
      message: error && error.message ? error.message : String(error)
    };
  }
}

/**
 * Вычисляет и записывает маску, если запись включена для листа.
 * При выключенном режиме не обращается к листам вообще.
 *
 * @param {Sheet} stateSheet
 * @param {Sheet|string} sheet
 * @param {boolean[]} filterVisibility
 * @param {Set<number>|number[]|null} manualRowsToHide
 * @param {number} lastRow
 * @return {{skipped: boolean, reason: (string|null), ok: boolean}}
 */
function syncEnergyMaskIfEnabled_(
  stateSheet,
  sheet,
  filterVisibility,
  manualRowsToHide,
  lastRow
) {
  if (!isEnergyMaskEnabledForSheet_(sheet)) {
    return {skipped: true, reason: 'DISABLED', ok: true};
  }

  const mask = computeEnergyMask_(
    filterVisibility,
    manualRowsToHide,
    CONFIG.firstRow
  );

  const result = writeEnergyMask_(stateSheet, sheet, mask, lastRow);

  return {
    skipped: false,
    reason: result.reason,
    ok: result.ok
  };
}

// ---------------------------------------------------------------------------
// Точки интеграции (R5.2). Все функции ниже не бросают исключений: ошибка
// Energy mask не должна ломать фильтрацию и сохранение ручных строк.
// При выключенном режиме для листа они не обращаются ни к одному листу.
// ---------------------------------------------------------------------------

/**
 * После успешного обновления FilterState строит маску из уже подготовленного
 * результата фильтрации и сохранённых ManualState-строк и записывает её.
 * Если маску построить или записать нельзя, помечает её STALE.
 *
 * @param {Sheet} sheet Игровой лист.
 * @param {Sheet|null} filterStateSheet Лист _FilterState (или null).
 * @param {boolean[]} filterVisibility Результат фильтров по строкам данных.
 * @param {Set<number>|number[]|null} manualRowsToHide Физические номера
 *   сохранённых ручных строк.
 * @param {number} lastRow Последняя строка данных.
 * @param {Object|null} perfTrace
 * @return {{skipped: boolean, ok: boolean, reason: (string|null)}}
 */
function syncEnergyMaskAfterFilters_(
  sheet,
  filterStateSheet,
  filterVisibility,
  manualRowsToHide,
  lastRow,
  perfTrace
) {
  if (!sheet || !isEnergyMaskEnabledForSheet_(sheet)) {
    return {skipped: true, ok: true, reason: 'DISABLED'};
  }

  try {
    const stateSheet = filterStateSheet || getEnergyMaskStateSheet_(sheet);

    if (!stateSheet) {
      throw new Error('Лист _FilterState не найден.');
    }

    const mask = computeEnergyMask_(
      filterVisibility,
      manualRowsToHide,
      CONFIG.firstRow
    );

    const result = writeEnergyMask_(stateSheet, sheet, mask, lastRow);

    if (!result.ok) {
      return failEnergyMaskSync_(
        sheet,
        stateSheet,
        result.reason,
        'ENERGY_MASK_SYNC_FAILED',
        perfTrace
      );
    }

    return {skipped: false, ok: true, reason: null};

  } catch (error) {
    return failEnergyMaskSync_(
      sheet,
      filterStateSheet,
      'ERROR: ' + (error && error.message ? error.message : String(error)),
      'ENERGY_MASK_SYNC_FAILED',
      perfTrace
    );
  }
}

/**
 * Помечает маску листа как STALE по месту вызова (досрочный выход или
 * исключение). Не бросает исключений.
 *
 * @param {Sheet|null} sheet
 * @param {string} reason
 * @param {Object|null} perfTrace
 * @param {Sheet|null} stateSheet Лист _FilterState, если уже известен.
 * @return {{skipped: boolean, ok: boolean, reason: (string|null)}}
 */
function markEnergyMaskStaleForSheet_(sheet, reason, perfTrace, stateSheet) {
  if (!sheet || !isEnergyMaskEnabledForSheet_(sheet)) {
    return {skipped: true, ok: true, reason: 'DISABLED'};
  }

  try {
    const target = stateSheet || getEnergyMaskStateSheet_(sheet);

    if (!target) {
      emitEnergyMaskEvent_(perfTrace, 'ENERGY_MASK_MARK_STALE_FAILED', {
        sheet: getEnergyMaskSheetName_(sheet),
        reason: reason,
        detail: 'NO_STATE_SHEET'
      });

      return {skipped: false, ok: false, reason: 'NO_STATE_SHEET'};
    }

    const result = markEnergyMaskStale_(target, sheet);

    emitEnergyMaskEvent_(
      perfTrace,
      result.ok ? 'ENERGY_MASK_MARKED_STALE' : 'ENERGY_MASK_MARK_STALE_FAILED',
      {
        sheet: getEnergyMaskSheetName_(sheet),
        reason: reason,
        detail: result.reason
      }
    );

    return {skipped: false, ok: result.ok, reason: result.reason};

  } catch (error) {
    return {
      skipped: false,
      ok: false,
      reason: 'ERROR: ' + (error && error.message ? error.message : String(error))
    };
  }
}

/**
 * Save Manual Hidden Rows: вызывается после проверок и ДО записи ManualState.
 * Если маска была актуальна (OK), помечает её STALE: при любом сбое Save
 * маска так и останется STALE и не будет выдана за актуальную. Если маска уже
 * STALE или не создана, Save её не трогает.
 *
 * @param {Sheet} sheet
 * @param {Sheet|null} filterStateSheet
 * @param {Object|null} perfTrace
 * @return {{active: boolean, reason: (string|null)}} Токен для
 *   finishEnergyMaskManualSave_.
 */
function beginEnergyMaskManualSave_(sheet, filterStateSheet, perfTrace) {
  if (!sheet || !isEnergyMaskEnabledForSheet_(sheet)) {
    return {active: false, reason: 'DISABLED'};
  }

  try {
    const stateSheet = filterStateSheet || getEnergyMaskStateSheet_(sheet);

    if (!stateSheet) {
      return {active: false, reason: 'NO_STATE_SHEET'};
    }

    const status = readEnergyMaskStatus_(stateSheet, sheet);

    if (!status.ok) {
      return {
        active: false,
        reason: status.stale ? 'MASK_STALE' : 'MASK_NOT_INITIALIZED'
      };
    }

    const marked = markEnergyMaskStale_(stateSheet, sheet);

    if (!marked.ok) {
      emitEnergyMaskEvent_(perfTrace, 'ENERGY_MASK_MARK_STALE_FAILED', {
        sheet: getEnergyMaskSheetName_(sheet),
        reason: 'MANUAL_SAVE_BEGIN',
        detail: marked.reason
      });

      return {active: false, reason: 'MARK_STALE_FAILED', detail: marked.reason};
    }

    return {active: true, reason: null};

  } catch (error) {
    return {
      active: false,
      reason: 'ERROR: ' + (error && error.message ? error.message : String(error))
    };
  }
}

/**
 * Save Manual Hidden Rows: вызывается ПОСЛЕ успешной записи ManualState.
 * Пересчитывает маску как filterVisible AND NOT savedManualRow и записывает
 * её со статусом OK. Не бросает исключений: ключи ManualState к этому моменту
 * уже сохранены и не зависят от результата.
 *
 * @param {{active: boolean, reason: (string|null)}} guard Токен из
 *   beginEnergyMaskManualSave_.
 * @param {Sheet} sheet
 * @param {Sheet|null} filterStateSheet
 * @param {boolean[]} filterVisibility Значения _FilterState!D на момент Save.
 * @param {number[]} savedManualRows Физические номера сохранённых строк.
 * @param {number} lastRow
 * @param {Object|null} perfTrace
 * @return {{skipped: boolean, ok: boolean, reason: (string|null)}}
 */
function finishEnergyMaskManualSave_(
  guard,
  sheet,
  filterStateSheet,
  filterVisibility,
  savedManualRows,
  lastRow,
  perfTrace
) {
  if (!guard || guard.reason === 'DISABLED') {
    return {skipped: true, ok: true, reason: 'DISABLED'};
  }

  if (!guard.active) {
    if (guard.reason === 'MASK_STALE') {
      notifyEnergyMaskStale_();
    }

    return {skipped: true, ok: true, reason: guard.reason};
  }

  try {
    const stateSheet = filterStateSheet || getEnergyMaskStateSheet_(sheet);

    if (!stateSheet) {
      throw new Error('Лист _FilterState не найден.');
    }

    const mask = computeEnergyMask_(
      filterVisibility,
      savedManualRows,
      CONFIG.firstRow
    );

    const result = writeEnergyMask_(stateSheet, sheet, mask, lastRow);

    if (!result.ok) {
      return failEnergyMaskSync_(
        sheet,
        stateSheet,
        result.reason,
        'ENERGY_MASK_SYNC_FAILED',
        perfTrace
      );
    }

    return {skipped: false, ok: true, reason: null};

  } catch (error) {
    return failEnergyMaskSync_(
      sheet,
      filterStateSheet,
      'ERROR: ' + (error && error.message ? error.message : String(error)),
      'ENERGY_MASK_SYNC_FAILED',
      perfTrace
    );
  }
}

function failEnergyMaskSync_(sheet, stateSheet, reason, eventName, perfTrace) {
  emitEnergyMaskEvent_(perfTrace, eventName, {
    sheet: getEnergyMaskSheetName_(sheet),
    reason: reason
  });

  console.warn(
    '[ENERGY MASK] ' + getEnergyMaskSheetName_(sheet) + ': ' + reason
  );

  // Маску нельзя считать актуальной: пытаемся пометить STALE.
  markEnergyMaskStaleForSheet_(sheet, reason, perfTrace, stateSheet || null);

  return {skipped: false, ok: false, reason: reason};
}

function getEnergyMaskStateSheet_(sheet) {
  try {
    const spreadsheet = typeof sheet.getParent === 'function'
      ? sheet.getParent()
      : SpreadsheetApp.getActiveSpreadsheet();

    return spreadsheet.getSheetByName(CONFIG.energyMask.stateSheetName) || null;
  } catch (error) {
    return null;
  }
}

function emitEnergyMaskEvent_(perfTrace, name, metadata) {
  try {
    if (perfTrace && typeof triggerPerfEvent_ === 'function') {
      triggerPerfEvent_(perfTrace, name, metadata);
    }
  } catch (error) {
    // Диагностика не должна влиять на работу.
  }
}

function notifyEnergyMaskStale_() {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Energy mask устарела: примените фильтры, чтобы обновить Energy.',
      'Manual exclusions',
      5
    );
  } catch (error) {
    // Уведомление необязательно.
  }
}

function getEnergyMaskSheetName_(sheet) {
  return typeof sheet === 'string' ? sheet : sheet.getName();
}

function getEnergyMaskVersionLabel_() {
  return `EnergyMask:${CONFIG.energyMask.schemaVersion}`;
}

function getEnergyMaskLastColumn_() {
  const columns = CONFIG.energyMask.columns;

  return Math.max.apply(
    null,
    Object.keys(columns).map(function(name) {
      return columns[name];
    })
  );
}

function isEmptyCellValue_(value) {
  return value === '' || value === null || value === undefined;
}


/**
 * MANUAL STATE
 *
 * Хранит результаты, скрытые пользователем вручную.
 *
 * Novice       -> J:K
 * Intermediate -> M:N
 * Advanced     -> P:Q
 *
 * Сохраняет идентификаторы записей по значениям C:L,
 * а не по номерам строк.
 */

function getManualStateColumns(sheet) {
  const sheetName = typeof sheet === 'string'
    ? sheet
    : sheet.getName();

  const columns = {
    Novice: 10,       // J:K
    Intermediate: 13, // M:N
    Advanced: 16      // P:Q
  };

  const startCol = columns[sheetName];

  if (!startCol) {
    throw new Error(
      `Нет блока ManualState для листа "${sheetName}".`
    );
  }

  return {
    keyCol: startCol,
    countCol: startCol + 1
  };
}


/**
 * Создаёт стабильный ключ записи по значениям C:L.
 */
function getManualRecordKey_(row) {
  const normalized = row.map(value => {
    if (value instanceof Date) {
      return ['date', value.toISOString()];
    }

    if (value === null || value === undefined) {
      return ['null', ''];
    }

    return [typeof value, value];
  });

  return JSON.stringify(normalized);
}



/**
 * Загружает ручные исключения в Map.
 *
 * Map хранит количество исключённых записей
 * с одинаковым содержимым.
 *
 * stateSheet и metadata можно передать заранее,
 * чтобы избежать повторного чтения метаданных.
 */
function loadManualHiddenKeyCounts_(
  sheet,
  stateSheet = null,
  metadata = null
) {
  stateSheet = stateSheet || ensureFilterStateLayout();

  const columns = getManualStateColumns(sheet);

  if (!metadata) {
    metadata = stateSheet
      .getRange(1, columns.keyCol, 2, 2)
      .getValues();
  }

  const count = Number(metadata[1][1]);

  if (
    !Number.isInteger(count) ||
    count <= 0
  ) {
    return new Map();
  }

  const keys = stateSheet
    .getRange(
      CONFIG.firstRow,
      columns.keyCol,
      count,
      1
    )
    .getValues();

  const counts = new Map();

  keys.forEach(row => {
    const key = row[0];

    if (typeof key !== 'string' || !key) return;

    counts.set(
      key,
      (counts.get(key) || 0) + 1
    );
  });

  return counts;
}


/**
 * Проверяет, является ли запись ручным исключением.
 *
 * При совпадении уменьшает оставшееся количество
 * одинаковых исключённых записей.
 */
function consumeManualHiddenKey_(row, counts) {
  const key = getManualRecordKey_(row);
  const count = counts.get(key) || 0;

  if (count <= 0) return false;

  if (count === 1) {
    counts.delete(key);
  } else {
    counts.set(key, count - 1);
  }

  return true;
}



/**
 * Сохраняет вручную скрытые строки.
 *
 * Важно:
 * перед запуском необходимо применить текущие фильтры,
 * чтобы FilterState отражал их актуальное состояние.
 *
 * Скрытые фильтрами строки не записываются в ManualState.
 */
function saveManualHiddenRows() {
  return withProjectLock(function() {
    const sheet = getTargetSheet();
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    // --------------------------------------------------
    // 1. Чтение данных и определение последней строки
    // --------------------------------------------------

    const sheetLastRow = sheet.getLastRow();

    if (sheetLastRow < CONFIG.firstRow) {
      spreadsheet.toast(
        'No data rows found.',
        'Manual exclusions',
        3
      );
      return;
    }

    // Одним запросом читаем C:L.
    const allData = sheet.getRange(
      CONFIG.firstRow,
      CONFIG.sortStartCol,
      sheetLastRow - CONFIG.firstRow + 1,
      CONFIG.sortWidth
    ).getValues();

    // Последняя игровая строка определяется по столбцу C.
    const lastRow = getLastScenarioRowFromData(allData);

    if (lastRow < CONFIG.firstRow) {
      spreadsheet.toast(
        'No data rows found.',
        'Manual exclusions',
        3
      );
      return;
    }

    const numRows = lastRow - CONFIG.firstRow + 1;

    // Убираем возможные пустые строки в конце.
    allData.length = numRows;

    const data = allData;

    // --------------------------------------------------
    // 2. Чтение FilterState
    // --------------------------------------------------

    const filterStateSheet = ensureFilterStateLayout();
    const filterColumns = getFilterStateColumns(sheet);

    const filterMetadata = filterStateSheet
      .getRange(
        1,
        filterColumns.visibleCol,
        2,
        2
      )
      .getValues();

    const savedLastRow = Number(filterMetadata[1][1]);

    if (
      !Number.isInteger(savedLastRow) ||
      savedLastRow < lastRow
    ) {
      throw new Error(
        'FilterState не содержит актуального состояния. ' +
        'Сначала примените фильтры и повторите сохранение.'
      );
    }

    const filterVisibility = filterStateSheet
      .getRange(
        CONFIG.firstRow,
        filterColumns.visibleCol,
        numRows,
        1
      )
      .getValues()
      .map(row => row[0]);

    if (
      !filterVisibility.every(
        value => value === true || value === false
      )
    ) {
      throw new Error(
        'FilterState содержит некорректные данные видимости.'
      );
    }

    // --------------------------------------------------
    // 3. Поиск вручную скрытых строк
    // --------------------------------------------------

    const hiddenKeys = [];

    for (let i = 0; i < numRows; i++) {
      const row = data[i];

      if (!String(row[0] || '').trim()) continue;

      // Строки, скрытые фильтрами, пропускаем.
      if (filterVisibility[i] !== true) continue;

      const sheetRow = CONFIG.firstRow + i;

      if (sheet.isRowHiddenByUser(sheetRow)) {
        hiddenKeys.push(
          getManualRecordKey_(row)
        );
      }
    }

    // --------------------------------------------------
    // 4. Сохранение ManualState
    // --------------------------------------------------

    const columns = getManualStateColumns(sheet);

    const previousMetadata = filterStateSheet
      .getRange(
        1,
        columns.keyCol,
        2,
        2
      )
      .getValues();

    const previousCount = Math.max(
      0,
      Number(previousMetadata[1][1]) || 0
    );

    filterStateSheet
      .getRange(
        1,
        columns.keyCol,
        2,
        2
      )
      .setValues([
        [
          sheet.getName() + ' Manual',
          'Count'
        ],
        [
          '',
          hiddenKeys.length
        ]
      ]);

    if (hiddenKeys.length > 0) {
      filterStateSheet
        .getRange(
          CONFIG.firstRow,
          columns.keyCol,
          hiddenKeys.length,
          1
        )
        .setValues(
          hiddenKeys.map(key => [key])
        );
    }

    // --------------------------------------------------
    // 5. Очистка устаревших ключей
    // --------------------------------------------------

    if (previousCount > hiddenKeys.length) {
      filterStateSheet
        .getRange(
          CONFIG.firstRow + hiddenKeys.length,
          columns.keyCol,
          previousCount - hiddenKeys.length,
          1
        )
        .clearContent();
    }

    // --------------------------------------------------
    // 6. Результат
    // --------------------------------------------------

    spreadsheet.toast(
      'Saved ' + hiddenKeys.length + ' manual exclusions.',
      'Manual exclusions',
      4
    );

    console.log('[MANUAL STATE] Saved', JSON.stringify({
      sheet: sheet.getName(),
      rowsChecked: numRows,
      hiddenRows: hiddenKeys.length
    }));

    return hiddenKeys.length;
  });
}


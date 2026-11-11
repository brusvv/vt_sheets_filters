/**
 * FILTER STATE
 *
 * Управление служебным листом _FilterState.
 *
 * Novice       -> A:B
 * Intermediate -> D:E
 * Advanced     -> G:H
 *
 * Содержит:
 * - кэш структуры;
 * - создание служебного листа;
 * - проверку заголовков;
 * - миграцию старого состояния;
 * - чтение предыдущей видимости;
 * - сохранение новой видимости.
 */

// Кэш структуры _FilterState в рамках одного выполнения скрипта.
var FILTER_STATE_LAYOUT_CACHE = null;


/**
 * Получает служебный лист состояния фильтров.
 */
function getFilterStateSheet() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const stateSheetName = '_FilterState';

  let stateSheet = spreadsheet.getSheetByName(stateSheetName);

  if (!stateSheet) {
    stateSheet = spreadsheet.insertSheet(stateSheetName);
    stateSheet.hideSheet();
  }

  return stateSheet;
}


/**
 * Возвращает столбцы состояния для конкретного листа.
 */
function getFilterStateColumns(sheet) {
  const sheetName = typeof sheet === 'string'
    ? sheet
    : sheet.getName();

  const columns = {
    Novice: 1,
    Intermediate: 4,
    Advanced: 7
  };

  const startCol = columns[sheetName];

  if (!startCol) {
    throw new Error(
      `Нет блока FilterState для листа "${sheetName}".`
    );
  }

  return {
    visibleCol: startCol,
    lastRowCol: startCol + 1
  };
}


/**
 * Подготавливает структуру _FilterState.
 *
 * Если передан metadataValues, использует уже прочитанный
 * массив A1:Q2 и не выполняет отдельное чтение A1:H2.
 *
 * При первом запуске переносит старое состояние Intermediate
 * из A:B в D:E.
 *
 * Все изменения метаданных синхронизируются с переданным
 * массивом, чтобы вызывающая функция использовала актуальные данные.
 *
 * Кэш структуры действует только в рамках одного выполнения.
 */
function ensureFilterStateLayout(
  stateSheet = null,
  metadataValues = null
) {
  const totalStarted = Date.now();

  // Структура уже проверена в текущем выполнении.
  if (FILTER_STATE_LAYOUT_CACHE) {
    console.log(
      '[FILTER STATE TIMING] Layout cache: hit'
    );

    return FILTER_STATE_LAYOUT_CACHE;
  }

  console.log(
    '[FILTER STATE TIMING] Layout cache: miss'
  );

  let stepStarted = Date.now();

  stateSheet = stateSheet || getFilterStateSheet();

  console.log(
    '[FILTER STATE TIMING] Get state sheet: ' +
    (Date.now() - stepStarted) +
    ' ms'
  );

  // Если массив не передан, сохраняем прежнее поведение.
  // Если передан — используем уже прочитанные A1:Q2.
  stepStarted = Date.now();

  const layout = metadataValues
    ? metadataValues.map(row => row.slice(0, 8))
    : stateSheet.getRange('A1:H2').getValues();

  console.log(
    '[FILTER STATE TIMING] Read layout A1:H2: ' +
    (
      metadataValues
        ? 'reused from A1:Q2'
        : (Date.now() - stepStarted) + ' ms'
    )
  );

  const headerRow = layout[0];

  const noviceHeader = layout[0][0];
  const oldLastRowHeader = layout[0][1];
  const intermediateHeader = layout[0][3];

  let layoutChanged = false;

  // --------------------------------------------------
  // Миграция старого состояния Intermediate
  // --------------------------------------------------

  if (
    noviceHeader === 'Visible' &&
    oldLastRowHeader === 'LastRow' &&
    !intermediateHeader
  ) {
    const oldLastRow = Number(layout[1][1]);

    if (
      Number.isInteger(oldLastRow) &&
      oldLastRow >= CONFIG.firstRow
    ) {
      const numRows =
        oldLastRow - CONFIG.firstRow + 1;

      stepStarted = Date.now();

      const oldVisibility = stateSheet
        .getRange(
          CONFIG.firstRow,
          1,
          numRows,
          1
        )
        .getValues();

      console.log(
        '[FILTER STATE TIMING] Migration read old visibility: ' +
        (Date.now() - stepStarted) +
        ' ms'
      );

      stepStarted = Date.now();

      stateSheet.getRange('D1:E2').setValues([
        ['Intermediate Visible', 'LastRow'],
        ['', oldLastRow]
      ]);

      console.log(
        '[FILTER STATE TIMING] Migration write metadata: ' +
        (Date.now() - stepStarted) +
        ' ms'
      );

      stepStarted = Date.now();

      stateSheet
        .getRange(
          CONFIG.firstRow,
          4,
          numRows,
          1
        )
        .setValues(oldVisibility);

      console.log(
        '[FILTER STATE TIMING] Migration write visibility: ' +
        (Date.now() - stepStarted) +
        ' ms'
      );

      console.log(
        '[FILTER STATE] Intermediate migrated: ' +
        numRows +
        ' rows.'
      );

      // Обновляем локальный массив A1:H2.
      headerRow[3] = 'Intermediate Visible';
      headerRow[4] = 'LastRow';

      layout[1][3] = '';
      layout[1][4] = oldLastRow;

      layoutChanged = true;
    }
  }

  // --------------------------------------------------
  // Проверка заголовков всех блоков
  // --------------------------------------------------

  stepStarted = Date.now();

  const headers = [
    { col: 0, name: 'Novice Visible' },
    { col: 3, name: 'Intermediate Visible' },
    { col: 6, name: 'Advanced Visible' }
  ];

  headers.forEach(item => {
    if (headerRow[item.col] !== item.name) {
      headerRow[item.col] = item.name;
      layoutChanged = true;
    }

    if (headerRow[item.col + 1] !== 'LastRow') {
      headerRow[item.col + 1] = 'LastRow';
      layoutChanged = true;
    }
  });

  console.log(
    '[FILTER STATE TIMING] Check headers: ' +
    (Date.now() - stepStarted) +
    ' ms'
  );

  // --------------------------------------------------
  // Запись заголовков
  // --------------------------------------------------

  if (layoutChanged) {
    stepStarted = Date.now();

    stateSheet
      .getRange('A1:H1')
      .setValues([headerRow]);

    console.log(
      '[FILTER STATE TIMING] Write headers A1:H1: ' +
      (Date.now() - stepStarted) +
      ' ms'
    );
  } else {
    console.log(
      '[FILTER STATE TIMING] Write headers A1:H1: skipped'
    );
  }

  // --------------------------------------------------
  // Синхронизация переданного массива A1:Q2
  // --------------------------------------------------

  if (metadataValues) {
    for (let col = 0; col < 8; col++) {
      metadataValues[0][col] = headerRow[col];
    }

    metadataValues[1][0] = layout[1][0];
    metadataValues[1][1] = layout[1][1];
    metadataValues[1][3] = layout[1][3];
    metadataValues[1][4] = layout[1][4];
    metadataValues[1][6] = layout[1][6];
    metadataValues[1][7] = layout[1][7];
  }

  // Кэшируем только после успешного завершения операций.
  FILTER_STATE_LAYOUT_CACHE = stateSheet;

  console.log(
    '[FILTER STATE TIMING] ENSURE LAYOUT TOTAL: ' +
    (Date.now() - totalStarted) +
    ' ms'
  );

  return stateSheet;
}



/**
 * Получает предыдущее состояние видимости из _FilterState.
 *
 * Если количество строк изменилось:
 * - сохраняет состояние существующих строк;
 * - новые строки считает видимыми;
 * - при сокращении диапазона использует только актуальную часть.
 *
 * Физическую видимость проверяет только при отсутствии
 * пригодного сохранённого состояния.
 *
 * Performance diagnostics:
 * - получение диапазона;
 * - чтение сохранённой видимости;
 * - преобразование массива;
 * - проверка корректности;
 * - дополнение новых строк.
 */
function getPreviousVisibility(
  sheet,
  numRows,
  lastRow,
  stateSheet = null,
  metadata = null
) {
  const totalStarted = Date.now();

  let stepStarted = Date.now();

  stateSheet = stateSheet || ensureFilterStateLayout();

  const columns = getFilterStateColumns(sheet);

  let metadataReadMs = 0;

  if (!metadata) {
    stepStarted = Date.now();

    const metadataRange = stateSheet.getRange(
      1,
      columns.visibleCol,
      2,
      2
    );

    const metadataRangeMs = Date.now() - stepStarted;

    stepStarted = Date.now();

    metadata = metadataRange.getValues();

    metadataReadMs = Date.now() - stepStarted;

    console.log('[FILTER STATE READ] Metadata', JSON.stringify({
      sheet: sheet.getName(),
      rangeMs: metadataRangeMs,
      getValuesMs: metadataReadMs
    }));
  }

  const savedLastRow = Number(metadata[1][1]);

  if (
    Number.isInteger(savedLastRow) &&
    savedLastRow >= CONFIG.firstRow - 1
  ) {
    const overlapRows = Math.max(
      0,
      Math.min(
        numRows,
        savedLastRow - CONFIG.firstRow + 1
      )
    );

    let savedVisibility = [];

    let rangeMs = 0;
    let getValuesMs = 0;
    let transformMs = 0;

    if (overlapRows > 0) {
      stepStarted = Date.now();

      const visibilityRange = stateSheet.getRange(
        CONFIG.firstRow,
        columns.visibleCol,
        overlapRows,
        1
      );

      rangeMs = Date.now() - stepStarted;

      stepStarted = Date.now();

      const rawVisibility = visibilityRange.getValues();

      getValuesMs = Date.now() - stepStarted;

      stepStarted = Date.now();

      savedVisibility = rawVisibility.map(row => row[0]);

      transformMs = Date.now() - stepStarted;
    }

    // Проверка корректности сохранённой видимости.
    stepStarted = Date.now();

    const validValuesCount = savedVisibility.filter(
      value => value === true || value === false
    ).length;

    const invalidValues = savedVisibility
      .map((value, index) => ({
        row: CONFIG.firstRow + index,
        value: value,
        type: typeof value
      }))
      .filter(item =>
        item.value !== true &&
        item.value !== false
      )
      .slice(0, 10);

    const isValid =
      overlapRows > 0 &&
      savedVisibility.length === overlapRows &&
      validValuesCount === overlapRows;

    const validationMs = Date.now() - stepStarted;

    console.log('[FILTER STATE READ] Breakdown', JSON.stringify({
      sheet: sheet.getName(),
      overlapRows: overlapRows,
      rangeMs: rangeMs,
      getValuesMs: getValuesMs,
      transformMs: transformMs,
      validationMs: validationMs,
      metadataReadMs: metadataReadMs,
      totalReadMs: rangeMs + getValuesMs + transformMs,
      isValid: isValid
    }));

    console.log(
      '[FILTER STATE VALIDATION] ' +
      JSON.stringify({
        sheet: sheet.getName(),
        savedLastRow: savedLastRow,
        currentLastRow: lastRow,
        numRows: numRows,
        overlapRows: overlapRows,
        savedVisibilityLength: savedVisibility.length,
        validValuesCount: validValuesCount,
        isValid: isValid,
        invalidValues: invalidValues
      })
    );

    if (isValid) {
      stepStarted = Date.now();

      while (savedVisibility.length < numRows) {
        savedVisibility.push(true);
      }

      const paddingMs = Date.now() - stepStarted;

      console.log('[FILTER STATE READ] Result', JSON.stringify({
        sheet: sheet.getName(),
        resultRows: savedVisibility.length,
        paddingRows: numRows - overlapRows,
        paddingMs: paddingMs,
        totalMs: Date.now() - totalStarted,
        source: 'FilterState'
      }));

      return savedVisibility;
    }
  }

  console.log(
    '[FILTER FALLBACK] Sheet: ' + sheet.getName() +
    ', SavedLastRow: ' + savedLastRow +
    ', CurrentLastRow: ' + lastRow +
    ', Rows: ' + numRows
  );

  const visibility = new Array(numRows);

  stepStarted = Date.now();

  for (let i = 0; i < numRows; i++) {
    visibility[i] = !sheet.isRowHiddenByUser(
      CONFIG.firstRow + i
    );
  }

  const fallbackReadMs = Date.now() - stepStarted;

  console.log('[FILTER STATE READ] Fallback', JSON.stringify({
    sheet: sheet.getName(),
    rows: numRows,
    durationMs: fallbackReadMs
  }));

  // Сохраняем резервно полученную видимость,
  // чтобы следующие операции не повторяли FILTER FALLBACK.
  stepStarted = Date.now();

  updateVisibleRowsState(
    visibility,
    lastRow,
    sheet,
    null,
    stateSheet,
    metadata
  );

  const saveMs = Date.now() - stepStarted;

  console.log('[FILTER STATE READ] Result', JSON.stringify({
    sheet: sheet.getName(),
    resultRows: visibility.length,
    fallbackReadMs: fallbackReadMs,
    saveMs: saveMs,
    totalMs: Date.now() - totalStarted,
    source: 'Physical visibility fallback'
  }));

  return visibility;
}


/**
 * Записывает актуальное состояние видимости
 * в соответствующий блок листа _FilterState.
 *
 * Если состояние не изменилось, повторная запись
 * массива видимости пропускается.
 *
 * Novice       -> A:B
 * Intermediate -> D:E
 * Advanced     -> G:H
 *
 * Очищает устаревшие статусы,
 * если количество строк на листе уменьшилось.
 */
function updateVisibleRowsState(
  visibility,
  lastRow,
  sheet,
  previousVisibility = null,
  stateSheet = null,
  metadata = null
) {
  const totalStarted = Date.now();

  let stepStarted = Date.now();

  if (!sheet) {
    sheet = getTargetSheet();
  }

  stateSheet = stateSheet || ensureFilterStateLayout();

  const columns = getFilterStateColumns(sheet);

  console.log(
    '[FILTER STATE TIMING] Resolve state sheet and columns: ' +
    (Date.now() - stepStarted) +
    ' ms'
  );

  // Используем переданные метаданные.
  stepStarted = Date.now();

  if (!metadata) {
    metadata = stateSheet
      .getRange(1, columns.visibleCol, 2, 2)
      .getValues();

    console.log(
      '[FILTER STATE TIMING] Read metadata: ' +
      (Date.now() - stepStarted) +
      ' ms'
    );
  } else {
    console.log(
      '[FILTER STATE TIMING] Read metadata: reused'
    );
  }

  const previousLastRow = Number(metadata[1][1]);

  const expectedMetadata = [
    [sheet.getName() + ' Visible', 'LastRow'],
    ['', lastRow]
  ];

  // Проверяем изменение метаданных.
  stepStarted = Date.now();

  const metadataChanged =
    metadata[0][0] !== expectedMetadata[0][0] ||
    metadata[0][1] !== expectedMetadata[0][1] ||
    metadata[1][0] !== expectedMetadata[1][0] ||
    metadata[1][1] !== expectedMetadata[1][1];

  console.log(
    '[FILTER STATE TIMING] Compare metadata: ' +
    (Date.now() - stepStarted) +
    ' ms'
  );

  if (metadataChanged) {
    stepStarted = Date.now();

    stateSheet
      .getRange(1, columns.visibleCol, 2, 2)
      .setValues(expectedMetadata);

    console.log(
      '[FILTER STATE TIMING] Write metadata: ' +
      (Date.now() - stepStarted) +
      ' ms'
    );
  } else {
    console.log(
      '[FILTER STATE TIMING] Write metadata: skipped'
    );
  }

  // Проверяем изменение видимости.
  stepStarted = Date.now();

  const visibilityUnchanged =
    Array.isArray(previousVisibility) &&
    previousVisibility.length === visibility.length &&
    previousVisibility.every(
      (value, index) => value === visibility[index]
    );

  console.log(
    '[FILTER STATE TIMING] Compare visibility: ' +
    (Date.now() - stepStarted) +
    ' ms'
  );

  // Записываем только при необходимости.
  if (!visibilityUnchanged && visibility.length > 0) {
    stepStarted = Date.now();

    stateSheet
      .getRange(
        CONFIG.firstRow,
        columns.visibleCol,
        visibility.length,
        1
      )
      .setValues(
        visibility.map(value => [value])
      );

    console.log(
      '[FILTER STATE TIMING] Write visibility: ' +
      (Date.now() - stepStarted) +
      ' ms'
    );
  } else {
    console.log(
      '[FILTER STATE TIMING] Write visibility: skipped'
    );
  }

  // Очищаем старое состояние,
  // если лист стал короче.
  if (
    previousLastRow > lastRow &&
    previousLastRow >= CONFIG.firstRow
  ) {
    const rowsToClear = previousLastRow - lastRow;

    stepStarted = Date.now();

    stateSheet
      .getRange(
        lastRow + 1,
        columns.visibleCol,
        rowsToClear,
        1
      )
      .clearContent();

    console.log(
      '[FILTER STATE TIMING] Clear old visibility (' +
      rowsToClear +
      ' rows): ' +
      (Date.now() - stepStarted) +
      ' ms'
    );
  } else {
    console.log(
      '[FILTER STATE TIMING] Clear old visibility: skipped'
    );
  }

  if (visibilityUnchanged) {
    console.log(
      '[FILTER STATE] Visibility unchanged. Write skipped: ' +
      sheet.getName()
    );
  }

  console.log(
    '[FILTER STATE TIMING] UPDATE VISIBLE ROWS TOTAL: ' +
    (Date.now() - totalStarted) +
    ' ms'
  );
}

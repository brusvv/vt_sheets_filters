/**
 * Общие вспомогательные функции проекта.
 */


/**
 * Возвращает рабочий лист.
 *
 * Если имя не передано:
 * 1. Использует активный игровой лист.
 * 2. Если активен другой лист — использует CONFIG.sheetName.
 */
function getTargetSheet(sheetName = null, spreadsheet = null) {
  // Используем переданный Spreadsheet,
  // чтобы не запрашивать ActiveSpreadsheet повторно.
  if (!spreadsheet) {
    spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  }

  let targetName = sheetName;

  if (!targetName) {
    const activeSheet = spreadsheet.getActiveSheet();
    const activeName = activeSheet.getName();

    targetName = RANK_TRACKED_SHEETS.includes(activeName)
      ? activeName
      : CONFIG.sheetName;
  }

  if (!RANK_TRACKED_SHEETS.includes(targetName)) {
    throw new Error(
      `Недопустимый целевой лист: "${targetName}". Ожидается один из игровых листов: ` +
      RANK_TRACKED_SHEETS.join(', ')
    );
  }

  const sheet = spreadsheet.getSheetByName(targetName);

  if (!sheet) {
    throw new Error(`Лист "${targetName}" не найден.`);
  }

  return sheet;
}


/**
 * Записывает время выполнения этапа в журнал.
 */
function logFilterTiming(label, startedAt) {
  console.log(
    '[FILTER TIMING] ' +
    label +
    ': ' +
    (Date.now() - startedAt) +
    ' ms'
  );
}


/**
 * Проверяет пересечение диапазона с указанной ячейкой.
 */
function rangeIntersectsCell(range, a1Notation) {
  const sheet = range.getSheet();
  const cell = sheet.getRange(a1Notation);

  return (
    range.getRow() <= cell.getLastRow() &&
    range.getLastRow() >= cell.getRow() &&
    range.getColumn() <= cell.getLastColumn() &&
    range.getLastColumn() >= cell.getColumn()
  );
}


/**
 * Читает настройки проекта с листа Options.
 *
 * Возвращает массив из пяти элементов:
 * [0] A5  — Rank Tiers
 * [1] A6  — Auto Sort
 * [2] A7  — Hide Empty Scores
 * [3] A10  — Columns Behind
 * [4] A11 — Columns Ahead
 *
 * Порядок сохранён для совместимости
 * с существующей логикой applyFilters()
 * и rankScroller().
 */


function getRankScrollerOptions_(spreadsheet) {
  const ss =
    spreadsheet || SpreadsheetApp.getActiveSpreadsheet();

  const optionsSheet =
    ss.getSheetByName(CONFIG.optionsSheetName);

  if (!optionsSheet) {
    throw new Error(
      'Options sheet not found: ' +
      CONFIG.optionsSheetName
    );
  }

  // Читаем A5:A11 одним запросом.
  const values = optionsSheet
    .getRange('A5:A11')
    .getValues();

  return [
    [values[0][0]], // A5 — Rank Tiers
    [values[1][0]], // A6 — Auto Sort
    [values[2][0]], // A7 — Hide Empty Scores
    [values[5][0]], // A10 — Columns Behind
    [values[6][0]]  // A11 — Columns Ahead
  ];
}



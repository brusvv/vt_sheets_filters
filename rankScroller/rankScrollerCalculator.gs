
/**
 * @OnlyCurrentDoc
 */

/**
 * RANK SCROLLER CALCULATOR
 *
 * Ответственность:
 * - определение целевой колонки по Energy;
 * - поиск подходящего интервала EnergyThreshold;
 * - обработка первого и последнего диапазонов.
 *
 * Файл содержит только вычисления.
 * Не обращается к Google Sheets API.
 * Не изменяет видимость колонок.
 * Не использует Performance Tracer.
 */


/**
 * Определяет целевую колонку Rank Scroller.
 *
 * @param {number} energyValue Текущее значение Energy.
 * @param {Array} energyValues Массив EnergyThreshold.
 * @param {number} startColumn Первая колонка диапазона.
 * @param {number} totalColumns Количество пороговых колонок.
 * @param {number} visibleColumns Количество отображаемых колонок.
 *
 * @return {{
 *   success: boolean,
 *   currentColumn: (number|undefined)
 * }}
 */
function calculateRankScrollerTargetColumn_(
  energyValue,
  energyValues,
  startColumn,
  totalColumns,
  visibleColumns
) {
  let currentColumn = startColumn;
  let foundColumn = false;

  // Первый диапазон.
  const firstThreshold =
    parseFloat(energyValues[0]);

  if (
    Number.isFinite(firstThreshold) &&
    energyValue < firstThreshold
  ) {
    currentColumn = startColumn;
    foundColumn = true;
  }

  // Поиск подходящего интервала.
  if (!foundColumn) {
    for (let j = 0; j < totalColumns - 1; j++) {
      const cellValue =
        parseFloat(energyValues[j]);

      const nextCellValue =
        parseFloat(energyValues[j + 1]);

      if (
        Number.isFinite(cellValue) &&
        Number.isFinite(nextCellValue) &&
        energyValue >= cellValue &&
        energyValue < nextCellValue
      ) {
        currentColumn = startColumn + j;
        foundColumn = true;
        break;
      }
    }
  }

  // Последний доступный диапазон.
  // Сохраняем исходное поведение:
  // эта проверка может переопределить ранее найденную колонку.
  const specialCaseIndex =
    totalColumns - visibleColumns;

  const specialCaseCell =
    parseFloat(energyValues[specialCaseIndex]);

  if (
    Number.isFinite(specialCaseCell) &&
    energyValue >= specialCaseCell
  ) {
    currentColumn =
      startColumn + specialCaseIndex;

    foundColumn = true;
  }

  if (!foundColumn) {
    return {
      success: false
    };
  }

  return {
    success: true,
    currentColumn: currentColumn
  };
}

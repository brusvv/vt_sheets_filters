
/**
 * @OnlyCurrentDoc
 */

/**
 * RANK SCROLLER VISIBILITY
 *
 * Ответственность:
 * - скрытие диапазона пороговых колонок;
 * - отображение целевого диапазона;
 * - отображение групп колонок для режима Rank Tiers = false.
 *
 * Файл не выполняет расчёты Energy.
 * Не обращается к Script Properties.
 * Не использует Performance Tracer.
 */


/**
 * Скрывает все пороговые колонки и показывает
 * рассчитанный диапазон.
 *
 * @param {Sheet} sheet Рабочий лист.
 * @param {number} startColumn Первая колонка диапазона.
 * @param {number} totalColumns Количество колонок диапазона.
 * @param {number} currentColumn Первая отображаемая колонка.
 * @param {number} visibleColumns Количество отображаемых колонок.
 *
 * @return {{
 *   success: boolean,
 *   currentColumn: number,
 *   visibleColumns: number
 * }}
 */
function showRankScrollerColumns_(
  sheet,
  startColumn,
  totalColumns,
  currentColumn,
  visibleColumns
) {
  if (
    !sheet ||
    !Number.isInteger(startColumn) ||
    !Number.isInteger(totalColumns) ||
    !Number.isInteger(currentColumn) ||
    !Number.isInteger(visibleColumns) ||
    totalColumns < 1 ||
    visibleColumns < 1 ||
    visibleColumns > totalColumns ||
    currentColumn < startColumn ||
    currentColumn + visibleColumns - 1 >
      startColumn + totalColumns - 1
  ) {
    return {
      success: false
    };
  }

  sheet.hideColumns(
    startColumn,
    totalColumns
  );

  sheet.showColumns(
    currentColumn,
    visibleColumns
  );

  return {
    success: true,
    currentColumn: currentColumn,
    visibleColumns: visibleColumns
  };
}


/**
 * Скрывает все пороговые колонки и показывает
 * только группы, значения которых делятся на 100.
 *
 * Соседние подходящие колонки объединяются в группы,
 * чтобы минимизировать количество вызовов showColumns().
 *
 * @param {Sheet} sheet Рабочий лист.
 * @param {number} startColumn Первая колонка диапазона.
 * @param {number} totalColumns Количество колонок диапазона.
 * @param {Array} energyValues Значения EnergyThreshold.
 *
 * @return {{
 *   success: boolean,
 *   groupsCount: number,
 *   groups: Array
 * }}
 */
function showRankScrollerTierColumns_(
  sheet,
  startColumn,
  totalColumns,
  energyValues
) {
  if (
    !sheet ||
    !Number.isInteger(startColumn) ||
    !Number.isInteger(totalColumns) ||
    totalColumns < 1 ||
    !Array.isArray(energyValues) ||
    energyValues.length < totalColumns
  ) {
    return {
      success: false,
      groupsCount: 0,
      groups: []
    };
  }

  // Сначала скрываем весь диапазон.
  sheet.hideColumns(
    startColumn,
    totalColumns
  );

  const visibleGroups = [];

  let groupStart = -1;
  let groupLength = 0;

  // Ищем непрерывные группы значений,
  // которые делятся на 100 без остатка.
  for (let i = 0; i < totalColumns; i++) {
    const value = energyValues[i];

    if (
      typeof value === 'number' &&
      value % 100 === 0
    ) {
      if (groupStart === -1) {
        groupStart = startColumn + i;
        groupLength = 1;
      } else {
        groupLength++;
      }

    } else if (groupStart !== -1) {
      visibleGroups.push({
        start: groupStart,
        length: groupLength
      });

      groupStart = -1;
      groupLength = 0;
    }
  }

  // Добавляем последнюю группу,
  // если подходящие значения находятся в конце массива.
  if (groupStart !== -1) {
    visibleGroups.push({
      start: groupStart,
      length: groupLength
    });
  }

  // Показываем найденные группы.
  visibleGroups.forEach(function(group) {
    sheet.showColumns(
      group.start,
      group.length
    );
  });

  return {
    success: true,
    groupsCount: visibleGroups.length,
    groups: visibleGroups
  };
}

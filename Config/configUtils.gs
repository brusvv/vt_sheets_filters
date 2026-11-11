/**
 * Вспомогательные функции для работы
 * с координатами и A1-диапазонами CONFIG.
 */


/**
 * Преобразует номер столбца в буквенное обозначение.
 *
 * 1  -> A
 * 26 -> Z
 * 27 -> AA
 * 28 -> AB
 */
function columnToLetter_(column) {
  if (
    !Number.isInteger(column) ||
    column < 1
  ) {
    throw new Error(
      `Некорректный номер столбца: ${column}`
    );
  }

  let result = '';

  while (column > 0) {
    const remainder = (column - 1) % 26;

    result =
      String.fromCharCode(65 + remainder) +
      result;

    column = Math.floor((column - 1) / 26);
  }

  return result;
}


/**
 * Преобразует буквенное обозначение столбца
 * в числовой индекс.
 *
 * A  -> 1
 * Z  -> 26
 * AA -> 27
 * AB -> 28
 */
function columnToIndex_(letters) {
  if (
    typeof letters !== 'string' ||
    !/^[A-Z]+$/i.test(letters)
  ) {
    throw new Error(
      `Некорректное обозначение столбца: ${letters}`
    );
  }

  let result = 0;

  letters = letters.toUpperCase();

  for (let i = 0; i < letters.length; i++) {
    result =
      result * 26 +
      letters.charCodeAt(i) - 64;
  }

  return result;
}


/**
 * Совместимость со старым именем функции.
 *
 * Если columnLetterToIndex() используется
 * в других файлах проекта, существующие вызовы
 * продолжат работать.
 */
function columnLetterToIndex(letter) {
  return columnToIndex_(letter);
}


/**
 * Формирует диапазон данных от указанной строки
 * до конца заданного столбца.
 *
 * Например:
 * getColumnDataRange_(5, 3) -> E3:E
 */
function getColumnDataRange_(
  column,
  firstRow = CONFIG.firstRow
) {
  const letter = columnToLetter_(column);

  return `${letter}${firstRow}:${letter}`;
}


/**
 * Формирует A1-адрес по номеру строки и столбца.
 *
 * Например:
 * getCellAddress_(3, 5) -> E3
 */
function getCellAddress_(row, column) {
  if (
    !Number.isInteger(row) ||
    row < 1
  ) {
    throw new Error(
      `Некорректный номер строки: ${row}`
    );
  }

  return `${columnToLetter_(column)}${row}`;
}


/**
 * Формирует A1-диапазон по номерам строк и столбцов.
 *
 * Например:
 * getA1Range_({
 *   firstRow: 3,
 *   lastRow: 70,
 *   firstCol: 2,
 *   lastCol: 3
 * }) -> B3:C70
 */
function getA1Range_({
  firstRow,
  lastRow,
  firstCol,
  lastCol
}) {
  return (
    `${getCellAddress_(firstRow, firstCol)}:` +
    `${getCellAddress_(lastRow, lastCol)}`
  );
}


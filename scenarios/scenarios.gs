
/**
 * SCENARIOS
 *
 * Вспомогательные функции для работы
 * со сценариями и объединёнными категориями.
 *
 * Комплексные пользовательские операции
 * находятся в operations.gs.
 */


/**
 * Получает объединённые диапазоны категорий в A:B и M.
 *
 * Поиск ограничен строками от CONFIG.firstRow до sourceRow,
 * поскольку нас интересуют только объединения,
 * которые могут заканчиваться на исходной строке.
 */
function getCategoryMerges(sheet, sourceRow) {
  const stepStarted = Date.now();

  const rowCount = sourceRow - CONFIG.firstRow + 1;

  if (rowCount <= 0) {
    return [];
  }

  const columns = [
    { start: 1, width: 2 },
    { start: 13, width: 1 }
  ];

  const mergedRanges = [];

  columns.forEach(column => {
    const range = sheet.getRange(
      CONFIG.firstRow,
      column.start,
      rowCount,
      column.width
    );

    range.getMergedRanges().forEach(mergedRange => {
      mergedRanges.push(mergedRange);
    });
  });

  console.log(
    '[DUPLICATE TIMING] Get category merges: ' +
    (Date.now() - stepStarted) +
    ' ms'
  );

  console.log(
    `[DUPLICATE TIMING] Found merged ranges: ${mergedRanges.length}`
  );

  return mergedRanges;
}


/**
 * Продлевает объединённые категории на новую строку,
 * если исходная строка входила в объединённый диапазон.
 */
function extendCategoryMerges(
  sheet,
  sourceRow,
  targetRow
) {
  const stepStarted = Date.now();

  const mergedRanges = getCategoryMerges(
    sheet,
    sourceRow
  );

  console.log(
    '[DUPLICATE TIMING] Prepare category merges: ' +
    (Date.now() - stepStarted) +
    ' ms'
  );

  const mergeStarted = Date.now();

  let extendedCount = 0;

  mergedRanges.forEach(mergedRange => {
    const startRow = mergedRange.getRow();

    const endRow =
      startRow + mergedRange.getNumRows() - 1;

    if (endRow === sourceRow) {
      const startCol = mergedRange.getColumn();
      const numCols = mergedRange.getNumColumns();
      const numRows = mergedRange.getNumRows();

      mergedRange.breakApart();

      sheet
        .getRange(
          startRow,
          startCol,
          numRows + 1,
          numCols
        )
        .merge();

      extendedCount++;
    }
  });

  console.log(
    '[DUPLICATE TIMING] Apply category merges: ' +
    (Date.now() - mergeStarted) +
    ' ms'
  );

  console.log(
    `[DUPLICATE TIMING] Extended merges: ${extendedCount}`
  );
}

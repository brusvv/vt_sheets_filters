/**
 * ROW VISIBILITY
 *
 * Низкоуровневое управление физической видимостью строк.
 *
 * Содержит:
 * - групповое отображение строк;
 * - групповое скрытие строк.
 */


/**
 * Показывает последовательные группы строк.
 *
 * Performance diagnostics:
 * - количество вызовов showRows();
 * - количество затронутых строк;
 * - количество групп;
 * - общее время выполнения.
 */
function showRowsInGroups(sheet, rows) {
  const totalStarted = Date.now();

  if (!rows.length) {
    console.log('[ROW VISIBILITY] SHOW: skipped (0 rows)');
    return;
  }

  let startRow = rows[0];
  let previousRow = rows[0];

  let groupCount = 0;
  let affectedRows = 0;

  for (let i = 1; i < rows.length; i++) {
    const currentRow = rows[i];

    if (currentRow === previousRow + 1) {
      previousRow = currentRow;
      continue;
    }

    const groupSize = previousRow - startRow + 1;

    sheet.showRows(startRow, groupSize);

    groupCount++;
    affectedRows += groupSize;

    startRow = currentRow;
    previousRow = currentRow;
  }

  // Последняя группа.
  const groupSize = previousRow - startRow + 1;

  sheet.showRows(startRow, groupSize);

  groupCount++;
  affectedRows += groupSize;

  console.log('[ROW VISIBILITY] SHOW', JSON.stringify({
    sheet: sheet.getName(),
    calls: groupCount,
    groups: groupCount,
    rows: affectedRows,
    durationMs: Date.now() - totalStarted
  }));
}


/**
 * Скрывает последовательные группы строк.
 *
 * Performance diagnostics:
 * - количество вызовов hideRows();
 * - количество затронутых строк;
 * - количество групп;
 * - общее время выполнения.
 */
function hideRowsInGroups(sheet, rowsToHide) {
  const totalStarted = Date.now();

  if (!rowsToHide.size) {
    console.log('[ROW VISIBILITY] HIDE: skipped (0 rows)');
    return;
  }

  const sortedRows = Array.from(rowsToHide)
    .sort((a, b) => a - b);

  let startRow = sortedRows[0];
  let previousRow = sortedRows[0];

  let groupCount = 0;
  let affectedRows = 0;

  for (let i = 1; i < sortedRows.length; i++) {
    const currentRow = sortedRows[i];

    if (currentRow === previousRow + 1) {
      previousRow = currentRow;
      continue;
    }

    const groupSize = previousRow - startRow + 1;

    sheet.hideRows(startRow, groupSize);

    groupCount++;
    affectedRows += groupSize;

    startRow = currentRow;
    previousRow = currentRow;
  }

  // Последняя группа.
  const groupSize = previousRow - startRow + 1;

  sheet.hideRows(startRow, groupSize);

  groupCount++;
  affectedRows += groupSize;

  console.log('[ROW VISIBILITY] HIDE', JSON.stringify({
    sheet: sheet.getName(),
    calls: groupCount,
    groups: groupCount,
    rows: affectedRows,
    durationMs: Date.now() - totalStarted
  }));
}

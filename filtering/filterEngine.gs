
/**
 * FILTER ENGINE
 *
 * Расчёт фильтрационной видимости.
 *
 * Содержит:
 * - применение Dots;
 * - применение TOP;
 * - обработку пустых High Score;
 * - учёт ManualState;
 * - формирование FilterState.
 */


/**
 * Строит результат Dots + TOP и учитывает ManualState.
 *
 * FilterState и ManualState остаются независимыми.
 */
function buildFilterVisibility_(
  data,
  selectedDots,
  topN,
  hideEmptyScores,
  sheet,
  manualHiddenCounts,
  perfTrace,
  resolvedManualHiddenRows = null
) {
  // --------------------------------------------------
  // Применение Dots и формирование групп
  // --------------------------------------------------

  let stepStarted = Date.now();

  const groups = new Map();

  const filterRowsToHide = new Set();
  const manualRowsToHide = new Set();
  const manuallyHiddenRows = new Set();

  data.forEach((row, index) => {
    const sheetRow = CONFIG.firstRow + index;

    const scenario = String(row[0] || '').trim();

    const score = row[
      CONFIG.scoreCol - CONFIG.scenarioCol
    ];

    const dots = String(
      row[CONFIG.dotsCol - CONFIG.scenarioCol] || ''
    ).trim();

    // Пустые строки скрываются фильтрацией.

    if (!scenario) {
      filterRowsToHide.add(sheetRow);
      return;
    }

    // Use physical positions from a validated resolver snapshot when available.
    // Other callers retain the established FIFO ManualState behavior.
    const isManualHidden = resolvedManualHiddenRows instanceof Set
      ? resolvedManualHiddenRows.has(sheetRow)
      : (
          manualHiddenCounts.size > 0 &&
          consumeManualHiddenKey_(row, manualHiddenCounts)
        );

    if (isManualHidden) {
      manualRowsToHide.add(sheetRow);
      manuallyHiddenRows.add(sheetRow);
      return;
    }

    // Dots.

    if (
      selectedDots.size > 0 &&
      !selectedDots.has(dots)
    ) {
      filterRowsToHide.add(sheetRow);
      return;
    }

    // Некорректный или пустой High Score.

    if (
      typeof score !== 'number' ||
      !Number.isFinite(score)
    ) {
      if (hideEmptyScores) {
        filterRowsToHide.add(sheetRow);
      }

      return;
    }

    if (!groups.has(scenario)) {
      groups.set(scenario, []);
    }

    groups.get(scenario).push({
      row: sheetRow,
      score: score
    });
  });

  filterPerfStep_(
    perfTrace,
    'Build scenario groups and apply Dots',
    stepStarted,
    {
      groups: groups.size,
      filterRowsToHide: filterRowsToHide.size,
      manualRowsToHide: manualRowsToHide.size
    }
  );

  // --------------------------------------------------
  // Применение TOP
  // --------------------------------------------------

  stepStarted = Date.now();

  if (topN !== null) {
    groups.forEach(results => {
      results.sort((a, b) => {
        if (b.score !== a.score) {
          return b.score - a.score;
        }

        return a.row - b.row;
      });

      results.forEach((result, index) => {
        if (index >= topN) {
          filterRowsToHide.add(result.row);
        }
      });
    });
  }

  filterPerfStep_(
    perfTrace,
    'Calculate TOP',
    stepStarted,
    {
      topN: topN,
      filterRowsToHide: filterRowsToHide.size
    }
  );

  // --------------------------------------------------
  // Формирование FilterState
  // --------------------------------------------------

  const filterVisibility = Array.from(
    { length: data.length },
    (_, index) => {
      const sheetRow = CONFIG.firstRow + index;

      return !filterRowsToHide.has(sheetRow);
    }
  );

  return {
    groups: groups,
    filterRowsToHide: filterRowsToHide,
    manualRowsToHide: manualRowsToHide,
    manuallyHiddenRows: manuallyHiddenRows,
    filterVisibility: filterVisibility
  };
}

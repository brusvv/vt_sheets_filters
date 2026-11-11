
/**
 * SORTING
 *
 * Непосредственная сортировка последовательных
 * блоков одинаковых сценариев.
 *
 * Комплексные операции ручной сортировки
 * находятся в SortOperations.gs.
 */


/**
 * Проверяет, отсортирован ли блок.
 *
 * Возвращает:
 * {
 *   sorted: boolean,
 *   reason: string,
 *   mismatchIndex: number | null
 * }
 *
 * Проверка консервативная:
 * - Score должен быть конечным числом.
 * - Dots должен быть непустой ASCII-строкой.
 * - При неоднозначных значениях блок передаётся Range.sort().
 */
function inspectScenarioBlockOrder_(
  values,
  sortType,
  scoreIndex,
  dotsIndex
) {
  if (values.length < 2) {
    return {
      sorted: true,
      reason: 'ALREADY_SORTED',
      mismatchIndex: null
    };
  }

  let previousScore;
  let previousDots;

  for (let i = 0; i < values.length; i++) {
    const score = values[i][scoreIndex];
    const dots = values[i][dotsIndex];

    // Проверка Dots
    if (typeof dots !== 'string' || dots.length === 0) {
      return {
        sorted: false,
        reason: 'EMPTY_DOTS',
        mismatchIndex: i
      };
    }

    if (!/^[\x00-\x7F]*$/.test(dots)) {
      return {
        sorted: false,
        reason: 'NON_ASCII_DOTS',
        mismatchIndex: i
      };
    }

    // Проверка Score для режимов scoreDesc и scoreAsc
    if (sortType !== 'dots') {
      if (
        typeof score !== 'number' ||
        !Number.isFinite(score)
      ) {
        return {
          sorted: false,
          reason: 'INVALID_SCORE',
          mismatchIndex: i
        };
      }
    }

    if (i === 0) {
      previousScore = score;
      previousDots = dots;
      continue;
    }

    let scoreEqual = true;

    if (sortType === 'scoreDesc') {
      if (previousScore < score) {
        return {
          sorted: false,
          reason: 'UNSORTED_SCORE_DESC',
          mismatchIndex: i
        };
      }

      scoreEqual = previousScore === score;

    } else if (sortType === 'scoreAsc') {
      if (previousScore > score) {
        return {
          sorted: false,
          reason: 'UNSORTED_SCORE_ASC',
          mismatchIndex: i
        };
      }

      scoreEqual = previousScore === score;

    } else if (sortType === 'dots') {
      scoreEqual = true;
    }

    // Dots проверяется:
    // - всегда для режима dots;
    // - только при равном Score для остальных режимов.
    if (scoreEqual && previousDots > dots) {
      return {
        sorted: false,
        reason: 'UNSORTED_DOTS',
        mismatchIndex: i
      };
    }

    previousScore = score;
    previousDots = dots;
  }

  return {
    sorted: true,
    reason: 'ALREADY_SORTED',
    mismatchIndex: null
  };
}


/**
 * Сортирует последовательные блоки одинаковых сценариев.
 *
 * sortType:
 * scoreDesc — High Score по убыванию, Dots A-Z при равенстве.
 * scoreAsc  — High Score по возрастанию, Dots A-Z.
 * dots      — только Dots A-Z.
 *
 * knownLastRow — заранее известная последняя строка.
 * Если не передана, определяется самостоятельно.
 *
 * knownScenarios — заранее прочитанные значения сценариев.
 * Если не переданы, считываются самостоятельно.
 */
function sortScenarioBlocks(
  sheet,
  sortType = 'scoreDesc',
  knownLastRow = null,
  knownScenarios = null
) {
  const totalStarted = Date.now();

  // 0. Проверка типа сортировки
  if (!['scoreDesc', 'scoreAsc', 'dots'].includes(sortType)) {
    throw new Error(
      'Неизвестный тип сортировки: ' + sortType
    );
  }

  const lastRow = knownLastRow !== null
    ? knownLastRow
    : sheet.getLastRow();

  if (lastRow < CONFIG.firstRow) return;

  const numRows = lastRow - CONFIG.firstRow + 1;

  // 1. Получение сценариев
  const scenarioReadStarted = Date.now();

  const scenarios = knownScenarios !== null
    ? knownScenarios
    : sheet
        .getRange(
          CONFIG.firstRow,
          CONFIG.scenarioCol,
          numRows,
          1
        )
        .getValues();

  const scenarioReadDuration =
    Date.now() - scenarioReadStarted;

  // 2. Пакетное чтение всего сортируемого диапазона
  const dataReadStarted = Date.now();

  const allSortValues = sheet
    .getRange(
      CONFIG.firstRow,
      CONFIG.sortStartCol,
      numRows,
      CONFIG.sortWidth
    )
    .getValues();

  const dataReadDuration = Date.now() - dataReadStarted;

  let sortCalls = 0;
  let sortedRows = 0;
  let sortDuration = 0;

  let checkedBlocks = 0;
  let skippedBlocks = 0;
  let skippedRows = 0;

  const blockDiagnostics = [];

  const scoreIndex =
    CONFIG.scoreCol - CONFIG.sortStartCol;

  const dotsIndex =
    CONFIG.dotsCol - CONFIG.sortStartCol;

  // Проверка корректности индексов внутри диапазона
  if (
    dotsIndex < 0 ||
    dotsIndex >= CONFIG.sortWidth
  ) {
    throw new Error(
      'CONFIG.dotsCol находится за пределами CONFIG.sortStartCol / sortWidth'
    );
  }

  if (
    sortType !== 'dots' &&
    (
      scoreIndex < 0 ||
      scoreIndex >= CONFIG.sortWidth
    )
  ) {
    throw new Error(
      'CONFIG.scoreCol находится за пределами CONFIG.sortStartCol / sortWidth'
    );
  }

  /**
   * Проверяет и при необходимости сортирует один блок.
   *
   * offset — индекс первой строки блока
   * относительно allSortValues.
   */
  function sortBlock(
    startRow,
    endRow,
    scenarioName,
    offset
  ) {
    const blockLength = endRow - startRow + 1;

    if (blockLength < 2) return;

    // Данные берутся из уже прочитанного массива.
    // Дополнительного обращения к Google Sheets нет.
    const blockValues = allSortValues.slice(
      offset,
      offset + blockLength
    );

    checkedBlocks++;

    const inspection = inspectScenarioBlockOrder_(
      blockValues,
      sortType,
      scoreIndex,
      dotsIndex
    );

    if (inspection.sorted) {
      skippedBlocks++;
      skippedRows += blockLength;
      return;
    }

    const range = sheet.getRange(
      startRow,
      CONFIG.sortStartCol,
      blockLength,
      CONFIG.sortWidth
    );

    const sortStarted = Date.now();

    if (sortType === 'scoreDesc') {
      range.sort([
        {
          column: CONFIG.scoreCol,
          ascending: false
        },
        {
          column: CONFIG.dotsCol,
          ascending: true
        }
      ]);

    } else if (sortType === 'scoreAsc') {
      range.sort([
        {
          column: CONFIG.scoreCol,
          ascending: true
        },
        {
          column: CONFIG.dotsCol,
          ascending: true
        }
      ]);

    } else if (sortType === 'dots') {
      range.sort({
        column: CONFIG.dotsCol,
        ascending: true
      });
    }

    sortDuration += Date.now() - sortStarted;
    sortCalls++;
    sortedRows += blockLength;

    // Подробная информация только о блоках,
    // которые действительно пришлось сортировать.
    blockDiagnostics.push({
      scenario: scenarioName,
      startRow: startRow,
      endRow: endRow,
      rows: blockLength,
      reason: inspection.reason,
      mismatchRow: inspection.mismatchIndex !== null
        ? startRow + inspection.mismatchIndex
        : null
    });
  }

  // 3. Поиск последовательных блоков сценариев
  let blockStart = null;
  let currentScenario = '';

  for (let i = 0; i < scenarios.length; i++) {
    const scenario = String(scenarios[i][0] || '').trim();
    const currentRow = CONFIG.firstRow + i;

    // Пустая строка завершает текущий блок
    if (!scenario) {
      if (blockStart !== null) {
        const offset = blockStart - CONFIG.firstRow;

        sortBlock(
          blockStart,
          currentRow - 1,
          currentScenario,
          offset
        );
      }

      blockStart = null;
      currentScenario = '';
      continue;
    }

    // Начало нового блока
    if (blockStart === null) {
      blockStart = currentRow;
      currentScenario = scenario;
      continue;
    }

    // Смена сценария — обрабатываем предыдущий блок
    if (scenario !== currentScenario) {
      const offset = blockStart - CONFIG.firstRow;

      sortBlock(
        blockStart,
        currentRow - 1,
        currentScenario,
        offset
      );

      blockStart = currentRow;
      currentScenario = scenario;
    }
  }

  // 4. Последний блок
  if (blockStart !== null) {
    const offset = blockStart - CONFIG.firstRow;

    sortBlock(
      blockStart,
      CONFIG.firstRow + scenarios.length - 1,
      currentScenario,
      offset
    );
  }

  // 5. Диагностика производительности
  console.log('[SORT DIAGNOSTICS]', JSON.stringify({
    sheet: sheet.getName(),
    sortType: sortType,
    rows: numRows,

    scenarioReadDurationMs: scenarioReadDuration,
    sortDataReadDurationMs: dataReadDuration,

    checkedBlocks: checkedBlocks,
    skippedBlocks: skippedBlocks,
    skippedRows: skippedRows,

    sortCalls: sortCalls,
    sortedRows: sortedRows,
    sortDurationMs: sortDuration,

    blockDiagnostics: blockDiagnostics,

    totalDurationMs: Date.now() - totalStarted
  }));
}

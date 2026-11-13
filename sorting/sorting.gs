
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
 * Если не переданы, при наличии knownSortData извлекаются из него.
 *
 * knownSortData — заранее прочитанные строки C:L, начиная с
 * CONFIG.firstRow. Используется только при совпадении количества
 * строк и ширины с ожидаемым диапазоном.
 *
 * @return {boolean} true, если хотя бы один блок был отсортирован.
 */
/**
 * Builds and validates a visible-slot permutation without mutating the sheet.
 *
 * Transfer is blocked until full C:L round-trip fidelity has been verified in
 * a real disposable Google workbook.
 */
/**
 * Builds and validates the visible-slot permutation without mutating the sheet.
 *
 * The caller must supply a freshly read visibility context. This function does
 * not infer the cause of hidden rows and never calls a visibility resolver.
 * Transfer is blocked until full C:L round-trip fidelity is verified live.
 */
function sortScenarioBlocksDetailed_(
  sheet,
  sortType = 'scoreDesc',
  knownLastRow = null,
  knownScenarios = null,
  knownSortData = null,
  knownVisibilityMask = null,
  perfTrace = null
) {
  if (!['scoreDesc', 'scoreAsc', 'dots'].includes(sortType)) {
    throw new Error('Неизвестный тип сортировки: ' + sortType);
  }

  if (!sheet) {
    return { status: 'BLOCKED_INPUT', applied: false, reason: 'SORT_SHEET_MISSING' };
  }

  const upperLastRow = knownLastRow !== null
    ? knownLastRow
    : sheet.getLastRow();

  if (!Number.isInteger(upperLastRow)) {
    return { status: 'BLOCKED_INPUT', applied: false, reason: 'INVALID_LAST_ROW' };
  }
  if (upperLastRow < CONFIG.firstRow) {
    return { status: 'NO_OP', applied: false, reason: 'NO_DATA' };
  }

  const upperRowCount = upperLastRow - CONFIG.firstRow + 1;
  const canReuseSnapshot =
    Array.isArray(knownSortData) &&
    knownSortData.length === upperRowCount &&
    knownSortData.every(row =>
      Array.isArray(row) && row.length >= CONFIG.sortWidth
    );

  const allData = canReuseSnapshot
    ? knownSortData.map(row => row.slice(0, CONFIG.sortWidth))
    : sheet.getRange(
        CONFIG.firstRow,
        CONFIG.sortStartCol,
        upperRowCount,
        CONFIG.sortWidth
      ).getValues();

  const lastScenarioRow = getLastScenarioRowFromData(allData);
  if (lastScenarioRow < CONFIG.firstRow) {
    return { status: 'NO_OP', applied: false, reason: 'NO_SCENARIO_DATA' };
  }

  const rowCount = lastScenarioRow - CONFIG.firstRow + 1;
  const data = allData.slice(0, rowCount);

  if (knownScenarios !== null) {
    if (
      !Array.isArray(knownScenarios) ||
      knownScenarios.length < rowCount ||
      !knownScenarios.slice(0, rowCount).every(row =>
        Array.isArray(row) && row.length > 0
      )
    ) {
      return {
        status: 'BLOCKED_INPUT',
        applied: false,
        reason: 'KNOWN_SCENARIOS_RANGE_MISMATCH'
      };
    }

    for (let index = 0; index < rowCount; index++) {
      const supplied = String(knownScenarios[index][0] || '').trim();
      const actual = String(data[index][0] || '').trim();

      if (supplied !== actual) {
        return {
          status: 'BLOCKED_INPUT',
          applied: false,
          reason: 'KNOWN_SCENARIOS_SNAPSHOT_MISMATCH',
          row: CONFIG.firstRow + index
        };
      }
    }
  }

  const visibilityValidationStarted = perfTrace ? Date.now() : 0;
  const visibilityValidation = validateSortVisibilityContext_(
    knownVisibilityMask,
    sheet,
    data,
    CONFIG.firstRow,
    lastScenarioRow
  );
  if (perfTrace) triggerPerfStep_(perfTrace, 'Validate Resolver context', visibilityValidationStarted, {
    ok: visibilityValidation.ok,
    reason: visibilityValidation.reason || null,
    rows: rowCount
  });

  if (!visibilityValidation.ok) {
    return {
      status: 'BLOCKED_VISIBILITY',
      applied: false,
      reason: visibilityValidation.reason,
      visibilityContext: knownVisibilityMask || null
    };
  }

  try {
    const plan = buildVisibleScenarioSortPlan_(
      data,
      CONFIG.firstRow,
      lastScenarioRow,
      knownVisibilityMask,
      sortType,
      perfTrace
    );

    if (!plan.changed) {
      return {
        status: 'NO_OP',
        applied: false,
        reason: 'ALREADY_SORTED_OR_NOT_ENOUGH_VISIBLE_ROWS',
        plan: plan,
        visibilityContext: knownVisibilityMask,
        fingerprint: knownVisibilityMask.stateFingerprint
      };
    }

    console.warn('[VISIBLE SORT] TRANSFER_NOT_VERIFIED', JSON.stringify({
      sheet: sheet.getName(),
      sortType: sortType,
      blocks: plan.blocks.length,
      visibleRows: plan.blocks.reduce(
        (total, block) => total + block.visibleRows.length,
        0
      ),
      reason: 'FULL_CELL_STAGING_ROUND_TRIP_NOT_VERIFIED'
    }));

    return {
      status: 'TRANSFER_NOT_VERIFIED',
      applied: false,
      reason: 'FULL_CELL_STAGING_ROUND_TRIP_NOT_VERIFIED',
      plan: plan,
      visibilityContext: knownVisibilityMask,
      fingerprint: knownVisibilityMask.stateFingerprint
    };

  } catch (error) {
    return {
      status: 'BLOCKED_PLAN',
      applied: false,
      reason: error && error.message ? error.message : String(error),
      visibilityContext: knownVisibilityMask
    };
  }
}

/**
 * Backward-compatible public Boolean entry point.
 *
 * The first five arguments keep their H1 meaning. A fresh, validated sixth
 * visibility snapshot is mandatory. Structured callers should use the
 * detailed function so they can handle NO_OP/abort statuses explicitly.
 */
function sortScenarioBlocks(
  sheet,
  sortType = 'scoreDesc',
  knownLastRow = null,
  knownScenarios = null,
  knownSortData = null,
  knownVisibilityMask = null,
  perfTrace = null
) {
  const result = sortScenarioBlocksDetailed_(
    sheet,
    sortType,
    knownLastRow,
    knownScenarios,
    knownSortData,
    knownVisibilityMask,
    perfTrace
  );

  if (result.status === 'NO_OP') return false;

  if (result.status === 'TRANSFER_NOT_VERIFIED') {
    const transferError = new Error(
      'Сортировка отменена: перенос C:L через staging не подтверждён реальным round-trip.'
    );
    transferError.code = result.reason;
    throw transferError;
  }

  if (result.status !== 'SUCCESS' || result.applied !== true) {
    const sortError = new Error(
      'Сортировка отменена preflight-проверкой: ' +
      (result.reason || result.status)
    );
    sortError.code = result.reason || result.status;
    throw sortError;
  }

  return true;
}

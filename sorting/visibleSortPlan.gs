/**
 * PURE VISIBLE SORT PLANNER.
 *
 * Computes a permutation only. It never reads or writes a spreadsheet.
 */
function validateVisibleSortMask_(mask, firstRow, lastRow) {
  if (!mask || typeof mask !== 'object' || mask.ok !== true) {
    throw new Error('VISIBILITY_MASK_MISSING_OR_UNVERIFIED');
  }
  if (mask.firstRow !== firstRow || mask.lastRow !== lastRow) {
    throw new Error('VISIBILITY_MASK_RANGE_MISMATCH');
  }

  const expectedLength = lastRow - firstRow + 1;
  if (
    !Array.isArray(mask.visibleByPhysicalPosition) ||
    mask.visibleByPhysicalPosition.length !== expectedLength ||
    !mask.visibleByPhysicalPosition.every(value => typeof value === 'boolean')
  ) {
    throw new Error('VISIBILITY_MASK_INVALID');
  }
}

function makeVisibleSortRecord_(row, physicalRow, originalOrdinal, sortType) {
  const dots = row[CONFIG.dotsCol - CONFIG.sortStartCol];
  const rawScore = row[CONFIG.scoreCol - CONFIG.sortStartCol];

  if (typeof dots !== 'string' || dots.length === 0) {
    throw new Error('INVALID_DOTS at row ' + physicalRow);
  }
  if (!/^[\x00-\x7F]+$/.test(dots)) {
    throw new Error('NON_ASCII_DOTS at row ' + physicalRow);
  }

  let score = null;
  let blankScore = false;

  if (sortType !== 'dots') {
    blankScore =
      rawScore === '' ||
      rawScore === null ||
      rawScore === undefined;

    if (!blankScore && (
      typeof rawScore !== 'number' ||
      !Number.isFinite(rawScore)
    )) {
      throw new Error('INVALID_SCORE at row ' + physicalRow);
    }

    if (!blankScore) score = rawScore;
  }

  return {
    sourceRow: physicalRow,
    originalOrdinal: originalOrdinal,
    score: score,
    blankScore: blankScore,
    dots: dots
  };
}

function compareVisibleSortRecords_(left, right, sortType) {
  if (sortType === 'scoreDesc' || sortType === 'scoreAsc') {
    // Blank Score goes last and stays stable, matching the specification.
    if (left.blankScore !== right.blankScore) {
      return left.blankScore ? 1 : -1;
    }
    if (left.blankScore && right.blankScore) {
      return left.originalOrdinal - right.originalOrdinal;
    }
    if (left.score !== right.score) {
      return sortType === 'scoreDesc'
        ? right.score - left.score
        : left.score - right.score;
    }
  } else if (sortType !== 'dots') {
    throw new Error('UNKNOWN_SORT_TYPE: ' + sortType);
  }

  if (left.dots !== right.dots) {
    return left.dots < right.dots ? -1 : 1;
  }
  return left.originalOrdinal - right.originalOrdinal;
}

/**
 * Hidden physical positions may define scenario boundaries but cannot be
 * sources or destinations.
 */
function buildVisibleScenarioSortPlan_(data, firstRow, lastRow, mask, sortType) {
  if (!['scoreDesc', 'scoreAsc', 'dots'].includes(sortType)) {
    throw new Error('UNKNOWN_SORT_TYPE: ' + sortType);
  }
  if (!Array.isArray(data) || data.length !== lastRow - firstRow + 1) {
    throw new Error('SORT_DATA_RANGE_MISMATCH');
  }
  if (!data.every(row => Array.isArray(row) && row.length >= CONFIG.sortWidth)) {
    throw new Error('SORT_DATA_WIDTH_INVALID');
  }

  validateVisibleSortMask_(mask, firstRow, lastRow);

  const visibleMask = mask.visibleByPhysicalPosition;
  const blocks = [];
  let blockStart = -1;
  let currentScenario = '';

  function flushBlock(endIndex) {
    if (blockStart < 0) return;

    const visibleRows = [];
    const records = [];

    for (let index = blockStart; index <= endIndex; index++) {
      if (!visibleMask[index]) continue;

      const physicalRow = firstRow + index;
      visibleRows.push(physicalRow);
      records.push(makeVisibleSortRecord_(
        data[index],
        physicalRow,
        index,
        sortType
      ));
    }

    if (visibleRows.length >= 2) {
      const sortedRecords = records.slice().sort(
        (left, right) => compareVisibleSortRecords_(left, right, sortType)
      );
      const sourceRows = sortedRecords.map(record => record.sourceRow);
      const changed = sourceRows.some((row, index) => row !== visibleRows[index]);

      if (changed) {
        blocks.push({
          scenario: currentScenario,
          startRow: firstRow + blockStart,
          endRow: firstRow + endIndex,
          visibleRows: visibleRows.slice(),
          sourceRows: sourceRows,
          destinations: visibleRows.map((destinationRow, index) => ({
            sourceRow: sourceRows[index],
            destinationRow: destinationRow
          }))
        });
      }
    }

    blockStart = -1;
    currentScenario = '';
  }

  for (let index = 0; index < data.length; index++) {
    const scenario = String(data[index][0] || '').trim();

    if (!scenario) {
      flushBlock(index - 1);
      continue;
    }

    if (blockStart < 0) {
      blockStart = index;
      currentScenario = scenario;
    } else if (scenario !== currentScenario) {
      flushBlock(index - 1);
      blockStart = index;
      currentScenario = scenario;
    }
  }

  flushBlock(data.length - 1);

  const plan = {
    firstRow: firstRow,
    lastRow: lastRow,
    sortType: sortType,
    blocks: blocks,
    changed: blocks.length > 0
  };

  assertVisibleSortPlan_(plan, mask);
  return plan;
}

function assertVisibleSortPlan_(plan, mask) {
  validateVisibleSortMask_(mask, plan.firstRow, plan.lastRow);

  const visibleRows = new Set();
  mask.visibleByPhysicalPosition.forEach((visible, index) => {
    if (visible) visibleRows.add(plan.firstRow + index);
  });

  const allSources = new Set();
  const allDestinations = new Set();

  plan.blocks.forEach(block => {
    const sourceSet = new Set();
    const destinationSet = new Set();

    if (
      block.sourceRows.length !== block.visibleRows.length ||
      block.destinations.length !== block.visibleRows.length
    ) {
      throw new Error('SORT_PLAN_LENGTH_MISMATCH');
    }

    block.sourceRows.forEach(row => {
      if (row < block.startRow || row > block.endRow) {
        throw new Error('SOURCE_OUTSIDE_SCENARIO_BLOCK: ' + row);
      }
      if (!visibleRows.has(row)) {
        throw new Error('FORBIDDEN_SOURCE_ROW: ' + row);
      }
      if (sourceSet.has(row) || allSources.has(row)) {
        throw new Error('DUPLICATE_SOURCE_ROW: ' + row);
      }
      sourceSet.add(row);
      allSources.add(row);
    });

    block.visibleRows.forEach(row => {
      if (row < block.startRow || row > block.endRow) {
        throw new Error('DESTINATION_OUTSIDE_SCENARIO_BLOCK: ' + row);
      }
      if (!visibleRows.has(row)) {
        throw new Error('FORBIDDEN_DESTINATION_ROW: ' + row);
      }
      if (destinationSet.has(row) || allDestinations.has(row)) {
        throw new Error('DUPLICATE_DESTINATION_ROW: ' + row);
      }
      destinationSet.add(row);
      allDestinations.add(row);
    });

    if (sourceSet.size !== destinationSet.size) {
      throw new Error('SORT_PLAN_NOT_A_PERMUTATION');
    }
    sourceSet.forEach(row => {
      if (!destinationSet.has(row)) {
        throw new Error('SORT_PLAN_SOURCE_DESTINATION_SET_MISMATCH');
      }
    });

    block.destinations.forEach((move, index) => {
      if (
        move.sourceRow !== block.sourceRows[index] ||
        move.destinationRow !== block.visibleRows[index]
      ) {
        throw new Error('SORT_PLAN_MAPPING_MISMATCH');
      }
    });
  });

  return true;
}

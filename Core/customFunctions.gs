function energy(score, requirements) {
  const reqs = requirements[0]; // Already a 2D array, get the first row

  const rankRow = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()
    .getRange(2, 10, 1, reqs.length)
    .getValues()[0]; // Fetch rank values once

  // --- Early out for score outside the bounds ---
  if (score < reqs[0]) {
    return interpolate(score, 0, reqs[0], 0, rankRow[0]);
  }

  if (score >= reqs[reqs.length - 1]) {
    const lastIndex = reqs.length - 1;

    return interpolate(
      score,
      reqs[lastIndex - 1],
      reqs[lastIndex],
      rankRow[lastIndex - 1],
      rankRow[lastIndex]
    );
  }

  // --- Optimized Binary Search (direct access of ranges) ---
  let low = 0;
  let high = reqs.length - 2;

  while (low <= high) {
    const mid = (low + high) >>> 1;

    if (score >= reqs[mid] && score < reqs[mid + 1]) {
      return interpolate(
        score,
        reqs[mid],
        reqs[mid + 1],
        rankRow[mid],
        rankRow[mid + 1]
      );

    } else if (score < reqs[mid]) {
      high = mid - 1;

    } else {
      low = mid + 1;
    }
  }

  return 'Error: No valid thresholds or ranks found.';
}


/**
 * Интерполяция между двумя порогами.
 */
function interpolate(
  score,
  currentThreshold,
  nextThreshold,
  currentRank,
  nextRank
) {
  return currentRank +
    (score - currentThreshold) /
    (nextThreshold - currentThreshold) *
    (nextRank - currentRank);
}



function getRank(
  energy,
  userScores,
  energyRequirements,
  scoreThresholds,
  headers,
  showCompleteRank,
  sourceLabel = 'UNKNOWN'
) {
  const started = Date.now();

  // Сохраняем размеры входных данных до преобразования массивов.
  const inputInfo = {
    source: sourceLabel,
    userScores: Array.isArray(userScores)
      ? userScores.length + 'x' + (userScores[0]?.length || 0)
      : 'scalar',
    energyRequirements: Array.isArray(energyRequirements)
      ? energyRequirements.length + 'x' + (energyRequirements[0]?.length || 0)
      : 'scalar',
    scoreThresholds: Array.isArray(scoreThresholds)
      ? scoreThresholds.length + 'x' + (scoreThresholds[0]?.length || 0)
      : 'scalar',
    headers: Array.isArray(headers)
      ? headers.length + 'x' + (headers[0]?.length || 0)
      : 'scalar',
    showCompleteRank: showCompleteRank,
    energy: energy
  };

  var rank = "Unranked";
  var energyIndex = -1;
  var highestCompleteRank = "Unranked";

  // Convert inputs to arrays
  userScores = userScores.flat();
  energyRequirements = energyRequirements.flat();
  scoreThresholds = scoreThresholds.map(row => row.flat());
  headers = headers.flat();

  let energyChecks = 0;
  let scoreChecks = 0;
  let completeRankChecks = 0;

  // Find the highest index of the energy requirement for the normal rank
  for (var i = 0; i < energyRequirements.length; i++) {
    energyChecks++;

    if (energy >= energyRequirements[i]) {
      rank = headers[i] + " ";
      energyIndex = i;
    }
  }

  // Check if all user scores meet the score thresholds for the complete rank
  var isComplete = true;

  if (energyIndex !== -1) {
    for (var i = 0; i < userScores.length; i++) {
      scoreChecks++;

      if (userScores[i] < scoreThresholds[i][energyIndex]) {
        isComplete = false;
        break;
      }
    }
  } else {
    isComplete = false;
  }

  // Determine the highest complete rank if the toggle is on
  if (showCompleteRank) {
    for (var i = 0; i < headers.length; i++) {
      var currentRank = headers[i];
      var complete = true;

      for (var j = 0; j < userScores.length; j++) {
        completeRankChecks++;

        if (userScores[j] < scoreThresholds[j][i]) {
          complete = false;
          break;
        }
      }

      if (complete) {
        highestCompleteRank = currentRank;
      }
    }
  }

  // Return the highest complete rank if the toggle is on and requirements are met
  var result;

  if (showCompleteRank && highestCompleteRank !== "Unranked") {
    result = highestCompleteRank + " Complete";

  } else if (isComplete && rank !== "Unranked") {
    result = rank + " Complete";

  } else if (rank !== "Unranked" && showCompleteRank !== true) {
    result = rank;

  } else {
    result = "Unranked";
  }

  // Diagnostics
  console.log('[GETRANK DIAGNOSTICS]', JSON.stringify({
    ...inputInfo,
    result: result,
    energyChecks: energyChecks,
    scoreChecks: scoreChecks,
    completeRankChecks: completeRankChecks,
    durationMs: Date.now() - started
  }));

  return result;
}


/**
 * Определяет цвет на основе значения Energy и порогов.
 */
function matchColor(
  energyVal,
  firstThreshold,
  lastThreshold,
  divisor,
  colorList,
  fallbackColor
) {
  if (energyVal < firstThreshold) {
    return fallbackColor;
  }

  // Google Sheets передаёт массив цветов как двумерный массив.
  const flatColors = colorList.flat();

  const diff = Math.min(energyVal, lastThreshold) - firstThreshold;
  let index = Math.floor(diff / divisor);

  // Ограничиваем индекс границами массива.
  if (index >= flatColors.length) {
    index = flatColors.length - 1;
  }

  if (index < 0) {
    index = 0;
  }

  return flatColors[index];
}


/**
 * Возвращает следующий применимый порог из горизонтального диапазона.
 * Если значение превышает все пороги, возвращает последний.
 */
function showNextRank(value, rankRange) {
  const ranks = rankRange[0];

  for (let i = 0; i < ranks.length; i++) {
    if (value < ranks[i]) {
      return ranks[i];
    }
  }

  return ranks[ranks.length - 1];
}

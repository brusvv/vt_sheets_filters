
/**
 * FILTER DATA
 *
 * Получение и подготовка данных для фильтрации.
 * Определение последней игровой строки.
 * Преобразование режима TOP в количество результатов.
 */


/**
 * Преобразует режим фильтрации в количество лучших результатов.
 *
 * BEST SCORES -> 1
 * TOP-5       -> 5
 * TOP-10      -> 10
 * 3           -> 3
 *
 * Если режим не распознан, возвращает null.
 */
function getTopN(mode) {
  const normalizedMode = String(mode || '')
    .trim()
    .toUpperCase();

  if (normalizedMode === 'BEST SCORES') return 1;

  const match = normalizedMode.match(/^TOP-(\d+)$/);
  const number = match ? Number(match[1]) : Number(normalizedMode);

  if (Number.isInteger(number) && number > 0) {
    return number;
  }

  return null;
}


/**
 * Определяет последнюю строку с игровыми данными.
 *
 * Использует getLastScenarioRowFromData().
 *
 * Оставлена для совместимости с другими функциями.
 */
function getLastScenarioRow(sheet) {
  const totalStarted = Date.now();

  let stepStarted = Date.now();

  const sheetLastRow = sheet.getLastRow();

  console.log(
    '[LAST SCENARIO TIMING] getLastRow: ' +
    (Date.now() - stepStarted) +
    ' ms'
  );

  if (sheetLastRow < CONFIG.firstRow) {
    console.log(
      '[LAST SCENARIO TIMING] TOTAL: ' +
      (Date.now() - totalStarted) +
      ' ms'
    );

    return CONFIG.firstRow - 1;
  }

  const numRows = sheetLastRow - CONFIG.firstRow + 1;

  stepStarted = Date.now();

  const range = sheet.getRange(
    CONFIG.firstRow,
    CONFIG.scenarioCol,
    numRows,
    CONFIG.dataWidth
  );

  console.log(
    '[LAST SCENARIO TIMING] getRange: ' +
    (Date.now() - stepStarted) +
    ' ms'
  );

  stepStarted = Date.now();

  const data = range.getValues();

  console.log(
    '[LAST SCENARIO TIMING] getValues: ' +
    (Date.now() - stepStarted) +
    ' ms'
  );

  stepStarted = Date.now();

  const result = getLastScenarioRowFromData(data);

  console.log(
    '[LAST SCENARIO TIMING] Search array: ' +
    (Date.now() - stepStarted) +
    ' ms'
  );

  console.log(
    '[LAST SCENARIO TIMING] TOTAL: ' +
    (Date.now() - totalStarted) +
    ' ms'
  );

  return result;
}


/**
 * Определяет последнюю игровую строку
 * по уже полученному массиву данных, начиная с C.
 *
 * Проверяет наличие названия сценария в C.
 */
function getLastScenarioRowFromData(data) {
  for (let i = data.length - 1; i >= 0; i--) {
    const scenario = String(data[i][0] || '').trim();

    if (scenario) {
      return CONFIG.firstRow + i;
    }
  }

  return CONFIG.firstRow - 1;
}


/**
 * Читает игровые данные C:L и настройки I1:J1.
 *
 * Раздельно измеряет:
 * - получение диапазона;
 * - getValues();
 * - локальную обработку массива.
 */
function readFilterInputData_(
  sheet,
  sheetLastRow,
  perfTrace
) {
  const stepStarted = Date.now();

  let readStarted = Date.now();

  const dataRange = sheet.getRange(
    1,
    CONFIG.scenarioCol,
    sheetLastRow,
    CONFIG.sortWidth
  );

  const getRangeMs = Date.now() - readStarted;

  readStarted = Date.now();

  const allData = dataRange.getValues();

  const getValuesMs = Date.now() - readStarted;

  readStarted = Date.now();

  const filterSettings = allData[0];

  const dotsFilter = String(
    filterSettings[
      CONFIG.dotsCol - CONFIG.scenarioCol
    ] || ''
  ).trim();

  const mode = String(
    filterSettings[
      CONFIG.dotsCol - CONFIG.scenarioCol + 1
    ] || ''
  ).trim();

  const data = allData.slice(CONFIG.firstRow - 1);

  const topN = getTopN(mode);

  const selectedDots = new Set(
    dotsFilter
      ? dotsFilter
          .split(',')
          .map(value => value.trim())
          .filter(Boolean)
      : []
  );

  const processingMs = Date.now() - readStarted;

  console.log(
    '[FILTER READ BREAKDOWN] ' +
    JSON.stringify({
      sheet: sheet.getName(),
      range: dataRange.getA1Notation(),
      rows: sheetLastRow,
      columns: CONFIG.sortWidth,
      getRangeMs: getRangeMs,
      getValuesMs: getValuesMs,
      processingMs: processingMs
    })
  );

  filterPerfStep_(
    perfTrace,
    'Read data and filter settings',
    stepStarted,
    {
      sheet: sheet.getName(),
      mode: mode,
      topN: topN,
      selectedDots: Array.from(selectedDots),
      getRangeMs: getRangeMs,
      getValuesMs: getValuesMs,
      processingMs: processingMs
    }
  );

  return {
    data: data,
    mode: mode,
    topN: topN,
    selectedDots: selectedDots
  };
}

/**
 * TRIGGERS
 *
 * Обработка событий Google Sheets.
 * Маршрутизация изменений к соответствующим операциям.
 *
 * Комплексные операции находятся в отдельных Operations-файлах.
 */


/**
 * Создание меню при открытии таблицы.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('Dots Filter')
    .addItem('Reset Filters', 'resetFilters')
    .addItem('Save Manual Hidden Rows', 'saveManualHiddenRows')
    .addSeparator()
    .addItem('Duplicate Scenario', 'duplicateScenario')
    .addToUi();

  ui.createMenu('Sorting')
    .addItem('Sort by Score (Descending)', 'manualSortByScoreDesc')
    .addItem('Sort by Score (Ascending)', 'manualSortByScoreAsc')
    .addItem('Sort by Dots', 'manualSortByDots')
    .addToUi();

  ui.createMenu('Rank Scrolling Setup')
    .addItem('Set Up Triggers', 'createTriggers')
    .addItem('Force Full Rank Sync', 'forceRankScroller')
    .addToUi();
}


/**
 * Проверяет пересечение изменённого диапазона
 * с указанным диапазоном без дополнительных
 * обращений к Google Sheets.
 */
function rankEditIntersects_(e, a1Notation) {
  if (!e || !e.range || !a1Notation) {
    return false;
  }

  function columnToNumber_(letters) {
    return columnToIndex_(
      letters.replace(/\$/g, '').toUpperCase()
    );
  }

  const match = a1Notation.match(
    /^(\$?[A-Z]+)(\$?\d*)?(?::(\$?[A-Z]+)(\$?\d*)?)?$/i
  );

  if (!match) {
    throw new Error(
      `Некорректная A1-нотация: ${a1Notation}`
    );
  }

  const targetFirstCol = columnToNumber_(match[1]);

  const targetLastCol = match[3]
    ? columnToNumber_(match[3])
    : targetFirstCol;

  const targetFirstRow = match[2]
    ? Number(match[2].replace(/\$/g, ''))
    : 1;

  const targetLastRow = match[4]
    ? Number(match[4].replace(/\$/g, ''))
    : match[3]
      ? Infinity
      : targetFirstRow;

  const range = e.range;

  const editFirstRow = range.getRow();

  const editLastRow =
    editFirstRow + range.getNumRows() - 1;

  const editFirstCol = range.getColumn();

  const editLastCol =
    editFirstCol + range.getNumColumns() - 1;

  return (
    editFirstRow <= targetLastRow &&
    editLastRow >= targetFirstRow &&
    editFirstCol <= targetLastCol &&
    editLastCol >= targetFirstCol
  );
}


/**
 * Основной обработчик изменений.
 */
function handleEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  const trace = triggerPerfStart_('handleEdit', {
    sheet: sheetName,
    range: e.range.getA1Notation(),
    rows: e.range.getNumRows(),
    columns: e.range.getNumColumns()
  });

  let status = 'SUCCESS';
  let errorMetadata = {};

  try {

    // -----------------------------------------
    // RANK REQUIREMENTS
    // -----------------------------------------

    if (sheetName === 'RankReqs') {
      const requirementsRange =
        CONFIG.triggers.rankRequirements;

      const requirementsChanged =
        rankEditIntersects_(
          e,
          getA1Range_(requirementsRange)
        );

      if (requirementsChanged) {
        triggerPerfEvent_(
          trace,
          'RANK_REQUIREMENTS_CHANGED'
        );

        markAllRankSheetsDirty();

        runRankScroller(false, trace);
      }

      return;
    }


    // -----------------------------------------
    // OPTIONS
    // -----------------------------------------

    if (sheetName === CONFIG.optionsSheetName) {
      const autoSortChanged =
        rankEditIntersects_(
          e,
          CONFIG.options.autoSort.cell
        );

      const hideEmptyScoresChanged =
        rankEditIntersects_(
          e,
          CONFIG.options.hideEmptyScores.cell
        );

      const manualStateSyncChanged =
        rankEditIntersects_(
          e,
          CONFIG.options.manualStateSync.cell
        );

      const rankOptionsChanged =
        rankEditIntersects_(
          e,
          CONFIG.options.rankTiers.cell
        ) ||
        rankEditIntersects_(
          e,
          CONFIG.options.columnsBehind.cell
        ) ||
        rankEditIntersects_(
          e,
          CONFIG.options.columnsAhead.cell
        );


      // ---------------------------------------
      // AUTO SORT
      // ---------------------------------------

      if (autoSortChanged) {
        triggerPerfEvent_(
          trace,
          'AUTO_SORT_OPTIONS_CHANGED'
        );

        markAllRankSheetsDirty();

        refreshAllFiltersAndRanks(trace);
      }


      // ---------------------------------------
      // HIDE EMPTY SCORES
      // ---------------------------------------

      if (hideEmptyScoresChanged) {
        triggerPerfEvent_(
          trace,
          'HIDE_EMPTY_SCORES_CHANGED'
        );

        refreshAllFiltersOnly(trace);
      }


      // ---------------------------------------
      // MANUAL STATE SYNC
      // ---------------------------------------

      if (manualStateSyncChanged) {
        triggerPerfEvent_(
          trace,
          'MANUAL_STATE_SYNC_MODE_CHANGED',
          {
            cell: CONFIG.options.manualStateSync.cell
          }
        );

        // Режим ManualState обрабатывается
        // отдельной логикой фильтров.
      }


      // ---------------------------------------
      // RANK OPTIONS
      // ---------------------------------------

      if (rankOptionsChanged) {
        triggerPerfEvent_(
          trace,
          'RANK_OPTIONS_CHANGED'
        );

        markAllRankSheetsDirty();

        runRankScroller(false, trace);
      }

      return;
    }


    // -----------------------------------------
    // ОСНОВНЫЕ ЛИСТЫ
    // -----------------------------------------

    if (!RANK_TRACKED_SHEETS.includes(sheetName)) {
      triggerPerfEvent_(
        trace,
        'IGNORED_SHEET'
      );

      return;
    }


    // -----------------------------------------
    // CLASSIFY CHANGES
    // -----------------------------------------

    const filterSettingsChanged =
      rankEditIntersects_(
        e,
        CONFIG.dotsFilterCell
      ) ||
      rankEditIntersects_(
        e,
        CONFIG.modeFilterCell
      );

    const scenarioChanged =
      rankEditIntersects_(
        e,
        getColumnDataRange_(
          CONFIG.scenarioCol
        )
      );

    const scoreChanged =
      rankEditIntersects_(
        e,
        getColumnDataRange_(
          CONFIG.scoreCol
        )
      );

    const dotsChanged =
      rankEditIntersects_(
        e,
        getColumnDataRange_(
          CONFIG.dotsCol
        )
      );


    // -----------------------------------------
    // FILTER ROUTING
    // -----------------------------------------

    const filterChanged =
      filterSettingsChanged ||
      scenarioChanged ||
      scoreChanged ||
      dotsChanged;


    // -----------------------------------------
    // SORT ROUTING
    // -----------------------------------------

    const sortRelevantChanged =
      scenarioChanged ||
      scoreChanged;


    // -----------------------------------------
    // RANK SCROLLER ROUTING
    // -----------------------------------------

    const rankScrollerConfig =
      CONFIG.triggers.rankScroller;

    const rankScrollerRelevantChanged =
      scenarioChanged ||
      scoreChanged ||

      rankScrollerConfig.fixedRanges.some(
        range => rankEditIntersects_(e, range)
      ) ||

      rankScrollerConfig.dataColumns.some(
        column => rankEditIntersects_(
          e,
          getColumnDataRange_(column)
        )
      );


    // -----------------------------------------
    // NO RELEVANT CHANGES
    // -----------------------------------------

    if (
      !filterChanged &&
      !rankScrollerRelevantChanged
    ) {
      triggerPerfEvent_(
        trace,
        'NO_RELEVANT_CHANGES'
      );

      return;
    }


    // -----------------------------------------
    // MARK RANK SHEET DIRTY
    // -----------------------------------------

    if (rankScrollerRelevantChanged) {
      markRankSheetDirty(sheetName);
    }


    // -----------------------------------------
    // FILTER UPDATE
    // -----------------------------------------

    if (filterChanged) {
      triggerPerfEvent_(
        trace,
        'FILTER_UPDATE_REQUIRED',
        {
          sheet: sheetName,
          sortRequired: sortRelevantChanged,
          rankScrollerRequired: rankScrollerRelevantChanged
        }
      );

      refreshFiltersAndRanks(
        !sortRelevantChanged,
        sheet,
        trace,
        rankScrollerRelevantChanged
      );

      return;
    }


    // -----------------------------------------
    // RANK SCROLLER ONLY
    // -----------------------------------------

    if (rankScrollerRelevantChanged) {
      triggerPerfEvent_(
        trace,
        'RANK_SCROLLER_ONLY',
        {
          sheet: sheetName
        }
      );

      runRankScroller(false, trace);
    }

  } catch (error) {
    status = 'ERROR';

    errorMetadata = {
      message: error.message,
      stack: error.stack
    };

    throw error;

  } finally {
    triggerPerfEnd_(
      trace,
      status,
      errorMetadata
    );
  }
}


/**
 * Установка триггеров.
 *
 * Удаляет старые управляемые триггеры
 * и создаёт один installable onEdit.
 */
function createTriggers() {
  const managedHandlers = [
    'onEdit',
    'handleEdit',
    'rankScroller'
  ];

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (
      managedHandlers.includes(
        trigger.getHandlerFunction()
      )
    ) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger('handleEdit')
    .forSpreadsheet(
      SpreadsheetApp.getActiveSpreadsheet()
    )
    .onEdit()
    .create();

  Logger.log('Triggers successfully installed.');
}


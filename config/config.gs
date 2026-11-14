/**
 * Основная конфигурация проекта.
 */

const CONFIG = {
  sheetName: 'Intermediate',

  // Первая строка с игровыми данными.
  firstRow: 3,

  // Столбцы.
  scenarioCol: 3, // C — название сценария
  scoreCol: 5,    // E — High Score
  dotsCol: 9,     // I — Dots

  // Ячейки фильтров.
  dotsFilterCell: 'I1',
  modeFilterCell: 'J1',

  // Диапазон данных для фильтрации: C:I.
  dataWidth: 7,

  // Столбцы, очищаемые при дублировании.
  clearScoreCol: 5,
  clearDotsStartCol: 9,
  clearDotsEndCol: 11,

  // Настройки сортировки.
  sortStartCol: 3,
  sortWidth: 10,

  // Безопасный Visible Sort Transfer остаётся выключенным до успешного
  // интеграционного теста на отдельной копии Google Spreadsheet.
  visibleSortTransferEnabled: false,

  // Смещения относительно диапазона C:L.
  sortScoreOffset: 3,
  sortDotsOffset: 7,

  sortScoreAscending: false,
  sortDotsAscending: true,

  // Настройки Options.
  optionsSheetName: 'Options',

  options: {
    rankTiers: {cell: 'A5'},
    autoSort: {cell: 'A6'},
    hideEmptyScores: {cell: 'A7'},
    manualStateSync: {cell: 'A8', defaultValue: 'Manual'},
    columnsBehind: {cell: 'A10'},
    columnsAhead: {cell: 'A11'}
  },

  // Настройки Rank Scroller.
  rankScroller: {
    sheets: {
      Novice: {
        energyRow: 2,
        energyColumn: 13,
        thresholdStartColumn: 14,
        thresholdColumns: 15,
        propertyKey: 'overallEnergyNovice'
      },

      Intermediate: {
        energyRow: 2,
        energyColumn: 13,
        thresholdStartColumn: 14,
        thresholdColumns: 16,
        propertyKey: 'overallEnergyIntermediate'
      },

      Advanced: {
        energyRow: 2,
        energyColumn: 13,
        thresholdStartColumn: 14,
        thresholdColumns: 16,
        propertyKey: 'overallEnergyAdvanced'
      }
    }
  },

  // Настройки триггеров.
  triggers: {
    rankRequirements: {
      firstRow: 4,
      lastRow: 70,
      firstCol: 2,
      lastCol: 3
    },

    rankScroller: {
      // Фиксированные диапазоны.
      fixedRanges: [
        'B3:C3',
        'M2',
        'AC2',
        'N2',
        'Q2',
        'T2',
        'W2',
        'Z2'
      ],

      // Столбцы с данными.
      dataColumns: [
        14, // N
        17, // Q
        20, // T
        23, // W
        26  // Z
      ]
    }
  },

  // Настройки Performance Tracer.
  perfModeCell: 'A12',
  perfDefaultMode: 'DEBUG'
};


/**
 * Листы, участвующие в Rank Scroller.
 */
const RANK_TRACKED_SHEETS = [
  'Novice',
  'Intermediate',
  'Advanced'
];


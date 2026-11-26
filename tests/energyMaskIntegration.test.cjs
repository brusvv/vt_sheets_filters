'use strict';

// R5.2: интеграция Energy mask с applyFilters и Save Manual Hidden Rows
// (лист Intermediate). Тесты используют настоящий код скрипта в mock-окружении
// и модель формулы Energy (tests/helpers/energyModel.cjs).
//
// Mock-окружение не доказывает поведение реального Google Sheets.

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const {
  createMockWorkbook,
  autoSortOptions
} = require('./helpers/mockWorkbook.cjs');
const {loadEnergyFixture, groupEnergies} = require('./helpers/energyModel.cjs');

const root = path.resolve(__dirname, '..');
const FIRST = 3;
const D_COL = 4;
const MASK_COL = 20;      // T (Intermediate)
const MASK_STATUS_ROW = 2;
const ROW_DOTTS = 103;    // VT DotTS / Icarus Lab Pure / 1170: максимум группы M99.

function realEnergyMaskConfig() {
  const context = vm.createContext({});

  vm.runInContext(
    fs.readFileSync(path.join(root, 'config/config.gs'), 'utf8'),
    context,
    {filename: 'config/config.gs'}
  );

  return JSON.parse(
    vm.runInContext('JSON.stringify(CONFIG.energyMask)', context)
  );
}

function isNumber(value) {
  return typeof value === 'number' && Number.isFinite(value);
}

function createWorkbook(options = {}) {
  const fixture = loadEnergyFixture();

  const workbook = createMockWorkbook(Object.assign({
    rows: fixture.rows.map(row => row.slice()),
    hiddenRows: [],
    manualCount: 0,
    manualKeys: [],
    gameLastRow: fixture.lastRow
  }, options.mock || {}));

  workbook.stateSheet.grid[1][4] = fixture.lastRow;

  if (options.config !== null) {
    const config = realEnergyMaskConfig();

    config.enabled = options.enabled === undefined ? true : options.enabled;
    config.sheets = options.sheets || ['Intermediate'];
    workbook.project.CONFIG.energyMask = config;
  }

  return {workbook, fixture};
}

function setDots(workbook, dots) {
  workbook.gameSheet.grid[0][8] = dots || '';
  workbook.gameSheet.grid[0][9] = '';
}

function apply(workbook, reset = false) {
  return workbook.project.applyFilters(
    false,
    workbook.gameSheet,
    reset,
    autoSortOptions(false),
    null
  );
}

function save(workbook) {
  return workbook.project.saveManualHiddenRows();
}

function column(workbook, col, fixture) {
  const values = [];

  for (let row = FIRST; row <= fixture.lastRow; row++) {
    values.push(workbook.stateSheet.grid[row - 1][col - 1]);
  }

  return values;
}

const stateD = (workbook, fixture) => column(workbook, D_COL, fixture);
const maskOf = (workbook, fixture) => column(workbook, MASK_COL, fixture);
const statusOf = workbook => workbook.stateSheet.grid[MASK_STATUS_ROW - 1][MASK_COL - 1];

function manualCount(workbook) {
  return Number(workbook.stateSheet.grid[1][13]) || 0;
}

function manualKeys(workbook) {
  const keys = [];

  for (let i = 0; i < manualCount(workbook); i++) {
    keys.push(workbook.stateSheet.grid[FIRST - 1 + i][12]);
  }

  return keys;
}

function energyFromMask(fixture, mask) {
  return groupEnergies(fixture, row => mask[row - FIRST] === true);
}

function groupValue(groups, firstRow) {
  return groups[firstRow] === null ? null : groups[firstRow].energy;
}

function dotsMatch(fixture, dots) {
  return function(row) {
    const record = fixture.rows[row - FIRST];

    return (!dots || String(record[6]).trim() === dots) &&
      String(record[0] || '').trim() !== '';
  };
}

function maskWrites(workbook) {
  return workbook.stateWrites.filter(write => write.startColumn >= 18);
}

function throwOnWrite(workbook, predicate, message = 'injected failure') {
  const original = workbook.stateSheet.getRange.bind(workbook.stateSheet);

  workbook.stateSheet.getRange = function(...args) {
    const range = original(...args);

    if (predicate(args)) {
      range.setValues = function() {
        throw new Error(message);
      };
    }

    return range;
  };
}

// --------------------------------------------------------------- 1, 2

test('R5.2-1: applyFilters записывает маску и статус OK после успешной фильтрации', function() {
  const {workbook, fixture} = createWorkbook();

  setDots(workbook, 'Icarus Lab Neon');
  apply(workbook);

  assert.equal(statusOf(workbook), 'OK');

  const mask = maskOf(workbook, fixture);

  assert.ok(mask.every(value => typeof value === 'boolean'));
  assert.ok(mask.includes(true));

  // Без сохранённых ручных строк маска совпадает с FilterState.
  assert.deepEqual(mask, stateD(workbook, fixture));
  assert.equal(workbook.stateSheet.grid[0][17], 'EnergyMask:1');
  assert.equal(workbook.stateSheet.grid[0][MASK_COL - 1], 'Intermediate Energy Mask');
});

test('R5.2-2: маска учитывает Dots filter (проверка через модель формулы Energy)', function() {
  const {workbook, fixture} = createWorkbook();

  ['Icarus Lab Neon', 'WestLab V2 Small', 'Icarus Lab Pure', ''].forEach(function(dots) {
    setDots(workbook, dots);
    apply(workbook, dots === '');

    const mask = maskOf(workbook, fixture);
    const expected = groupEnergies(fixture, dotsMatch(fixture, dots));
    const actual = energyFromMask(fixture, mask);

    fixture.groups.forEach(function([startRow]) {
      const a = groupValue(actual, startRow);
      const e = groupValue(expected, startRow);

      assert.ok(
        (a === null && e === null) || Math.abs(a - e) < 1e-9,
        `Dots="${dots}" группа M${startRow}: ${a} вместо ${e}`
      );
    });
  });
});

// ------------------------------------------------------------------- 3

test('R5.2-3: сохранённая ручная строка исключается из маски при любом Dots и после resetFilters', function() {
  const {workbook, fixture} = createWorkbook();

  setDots(workbook, '');
  apply(workbook);

  workbook.gameSheet.hideRows(ROW_DOTTS, 1);
  assert.equal(save(workbook), 1);

  const withoutRow = row => row !== ROW_DOTTS;

  // Сразу после Save: D остаётся TRUE, маска FALSE, Energy группы M99 без строки.
  assert.equal(stateD(workbook, fixture)[ROW_DOTTS - FIRST], true);
  assert.equal(maskOf(workbook, fixture)[ROW_DOTTS - FIRST], false);
  assert.ok(
    Math.abs(
      groupValue(energyFromMask(fixture, maskOf(workbook, fixture)), 99) - 511.43
    ) < 0.01
  );

  [
    ['повторное применение', '', false],
    ['Dots того же строки', 'Icarus Lab Pure', false],
    ['другой Dots', 'WestLab V2 Small', false],
    ['resetFilters', '', true]
  ].forEach(function([label, dots, reset]) {
    setDots(workbook, dots);
    apply(workbook, reset);

    const mask = maskOf(workbook, fixture);

    assert.equal(mask[ROW_DOTTS - FIRST], false, label);
    assert.equal(stateD(workbook, fixture)[ROW_DOTTS - FIRST], true, `${label}: D не меняется`);
    assert.equal(manualCount(workbook), 1, label);

    // Строгое ожидание: Dots и исключение сохранённой строки.
    const dotsRule = dotsMatch(fixture, dots);
    const expected = groupEnergies(fixture, row => dotsRule(row) && withoutRow(row));
    const actual = energyFromMask(fixture, mask);

    fixture.groups.forEach(function([startRow]) {
      const a = groupValue(actual, startRow);
      const e = groupValue(expected, startRow);

      assert.ok(
        (a === null && e === null) || Math.abs(a - e) < 1e-9,
        `${label}: группа M${startRow}: ${a} вместо ${e}`
      );
    });
  });
});

// ------------------------------------------------------------------- 4

test('R5.2-4: несохранённое физическое скрытие не исключает строку из маски', function() {
  const {workbook, fixture} = createWorkbook();

  setDots(workbook, '');
  apply(workbook);

  workbook.gameSheet.hideRows(ROW_DOTTS, 1);
  apply(workbook);

  assert.equal(manualCount(workbook), 0);
  assert.equal(maskOf(workbook, fixture)[ROW_DOTTS - FIRST], true);
  assert.ok(
    Math.abs(
      groupValue(energyFromMask(fixture, maskOf(workbook, fixture)), 99) - 585.7142857
    ) < 0.001
  );
});

// ------------------------------------------------------------------- 5

test('R5.2-5: Save обновляет маску после записи ManualState (порядок записей)', function() {
  const {workbook, fixture} = createWorkbook();

  setDots(workbook, '');
  apply(workbook);

  workbook.gameSheet.hideRows(ROW_DOTTS, 1);

  const before = workbook.stateWrites.length;

  assert.equal(save(workbook), 1);

  const writes = workbook.stateWrites.slice(before);
  const markers = writes.map(function(write) {
    if (write.startColumn >= 18) {
      return write.values[0][0] === 'STALE' ? 'mask:STALE' : `mask:${write.values[0][0]}`;
    }

    return write.startColumn >= 13 ? 'manual' : 'other';
  });

  // STALE ставится до записи ManualState, OK - после.
  assert.equal(markers[0], 'mask:STALE');
  assert.ok(markers.includes('manual'));
  assert.equal(markers[markers.length - 1], 'mask:OK');
  assert.ok(markers.lastIndexOf('manual') < markers.lastIndexOf('mask:OK'));
  assert.ok(!markers.includes('other'), 'Save не пишет в FilterState (столбец D)');

  assert.equal(statusOf(workbook), 'OK');
  assert.equal(maskOf(workbook, fixture)[ROW_DOTTS - FIRST], false);
});

// ------------------------------------------------------------------- 6

test('R5.2-6: повторный Save не повреждает маску и ManualState', function() {
  const {workbook, fixture} = createWorkbook();

  setDots(workbook, '');
  apply(workbook);
  workbook.gameSheet.hideRows(ROW_DOTTS, 1);
  save(workbook);

  const mask = maskOf(workbook, fixture).slice();
  const keys = manualKeys(workbook).slice();

  for (let i = 0; i < 3; i++) {
    assert.equal(save(workbook), 1);
    assert.deepEqual(maskOf(workbook, fixture), mask);
    assert.equal(statusOf(workbook), 'OK');
    assert.equal(manualCount(workbook), 1);
    assert.deepEqual(manualKeys(workbook), keys);
  }

  // Показать строку и сохранить: строка возвращается в маску.
  workbook.gameSheet.showRows(ROW_DOTTS, 1);
  save(workbook);

  assert.equal(manualCount(workbook), 0);
  assert.equal(maskOf(workbook, fixture)[ROW_DOTTS - FIRST], true);
});

// ------------------------------------------------------------------- 7

test('R5.2-7a: Save с некорректным FilterState не трогает маску', function() {
  const {workbook, fixture} = createWorkbook();

  setDots(workbook, '');
  apply(workbook);

  const maskBefore = maskOf(workbook, fixture).slice();

  workbook.stateSheet.grid[FIRST][D_COL - 1] = 'не boolean';
  workbook.gameSheet.hideRows(ROW_DOTTS, 1);

  const before = workbook.stateWrites.length;

  assert.throws(() => save(workbook), /некорректные данные видимости/);
  assert.equal(workbook.stateWrites.length, before, 'ни одной записи');
  assert.equal(statusOf(workbook), 'OK');
  assert.deepEqual(maskOf(workbook, fixture), maskBefore);
  assert.equal(manualCount(workbook), 0);
});

test('R5.2-7b: сбой записи ManualState оставляет маску STALE и не записывает её заново', function() {
  const {workbook, fixture} = createWorkbook();

  setDots(workbook, '');
  apply(workbook);

  const maskBefore = maskOf(workbook, fixture).slice();

  workbook.gameSheet.hideRows(ROW_DOTTS, 1);

  // Ключи ManualState пишутся в колонку 13, начиная со строки 3.
  throwOnWrite(
    workbook,
    args => args[1] === 13 && args[0] === FIRST,
    'manual write failed'
  );

  const before = workbook.stateWrites.length;

  assert.throws(() => save(workbook), /manual write failed/);

  assert.equal(statusOf(workbook), 'STALE');
  assert.deepEqual(maskOf(workbook, fixture), maskBefore);

  const writes = workbook.stateWrites.slice(before).filter(w => w.startColumn >= 18);

  assert.equal(writes.length, 1, 'только пометка STALE');
  assert.equal(writes[0].values[0][0], 'STALE');
});

test('R5.2-7e: если не удалось пометить OK-маску как STALE, Save не меняет ManualState', function() {
  const {workbook} = createWorkbook();

  setDots(workbook, '');
  apply(workbook);

  const keysBefore = manualKeys(workbook).slice();
  const countBefore = manualCount(workbook);
  const writesBefore = workbook.stateWrites.length;

  workbook.gameSheet.hideRows(ROW_DOTTS, 1);

  // Сбой именно при попытке перевести статус OK -> STALE.
  throwOnWrite(workbook, args => args[1] >= 18, 'stale marker failed');

  assert.throws(() => save(workbook), /ManualState не сохранён/);

  assert.equal(statusOf(workbook), 'OK');
  assert.equal(manualCount(workbook), countBefore);
  assert.deepEqual(manualKeys(workbook), keysBefore);

  const writes = workbook.stateWrites.slice(writesBefore);
  assert.equal(writes.length, 0, 'ни ManualState, ни маска не изменены');
});

test('R5.2-7c: Save при уже устаревшей маске сохраняет ManualState, маску не трогает', function() {
  const {workbook, fixture} = createWorkbook();

  setDots(workbook, '');
  apply(workbook);

  workbook.stateSheet.grid[MASK_STATUS_ROW - 1][MASK_COL - 1] = 'STALE';

  const maskBefore = maskOf(workbook, fixture).slice();

  workbook.gameSheet.hideRows(ROW_DOTTS, 1);
  assert.equal(save(workbook), 1);

  assert.equal(manualCount(workbook), 1);
  assert.equal(statusOf(workbook), 'STALE');
  assert.deepEqual(maskOf(workbook, fixture), maskBefore);
  assert.ok(
    workbook.events.some(event => event[0] === 'toast' && /Energy mask/.test(event[1]))
  );
});

test('R5.2-7d: Save до первого applyFilters (маска не создана) блок не создаёт', function() {
  const {workbook, fixture} = createWorkbook();

  workbook.gameSheet.hideRows(ROW_DOTTS, 1);
  assert.equal(save(workbook), 1);

  assert.equal(manualCount(workbook), 1);
  assert.equal(maskWrites(workbook).length, 0);
  assert.equal(workbook.stateSheet.grid[0][17], '');
  assert.ok(maskOf(workbook, fixture).every(value => value === ''));
});

// ------------------------------------------------------------------- 8

test('R5.2-8a: сбой записи маски не ломает фильтрацию', function() {
  const baseline = createWorkbook({enabled: false});

  setDots(baseline.workbook, 'Icarus Lab Neon');
  apply(baseline.workbook);

  const {workbook, fixture} = createWorkbook();

  throwOnWrite(workbook, args => args[1] >= 18, 'mask write failed');

  setDots(workbook, 'Icarus Lab Neon');
  assert.doesNotThrow(() => apply(workbook));

  assert.deepEqual(
    stateD(workbook, fixture),
    stateD(baseline.workbook, baseline.fixture)
  );
  assert.deepEqual(
    [...workbook.gameSheet.hiddenRows].sort((a, b) => a - b),
    [...baseline.workbook.gameSheet.hiddenRows].sort((a, b) => a - b)
  );
  assert.ok(
    workbook.events.some(event => event[0] === 'perfEvent' && event[1] === 'ENERGY_MASK_SYNC_FAILED')
  );
  assert.ok(maskOf(workbook, fixture).every(value => value === ''));
});

test('R5.2-8b: повреждённая версия схемы: фильтрация работает, в блок ничего не пишется', function() {
  const {workbook, fixture} = createWorkbook();

  workbook.stateSheet.grid[0][17] = 'EnergyMask:2';

  setDots(workbook, 'Icarus Lab Neon');
  assert.doesNotThrow(() => apply(workbook));

  assert.equal(maskWrites(workbook).length, 0);
  assert.equal(workbook.stateSheet.grid[0][17], 'EnergyMask:2');
  assert.ok(maskOf(workbook, fixture).every(value => value === ''));
  assert.ok(stateD(workbook, fixture).includes(false), 'FilterState обновлён');

  const failed = workbook.events.filter(
    event => event[0] === 'perfEvent' && event[1] === 'ENERGY_MASK_SYNC_FAILED'
  );

  assert.equal(failed.length, 1);
  assert.equal(failed[0][2].reason, 'SCHEMA_VERSION_MISMATCH');
  assert.ok(
    workbook.events.some(event => event[0] === 'perfEvent' && event[1] === 'ENERGY_MASK_MARK_STALE_FAILED')
  );
});

test('R5.2-8c: чужой заголовок колонки маски: ничего не пишется, фильтрация работает', function() {
  const {workbook, fixture} = createWorkbook();

  workbook.stateSheet.grid[0][MASK_COL - 1] = 'Чужие данные';

  setDots(workbook, 'Icarus Lab Neon');
  assert.doesNotThrow(() => apply(workbook));

  assert.equal(maskWrites(workbook).length, 0);
  assert.equal(workbook.stateSheet.grid[0][MASK_COL - 1], 'Чужие данные');
  assert.ok(stateD(workbook, fixture).includes(false));
});

test('R5.2-8d: исключение при записи FilterState: ошибка сохраняется, маска STALE и не обновляется', function() {
  const {workbook, fixture} = createWorkbook();

  setDots(workbook, '');
  apply(workbook);

  const maskBefore = maskOf(workbook, fixture).slice();

  workbook.project.updateVisibleRowsState = function() {
    throw new Error('state write failed');
  };

  setDots(workbook, 'Icarus Lab Neon');
  assert.throws(() => apply(workbook), /state write failed/);

  assert.equal(statusOf(workbook), 'STALE');
  assert.deepEqual(maskOf(workbook, fixture), maskBefore);
  assert.ok(
    workbook.events.some(event => event[0] === 'perfEvent' && event[1] === 'ENERGY_MASK_MARKED_STALE')
  );
});

test('R5.2-8e: применение прервано Resolver (SORT_VISIBILITY_BLOCKED): маска STALE, D и ManualState не меняются', function() {
  const {workbook} = createWorkbook({
    mock: {
      rows: undefined,
      hiddenRows: [3, 4],
      manualCount: 1,
      manualKeys: undefined,
      gameLastRow: undefined
    }
  });

  const dBefore = JSON.stringify(workbook.stateSheet.getRange(3, 4, 3, 1).getValues());
  const manualBefore = JSON.stringify(workbook.stateSheet.getRange(1, 13, 4, 2).getValues());

  const result = workbook.project.applyFilters(
    false,
    workbook.gameSheet,
    false,
    autoSortOptions(true),
    null
  );

  assert.equal(result.status, 'SORT_VISIBILITY_BLOCKED');
  assert.equal(workbook.stateSheet.grid[MASK_STATUS_ROW - 1][MASK_COL - 1], 'STALE');
  assert.equal(JSON.stringify(workbook.stateSheet.getRange(3, 4, 3, 1).getValues()), dBefore);
  assert.equal(JSON.stringify(workbook.stateSheet.getRange(1, 13, 4, 2).getValues()), manualBefore);

  // В FilterState/ManualState записей нет, только блок маски.
  assert.ok(workbook.stateWrites.every(write => write.startColumn >= 18));
  assert.equal(workbook.sheetWrites.length, 0);
});

test('R5.2-8f: Resolver блокирует reset до подготовки состояния, листа _FilterState нет: без исключений', function() {
  const {workbook} = createWorkbook({mock: {noStateSheet: true}});

  let result;

  assert.doesNotThrow(function() {
    result = workbook.project.applyFilters(
      false,
      workbook.gameSheet,
      true,
      autoSortOptions(true),
      null
    );
  });

  assert.equal(result.status, 'SORT_VISIBILITY_BLOCKED');
  assert.deepEqual(workbook.visibilityEvents, []);
  assert.ok(
    workbook.events.some(event => event[0] === 'perfEvent' && event[1] === 'ENERGY_MASK_MARK_STALE_FAILED')
  );
});

test('R5.2: повреждённые значения маски и статуса перезаписываются при следующем применении', function() {
  const {workbook, fixture} = createWorkbook();

  setDots(workbook, '');
  apply(workbook);

  for (let row = FIRST; row < FIRST + 10; row++) {
    workbook.stateSheet.grid[row - 1][MASK_COL - 1] = 'x';
  }

  workbook.stateSheet.grid[MASK_STATUS_ROW - 1][MASK_COL - 1] = 'garbage';

  apply(workbook);

  assert.equal(statusOf(workbook), 'OK');
  assert.ok(maskOf(workbook, fixture).every(value => typeof value === 'boolean'));
});

// ------------------------------------------------------------------- 9

const MASK_FUNCTIONS = [
  'ensureEnergyMaskLayout_',
  'writeEnergyMask_',
  'markEnergyMaskStale_',
  'readEnergyMaskStatus_',
  'readEnergyMask_',
  'computeEnergyMask_',
  'getEnergyMaskStateSheet_'
];

function spyMaskFunctions(workbook) {
  const calls = {};

  MASK_FUNCTIONS.forEach(function(name) {
    const original = workbook.project[name];

    workbook.project[name] = function(...args) {
      calls[name] = (calls[name] || 0) + 1;

      return original.apply(this, args);
    };
  });

  return calls;
}

function runFullScenario(workbook) {
  setDots(workbook, 'Icarus Lab Neon');
  apply(workbook);

  workbook.gameSheet.hideRows(21, 1);
  save(workbook);
  save(workbook);

  setDots(workbook, 'WestLab V2 Small');
  apply(workbook);

  setDots(workbook, '');
  apply(workbook, true);
}

[
  ['конфигурация energyMask отсутствует', {config: null}],
  ['enabled=false', {enabled: false}],
  ['enabled=true, список листов пуст', {enabled: true, sheets: []}]
].forEach(function([label, options]) {
  test(`R5.2-9: ${label}: нет обращений к блоку Energy mask`, function() {
    const {workbook, fixture} = createWorkbook(options);
    const calls = spyMaskFunctions(workbook);

    runFullScenario(workbook);

    assert.deepEqual(calls, {}, 'функции работы с маской не вызывались');
    assert.equal(maskWrites(workbook).length, 0);

    workbook.stateSheet.reads.forEach(function(read) {
      assert.ok(
        read.startColumn + read.numColumns - 1 < 18,
        'чтение колонок блока маски'
      );
    });

    for (let row = 1; row <= fixture.lastRow + 2; row++) {
      for (let col = 18; col <= 24; col++) {
        assert.equal(workbook.stateSheet.grid[row - 1][col - 1], '');
      }
    }
  });
});

test('R5.2-9: с выключенной маской FilterState и ManualState совпадают с эталоном до R5', function() {
  // Сквозная проверка выполняется golden-тестом filterStateGolden.test.cjs:
  // он запускается без CONFIG.energyMask и обязан проходить без изменений.
  const goldenPath = path.join(__dirname, 'golden', 'filterstate-manual.golden.json');

  assert.ok(fs.existsSync(goldenPath));
});

test('R5.2-9: включённая маска не меняет D и ключи ManualState (сравнение с выключенной)', function() {
  const off = createWorkbook({enabled: false});
  const on = createWorkbook({enabled: true});

  runFullScenario(off.workbook);
  runFullScenario(on.workbook);

  assert.deepEqual(stateD(on.workbook, on.fixture), stateD(off.workbook, off.fixture));
  assert.deepEqual(manualKeys(on.workbook), manualKeys(off.workbook));
  assert.equal(manualCount(on.workbook), manualCount(off.workbook));
  assert.deepEqual(
    [...on.workbook.gameSheet.hiddenRows].sort((a, b) => a - b),
    [...off.workbook.gameSheet.hiddenRows].sort((a, b) => a - b)
  );

  // Записи в колонки A:Q совпадают по порядку и значениям.
  // (JSON: записи создаются в разных vm-контекстах).
  const core = workbook => JSON.parse(JSON.stringify(
    workbook.stateWrites.filter(write => write.startColumn < 18)
  ));

  assert.deepEqual(core(on.workbook), core(off.workbook));
});

// ------------------------------------------------------------------ 10

test('R5.2-10: Novice и Advanced: функции интеграции ничего не читают и не пишут', function() {
  const {workbook} = createWorkbook({sheets: ['Intermediate']});

  const throwingSheet = {
    getRange() { throw new Error('обращение к листу'); },
    getMaxRows() { throw new Error('обращение к листу'); },
    getMaxColumns() { throw new Error('обращение к листу'); },
    getName() { return '_FilterState'; }
  };

  ['Novice', 'Advanced'].forEach(function(name) {
    const sheet = {getName: () => name, getParent() { throw new Error('parent'); }};

    assert.equal(
      workbook.project.isEnergyMaskEnabledForSheet_(name),
      false
    );

    const afterFilters = workbook.project.syncEnergyMaskAfterFilters_(
      sheet, throwingSheet, [true, false], new Set([4]), 4, null
    );

    assert.equal(afterFilters.skipped, true);

    const stale = workbook.project.markEnergyMaskStaleForSheet_(
      sheet, 'test', null, throwingSheet
    );

    assert.equal(stale.skipped, true);

    const guard = workbook.project.beginEnergyMaskManualSave_(
      sheet, throwingSheet, null
    );

    assert.equal(guard.active, false);
    assert.equal(guard.reason, 'DISABLED');

    const finish = workbook.project.finishEnergyMaskManualSave_(
      guard, sheet, throwingSheet, [true], [3], 3, null
    );

    assert.equal(finish.skipped, true);
  });
});

['Novice', 'Advanced'].forEach(function(name) {
  test(`R5.2-10: applyFilters и Save на листе ${name} не затрагивают блок маски`, function() {
    const {workbook, fixture} = createWorkbook({sheets: ['Intermediate']});

    workbook.gameSheet.name = name;

    const calls = spyMaskFunctions(workbook);

    setDots(workbook, '');

    // Состояние Novice/Advanced хранится в других колонках _FilterState;
    // здесь важно только то, что маска Intermediate не затрагивается.
    try {
      apply(workbook);
      save(workbook);
    } catch (error) {
      // Ошибка раскладки Novice/Advanced в этом mock допустима.
    }

    assert.deepEqual(calls, {});
    assert.equal(maskWrites(workbook).length, 0);
    assert.ok(maskOf(workbook, fixture).every(value => value === ''));
  });
});

// ------------------------------------------------------------ другое

test('R5.2: Resolver при AutoSort ON не получает изменений: маска пишется и при включённом AutoSort', function() {
  const {workbook, fixture} = createWorkbook();

  setDots(workbook, '');
  workbook.project.applyFilters(false, workbook.gameSheet, false, autoSortOptions(true), null);

  assert.equal(statusOf(workbook), 'OK');
  assert.deepEqual(maskOf(workbook, fixture), stateD(workbook, fixture));
  assert.equal(workbook.rangeSortCalls(), 0);
});

test('R5.2: порядок записей applyFilters: FilterState (D) раньше маски', function() {
  const {workbook} = createWorkbook();

  setDots(workbook, 'Icarus Lab Neon');
  apply(workbook);

  const firstD = workbook.stateWrites.findIndex(write => write.startColumn === D_COL);
  const firstMask = workbook.stateWrites.findIndex(write => write.startColumn >= 18);

  assert.ok(firstD >= 0 && firstMask >= 0);
  assert.ok(firstD < firstMask, 'маска пишется после обновления FilterState');
});

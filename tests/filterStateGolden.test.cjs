'use strict';

// Golden-тест инварианта I-1: столбец _FilterState!D и ключи ManualState
// записываются так же, как до внедрения Energy mask (R5).
//
// Сценарии прогоняются настоящим кодом applyFilters и saveManualHiddenRows
// на данных листа Intermediate (tests/fixtures/energy-intermediate.json).
// Эталон зафиксирован по состоянию кода до R5 (c9113ea). Он намеренно
// сохраняет текущее поведение: у сохранённой ручной строки D остаётся TRUE.
// Это поведение Energy mask не меняет, маска пишется в отдельную колонку.
//
// Обновление эталона (только осознанно, при изменении семантики D):
//   WRITE_GOLDEN=1 node --test tests/filterStateGolden.test.cjs

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');

const {
  createMockWorkbook,
  autoSortOptions
} = require('./helpers/mockWorkbook.cjs');
const {loadEnergyFixture} = require('./helpers/energyModel.cjs');

const goldenPath = path.join(__dirname, 'golden', 'filterstate-manual.golden.json');
const FIRST = 3;

function isNumber(value) {
  return typeof value === 'number' && Number.isFinite(value);
}

function freshWorkbook(fixture) {
  const workbook = createMockWorkbook({
    rows: fixture.rows.map(row => row.slice()),
    hiddenRows: [],
    manualCount: 0,
    manualKeys: [],
    gameLastRow: fixture.lastRow
  });

  workbook.stateSheet.grid[1][4] = fixture.lastRow;

  return workbook;
}

function setDots(workbook, dots) {
  workbook.gameSheet.grid[0][8] = dots || '';
  workbook.gameSheet.grid[0][9] = '';
}

function apply(workbook, autoSort, reset = false) {
  return workbook.project.applyFilters(
    false,
    workbook.gameSheet,
    reset,
    autoSortOptions(autoSort),
    null
  );
}

function snapshot(label, workbook, fixture) {
  const grid = workbook.stateSheet.grid;
  const count = Number(grid[1][13]) || 0;
  const keys = [];

  for (let i = 0; i < count; i++) {
    keys.push(grid[FIRST - 1 + i][12]);
  }

  const visibility = fixture.rows
    .map(function(_, index) {
      return grid[FIRST - 1 + index][3] === true ? 'T' : 'F';
    })
    .join('');

  const hidden = [...workbook.gameSheet.hiddenRows]
    .filter(row => row >= FIRST && row <= fixture.lastRow)
    .sort((a, b) => a - b);

  return {
    step: label,
    filterStateD: visibility,
    filterStateLastRow: grid[1][4],
    manualStateKeys: keys,
    manualStateCount: grid[1][13],
    physicallyHiddenRows: hidden
  };
}

function runScenarios() {
  const fixture = loadEnergyFixture();

  const sameDotsRow = fixture.rows.findIndex(function(row) {
    return String(row[6]).trim() === 'Icarus Lab Neon' && isNumber(row[2]);
  }) + FIRST;

  const otherDotsRow = 4; // VT Pasu / PureAim P-Dots / 750.08.

  const result = {};

  [false, true].forEach(function(autoSort) {
    const suffix = `autoSort=${autoSort}`;

    // S1: Dots выбран первым, затем скрывается и сохраняется строка того же Dots.
    let w = freshWorkbook(fixture);
    const s1 = [];

    setDots(w, 'Icarus Lab Neon');
    apply(w, autoSort);
    s1.push(snapshot('apply Dots=Neon', w, fixture));

    w.gameSheet.hideRows(sameDotsRow, 1);
    w.project.saveManualHiddenRows();
    s1.push(snapshot('hide+Save', w, fixture));

    apply(w, autoSort);
    s1.push(snapshot('re-apply', w, fixture));

    setDots(w, 'WestLab V2 Small');
    apply(w, autoSort);
    s1.push(snapshot('Dots=WestLab', w, fixture));

    setDots(w, '');
    apply(w, autoSort, true);
    s1.push(snapshot('resetFilters', w, fixture));
    result[`S1 ${suffix}`] = s1;

    // S2: скрыта строка другого Dots при "All Results", затем выбран Dots.
    w = freshWorkbook(fixture);
    const s2 = [];

    setDots(w, '');
    apply(w, autoSort);
    w.gameSheet.hideRows(otherDotsRow, 1);
    w.project.saveManualHiddenRows();
    s2.push(snapshot('hide+Save', w, fixture));

    setDots(w, 'Icarus Lab Neon');
    apply(w, autoSort);
    s2.push(snapshot('Dots=Neon', w, fixture));

    setDots(w, 'WestLab V2 Small');
    apply(w, autoSort);
    s2.push(snapshot('Dots=WestLab', w, fixture));

    setDots(w, '');
    apply(w, autoSort, true);
    s2.push(snapshot('resetFilters', w, fixture));
    result[`S2 ${suffix}`] = s2;

    // S3: Dots выбран первым, строка другого Dots уже скрыта фильтром.
    w = freshWorkbook(fixture);
    const s3 = [];

    setDots(w, 'Icarus Lab Neon');
    apply(w, autoSort);
    w.project.saveManualHiddenRows();
    s3.push(snapshot('Dots=Neon + Save without manual rows', w, fixture));
    result[`S3 ${suffix}`] = s3;

    // S4 (T9): повторный Save не уменьшает число ключей ManualState.
    w = freshWorkbook(fixture);
    const s4 = [];

    setDots(w, '');
    apply(w, autoSort);
    w.gameSheet.hideRows(otherDotsRow, 1);

    for (let i = 1; i <= 3; i++) {
      w.project.saveManualHiddenRows();
      s4.push(snapshot(`Save #${i}`, w, fixture));
    }

    apply(w, autoSort);
    s4.push(snapshot('apply after repeated Save', w, fixture));
    result[`S4 ${suffix}`] = s4;
  });

  return JSON.parse(JSON.stringify(result));
}

test('I-1: _FilterState!D и ключи ManualState совпадают с эталоном до R5', function() {
  const actual = runScenarios();

  if (process.env.WRITE_GOLDEN === '1') {
    fs.writeFileSync(goldenPath, JSON.stringify(actual, null, 1) + '\n', 'utf8');

    return;
  }

  const golden = JSON.parse(fs.readFileSync(goldenPath, 'utf8'));

  assert.deepEqual(actual, golden);
});

test('T9: повторный Save не уменьшает число ключей ManualState', function() {
  const actual = runScenarios();

  ['autoSort=false', 'autoSort=true'].forEach(function(suffix) {
    const steps = actual[`S4 ${suffix}`];

    steps.forEach(function(step) {
      assert.equal(step.manualStateCount, 1, `${suffix}: ${step.step}`);
      assert.equal(step.manualStateKeys.length, 1);
      assert.deepEqual(step.physicallyHiddenRows, [4]);
    });
  });
});

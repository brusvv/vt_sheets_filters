'use strict';

// Модель формулы Energy (Intermediate!M3 и аналогичных ячеек групп).
//
// Формулы Google Sheets в Node не выполняются, поэтому логика формулы
// воспроизведена здесь и проверяется на кешированных значениях таблицы
// (тест T0). Модель описывает формулу, а не код скрипта:
//   1. берутся строки, где Score - число и строка учитывается (isCounted);
//   2. для каждого сценария группы выбирается лучший Score;
//   3. по порогам строки (N, Q, T, W, Z) и энергиям строки 2 линейной
//      интерполяцией считается Energy сценария;
//   4. Energy группы - максимум по сценариям группы.
// Общая Energy M2 = TRUNC(HARMEAN(групп)).

const fs = require('node:fs');
const path = require('node:path');

function loadEnergyFixture() {
  return JSON.parse(
    fs.readFileSync(
      path.join(__dirname, '..', 'fixtures', 'energy-intermediate.json'),
      'utf8'
    )
  );
}

function isNumber(value) {
  return typeof value === 'number' && Number.isFinite(value);
}

function scenarioEnergy(score, thresholds, tierEnergy) {
  const [t1, t2, t3, t4, t5] = thresholds;
  const marks = [0, t1, t2, t3, t4, t5];

  let tier = 0;

  for (let i = 0; i < marks.length; i++) {
    if (score >= marks[i]) {
      tier = i + 1;
    }
  }

  const e = [
    tierEnergy.N2,
    tierEnergy.Q2,
    tierEnergy.T2,
    tierEnergy.W2,
    tierEnergy.Z2
  ];

  const base = [0, e[0], e[1], e[2], e[3], e[4]][tier - 1];
  const lower = marks[tier - 1];

  const width = [
    t1,
    t2 - t1,
    t3 - t2,
    t4 - t3,
    t5 - t4,
    t5 - t4
  ][tier - 1];

  const step = [
    e[0],
    e[1] - e[0],
    e[2] - e[1],
    e[3] - e[2],
    e[4] - e[3],
    e[4] - e[3]
  ][tier - 1];

  return base + ((score - lower) / width) * step;
}

/**
 * Energy групп.
 *
 * @param {Object} fixture
 * @param {function(number): boolean} isCounted Учитывается ли физическая
 *   строка (номер строки листа).
 * @return {Object<number, ?{energy: number, scenario: string, row: number,
 *   score: number}>} Ключ - первая строка группы.
 */
function groupEnergies(fixture, isCounted) {
  const result = {};

  fixture.groups.forEach(function([startRow, endRow]) {
    const best = new Map();

    for (
      let row = startRow;
      row <= Math.min(endRow, fixture.lastRow);
      row++
    ) {
      const index = row - fixture.firstRow;
      const scenario = String(fixture.rows[index][0] || '').trim();
      const score = fixture.rows[index][2];

      if (!isNumber(score) || !isCounted(row)) {
        continue;
      }

      if (!best.has(scenario) || score > best.get(scenario).score) {
        best.set(scenario, {
          score: score,
          row: row,
          thresholds: fixture.thresholds[index]
        });
      }
    }

    let top = null;

    best.forEach(function(entry, scenario) {
      const energy = scenarioEnergy(
        entry.score,
        entry.thresholds,
        fixture.tierEnergy
      );

      if (top === null || energy > top.energy) {
        top = {
          energy: energy,
          scenario: scenario,
          row: entry.row,
          score: entry.score
        };
      }
    });

    result[startRow] = top;
  });

  return result;
}

function overallEnergy(groups) {
  const values = Object.keys(groups)
    .map(function(key) {
      return groups[key] === null ? null : groups[key].energy;
    })
    .filter(function(value) {
      return value !== null && value > 0;
    });

  if (values.length === 0) {
    return null;
  }

  const harmonic = values.length /
    values.reduce(function(sum, value) {
      return sum + 1 / value;
    }, 0);

  return Math.trunc(harmonic);
}

module.exports = {
  loadEnergyFixture,
  scenarioEnergy,
  groupEnergies,
  overallEnergy
};

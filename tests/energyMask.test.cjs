'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const {
  loadEnergyFixture,
  groupEnergies,
  overallEnergy
} = require('./helpers/energyModel.cjs');

const root = path.resolve(__dirname, '..');

// ---------------------------------------------------------------- helpers

function loadProject(options = {}) {
  const context = vm.createContext({});

  ['config/config.gs', 'filtering/energyMask.gs'].forEach(function(file) {
    vm.runInContext(
      fs.readFileSync(path.join(root, file), 'utf8'),
      context,
      {filename: file}
    );
  });

  if (options.enabled) {
    vm.runInContext(
      'CONFIG.energyMask.enabled = true;' +
      'CONFIG.energyMask.sheets = ' +
      JSON.stringify(options.sheets || ['Intermediate']) + ';',
      context
    );
  }

  return context;
}

// Лист _FilterState: сетка значений и журнал вызовов.
function createStateSheet(options = {}) {
  const rows = options.rows || 40;
  const columns = options.columns || 26;

  const grid = Array.from({length: rows}, function() {
    return Array(columns).fill('');
  });

  const calls = [];

  function range(row, column, numRows = 1, numColumns = 1) {
    return {
      getValues() {
        calls.push({op: 'getValues', row, column, numRows, numColumns});

        return grid
          .slice(row - 1, row - 1 + numRows)
          .map(function(line) {
            return line.slice(column - 1, column - 1 + numColumns);
          });
      },
      getValue() {
        return this.getValues()[0][0];
      },
      setValues(values) {
        calls.push({op: 'setValues', row, column, numRows, numColumns, values});

        assert.equal(values.length, numRows);

        values.forEach(function(line, r) {
          assert.equal(line.length, numColumns);

          line.forEach(function(value, c) {
            grid[row - 1 + r][column - 1 + c] = value;
          });
        });

        return this;
      },
      setValue(value) {
        calls.push({op: 'setValue', row, column, value});
        grid[row - 1][column - 1] = value;

        return this;
      },
      clearContent() {
        calls.push({op: 'clearContent', row, column, numRows, numColumns});

        for (let r = 0; r < numRows; r++) {
          for (let c = 0; c < numColumns; c++) {
            grid[row - 1 + r][column - 1 + c] = '';
          }
        }

        return this;
      }
    };
  }

  return {
    grid,
    calls,
    callsOf(op) {
      return calls.filter(function(call) {
        return call.op === op;
      });
    },
    getName() {
      return '_FilterState';
    },
    getMaxRows() {
      return grid.length;
    },
    getMaxColumns() {
      return grid[0].length;
    },
    insertColumnsAfter(after, count) {
      calls.push({op: 'insertColumnsAfter', after, count});

      grid.forEach(function(line) {
        for (let i = 0; i < count; i++) {
          line.push('');
        }
      });
    },
    getRange: range
  };
}

function gameSheet(name) {
  return {getName: () => name};
}

function plain(value) {
  return JSON.parse(JSON.stringify(value));
}

function maskColumn(sheet, column, fromRow, toRow) {
  const values = [];

  for (let row = fromRow; row <= toRow; row++) {
    values.push(sheet.grid[row - 1][column - 1]);
  }

  return values;
}

const INTERMEDIATE_COLUMN = 20; // T

// ------------------------------------------------------------------- T0

test('T0: модель формулы Energy воспроизводит кешированные значения таблицы', function() {
  const fixture = loadEnergyFixture();
  const groups = groupEnergies(fixture, () => true);

  fixture.groups.forEach(function([startRow]) {
    assert.ok(
      Math.abs(groups[startRow].energy - fixture.cachedGroupEnergy[startRow]) < 1e-6,
      `группа M${startRow}`
    );
  });

  assert.equal(overallEnergy(groups), fixture.cachedM2);
  assert.equal(fixture.cachedM2, 630);
});

test('T0: строки, определяющие максимум группы, и Energy без них', function() {
  const fixture = loadEnergyFixture();
  const all = groupEnergies(fixture, () => true);

  // [первая строка группы]: [строка максимума, Energy группы без неё].
  const expected = {
    3: [17, 688.89],
    19: [28, 633.33],
    45: [54, 660.0],
    55: [63, 726.18],
    69: [88, 611.14],
    89: [98, 558.82],
    99: [103, 511.43],
    112: [117, 552.4],
    125: [130, 542.5]
  };

  Object.keys(expected).forEach(function(group) {
    const [topRow, without] = expected[group];

    assert.equal(all[group].row, topRow, `строка максимума группы M${group}`);

    const reduced = groupEnergies(fixture, row => row !== topRow)[group];

    assert.ok(
      Math.abs(reduced.energy - without) < 0.01,
      `группа M${group}: ${reduced.energy} вместо ${without}`
    );
  });
});

// ------------------------------------------------------- конфигурация

test('конфигурация: по умолчанию выключено, блок не пересекается с A:Q', function() {
  const context = loadProject();
  const settings = plain(
    JSON.parse(vm.runInContext('JSON.stringify(CONFIG.energyMask)', context))
  );

  assert.equal(settings.enabled, false);
  assert.deepEqual(settings.sheets, []);
  assert.equal(settings.schemaVersion, 1);

  const maskColumns = Object.values(settings.columns);

  assert.equal(new Set(maskColumns).size, maskColumns.length);

  // FilterState занимает A:H, ManualState J:Q.
  maskColumns.forEach(function(column) {
    assert.ok(column > 17, `колонка ${column} пересекается с A:Q`);
  });

  assert.ok(settings.versionCol > 17);
  assert.ok(settings.versionCol < Math.min(...maskColumns));

  assert.equal(
    settings.statusRow + 1,
    vm.runInContext('CONFIG.firstRow', context)
  );
});

// -------------------------------------------------------- computeEnergyMask_

test('computeEnergyMask_: ручные строки определяются по физической позиции', function() {
  const context = loadProject();

  const mask = context.computeEnergyMask_(
    [true, true, false, true, true],
    new Set([4, 5]),
    3
  );

  // Строки 3..7: 4 и 5 сохранены вручную (строка 5 при этом уже скрыта
  // фильтром), строки 3, 6 и 7 учитываются.
  assert.deepEqual(Array.from(mask), [true, false, false, true, true]);

  const manualOnlyFilterHidden = context.computeEnergyMask_(
    [false, true],
    [3],
    3
  );

  assert.deepEqual(Array.from(manualOnlyFilterHidden), [false, true]);
});

test('computeEnergyMask_: нестрогие значения не считаются видимыми', function() {
  const context = loadProject();

  const mask = context.computeEnergyMask_(
    [true, 'TRUE', 1, null, undefined],
    null,
    3
  );

  assert.deepEqual(Array.from(mask), [true, false, false, false, false]);
});

test('computeEnergyMask_: проверка аргументов', function() {
  const context = loadProject();

  assert.throws(() => context.computeEnergyMask_('x', null, 3), /массивом/);
  assert.throws(() => context.computeEnergyMask_([], null, 0), /firstRow/);
});

// ------------------------------------------------------------------ T12

test('T12: блок маски отсутствует: создаётся, FilterState и ManualState не затрагиваются', function() {
  const context = loadProject({enabled: true});
  const sheet = createStateSheet({columns: 17});

  // Значения столбцов A:Q (FilterState и ManualState) до записи.
  for (let row = 1; row <= 12; row++) {
    for (let column = 1; column <= 17; column++) {
      sheet.grid[row - 1][column - 1] = `r${row}c${column}`;
    }
  }

  const before = plain(sheet.grid.map(line => line.slice(0, 17)));

  const result = plain(
    context.writeEnergyMask_(
      sheet,
      gameSheet('Intermediate'),
      [true, false, true, true],
      6
    )
  );

  assert.equal(result.ok, true);
  assert.equal(result.layoutCreated, true);
  assert.equal(result.rowsWritten, 4);

  // Добавлены недостающие колонки до U.
  assert.deepEqual(
    plain(sheet.callsOf('insertColumnsAfter')),
    [{op: 'insertColumnsAfter', after: 17, count: 4}]
  );

  assert.equal(sheet.grid[0][17], 'EnergyMask:1');                 // R1
  assert.equal(sheet.grid[0][INTERMEDIATE_COLUMN - 1], 'Intermediate Energy Mask');
  assert.equal(sheet.grid[1][INTERMEDIATE_COLUMN - 1], 'OK');
  assert.deepEqual(
    maskColumn(sheet, INTERMEDIATE_COLUMN, 3, 6),
    [true, false, true, true]
  );

  assert.deepEqual(
    plain(sheet.grid.map(line => line.slice(0, 17))),
    before,
    'столбцы A:Q не должны меняться'
  );
});

test('T12: достаточная ширина листа: колонки не добавляются', function() {
  const context = loadProject({enabled: true});
  const sheet = createStateSheet({columns: 26});

  context.writeEnergyMask_(sheet, 'Intermediate', [true], 3);

  assert.equal(sheet.callsOf('insertColumnsAfter').length, 0);
});

// ------------------------------------------------------------------ T13

test('T13: небулевые значения маски полностью перезаписываются', function() {
  const context = loadProject({enabled: true});
  const sheet = createStateSheet();

  sheet.grid[1][INTERMEDIATE_COLUMN - 1] = 'garbage';

  ['x', 1, '', null, 'TRUE', 0, undefined, 'FALSE', 2, 'y']
    .forEach(function(value, index) {
      sheet.grid[2 + index][INTERMEDIATE_COLUMN - 1] = value;
    });

  const before = plain(context.readEnergyMask_(sheet, 'Intermediate', 12));

  assert.equal(before.valid, false);

  const mask = [true, true, false, true, false, false, true, true, false, true];

  const result = plain(
    context.writeEnergyMask_(sheet, 'Intermediate', mask, 12)
  );

  assert.equal(result.ok, true);

  const after = plain(context.readEnergyMask_(sheet, 'Intermediate', 12));

  assert.equal(after.valid, true);
  assert.deepEqual(after.values, mask);
  assert.equal(
    plain(context.readEnergyMaskStatus_(sheet, 'Intermediate')).ok,
    true
  );
});

test('T13: запись небулевой маски и неверной длины отклоняется до обращения к листу', function() {
  const context = loadProject({enabled: true});
  const sheet = createStateSheet();

  assert.throws(
    () => context.writeEnergyMask_(sheet, 'Intermediate', [true, 'x'], 4),
    /boolean/
  );

  assert.throws(
    () => context.writeEnergyMask_(sheet, 'Intermediate', [true], 5),
    /Длина маски/
  );

  assert.equal(sheet.calls.length, 0);
});

// ------------------------------------------------------------------ T14

test('T14: уменьшение LastRow очищает значения ниже', function() {
  const context = loadProject({enabled: true});
  const sheet = createStateSheet({rows: 40});

  context.writeEnergyMask_(
    sheet,
    'Intermediate',
    Array(10).fill(true),
    12
  );

  assert.deepEqual(
    maskColumn(sheet, INTERMEDIATE_COLUMN, 3, 12),
    Array(10).fill(true)
  );

  const result = plain(
    context.writeEnergyMask_(
      sheet,
      'Intermediate',
      Array(6).fill(false),
      8
    )
  );

  assert.equal(result.ok, true);
  assert.equal(result.rowsWritten, 6);

  // Ниже последней строки данных (строки 9..40) всё очищено.
  assert.equal(result.rowsCleared, 40 - 9 + 1);
  assert.deepEqual(
    maskColumn(sheet, INTERMEDIATE_COLUMN, 9, 40),
    Array(32).fill('')
  );
  assert.deepEqual(
    maskColumn(sheet, INTERMEDIATE_COLUMN, 3, 8),
    Array(6).fill(false)
  );
});

test('T14: нет строк данных: пишется только статус, данные ниже очищаются', function() {
  const context = loadProject({enabled: true});
  const sheet = createStateSheet({rows: 20});

  context.writeEnergyMask_(sheet, 'Intermediate', [true, true], 4);

  const result = plain(
    context.writeEnergyMask_(sheet, 'Intermediate', [], 2)
  );

  assert.equal(result.ok, true);
  assert.equal(result.rowsWritten, 0);
  assert.equal(sheet.grid[1][INTERMEDIATE_COLUMN - 1], 'OK');
  assert.deepEqual(
    maskColumn(sheet, INTERMEDIATE_COLUMN, 3, 20),
    Array(18).fill('')
  );
});

// ------------------------------------------------------------------ T17

test('T17: по умолчанию маска выключена, обращений к листам нет', function() {
  const context = loadProject();
  const sheet = createStateSheet();

  assert.equal(context.isEnergyMaskEnabledForSheet_('Intermediate'), false);

  const result = plain(
    context.syncEnergyMaskIfEnabled_(
      sheet,
      gameSheet('Intermediate'),
      [true, false],
      new Set([4]),
      4
    )
  );

  assert.deepEqual(result, {skipped: true, reason: 'DISABLED', ok: true});
  assert.equal(sheet.calls.length, 0);
});

test('T17: включено, но лист не в списке: ничего не пишется', function() {
  const sheet = createStateSheet();

  const emptyList = loadProject({enabled: true, sheets: []});

  assert.equal(emptyList.isEnergyMaskEnabledForSheet_('Intermediate'), false);

  const otherSheet = loadProject({enabled: true, sheets: ['Advanced']});

  assert.equal(otherSheet.isEnergyMaskEnabledForSheet_('Intermediate'), false);
  assert.equal(otherSheet.isEnergyMaskEnabledForSheet_('Advanced'), true);

  const result = plain(
    otherSheet.syncEnergyMaskIfEnabled_(
      sheet,
      gameSheet('Intermediate'),
      [true],
      null,
      3
    )
  );

  assert.equal(result.skipped, true);
  assert.equal(sheet.calls.length, 0);
});

test('T17: включено для листа: sync вычисляет и пишет маску', function() {
  const context = loadProject({enabled: true});
  const sheet = createStateSheet();

  const result = plain(
    context.syncEnergyMaskIfEnabled_(
      sheet,
      gameSheet('Intermediate'),
      [true, true, false, true],
      new Set([4]),
      6
    )
  );

  assert.equal(result.skipped, false);
  assert.equal(result.ok, true);
  assert.deepEqual(
    maskColumn(sheet, INTERMEDIATE_COLUMN, 3, 6),
    [true, false, false, true]
  );
});

// ------------------------------------------------------------------ TS1

test('TS1: маска и статус пишутся одной операцией setValues', function() {
  const context = loadProject({enabled: true});
  const sheet = createStateSheet();

  context.writeEnergyMask_(sheet, 'Intermediate', [true, false, true], 5);

  // Первая запись: создание блока (R1 и заголовок) и одна запись данных.
  const writes = sheet.callsOf('setValues');

  assert.equal(writes.length, 1);
  assert.equal(writes[0].row, 2);
  assert.equal(writes[0].column, INTERMEDIATE_COLUMN);
  assert.deepEqual(
    plain(writes[0].values),
    [['OK'], [true], [false], [true]]
  );

  // Блок уже создан: повторная запись не делает отдельных setValue.
  const callsBefore = sheet.calls.length;

  context.writeEnergyMask_(sheet, 'Intermediate', [false, false, false], 5);

  const secondCalls = sheet.calls.slice(callsBefore);

  assert.equal(
    secondCalls.filter(call => call.op === 'setValue').length,
    0
  );
  assert.equal(
    secondCalls.filter(call => call.op === 'setValues').length,
    1
  );
});

// ------------------------------------------------------------------ TS2

test('TS2: markEnergyMaskStale_ меняет только ячейку статуса', function() {
  const context = loadProject({enabled: true});
  const sheet = createStateSheet();

  context.writeEnergyMask_(sheet, 'Intermediate', [true, false, true], 5);

  const maskBefore = maskColumn(sheet, INTERMEDIATE_COLUMN, 3, 5);
  const callsBefore = sheet.calls.length;

  const result = plain(context.markEnergyMaskStale_(sheet, 'Intermediate'));

  assert.deepEqual(result, {ok: true, reason: null});

  const writes = sheet.calls
    .slice(callsBefore)
    .filter(call => call.op === 'setValue' || call.op === 'setValues' || call.op === 'clearContent');

  assert.deepEqual(
    plain(writes),
    [{op: 'setValue', row: 2, column: INTERMEDIATE_COLUMN, value: 'STALE'}]
  );

  assert.deepEqual(maskColumn(sheet, INTERMEDIATE_COLUMN, 3, 5), maskBefore);

  const status = plain(context.readEnergyMaskStatus_(sheet, 'Intermediate'));

  assert.equal(status.stale, true);
  assert.equal(status.ok, false);

  // Успешная запись возвращает статус OK.
  context.writeEnergyMask_(sheet, 'Intermediate', [true, true, true], 5);

  assert.equal(
    plain(context.readEnergyMaskStatus_(sheet, 'Intermediate')).ok,
    true
  );
});

test('TS2: markEnergyMaskStale_ не бросает исключений при сбое листа', function() {
  const context = loadProject({enabled: true});
  const broken = {
    getMaxColumns() {
      throw new Error('boom');
    }
  };

  const result = plain(context.markEnergyMaskStale_(broken, 'Intermediate'));

  assert.equal(result.ok, false);
  assert.equal(result.reason, 'ERROR');
  assert.equal(result.message, 'boom');
});

// ------------------------------------------------------------------ TS3

test('TS3: неизвестная версия схемы: ничего не записывается', function() {
  const context = loadProject({enabled: true});
  const sheet = createStateSheet();

  sheet.grid[0][17] = 'EnergyMask:2';

  const layout = plain(
    context.ensureEnergyMaskLayout_(sheet, 'Intermediate')
  );

  assert.deepEqual(layout, {
    ok: false,
    created: false,
    reason: 'SCHEMA_VERSION_MISMATCH'
  });

  const write = plain(
    context.writeEnergyMask_(sheet, 'Intermediate', [true], 3)
  );

  assert.equal(write.ok, false);
  assert.equal(write.reason, 'SCHEMA_VERSION_MISMATCH');
  assert.equal(write.rowsWritten, 0);

  const stale = plain(context.markEnergyMaskStale_(sheet, 'Intermediate'));

  assert.equal(stale.ok, false);
  assert.equal(stale.reason, 'SCHEMA_VERSION_MISMATCH');

  assert.equal(sheet.callsOf('setValues').length, 0);
  assert.equal(sheet.callsOf('setValue').length, 0);
  assert.equal(sheet.callsOf('clearContent').length, 0);
});

test('TS3: чужой заголовок колонки маски: ничего не записывается', function() {
  const context = loadProject({enabled: true});
  const sheet = createStateSheet();

  sheet.grid[0][INTERMEDIATE_COLUMN - 1] = 'Чужие данные';

  const write = plain(
    context.writeEnergyMask_(sheet, 'Intermediate', [true], 3)
  );

  assert.equal(write.ok, false);
  assert.equal(write.reason, 'HEADER_MISMATCH');
  assert.equal(sheet.callsOf('setValues').length, 0);
  assert.equal(sheet.callsOf('setValue').length, 0);
});

test('Блок листа без настройки: понятная ошибка', function() {
  const context = loadProject({enabled: true});

  assert.throws(
    () => context.getEnergyMaskColumn_('Elite'),
    /Нет блока Energy mask для листа "Elite"/
  );
});

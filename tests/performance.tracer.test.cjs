const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

const root = path.resolve(__dirname, '..');

function loadTracer() {
  const sandbox = {
    console: { log() {}, warn() {}, error() {} },
    Utilities: { getUuid: () => '12345678-1234-1234-1234-123456789abc' },
    SpreadsheetApp: { getActiveSpreadsheet: () => null },
    Date,
    JSON,
    Math,
    Number,
    String,
    Object,
    Array
  };
  vm.createContext(sandbox);
  for (const file of ['performance/perf.gs', 'core/lock.gs']) {
    vm.runInContext(
      fs.readFileSync(path.join(root, file), 'utf8'),
      sandbox,
      { filename: file }
    );
  }
  return sandbox;
}

function createLockFixture(options = {}) {
  const context = loadTracer();
  const state = {
    waitCalls: 0,
    releaseCalls: 0,
    callbackCalls: 0
  };

  context.LockService = {
    getScriptLock() {
      return {
        waitLock(timeout) {
          state.waitCalls++;
          assert.equal(timeout, 30000);
          if (options.waitError) throw options.waitError;
        },
        releaseLock() {
          state.releaseCalls++;
          if (options.releaseError) throw options.releaseError;
        }
      };
    }
  };

  // Deliberately replace every entry point with a throwing diagnostic hook.
  // The lock manager itself must still complete its lifecycle.
  const diagnosticError = new Error('injected tracer failure');
  context.triggerPerfStart_ = () => { throw diagnosticError; };
  context.triggerPerfEvent_ = () => { throw diagnosticError; };
  context.triggerPerfStep_ = () => { throw diagnosticError; };
  context.triggerPerfEnd_ = () => { throw diagnosticError; };

  return { context, state, diagnosticError };
}

test('all compatibility wrappers swallow injected tracer failures', () => {
  const context = loadTracer();
  const injected = new Error('tracer method failed');

  context.perfStart_ = () => { throw injected; };
  context.perfEvent_ = () => { throw injected; };
  context.perfStep_ = () => { throw injected; };
  context.perfEnd_ = () => { throw injected; };

  assert.equal(context.triggerPerfStart_('test'), null);
  assert.doesNotThrow(() => context.triggerPerfEvent_({ mode: 'INFO' }, 'event'));
  assert.doesNotThrow(() => context.triggerPerfStep_({ mode: 'DEBUG' }, 'step', Date.now()));
  assert.doesNotThrow(() => context.triggerPerfEnd_({ mode: 'DEBUG' }));
});

test('perfLockWait is fail-safe even when the event wrapper itself throws', () => {
  const context = loadTracer();
  context.triggerPerfEvent_ = () => { throw new Error('wrapper failure'); };

  assert.doesNotThrow(() => context.perfLockWait_({ mode: 'INFO' }, 12));
});

test('tracer wrappers remain callable in OFF, ERROR, INFO and DEBUG modes', () => {
  const context = loadTracer();

  for (const mode of ['OFF', 'ERROR', 'INFO', 'DEBUG']) {
    context.PERF_MODE_OVERRIDE = mode;
    const trace = context.triggerPerfStart_('mode-test');
    if (mode === 'OFF') {
      assert.equal(trace, null);
      continue;
    }

    assert.equal(trace.mode, mode);
    assert.doesNotThrow(() => context.triggerPerfStep_(trace, 'step', Date.now()));
    assert.doesNotThrow(() => context.triggerPerfEvent_(trace, 'event'));
    assert.doesNotThrow(() => context.triggerPerfEnd_(trace));
  }
});

test('tracer failures cannot prevent lock acquisition, callback result or release', () => {
  const fixture = createLockFixture();
  const expectedResult = { ok: true, value: 42 };

  const result = fixture.context.withProjectLock(() => {
    fixture.state.callbackCalls++;
    return expectedResult;
  });

  assert.equal(result, expectedResult);
  assert.equal(fixture.state.waitCalls, 1);
  assert.equal(fixture.state.callbackCalls, 1);
  assert.equal(fixture.state.releaseCalls, 1);
});

test('tracer failures cannot replace the exact callback exception or prevent release', () => {
  const fixture = createLockFixture();
  const callbackError = new Error('business callback failed');

  assert.throws(
    () => fixture.context.withProjectLock(() => {
      fixture.state.callbackCalls++;
      throw callbackError;
    }),
    error => error === callbackError
  );

  assert.equal(fixture.state.callbackCalls, 1);
  assert.equal(fixture.state.releaseCalls, 1);
});

test('waitLock business error is preserved and callback/release are skipped', () => {
  const waitError = new Error('waitLock failed');
  const fixture = createLockFixture({ waitError });

  assert.throws(
    () => fixture.context.withProjectLock(() => {
      fixture.state.callbackCalls++;
    }),
    error => error === waitError
  );

  assert.equal(fixture.state.waitCalls, 1);
  assert.equal(fixture.state.callbackCalls, 0);
  assert.equal(fixture.state.releaseCalls, 0);
});

test('releaseLock failure is preserved when callback succeeds', () => {
  const releaseError = new Error('releaseLock failed');
  const fixture = createLockFixture({ releaseError });

  assert.throws(
    () => fixture.context.withProjectLock(() => {
      fixture.state.callbackCalls++;
      return 'ok';
    }),
    error => error === releaseError
  );

  assert.equal(fixture.state.callbackCalls, 1);
  assert.equal(fixture.state.releaseCalls, 1);
});

test('callback exception takes precedence over a subsequent releaseLock exception', () => {
  const callbackError = new Error('callback failed first');
  const releaseError = new Error('release failed second');
  const fixture = createLockFixture({ releaseError });

  assert.throws(
    () => fixture.context.withProjectLock(() => {
      fixture.state.callbackCalls++;
      throw callbackError;
    }),
    error => error === callbackError
  );

  assert.equal(fixture.state.releaseCalls, 1);
});

test('planner emits detailed stage measurements without changing the plan', () => {
  const steps = [];
  const context = {
    CONFIG: {
      sortStartCol: 3,
      sortWidth: 10,
      scoreCol: 5,
      dotsCol: 9
    },
    triggerPerfStep_(trace, name, startedAt, metadata) {
      steps.push({ name, startedAt, metadata });
    }
  };
  vm.createContext(context);
  const plannerFile = 'sorting/visibleSortPlan.gs';
  vm.runInContext(
    fs.readFileSync(path.join(root, plannerFile), 'utf8'),
    context,
    { filename: plannerFile }
  );

  function row(scenario, score, dots) {
    const result = Array(10).fill('');
    result[0] = scenario;
    result[2] = score;
    result[6] = dots;
    return result;
  }

  const data = [
    row('A', 10, 'A'),
    row('A', 30, 'C'),
    row('A', 20, 'B'),
    row('B', 5, 'A'),
    row('B', 15, 'B')
  ];
  const mask = {
    ok: true,
    firstRow: 3,
    lastRow: 7,
    visibleByPhysicalPosition: [true, false, true, true, true]
  };

  const withoutTrace = context.buildVisibleScenarioSortPlan_(
    data, 3, 7, mask, 'scoreDesc'
  );
  const withTrace = context.buildVisibleScenarioSortPlan_(
    data, 3, 7, mask, 'scoreDesc', { mode: 'DEBUG' }
  );

  assert.deepEqual(
    JSON.parse(JSON.stringify(withTrace)),
    JSON.parse(JSON.stringify(withoutTrace))
  );
  const stepNames = steps.map(step => step.name);
  for (const expected of [
    'Sort planner: Validate input and visibility mask',
    'Sort planner: Collect visible rows',
    'Sort planner: Compute row order',
    'Sort planner: Build changed block',
    'Sort planner: Build scenario blocks',
    'Sort planner: Validate plan'
  ]) {
    assert.ok(stepNames.includes(expected), 'missing planner step: ' + expected);
  }
  assert.equal(stepNames.filter(name => name === 'Sort planner: Collect visible rows').length, 2);
  assert.equal(stepNames.filter(name => name === 'Sort planner: Compute row order').length, 2);
});

test('manual sort keeps its last-row timer before getLastRow and retains guarded transfer status', () => {
  const source = fs.readFileSync(
    path.join(root, 'sorting/sortOperations.gs'),
    'utf8'
  );
  const timerAt = source.indexOf("const lastRowStarted = Date.now();");
  const readAt = source.indexOf("const sheetLastRow = sheet.getLastRow();");
  assert.ok(timerAt >= 0 && readAt > timerAt);
  assert.match(source, /'Get last row',\s*lastRowStarted/);

  const sortingSource = fs.readFileSync(
    path.join(root, 'sorting/sorting.gs'),
    'utf8'
  );
  assert.match(sortingSource, /status:\s*'TRANSFER_NOT_VERIFIED'/);
  assert.match(sortingSource, /applied:\s*false/);
});


test('Rank Scroller measures hide/show calls without changing the allow-list or call order', () => {
  const steps = [];
  const calls = [];
  const context = {
    isRankScrollerSheetAllowed_(sheet) {
      return ['Novice', 'Intermediate', 'Advanced'].includes(
        typeof sheet === 'string' ? sheet : sheet.getName()
      );
    },
    triggerPerfStep_(trace, name, startedAt, metadata) {
      steps.push({ name, metadata });
    }
  };
  vm.createContext(context);
  const visibilityFile = 'rankscroller/rankScrollerVisibility.gs';
  vm.runInContext(
    fs.readFileSync(path.join(root, visibilityFile), 'utf8'),
    context,
    { filename: visibilityFile }
  );

  const sheet = {
    getName: () => 'Intermediate',
    hideColumns(start, count) { calls.push(['hideColumns', start, count]); },
    showColumns(start, count) { calls.push(['showColumns', start, count]); }
  };

  const result = context.showRankScrollerColumns_(
    sheet, 14, 16, 18, 4, { mode: 'DEBUG' }
  );
  assert.equal(result.success, true);
  assert.deepEqual(calls, [
    ['hideColumns', 14, 16],
    ['showColumns', 18, 4]
  ]);
  assert.deepEqual(steps.map(step => step.name), [
    'Rank Scroller: hideColumns',
    'Rank Scroller: showColumns'
  ]);

  calls.length = 0;
  steps.length = 0;
  const tiers = context.showRankScrollerTierColumns_(
    sheet, 14, 5, [100, 200, 350, 400, 500], { mode: 'DEBUG' }
  );
  assert.equal(tiers.success, true);
  assert.equal(tiers.groupsCount, 2);
  assert.deepEqual(calls, [
    ['hideColumns', 14, 5],
    ['showColumns', 14, 2],
    ['showColumns', 17, 2]
  ]);
  assert.equal(steps[0].name, 'Rank Scroller: hideColumns');
  assert.equal(steps[1].name, 'Rank Scroller: showColumns groups');
  assert.equal(steps[1].metadata.calls, 2);
});

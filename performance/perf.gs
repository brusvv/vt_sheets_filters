
/**
 * PERFORMANCE TRACER 2.2
 *
 * Единая система диагностики производительности.
 *
 * Режимы:
 * OFF   — диагностика отключена.
 * ERROR — только ошибки.
 * INFO  — START, END и ключевые EVENT.
 * DEBUG — полная трассировка, включая STEP.
 *
 * Поддерживает:
 * - traceId и parentId;
 * - измерение длительности операций и этапов;
 * - автоматическое определение вложенных этапов;
 * - общий и эксклюзивный временной охват;
 * - анализ этапов верхнего уровня;
 * - анализ неучтённого времени;
 * - статусы выполнения и metadata.
 *
 * Не изменяет бизнес-логику проекта.
 */

const PERF_CONFIG = {
  logPrefix: '[PERF]',
  validModes: ['OFF', 'ERROR', 'INFO', 'DEBUG']
};

// Кэш режима действует только в рамках одного выполнения скрипта.
var PERF_MODE_CACHE = null;

// Позволяет временно переопределить режим программно.
var PERF_MODE_OVERRIDE = null;


/**
 * Возвращает текущий режим диагностики.
 */
function perfGetMode_() {
  if (PERF_MODE_OVERRIDE !== null) {
    return PERF_MODE_OVERRIDE;
  }

  if (PERF_MODE_CACHE !== null) {
    return PERF_MODE_CACHE;
  }

  let mode = CONFIG.perfDefaultMode || 'DEBUG';

  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

    const optionsSheet = spreadsheet.getSheetByName(
      CONFIG.optionsSheetName || 'Options'
    );

    if (optionsSheet) {
      const configuredMode = String(
        optionsSheet.getRange(
          CONFIG.perfModeCell || 'A10'
        ).getValue()
      ).trim().toUpperCase();

      if (PERF_CONFIG.validModes.includes(configuredMode)) {
        mode = configuredMode;
      }
    }

  } catch (error) {
    console.error(
      '[PERF] Failed to read tracer mode: ' + error.message
    );
  }

  if (!PERF_CONFIG.validModes.includes(mode)) {
    mode = 'DEBUG';
  }

  PERF_MODE_CACHE = mode;

  return mode;
}


/**
 * Проверяет, разрешено ли логирование указанного уровня.
 */
function perfIsEnabled_(level) {
  const mode = perfGetMode_();

  const levels = {
    OFF: 0,
    ERROR: 1,
    INFO: 2,
    DEBUG: 3
  };

  return levels[mode] >= levels[level];
}


/**
 * Создаёт ID диагностического запуска.
 */
function perfCreateId_() {
  return Utilities.getUuid().slice(0, 8);
}


/**
 * Начинает трассировку.
 *
 * parentTrace — необязательная родительская трассировка.
 *
 * Старые вызовы:
 * perfStart_(operation, metadata)
 *
 * продолжают работать без изменений.
 */
function perfStart_(
  operation,
  metadata = {},
  parentTrace = null
) {
  const mode = perfGetMode_();

  if (mode === 'OFF') {
    return null;
  }

  const trace = {
    id: perfCreateId_(),
    parentId: parentTrace ? parentTrace.id : null,
    operation: operation,
    startedAt: Date.now(),
    startedIso: new Date().toISOString(),
    metadata: metadata,
    steps: [],
    mode: mode
  };

  if (mode === 'INFO' || mode === 'DEBUG') {
    perfLog_('START', trace, {
      metadata: metadata
    });
  }

  return trace;
}


/**
 * Фиксирует отдельный этап.
 *
 * startedAt — результат Date.now() в начале этапа.
 *
 * Время начала и завершения сохраняется
 * для последующего анализа вложенности.
 */
function perfStep_(
  trace,
  name,
  startedAt,
  metadata = {}
) {
  if (!trace || trace.mode !== 'DEBUG') return;

  const endedAt = Date.now();

  const validStart = Number.isFinite(startedAt)
    ? startedAt
    : endedAt;

  const duration = Math.max(
    0,
    endedAt - validStart
  );

  const step = {
    name: name,
    startedAt: validStart,
    endedAt: endedAt,
    durationMs: duration,
    metadata: metadata
  };

  trace.steps.push(step);

  perfLog_('STEP', trace, {
    name: name,
    durationMs: duration,
    metadata: metadata
  });
}


/**
 * Объединяет пересекающиеся интервалы.
 *
 * Возвращает:
 * - intervals — объединённые интервалы;
 * - durationMs — их суммарную длительность.
 */
function perfMergeIntervals_(intervals) {
  if (!intervals.length) {
    return {
      intervals: [],
      durationMs: 0
    };
  }

  const sorted = intervals.slice().sort((a, b) => {
    if (a.start !== b.start) {
      return a.start - b.start;
    }

    return a.end - b.end;
  });

  const merged = [];

  let current = {
    start: sorted[0].start,
    end: sorted[0].end
  };

  for (let i = 1; i < sorted.length; i++) {
    const interval = sorted[i];

    if (interval.start <= current.end) {
      current.end = Math.max(
        current.end,
        interval.end
      );

    } else {
      merged.push(current);

      current = {
        start: interval.start,
        end: interval.end
      };
    }
  }

  merged.push(current);

  const durationMs = merged.reduce(
    (sum, interval) => {
      return sum + interval.end - interval.start;
    },
    0
  );

  return {
    intervals: merged,
    durationMs: durationMs
  };
}


/**
 * Определяет этапы верхнего уровня.
 *
 * Этап считается вложенным, если его интервал
 * полностью находится внутри другого этапа.
 *
 * Для одинаковых интервалов один этап считается
 * родительским, чтобы не учитывать оба повторно.
 */
function perfGetRootIntervals_(intervals) {
  return intervals.filter((candidate, index) => {
    return !intervals.some((other, otherIndex) => {
      if (index === otherIndex) {
        return false;
      }

      const contains =
        other.start <= candidate.start &&
        other.end >= candidate.end;

      if (!contains) {
        return false;
      }

      const strictlyContains =
        other.start < candidate.start ||
        other.end > candidate.end;

      if (strictlyContains) {
        return true;
      }

      // При одинаковых границах оставляем
      // только первый зарегистрированный этап.
      return otherIndex < index;
    });
  });
}


/**
 * Рассчитывает временные показатели этапов.
 *
 * totalStepTimeMs:
 * сумма длительностей всех этапов.
 *
 * exclusiveStepTimeMs:
 * временной охват всех этапов без пересечений.
 *
 * rootStepTimeMs:
 * временной охват этапов верхнего уровня.
 *
 * overlapTimeMs:
 * повторно учтённое время.
 *
 * unaccountedTimeMs:
 * время операции вне всех этапов.
 */
function perfCalculateStepMetrics_(trace, duration) {
  const steps = trace.steps || [];

  const operationStart = trace.startedAt;
  const operationEnd = operationStart + duration;

  if (!steps.length) {
    return {
      totalStepTimeMs: 0,
      exclusiveStepTimeMs: 0,
      rootStepTimeMs: 0,
      overlapTimeMs: 0,
      unaccountedTimeMs: duration,
      unaccountedBeforeFirstStepMs: duration,
      unaccountedBetweenStepsMs: 0,
      unaccountedAfterLastStepMs: 0,
      rootStepsCount: 0,
      stepsOutsideTraceCount: 0
    };
  }

  const intervals = [];

  let totalStepTime = 0;
  let stepsOutsideTraceCount = 0;

  steps.forEach(step => {
    const rawStart = step.startedAt;
    const rawEnd = step.endedAt;

    if (
      !Number.isFinite(rawStart) ||
      !Number.isFinite(rawEnd) ||
      rawEnd < rawStart
    ) {
      stepsOutsideTraceCount++;
      return;
    }

    const start = Math.max(
      operationStart,
      rawStart
    );

    const end = Math.min(
      operationEnd,
      rawEnd
    );

    const clippedDuration = Math.max(
      0,
      end - start
    );

    if (
      rawStart < operationStart ||
      rawEnd > operationEnd
    ) {
      stepsOutsideTraceCount++;
    }

    if (clippedDuration <= 0) {
      return;
    }

    totalStepTime += clippedDuration;

    intervals.push({
      start: start,
      end: end
    });
  });

  // Общее покрытие всех этапов.
  const merged = perfMergeIntervals_(intervals);

  const exclusiveStepTime = merged.durationMs;

  // Этапы верхнего уровня.
  const rootIntervals = perfGetRootIntervals_(intervals);

  const rootMerged = perfMergeIntervals_(
    rootIntervals
  );

  // Повторно учтённое время.
  const overlapTime = Math.max(
    0,
    totalStepTime - exclusiveStepTime
  );

  // Неучтённое время.
  const gaps = {
    before: 0,
    between: 0,
    after: 0
  };

  if (merged.intervals.length) {
    const first = merged.intervals[0];
    const last = merged.intervals[
      merged.intervals.length - 1
    ];

    gaps.before = Math.max(
      0,
      first.start - operationStart
    );

    gaps.after = Math.max(
      0,
      operationEnd - last.end
    );

    for (
      let i = 1;
      i < merged.intervals.length;
      i++
    ) {
      gaps.between += Math.max(
        0,
        merged.intervals[i].start -
        merged.intervals[i - 1].end
      );
    }

  } else {
    gaps.before = duration;
  }

  const unaccountedTime =
    gaps.before +
    gaps.between +
    gaps.after;

  return {
    totalStepTimeMs: totalStepTime,
    exclusiveStepTimeMs: exclusiveStepTime,
    rootStepTimeMs: rootMerged.durationMs,
    overlapTimeMs: overlapTime,
    unaccountedTimeMs: unaccountedTime,

    unaccountedBeforeFirstStepMs: gaps.before,
    unaccountedBetweenStepsMs: gaps.between,
    unaccountedAfterLastStepMs: gaps.after,

    rootStepsCount: rootIntervals.length,
    stepsOutsideTraceCount: stepsOutsideTraceCount
  };
}


/**
 * Фиксирует событие без измерения длительности.
 */
function perfEvent_(trace, name, metadata = {}) {
  if (!trace || !['INFO', 'DEBUG'].includes(trace.mode)) {
    return;
  }

  perfLog_('EVENT', trace, {
    name: name,
    metadata: metadata
  });
}


/**
 * Завершает трассировку.
 *
 * В режиме ERROR записывает ошибки,
 * включая LOCK_ERROR и LOCK_RELEASE_ERROR.
 */
function perfEnd_(
  trace,
  status = 'SUCCESS',
  metadata = {}
) {
  if (!trace) return;

  const normalizedStatus = String(status).toUpperCase();

  const isError =
    normalizedStatus === 'ERROR' ||
    normalizedStatus.endsWith('_ERROR') ||
    normalizedStatus.endsWith('_FAILED') ||
    normalizedStatus === 'TIMEOUT';

  if (trace.mode === 'OFF') return;

  if (trace.mode === 'ERROR' && !isError) {
    return;
  }

  const duration = Math.max(
    0,
    Date.now() - trace.startedAt
  );

  perfLog_('END', trace, {
    status: normalizedStatus,
    durationMs: duration,
    metadata: metadata
  });

  if (trace.mode === 'DEBUG') {
    const metrics = perfCalculateStepMetrics_(
      trace,
      duration
    );

    perfLog_('SUMMARY', trace, {
      status: normalizedStatus,
      durationMs: duration,
      stepsCount: trace.steps.length,

      totalStepTimeMs: metrics.totalStepTimeMs,
      exclusiveStepTimeMs: metrics.exclusiveStepTimeMs,
      rootStepTimeMs: metrics.rootStepTimeMs,
      overlapTimeMs: metrics.overlapTimeMs,

      unaccountedTimeMs: metrics.unaccountedTimeMs,

      unaccountedBeforeFirstStepMs:
        metrics.unaccountedBeforeFirstStepMs,

      unaccountedBetweenStepsMs:
        metrics.unaccountedBetweenStepsMs,

      unaccountedAfterLastStepMs:
        metrics.unaccountedAfterLastStepMs,

      rootStepsCount: metrics.rootStepsCount,

      stepsOutsideTraceCount:
        metrics.stepsOutsideTraceCount
    });
  }
}


/**
 * Записывает структурированное событие.
 */
function perfLog_(type, trace, data = {}) {
  if (!trace) return;

  const record = {
    type: type,
    traceId: trace.id,
    parentId: trace.parentId,
    operation: trace.operation,
    timestamp: new Date().toISOString(),
    ...data
  };

  console.log(
    PERF_CONFIG.logPrefix + ' ' +
    JSON.stringify(record)
  );
}


/**
 * Диагностика ожидания ScriptLock.
 */
function perfLockWait_(trace, duration) {
  if (!trace) return;

  try {
    if (typeof perfEvent_ !== 'function') return;
    perfEvent_(trace, 'LOCK_ACQUIRED', {
      waitMs: duration
    });
  } catch (error) {
    // Lock diagnostics are best-effort and must never escape.
  }
}


/**
 * Включение / отключение диагностики.
 *
 * true  → DEBUG
 * false → OFF
 */
function setPerformanceTracer(enabled) {
  PERF_MODE_OVERRIDE = enabled ? 'DEBUG' : 'OFF';

  console.log(
    '[PERF] Tracer mode: ' + PERF_MODE_OVERRIDE
  );
}


/**
 * Программная смена режима.
 */
function setPerformanceMode(mode) {
  mode = String(mode).trim().toUpperCase();

  if (!PERF_CONFIG.validModes.includes(mode)) {
    throw new Error(
      'Unknown Performance Tracer mode: ' + mode
    );
  }

  PERF_MODE_OVERRIDE = mode;

  console.log('[PERF] Tracer mode: ' + mode);
}


/**
 * Сбрасывает кэш режима.
 */
function resetPerformanceModeCache() {
  PERF_MODE_CACHE = null;
  PERF_MODE_OVERRIDE = null;

  console.log('[PERF] Mode cache reset.');
}


/**
 * Ручное тестирование логгера.
 */
function testPerformanceTracer() {
  const trace = perfStart_('Tracer self-test', {
    source: 'manual'
  });

  if (!trace) {
    console.log('[PERF] Tracer is OFF.');
    return;
  }

  const parentStarted = Date.now();

  const firstStarted = Date.now();

  Utilities.sleep(100);

  perfStep_(
    trace,
    'First nested step',
    firstStarted
  );

  const secondStarted = Date.now();

  Utilities.sleep(100);

  perfStep_(
    trace,
    'Second nested step',
    secondStarted
  );

  perfStep_(
    trace,
    'Parent step',
    parentStarted
  );

  perfEvent_(
    trace,
    'TEST_EVENT',
    {
      message: 'Tracer test event'
    }
  );

  perfEnd_(trace);
}



/**
 * COMPATIBILITY WRAPPERS
 *
 * Безопасная интеграция Performance Tracer.
 *
 * Позволяют остальным файлам проекта работать
 * даже в том случае, если Tracer недоступен.
 *
 * Все вызовы централизованы здесь.
 */


/**
 * Начинает трассировку.
 *
 * Поддерживает parentTrace для вложенных операций.
 */
function triggerPerfStart_(
  operation,
  metadata = {},
  parentTrace = null
) {
  if (typeof perfStart_ !== 'function') {
    return null;
  }

  try {
    return perfStart_(
      operation,
      metadata,
      parentTrace
    );
  } catch (error) {
    // Tracing is best-effort and must not interrupt business operations.
    return null;
  }
}


/**
 * Записывает событие.
 */
function triggerPerfEvent_(
  trace,
  name,
  metadata = {}
) {
  if (
    !trace ||
    typeof perfEvent_ !== 'function'
  ) {
    return;
  }

  try {
    perfEvent_(
      trace,
      name,
      metadata
    );
  } catch (error) {
    // Diagnostic failures must never escape into business code.
  }
}


/**
 * Записывает этап выполнения.
 */
function triggerPerfStep_(
  trace,
  name,
  startedAt,
  metadata = {}
) {
  if (
    !trace ||
    typeof perfStep_ !== 'function'
  ) {
    return;
  }

  try {
    perfStep_(
      trace,
      name,
      startedAt,
      metadata
    );
  } catch (error) {
    // Diagnostic failures must never escape into business code.
  }
}


/**
 * Завершает трассировку.
 */
function triggerPerfEnd_(
  trace,
  status = 'SUCCESS',
  metadata = {}
) {
  if (
    !trace ||
    typeof perfEnd_ !== 'function'
  ) {
    return;
  }

  try {
    perfEnd_(
      trace,
      status,
      metadata
    );
  } catch (error) {
    // Diagnostic failures must never escape into business code.
  }
}


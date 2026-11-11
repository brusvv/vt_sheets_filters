
/**
 * LOCK MANAGER 2.1
 *
 * Централизованное управление ScriptLock.
 *
 * Поддерживает:
 * - ожидание блокировки;
 * - измерение времени ожидания;
 * - измерение callback;
 * - измерение освобождения блокировки;
 * - обработку ошибок;
 * - сохранение исходной ошибки;
 * - автоматическое освобождение блокировки;
 * - интеграцию с Performance Tracer;
 * - передачу родительской трассировки.
 */


/**
 * Блокировка параллельного выполнения скриптов.
 *
 * Если parentTrace передан:
 * использует существующую трассировку.
 *
 * Если parentTrace отсутствует:
 * создаёт собственную трассировку.
 */
function withProjectLock(callback, parentTrace = null) {
  const ownsTrace = !parentTrace;

  const trace = parentTrace ||
    triggerPerfStart_('withProjectLock');

  const lock = LockService.getScriptLock();

  let lockAcquired = false;

  let status = 'SUCCESS';

  let errorMetadata = {};

  let operationError = null;

  let result;

  /**
   * Ожидание блокировки.
   */
  const waitStarted = Date.now();

  try {
    lock.waitLock(30000);

    lockAcquired = true;

  } catch (error) {
    status = 'LOCK_ERROR';

    operationError = error;

    errorMetadata = {
      message: error.message,
      stack: error.stack
    };

  } finally {
    const waitDuration = Date.now() - waitStarted;

    // Событие успешного получения блокировки.
    if (lockAcquired) {
      triggerPerfEvent_(
        trace,
        'LOCK_ACQUIRED',
        {
          waitMs: waitDuration
        }
      );
    } else {
      triggerPerfEvent_(
        trace,
        'LOCK_WAIT_FAILED',
        {
          waitMs: waitDuration,
          message: operationError
            ? operationError.message
            : 'Unknown lock error'
        }
      );
    }

    // Этап ожидания фиксируется даже при ошибке.
    triggerPerfStep_(
      trace,
      'Lock wait',
      waitStarted,
      {
        status: lockAcquired ? 'ACQUIRED' : 'FAILED',
        waitMs: waitDuration
      }
    );
  }


  /**
   * Выполнение callback.
   */
  if (lockAcquired && !operationError) {
    const executionStarted = Date.now();

    try {
      result = callback();

    } catch (error) {
      status = 'ERROR';

      operationError = error;

      errorMetadata = {
        message: error.message,
        stack: error.stack
      };

    } finally {
      const executionDuration =
        Date.now() - executionStarted;

      triggerPerfStep_(
        trace,
        'Locked callback execution',
        executionStarted,
        {
          status: status,
          durationMs: executionDuration
        }
      );
    }
  }


  /**
   * Освобождение блокировки.
   */
  if (lockAcquired) {
    const releaseStarted = Date.now();

    try {
      lock.releaseLock();

    } catch (error) {
      // Не заменяем первоначальную ошибку,
      // если callback или waitLock уже завершились ошибкой.
      if (!operationError) {
        status = 'LOCK_RELEASE_ERROR';

        operationError = error;

        errorMetadata = {
          message: error.message,
          stack: error.stack
        };

      } else {
        errorMetadata.releaseError = {
          message: error.message,
          stack: error.stack
        };
      }

    } finally {
      triggerPerfStep_(
        trace,
        'Lock release',
        releaseStarted,
        {
          status: status
        }
      );
    }
  }


  /**
   * Завершение собственной трассировки.
   *
   * Если parentTrace передан, закрытие остаётся
   * ответственностью родительской функции.
   */
  if (ownsTrace) {
    triggerPerfEnd_(
      trace,
      status,
      errorMetadata
    );
  }


  /**
   * Повторно выбрасываем исходную ошибку
   * только после выполнения очистки.
   */
  if (operationError) {
    throw operationError;
  }

  return result;
}

/**
 * OPERATIONS
 *
 * Общие вспомогательные функции комплексных операций.
 */


/**
 * Записывает этап комплексной операции
 * в Performance Tracer.
 */
function operationPerfStep_(
  trace,
  label,
  startedAt,
  metadata = {}
) {
  triggerPerfStep_(
    trace,
    label,
    startedAt,
    metadata
  );
}


/**
 * Завершает операцию с учётом возможной ошибки.
 */
function operationPerfEnd_(
  trace,
  status,
  error = null
) {
  triggerPerfEnd_(
    trace,
    status,
    error
      ? {
          message: error.message,
          stack: error.stack
        }
      : {}
  );
}

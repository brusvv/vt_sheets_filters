
/**
 * FILTER OPTIONS
 *
 * Работа с настройками листа Options
 * и вспомогательные функции фильтрации.
 *
 * Содержит:
 * - чтение Boolean-настроек;
 * - проверку автоматической сортировки;
 * - логирование этапов фильтрации.
 */


/**
 * Записывает этап одновременно в старый журнал
 * и в Performance Tracer.
 */
function filterPerfStep_(
  trace,
  label,
  startedAt,
  metadata = {}
) {
  const duration = Date.now() - startedAt;

  logFilterTiming(label, startedAt);

  triggerPerfStep_(
    trace,
    label,
    startedAt,
    {
      ...metadata,
      durationMs: duration
    }
  );
}


/**
 * Читает Boolean-настройку с листа Options.
 */
function getOptionBoolean_(optionName) {
  const option = CONFIG.options[optionName];

  if (!option || !option.cell) {
    throw new Error(
      `Настройка "${optionName}" не найдена в CONFIG.options.`
    );
  }

  const optionsSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(CONFIG.optionsSheetName);

  if (!optionsSheet) {
    throw new Error(
      `Лист "${CONFIG.optionsSheetName}" не найден.`
    );
  }

  return optionsSheet
    .getRange(option.cell)
    .getValue() === true;
}


/**
 * Проверяет, включена ли автоматическая сортировка.
 */
function isAutoSortEnabled() {
  return getOptionBoolean_('autoSort');
}

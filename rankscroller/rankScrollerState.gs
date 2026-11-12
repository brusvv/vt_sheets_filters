
/**
 * RANK SCROLLER STATE
 *
 * Централизованное управление Script Properties:
 * - чтение состояния;
 * - сохранение изменений;
 * - получение dirty-флагов;
 * - установка dirty-флагов;
 * - сброс dirty-флагов.
 *
 * Конфигурация листов берётся из CONFIG.rankScroller.sheets.
 */


/**
 * Возвращает Script Properties для Rank Scroller.
 *
 * Читает все свойства одним запросом.
 */
function getRankScrollerProperties_() {
  return PropertiesService
    .getScriptProperties()
    .getProperties();
}


/**
 * Сохраняет переданные свойства.
 *
 * Пустой объект не вызывает setProperties().
 *
 * @param {Object} propertiesToSave
 * @return {{saved: boolean, keys: string[]}}
 */
function saveRankScrollerProperties_(propertiesToSave) {
  const keys = Object.keys(propertiesToSave || {});

  if (keys.length === 0) {
    return {
      saved: false,
      keys: []
    };
  }

  PropertiesService
    .getScriptProperties()
    .setProperties(propertiesToSave);

  return {
    saved: true,
    keys: keys
  };
}


/**
 * Возвращает список листов Rank Scroller
 * на основании CONFIG.
 *
 * @return {string[]}
 */
function getRankScrollerSheetNames_() {
  if (
    !CONFIG.rankScroller ||
    !CONFIG.rankScroller.sheets
  ) {
    throw new Error(
      'CONFIG.rankScroller.sheets is not configured.'
    );
  }

  return Object.keys(CONFIG.rankScroller.sheets);
}


/**
 * Проверяет, помечен ли лист как dirty.
 *
 * @param {Object} properties
 * @param {string} sheetName
 * @return {boolean}
 */
function isRankSheetDirty_(properties, sheetName) {
  const dirtyKey = 'rankDirty_' + sheetName;

  return properties[dirtyKey] === 'true';
}


/**
 * Помечает конкретный лист как требующий
 * обновления Rank Scroller.
 */
function markRankSheetDirty(sheetName) {
  const sheetNames = getRankScrollerSheetNames_();

  if (!sheetNames.includes(sheetName)) {
    return;
  }

  PropertiesService
    .getScriptProperties()
    .setProperty(
      'rankDirty_' + sheetName,
      'true'
    );
}


/**
 * Помечает все листы Rank Scroller
 * как требующие обновления.
 *
 * Записывает свойства одним запросом.
 */
function markAllRankSheetsDirty() {
  const properties = {};

  getRankScrollerSheetNames_().forEach(sheetName => {
    properties['rankDirty_' + sheetName] = 'true';
  });

  if (Object.keys(properties).length === 0) {
    return;
  }

  PropertiesService
    .getScriptProperties()
    .setProperties(properties);
}


/**
 * Подготавливает сброс dirty-флага
 * после успешной обработки листа.
 *
 * Функция только формирует объект изменений.
 * Фактическая запись выполняется общей функцией сохранения.
 *
 * @param {Object} propertiesToSave
 * @param {string} sheetName
 */
function resetRankSheetDirty_(propertiesToSave, sheetName) {
  propertiesToSave['rankDirty_' + sheetName] = 'false';
}

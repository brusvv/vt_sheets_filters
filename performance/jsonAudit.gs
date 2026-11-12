
/**
 * READ-ONLY AUDIT EXPORT
 *
 * Читает данные всех листов и создаёт отдельный JSON-файл
 * на Google Drive.
 *
 * Сама таблица не изменяется:
 * нет записи в ячейки, сортировки, фильтрации,
 * изменения видимости строк или настроек листов.
 */

function exportReadOnlyAuditJson() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const payload = {
    exportedAt: new Date().toISOString(),
    spreadsheet: {
      name: ss.getName(),
      id: ss.getId(),
      url: ss.getUrl(),
      timezone: ss.getSpreadsheetTimeZone(),
      locale: ss.getSpreadsheetLocale(),
      activeSheet: ss.getActiveSheet().getName()
    },

    namedRanges: ss.getNamedRanges().map(nr => {
      const range = nr.getRange();

      return {
        name: nr.getName(),
        sheet: range.getSheet().getName(),
        a1Notation: range.getA1Notation()
      };
    }),

    sheets: ss.getSheets().map(auditExportSheet_)
  };

  const json = JSON.stringify(payload, null, 2);

  const timestamp = Utilities.formatDate(
    new Date(),
    ss.getSpreadsheetTimeZone(),
    'yyyyMMdd_HHmmss'
  );

  const filename = 'Sheets_Audit_' + timestamp + '.json';

  // Создаётся отдельный файл на Google Drive.
  // Содержимое самой таблицы не изменяется.
  const file = DriveApp.createFile(filename, json);

  Logger.log('JSON file: ' + file.getUrl());
  Logger.log('File size: ' + json.length + ' characters');

  return file.getUrl();
}


/**
 * Выгрузка одного листа.
 */
function auditExportSheet_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();

  const rowCount = Math.max(lastRow, 1);
  const columnCount = Math.max(lastColumn, 1);

  const range = sheet.getRange(
    1, 1, rowCount, columnCount
  );

  const values = range.getValues();

  return {
    name: sheet.getName(),
    sheetId: sheet.getSheetId(),
    type: String(sheet.getType()),
    hidden: sheet.isSheetHidden(),

    lastRow: lastRow,
    lastColumn: lastColumn,

    frozenRows: sheet.getFrozenRows(),
    frozenColumns: sheet.getFrozenColumns(),

    gridRange: range.getA1Notation(),

    // Значения, полученные из ячеек.
    effectiveValues: auditNormalizeMatrix_(values),

    // Отображаемые значения с учётом формата.
    displayValues: range.getDisplayValues(),

    // Формулы в A1 и R1C1.
    formulas: range.getFormulas(),
    formulasR1C1: range.getFormulasR1C1(),

    // Примечания.
    notes: range.getNotes(),

    // Rich Text и гиперссылки.
    richText: auditRichTextMatrix_(range.getRichTextValues()),

    // Проверки данных.
    dataValidations: auditValidationMatrix_(
      range.getDataValidations()
    ),

    // Форматы ячеек.
    formatting: {
      numberFormats: range.getNumberFormats(),
      backgrounds: range.getBackgrounds(),
      fontColors: range.getFontColors(),
      fontFamilies: range.getFontFamilies(),
      fontSizes: range.getFontSizes(),
      fontWeights: range.getFontWeights(),
      fontStyles: range.getFontStyles(),
      horizontalAlignments: range.getHorizontalAlignments(),
      verticalAlignments: range.getVerticalAlignments(),
      wrapStrategies: auditNormalizeMatrix_(
        range.getWrapStrategies()
      )
    },

    // Объединённые диапазоны.
    mergedRanges: range.getMergedRanges().map(r => ({
      a1Notation: r.getA1Notation(),
      row: r.getRow(),
      column: r.getColumn(),
      rows: r.getNumRows(),
      columns: r.getNumColumns()
    })),

    // Состояние строк.
    rows: Array.from(
      { length: lastRow },
      (_, index) => {
        const row = index + 1;

        return {
          row: row,
          height: sheet.getRowHeight(row),
          hiddenByUser: sheet.isRowHiddenByUser(row),
          hiddenByFilter: sheet.isRowHiddenByFilter(row),
          groupDepth: sheet.getRowGroupDepth(row)
        };
      }
    ),

    // Состояние столбцов.
    columns: Array.from(
      { length: lastColumn },
      (_, index) => {
        const column = index + 1;

        return {
          column: column,
          width: sheet.getColumnWidth(column),
          hiddenByUser: sheet.isColumnHiddenByUser(column)
        };
      }
    ),

    // Правила условного форматирования.
    conditionalFormatting: sheet
      .getConditionalFormatRules()
      .map(auditConditionalRule_),

    // Фильтр листа.
    filter: auditFilter_(sheet)
  };
}


/**
 * Преобразование значений в формат JSON.
 */
function auditNormalizeMatrix_(matrix) {
  return matrix.map(row =>
    row.map(value => auditNormalizeValue_(value))
  );
}


function auditNormalizeValue_(value) {
  if (value instanceof Date) {
    return {
      type: 'Date',
      iso: value.toISOString()
    };
  }

  if (value === null || value === undefined) {
    return null;
  }

  if (Array.isArray(value)) {
    return value.map(auditNormalizeValue_);
  }

  if (value && typeof value.getA1Notation === 'function') {
    return {
      type: 'Range',
      sheet: value.getSheet().getName(),
      a1Notation: value.getA1Notation()
    };
  }

  if (
    typeof value === 'string' ||
    typeof value === 'number' ||
    typeof value === 'boolean'
  ) {
    return value;
  }

  return String(value);
}


/**
 * Rich Text, ссылки и форматирование отдельных фрагментов.
 */
function auditRichTextMatrix_(matrix) {
  return matrix.map(row =>
    row.map(value => {
      if (!value) return null;

      return {
        text: value.getText(),
        linkUrl: value.getLinkUrl(),

        runs: value.getRuns().map(run => {
          const style = run.getTextStyle();

          return {
            text: run.getText(),
            startIndex: run.getStartIndex(),
            endIndex: run.getEndIndex(),

            style: {
              bold: style.isBold(),
              italic: style.isItalic(),
              underline: style.isUnderline(),
              fontFamily: style.getFontFamily(),
              fontSize: style.getFontSize(),
              foregroundColor: style.getForegroundColor(),
              linkUrl: style.getLinkUrl()
            }
          };
        })
      };
    })
  );
}


/**
 * Правила проверки данных.
 */
function auditValidationMatrix_(matrix) {
  return matrix.map(row =>
    row.map(rule => {
      if (!rule) return null;

      return {
        criteriaType: String(rule.getCriteriaType()),

        criteriaValues: rule.getCriteriaValues().map(
          auditNormalizeValue_
        ),

        allowInvalid: rule.getAllowInvalid(),
        helpText: rule.getHelpText()
      };
    })
  );
}


/**
 * Условное форматирование.
 */
function auditRichTextMatrix_(matrix) {
  return matrix.map(row =>
    row.map(value => {
      if (!value) return null;

      return {
        text: value.getText(),
        linkUrl: value.getLinkUrl(),

        runs: value.getRuns().map(run => {
          const style = run.getTextStyle();

          return {
            text: run.getText(),
            startIndex: run.getStartIndex(),
            endIndex: run.getEndIndex(),

            // Ссылка конкретного текстового фрагмента.
            linkUrl: run.getLinkUrl(),

            style: {
              bold: style.isBold(),
              italic: style.isItalic(),
              underline: style.isUnderline(),
              fontFamily: style.getFontFamily(),
              fontSize: style.getFontSize(),
              foregroundColor: style.getForegroundColor()
            }
          };
        })
      };
    })
  );
}



/**
 * Фильтр листа и критерии по столбцам.
 */
function auditFilter_(sheet) {
  const filter = sheet.getFilter();

  if (!filter) return null;

  const range = filter.getRange();
  const criteria = {};

  for (
    let column = range.getColumn();
    column <= range.getLastColumn();
    column++
  ) {
    const rule = filter.getColumnFilterCriteria(column);

    if (!rule) continue;

    criteria[column] = {
      criteriaType: String(rule.getCriteriaType()),

      criteriaValues: rule.getCriteriaValues()
        .map(auditNormalizeValue_),

      hiddenValues: rule.getHiddenValues()
    };
  }

  return {
    range: range.getA1Notation(),
    columnCriteria: criteria
  };
}



/**
 * Выгрузка правил условного форматирования.
 */
function auditConditionalRule_(rule) {
  const booleanCondition = rule.getBooleanCondition();
  const gradientCondition = rule.getGradientCondition();

  return {
    ranges: rule.getRanges().map(
      range => range.getA1Notation()
    ),

    booleanCondition: booleanCondition
      ? {
          type: String(
            booleanCondition.getCriteriaType()
          ),

          values: booleanCondition
            .getCriteriaValues()
            .map(auditNormalizeValue_)
        }
      : null,

    hasGradientCondition: Boolean(gradientCondition)
  };
}


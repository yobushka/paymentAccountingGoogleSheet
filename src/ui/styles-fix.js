/**
 * @fileoverview Визуальная починка и восстановление стилей таблиц
 * @version 2.0
 */

// ============================================================================
// Цветовая схема
// ============================================================================

const STYLE_COLORS = {
  // Заголовки
  HEADER_BG: '#4a86e8',
  HEADER_TEXT: '#ffffff',
  
  // Зебра-полосы
  ZEBRA_EVEN: '#ffffff',
  ZEBRA_ODD: '#f8f9fa',
  
  // Статусы
  STATUS_OPEN: '#c8e6c9',      // зелёный
  STATUS_CLOSED: '#e0e0e0',    // серый
  STATUS_CANCELLED: '#ffcdd2', // красный
  STATUS_YES: '#c8e6c9',       // зелёный
  STATUS_NO: '#ffcdd2',        // красный
  
  // Баланс
  POSITIVE: '#c8e6c9',         // зелёный (переплата)
  NEGATIVE: '#ffcdd2',         // красный (долг)
  NEUTRAL: '#fff9c4',          // жёлтый (ноль)
  
  // Неактивные строки
  INACTIVE_BG: '#f5f5f5',
  INACTIVE_TEXT: '#9e9e9e',
  
  // Границы
  BORDER: '#dadce0'
};

// ============================================================================
// Главные функции починки стилей
// ============================================================================

/**
 * Полная починка стилей всех листов
 * Точка входа из меню
 */
function fixAllSheetsStyles() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Починка стилей',
    'Будут восстановлены:\n\n' +
    '• Форматирование заголовков\n' +
    '• Чередующиеся строки (zebra)\n' +
    '• Условное форматирование\n' +
    '• Числовые форматы\n' +
    '• Ширины колонок\n' +
    '• Фильтры и закреплённые строки\n' +
    '• Примечания к заголовкам\n\n' +
    'Продолжить?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    const ss = SpreadsheetApp.getActive();
    const specs = getSheetsSpec();
    const version = detectVersion();
    let fixed = 0;
    
    specs.forEach(spec => {
      // Пропускаем legacy/новый лист в зависимости от версии
      if (version === 'v2' && spec.name === SHEET_NAMES.COLLECTIONS) return;
      if (version === 'v1' && spec.name === SHEET_NAMES.GOALS) return;
      
      const sh = ss.getSheetByName(spec.name);
      if (sh) {
        fixSheetStyles_(sh, spec);
        fixed++;
      }
    });
    
    // Добавляем примечания
    addHeaderNotes_();
    
    ui.alert(
      'Стили восстановлены',
      `Обработано листов: ${fixed}`,
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Ошибка', e.message, ui.ButtonSet.OK);
    Logger.log('fixAllSheetsStyles error: ' + e.message);
  }
}

/**
 * Починка стилей текущего листа
 * Точка входа из меню
 */
function fixCurrentSheetStyles() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  const sheetName = sh.getName();
  
  const specs = getSheetsSpec();
  const spec = specs.find(s => s.name === sheetName);
  
  if (!spec) {
    SpreadsheetApp.getUi().alert(
      'Лист не распознан',
      `Лист «${sheetName}» не является системным листом.`,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  fixSheetStyles_(sh, spec);
  SpreadsheetApp.getActive().toast(`Стили листа «${sheetName}» восстановлены.`, 'Funds');
}

/**
 * Полная починка стилей одного листа
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {{name: string, headers: string[], colWidths: number[], dateCols?: number[]}} spec
 */
function fixSheetStyles_(sh, spec) {
  const sheetName = spec.name;
  Logger.log(`Fixing styles for: ${sheetName}`);
  
  // 1. Очищаем все существующие стили
  clearAllStyles_(sh);
  
  // 2. Базовые стили (заголовок, zebra, ширины)
  applyBaseStyles_(sh, spec);
  
  // 3. Специфичные стили для листа
  applySheetSpecificStyles_(sh, sheetName);
  
  // 4. Фильтры
  applyAutoFilter_(sh);
  
  Logger.log(`Styles fixed for: ${sheetName}`);
}

// ============================================================================
// Очистка стилей
// ============================================================================

/**
 * Очищает все стили листа
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function clearAllStyles_(sh) {
  const lastRow = Math.max(sh.getLastRow(), 1);
  const lastCol = Math.max(sh.getLastColumn(), 1);
  
  // Удаляем условное форматирование
  sh.setConditionalFormatRules([]);
  
  // Удаляем banding (чередующиеся строки)
  const bandings = sh.getBandings();
  bandings.forEach(b => b.remove());
  
  // Удаляем фильтры
  const filter = sh.getFilter();
  if (filter) filter.remove();
  
  // Сбрасываем форматирование данных
  const dataRange = sh.getRange(1, 1, lastRow, lastCol);
  dataRange
    .setBackground(null)
    .setFontColor('#000000')
    .setFontWeight('normal')
    .setFontStyle('normal')
    .setHorizontalAlignment('left')
    .setBorder(false, false, false, false, false, false);
  
  // Удаляем примечания с заголовков
  if (lastCol > 0) {
    sh.getRange(1, 1, 1, lastCol).clearNote();
  }
}

// ============================================================================
// Базовые стили
// ============================================================================

/**
 * Применяет базовые стили к листу
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {{name: string, headers: string[], colWidths: number[], dateCols?: number[]}} spec
 */
function applyBaseStyles_(sh, spec) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  
  if (lastCol < 1) return;
  
  // 1. Стили заголовка
  const headerRange = sh.getRange(1, 1, 1, lastCol);
  headerRange
    .setFontWeight('bold')
    .setBackground(STYLE_COLORS.HEADER_BG)
    .setFontColor(STYLE_COLORS.HEADER_TEXT)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrap(true);
  
  // Закрепляем первую строку
  sh.setFrozenRows(1);
  
  // 2. Ширины колонок
  spec.colWidths.forEach((w, i) => {
    if (w && i < lastCol) {
      sh.setColumnWidth(i + 1, w);
    }
  });
  
  // 3. Чередующиеся строки (zebra)
  if (lastRow > 1) {
    applyZebraStripesEnhanced_(sh, 2, lastRow, lastCol);
  }
  
  // 4. Форматы дат
  if (spec.dateCols && lastRow > 1) {
    spec.dateCols.forEach(col => {
      if (col <= lastCol) {
        sh.getRange(2, col, lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
      }
    });
  }
  
  // 5. Границы данных
  if (lastRow > 0) {
    const allRange = sh.getRange(1, 1, lastRow, lastCol);
    allRange.setBorder(true, true, true, true, true, true, STYLE_COLORS.BORDER, SpreadsheetApp.BorderStyle.SOLID);
  }
}

/**
 * Применяет улучшенные чередующиеся строки
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {number} startRow
 * @param {number} endRow
 * @param {number} lastCol
 */
function applyZebraStripesEnhanced_(sh, startRow, endRow, lastCol) {
  if (endRow < startRow || lastCol < 1) return;
  
  const dataRange = sh.getRange(startRow, 1, endRow - startRow + 1, lastCol);
  
  try {
    // Удаляем старые bandings
    const bandings = sh.getBandings();
    bandings.forEach(b => b.remove());
    
    // Применяем banding
    dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
  } catch (e) {
    // Fallback: ручное чередование
    for (let row = startRow; row <= endRow; row++) {
      const color = (row - startRow) % 2 === 0 ? STYLE_COLORS.ZEBRA_EVEN : STYLE_COLORS.ZEBRA_ODD;
      sh.getRange(row, 1, 1, lastCol).setBackground(color);
    }
  }
}

// ============================================================================
// Специфичные стили для листов
// ============================================================================

/**
 * Применяет специфичные стили в зависимости от типа листа
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {string} sheetName
 */
function applySheetSpecificStyles_(sh, sheetName) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  switch (sheetName) {
    case SHEET_NAMES.FAMILIES:
      applyFamiliesStyles_(sh, lastRow);
      break;
    case SHEET_NAMES.GOALS:
      applyGoalsStyles_(sh, lastRow);
      break;
    case SHEET_NAMES.COLLECTIONS:
      applyCollectionsStyles_(sh, lastRow);
      break;
    case SHEET_NAMES.PAYMENTS:
      applyPaymentsStyles_(sh, lastRow);
      break;
    case SHEET_NAMES.PARTICIPATION:
      applyParticipationStyles_(sh, lastRow);
      break;
    case SHEET_NAMES.BALANCE:
      applyBalanceStyles_(sh, lastRow);
      break;
    case SHEET_NAMES.DETAIL:
      applyDetailStyles_(sh, lastRow);
      break;
    case SHEET_NAMES.SUMMARY:
      applySummaryStyles_(sh, lastRow);
      break;
    case SHEET_NAMES.ISSUES:
      applyIssuesStyles_(sh, lastRow);
      break;
    case SHEET_NAMES.ISSUE_STATUS:
      applyIssueStatusStyles_(sh, lastRow);
      break;
  }
}

/**
 * Стили для листа «Семьи»
 */
function applyFamiliesStyles_(sh, lastRow) {
  const map = getHeaderMap_(sh);
  const rules = [];
  
  // Формат дат
  ['День рождения', 'Членство с', 'Членство по'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });
  
  // Условное форматирование: неактивные — серый фон
  const activeCol = map['Активен'];
  if (activeCol) {
    const lastCol = sh.getLastColumn();
    const rowRange = sh.getRange(2, 1, lastRow - 1, lastCol);
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${colToLetter_(activeCol)}2="${ACTIVE_STATUS.NO}"`)
      .setBackground(STYLE_COLORS.INACTIVE_BG)
      .setFontColor(STYLE_COLORS.INACTIVE_TEXT)
      .setRanges([rowRange])
      .build());
  }
  
  sh.setConditionalFormatRules(rules);
}

/**
 * Стили для листа «Цели»
 */
function applyGoalsStyles_(sh, lastRow) {
  const map = getHeaderMap_(sh);
  const rules = [];
  
  // Числовые форматы
  ['Параметр суммы', 'Фиксированный x', 'Возмещено'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('#,##0.00');
    }
  });
  
  // Условное форматирование: Статус
  const statusCol = map['Статус'];
  if (statusCol) {
    const statusRange = sh.getRange(2, statusCol, lastRow - 1, 1);
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(GOAL_STATUS.OPEN)
      .setBackground(STYLE_COLORS.STATUS_OPEN)
      .setRanges([statusRange])
      .build());
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(GOAL_STATUS.CLOSED)
      .setBackground(STYLE_COLORS.STATUS_CLOSED)
      .setRanges([statusRange])
      .build());
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(GOAL_STATUS.CANCELLED)
      .setBackground(STYLE_COLORS.STATUS_CANCELLED)
      .setRanges([statusRange])
      .build());
  }
  
  sh.setConditionalFormatRules(rules);
}

/**
 * Стили для листа «Сборы» (legacy v1.x)
 */
function applyCollectionsStyles_(sh, lastRow) {
  const map = getHeaderMap_(sh);
  const rules = [];
  
  ['Параметр суммы', 'Фиксированный x', 'Возмещено'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('#,##0.00');
    }
  });
  
  const statusCol = map['Статус'];
  if (statusCol) {
    const statusRange = sh.getRange(2, statusCol, lastRow - 1, 1);
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(COLLECTION_STATUS_V1.OPEN)
      .setBackground(STYLE_COLORS.STATUS_OPEN)
      .setRanges([statusRange])
      .build());
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(COLLECTION_STATUS_V1.CLOSED)
      .setBackground(STYLE_COLORS.STATUS_CLOSED)
      .setRanges([statusRange])
      .build());
  }
  
  sh.setConditionalFormatRules(rules);
}

/**
 * Стили для листа «Платежи»
 */
function applyPaymentsStyles_(sh, lastRow) {
  const map = getHeaderMap_(sh);
  
  // Формат суммы
  if (map['Сумма']) {
    sh.getRange(2, map['Сумма'], lastRow - 1, 1).setNumberFormat('#,##0.00');
  }
  
  // Выравнивание
  if (map['payment_id']) {
    sh.getRange(2, map['payment_id'], lastRow - 1, 1).setHorizontalAlignment('center');
  }
}

/**
 * Стили для листа «Участие»
 */
function applyParticipationStyles_(sh, lastRow) {
  const map = getHeaderMap_(sh);
  const rules = [];
  
  const statusCol = map['Статус'];
  if (statusCol) {
    const statusRange = sh.getRange(2, statusCol, lastRow - 1, 1);
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(PARTICIPATION_STATUS.PARTICIPATES)
      .setBackground(STYLE_COLORS.STATUS_YES)
      .setRanges([statusRange])
      .build());
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(PARTICIPATION_STATUS.NOT_PARTICIPATES)
      .setBackground(STYLE_COLORS.STATUS_NO)
      .setRanges([statusRange])
      .build());
  }
  
  sh.setConditionalFormatRules(rules);
}

/**
 * Стили для листа «Баланс»
 */
function applyBalanceStyles_(sh, lastRow) {
  const map = getHeaderMap_(sh);
  const rules = [];
  
  // Числовые форматы
  ['Внесено всего', 'Списано всего', 'Зарезервировано', 'Свободный остаток', 'Задолженность'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('#,##0.00');
    }
  });
  
  // Условное форматирование: Задолженность > 0
  const debtCol = map['Задолженность'];
  if (debtCol) {
    const debtRange = sh.getRange(2, debtCol, lastRow - 1, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(STYLE_COLORS.NEGATIVE)
      .setRanges([debtRange])
      .build());
  }
  
  // Условное форматирование: Свободный остаток < 0
  const freeCol = map['Свободный остаток'];
  if (freeCol) {
    const freeRange = sh.getRange(2, freeCol, lastRow - 1, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(STYLE_COLORS.NEGATIVE)
      .setRanges([freeRange])
      .build());
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(STYLE_COLORS.POSITIVE)
      .setRanges([freeRange])
      .build());
  }
  
  sh.setConditionalFormatRules(rules);
}

/**
 * Стили для листа «Детализация»
 */
function applyDetailStyles_(sh, lastRow) {
  const map = getHeaderMap_(sh);
  const rules = [];
  
  ['Оплачено', 'Начислено', 'Разность (±)'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('#,##0.00');
    }
  });
  
  // Условное форматирование: Разность
  const diffCol = map['Разность (±)'];
  if (diffCol) {
    const diffRange = sh.getRange(2, diffCol, lastRow - 1, 1);
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(STYLE_COLORS.POSITIVE)
      .setRanges([diffRange])
      .build());
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground(STYLE_COLORS.NEGATIVE)
      .setRanges([diffRange])
      .build());
  }
  
  sh.setConditionalFormatRules(rules);
}

/**
 * Стили для листа «Сводка»
 */
function applySummaryStyles_(sh, lastRow) {
  const map = getHeaderMap_(sh);
  
  ['Сумма цели', 'Собрано', 'Остаток до цели', 'Переплата'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('#,##0.00');
    }
  });
}

/**
 * Стили для листа «Выдача»
 */
function applyIssuesStyles_(sh, lastRow) {
  const map = getHeaderMap_(sh);
  const rules = [];
  
  // Условное форматирование: Выдано
  const issuedCol = map['Выдано'];
  if (issuedCol) {
    const issuedRange = sh.getRange(2, issuedCol, lastRow - 1, 1);
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(ACTIVE_STATUS.YES)
      .setBackground(STYLE_COLORS.STATUS_YES)
      .setRanges([issuedRange])
      .build());
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(ACTIVE_STATUS.NO)
      .setBackground(STYLE_COLORS.STATUS_NO)
      .setRanges([issuedRange])
      .build());
  }
  
  sh.setConditionalFormatRules(rules);
}

/**
 * Стили для листа «Статус выдачи»
 */
function applyIssueStatusStyles_(sh, lastRow) {
  const map = getHeaderMap_(sh);
  const rules = [];
  
  if (map['x (цена)']) {
    sh.getRange(2, map['x (цена)'], lastRow - 1, 1).setNumberFormat('#,##0.00');
  }
  
  // Условное форматирование: Остаток > 0
  const remainCol = map['Остаток (шт)'];
  if (remainCol) {
    const remainRange = sh.getRange(2, remainCol, lastRow - 1, 1);
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground(STYLE_COLORS.NEUTRAL)
      .setRanges([remainRange])
      .build());
  }
  
  sh.setConditionalFormatRules(rules);
}

// ============================================================================
// Вспомогательные функции
// ============================================================================

/**
 * Применяет автофильтр к листу
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function applyAutoFilter_(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  
  if (lastRow < 1 || lastCol < 1) return;
  
  // Удаляем существующий фильтр
  const existingFilter = sh.getFilter();
  if (existingFilter) existingFilter.remove();
  
  // Создаём новый
  try {
    sh.getRange(1, 1, lastRow, lastCol).createFilter();
  } catch (e) {
    // Фильтр уже существует или ошибка
    Logger.log(`Filter error for ${sh.getName()}: ${e.message}`);
  }
}

/**
 * Быстрая починка стилей всех листов (без диалога)
 */
function quickFixAllStyles() {
  const ss = SpreadsheetApp.getActive();
  const specs = getSheetsSpec();
  const version = detectVersion();
  
  specs.forEach(spec => {
    if (version === 'v2' && spec.name === SHEET_NAMES.COLLECTIONS) return;
    if (version === 'v1' && spec.name === SHEET_NAMES.GOALS) return;
    
    const sh = ss.getSheetByName(spec.name);
    if (sh) {
      fixSheetStyles_(sh, spec);
    }
  });
  
  addHeaderNotes_();
  SpreadsheetApp.getActive().toast('Стили всех листов обновлены.', 'Funds');
}

/**
 * Сбрасывает все стили листа к базовым
 * Точка входа из меню
 */
function resetCurrentSheetStyles() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  
  clearAllStyles_(sh);
  SpreadsheetApp.getActive().toast(`Стили листа «${sh.getName()}» сброшены.`, 'Funds');
}

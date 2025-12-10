/**
 * @fileoverview Стили и оформление листов
 */

/**
 * Применяет стили ко всей книге
 */
function styleWorkbook_() {
  const ss = SpreadsheetApp.getActive();
  const sheets = getSheetsSpec();
  
  sheets.forEach(spec => {
    const sh = ss.getSheetByName(spec.name);
    if (!sh) return;
    
    styleSheetHeader_(sh);
    
    // Применяем специфичные стили для каждого листа
    switch (spec.name) {
      case SHEET_NAMES.BALANCE:
        styleBalanceSheet_(sh);
        break;
      case SHEET_NAMES.PAYMENTS:
        stylePaymentsSheet_(sh);
        break;
      case SHEET_NAMES.GOALS:
        styleGoalsSheet_(sh);
        break;
      case SHEET_NAMES.COLLECTIONS:
        styleCollectionsSheet_(sh);
        break;
      case SHEET_NAMES.FAMILIES:
        styleFamiliesSheet_(sh);
        break;
      case SHEET_NAMES.PARTICIPATION:
        styleParticipationSheet_(sh);
        break;
      case SHEET_NAMES.DETAIL:
        styleDetailSheet_(sh);
        break;
      case SHEET_NAMES.SUMMARY:
        styleSummarySheet_(sh);
        break;
      case SHEET_NAMES.ISSUE_STATUS:
        styleIssueStatusSheet_(sh);
        break;
    }
  });
}

/**
 * Стилизует заголовок листа (первая строка)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleSheetHeader_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return;
  
  const headerRange = sh.getRange(1, 1, 1, lastCol);
  headerRange
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  sh.setFrozenRows(1);
}

/**
 * Стилизует лист «Баланс»
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleBalanceSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  const lastCol = sh.getLastColumn();
  const dataRange = sh.getRange(2, 1, lastRow - 1, lastCol);
  
  // Чередующиеся строки (zebra)
  applyZebraStripes_(sh, 2, lastRow);
  
  // Числовой формат для денежных колонок (C:G)
  if (lastCol >= 7) {
    sh.getRange(2, 3, lastRow - 1, 5).setNumberFormat('#,##0.00');
  }
  
  // Условное форматирование: Задолженность > 0 — красный фон
  const rules = sh.getConditionalFormatRules();
  const debtCol = 7; // Задолженность
  const debtRange = sh.getRange(2, debtCol, lastRow - 1, 1);
  const debtRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#ffcdd2')
    .setRanges([debtRange])
    .build();
  rules.push(debtRule);
  sh.setConditionalFormatRules(rules);
  
  // Авто-фильтр
  try { sh.getRange(1, 1, lastRow, lastCol).createFilter(); } catch (_) {}
}

/**
 * Стилизует лист «Платежи»
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function stylePaymentsSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  
  // Формат даты
  if (map['Дата']) {
    sh.getRange(2, map['Дата'], lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
  }
  
  // Формат суммы
  if (map['Сумма']) {
    sh.getRange(2, map['Сумма'], lastRow - 1, 1).setNumberFormat('#,##0.00');
  }
  
  // Авто-фильтр
  try { sh.getDataRange().createFilter(); } catch (_) {}
}

/**
 * Стилизует лист «Цели» (v2.0)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleGoalsSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  
  // Форматы дат
  ['Дата начала', 'Дедлайн'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });
  
  // Форматы чисел
  ['Параметр суммы', 'Фиксированный x'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('#,##0.00');
    }
  });
  
  // Условное форматирование: Статус
  const statusCol = map['Статус'];
  if (statusCol) {
    const rules = sh.getConditionalFormatRules();
    const statusRange = sh.getRange(2, statusCol, lastRow - 1, 1);
    
    // Открыта — зелёный
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(GOAL_STATUS.OPEN)
      .setBackground('#c8e6c9')
      .setRanges([statusRange])
      .build());
    
    // Закрыта — серый
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(GOAL_STATUS.CLOSED)
      .setBackground('#e0e0e0')
      .setRanges([statusRange])
      .build());
    
    // Отменена — красный
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(GOAL_STATUS.CANCELLED)
      .setBackground('#ffcdd2')
      .setRanges([statusRange])
      .build());
    
    sh.setConditionalFormatRules(rules);
  }
  
  try { sh.getDataRange().createFilter(); } catch (_) {}
}

/**
 * Стилизует лист «Сборы» (v1.x legacy)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleCollectionsSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  
  ['Дата начала', 'Дедлайн'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });
  
  ['Параметр суммы', 'Фиксированный x'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('#,##0.00');
    }
  });
  
  const statusCol = map['Статус'];
  if (statusCol) {
    const rules = sh.getConditionalFormatRules();
    const statusRange = sh.getRange(2, statusCol, lastRow - 1, 1);
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(COLLECTION_STATUS_V1.OPEN)
      .setBackground('#c8e6c9')
      .setRanges([statusRange])
      .build());
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(COLLECTION_STATUS_V1.CLOSED)
      .setBackground('#e0e0e0')
      .setRanges([statusRange])
      .build());
    
    sh.setConditionalFormatRules(rules);
  }
  
  try { sh.getDataRange().createFilter(); } catch (_) {}
}

/**
 * Стилизует лист «Семьи»
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleFamiliesSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  
  if (map['День рождения']) {
    sh.getRange(2, map['День рождения'], lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
  }
  
  // Условное форматирование: неактивные — серый фон
  const activeCol = map['Активен'];
  if (activeCol) {
    const rules = sh.getConditionalFormatRules();
    const lastCol = sh.getLastColumn();
    const rowRange = sh.getRange(2, 1, lastRow - 1, lastCol);
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${colToLetter_(activeCol)}2="Нет"`)
      .setBackground('#f5f5f5')
      .setFontColor('#9e9e9e')
      .setRanges([rowRange])
      .build());
    
    sh.setConditionalFormatRules(rules);
  }
  
  try { sh.getDataRange().createFilter(); } catch (_) {}
}

/**
 * Стилизует лист «Участие»
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleParticipationSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  const statusCol = map['Статус'];
  
  if (statusCol) {
    const rules = sh.getConditionalFormatRules();
    const statusRange = sh.getRange(2, statusCol, lastRow - 1, 1);
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(PARTICIPATION_STATUS.PARTICIPATES)
      .setBackground('#c8e6c9')
      .setRanges([statusRange])
      .build());
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(PARTICIPATION_STATUS.NOT_PARTICIPATES)
      .setBackground('#ffcdd2')
      .setRanges([statusRange])
      .build());
    
    sh.setConditionalFormatRules(rules);
  }
  
  try { sh.getDataRange().createFilter(); } catch (_) {}
}

/**
 * Стилизует лист «Детализация»
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleDetailSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  
  ['Оплачено', 'Начислено', 'Разность (±)'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('#,##0.00');
    }
  });
  
  // Условное форматирование: Разность
  const diffCol = map['Разность (±)'];
  if (diffCol) {
    const rules = sh.getConditionalFormatRules();
    const diffRange = sh.getRange(2, diffCol, lastRow - 1, 1);
    
    // Положительная разность (переплата) — зелёный
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground('#c8e6c9')
      .setRanges([diffRange])
      .build());
    
    // Отрицательная разность (недоплата) — красный
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#ffcdd2')
      .setRanges([diffRange])
      .build());
    
    sh.setConditionalFormatRules(rules);
  }
  
  try { sh.getDataRange().createFilter(); } catch (_) {}
}

/**
 * Стилизует лист «Сводка»
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleSummarySheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  
  ['Сумма цели', 'Собрано', 'Остаток до цели'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('#,##0.00');
    }
  });
}

/**
 * Стилизует лист «Статус выдачи»
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleIssueStatusSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  
  if (map['x (цена)']) {
    sh.getRange(2, map['x (цена)'], lastRow - 1, 1).setNumberFormat('#,##0.00');
  }
}

/**
 * Применяет чередующиеся цвета строк (zebra stripes)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {number} startRow
 * @param {number} endRow
 */
function applyZebraStripes_(sh, startRow, endRow) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1 || endRow < startRow) return;
  
  const dataRange = sh.getRange(startRow, 1, endRow - startRow + 1, lastCol);
  
  // Удаляем старое banding
  const bandings = sh.getBandings();
  bandings.forEach(b => b.remove());
  
  // Применяем новое
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
}

/**
 * Очищает книгу: удаляет пустые строки/колонки, обновляет стили
 */
function cleanupWorkbook_() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  
  sheets.forEach(sh => {
    // Не трогаем скрытый Lists
    if (sh.getName() === SHEET_NAMES.LISTS) return;
    
    // Удаляем лишние пустые строки
    const lastRow = sh.getLastRow();
    const maxRows = sh.getMaxRows();
    if (maxRows > lastRow + 10) {
      sh.deleteRows(lastRow + 11, maxRows - lastRow - 10);
    }
    
    // Удаляем лишние пустые колонки
    const lastCol = sh.getLastColumn();
    const maxCols = sh.getMaxColumns();
    if (maxCols > lastCol + 3) {
      sh.deleteColumns(lastCol + 4, maxCols - lastCol - 3);
    }
  });
  
  styleWorkbook_();
  SpreadsheetApp.getActive().toast('Cleanup complete.', 'Funds');
}

/**
 * Добавляет примечания к заголовкам листов
 */
function addHeaderNotes_() {
  const ss = SpreadsheetApp.getActive();
  
  // Примечания для листа «Цели»
  const goalsNotes = {
    'Название цели': 'Название цели/сбора',
    'Тип': 'разовая / регулярная',
    'Статус': 'Открыта / Закрыта / Отменена',
    'Начисление': 'Режим начисления:\n• static_per_family\n• shared_total_all\n• shared_total_by_payers\n• dynamic_by_payers\n• proportional_by_payers\n• unit_price\n• voluntary',
    'Параметр суммы': 'T — сумма цели или ставка на семью',
    'Фиксированный x': 'Для dynamic_by_payers — cap после закрытия.\nДля unit_price — цена единицы.',
    'К выдаче детям': 'Да — включает учёт выдачи для unit_price',
    'Периодичность': 'Для регулярных целей: ежемесячно / ежеквартально / ежегодно',
    'goal_id': 'ID цели (G001, G002, ...)'
  };
  
  setHeaderNotes_(ss.getSheetByName(SHEET_NAMES.GOALS), goalsNotes);
  
  // Примечания для листа «Семьи»
  const familyNotes = {
    'Активен': 'Да — семья участвует по умолчанию во всех целях',
    'family_id': 'ID семьи (F001, F002, ...)'
  };
  
  setHeaderNotes_(ss.getSheetByName(SHEET_NAMES.FAMILIES), familyNotes);
  
  // Примечания для листа «Участие»
  const partNotes = {
    'Статус': 'Участвует / Не участвует.\nЕсли есть хотя бы один «Участвует» — участвуют только отмеченные.\n«Не участвует» всегда исключает.'
  };
  
  setHeaderNotes_(ss.getSheetByName(SHEET_NAMES.PARTICIPATION), partNotes);
  
  // Примечания для листа «Платежи»
  const payNotes = {
    'Дата': 'Справочная дата (не влияет на расчёты)',
    'goal_id (label)': 'Выберите цель из выпадающего списка.\nПустое — свободный платёж (v2.0)',
    'Сумма': 'Сумма платежа > 0',
    'payment_id': 'ID платежа (PMT001, PMT002, ...)'
  };
  
  setHeaderNotes_(ss.getSheetByName(SHEET_NAMES.PAYMENTS), payNotes);
  
  // Примечания для листа «Баланс»
  const balanceNotes = {
    'Внесено всего': 'Сумма всех платежей семьи',
    'Списано всего': 'Сумма начислений по всем целям',
    'Зарезервировано': 'Зарезервировано под открытые цели',
    'Свободный остаток': 'Внесено - Списано - Резерв',
    'Задолженность': 'max(0, Списано - Внесено)'
  };
  
  setHeaderNotes_(ss.getSheetByName(SHEET_NAMES.BALANCE), balanceNotes);
}

/**
 * Устанавливает примечания к заголовкам листа
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {Object<string, string>} notes — {headerName: noteText}
 */
function setHeaderNotes_(sh, notes) {
  if (!sh) return;
  
  const map = getHeaderMap_(sh);
  
  Object.entries(notes).forEach(([header, note]) => {
    const col = map[header];
    if (col) {
      sh.getRange(1, col).setNote(note);
    }
  });
}

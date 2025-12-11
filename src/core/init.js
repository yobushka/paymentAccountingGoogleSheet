/**
 * @fileoverview Инициализация структуры таблицы
 */

/**
 * Инициализирует или пересоздаёт структуру таблицы
 * Точка входа из меню
 */
function init() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  
  // Определяем версию
  const version = detectVersion();
  
  if (version === 'v1') {
    // Предлагаем миграцию
    const response = ui.alert(
      'Обнаружена версия 1.x',
      'Таблица использует старую версию (Сборы/collection_id).\n\n' +
      'Выполнить миграцию на v2.0 (Цели/goal_id)?\n\n' +
      'Нажмите «Да» для миграции или «Нет» для работы в режиме совместимости.',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (response === ui.Button.YES) {
      migrateToV2();
      return;
    } else if (response === ui.Button.CANCEL) {
      return;
    }
    // Продолжаем с v1 структурой
    initV1Structure_(ss);
  } else if (version === 'v2') {
    // Обновляем v2 структуру
    initV2Structure_(ss);
  } else {
    // Новая таблица — создаём v2
    initV2Structure_(ss);
  }
  
  // Общие настройки
  setupListsSheet();
  setupNamedRanges_();
  rebuildValidations();
  setupBalanceExamples();
  addHeaderNotes_();
  styleWorkbook_();
  
  SpreadsheetApp.getActive().toast('Structure initialized.', 'Funds');
}

/**
 * Инициализирует структуру для v2.0
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function initV2Structure_(ss) {
  const specs = getSheetsSpec();
  
  specs.forEach(spec => {
    // Пропускаем legacy лист «Сборы» в v2
    if (spec.name === SHEET_NAMES.COLLECTIONS) return;
    
    const sh = getOrCreateSheet(ss, spec.name);
    
    // Расширяем лист если нужно больше столбцов
    const requiredCols = spec.headers.length;
    const currentCols = sh.getMaxColumns();
    if (currentCols < requiredCols) {
      sh.insertColumnsAfter(currentCols, requiredCols - currentCols);
    }
    
    // Заголовки — всегда перезаписываем для соответствия спецификации
    const headerRange = sh.getRange(1, 1, 1, spec.headers.length);
    headerRange.setValues([spec.headers]);
    
    // Ширины колонок
    spec.colWidths.forEach((w, i) => {
      if (w) sh.setColumnWidth(i + 1, w);
    });
    
    // Форматы дат
    if (spec.dateCols) {
      spec.dateCols.forEach(col => {
        sh.getRange(2, col, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
      });
    }
    
    sh.setFrozenRows(1);
  });
  
  // Инструкция для v2.0
  setupInstructionSheet();
}

/**
 * Инициализирует структуру для v1.x (режим совместимости)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function initV1Structure_(ss) {
  const specs = getSheetsSpecV1();
  
  specs.forEach(spec => {
    // Пропускаем новый лист «Цели» в v1
    if (spec.name === SHEET_NAMES.GOALS) return;
    
    const sh = getOrCreateSheet(ss, spec.name);
    
    // Расширяем лист если нужно больше столбцов
    const requiredCols = spec.headers.length;
    const currentCols = sh.getMaxColumns();
    if (currentCols < requiredCols) {
      sh.insertColumnsAfter(currentCols, requiredCols - currentCols);
    }
    
    // Заголовки — всегда перезаписываем для соответствия спецификации
    const headerRange = sh.getRange(1, 1, 1, spec.headers.length);
    headerRange.setValues([spec.headers]);
    
    spec.colWidths.forEach((w, i) => {
      if (w) sh.setColumnWidth(i + 1, w);
    });
    
    if (spec.dateCols) {
      spec.dateCols.forEach(col => {
        sh.getRange(2, col, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
      });
    }
    
    sh.setFrozenRows(1);
  });
  
  setupInstructionSheetV1_();
}

/**
 * Настраивает именованные диапазоны
 */
function setupNamedRanges_() {
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  
  // Диапазоны для меток (labels) — всегда из Lists
  ensureNamedRange(NAMED_RANGES.FAMILIES_LABELS, 'Lists!D2:D');
  ensureNamedRange(NAMED_RANGES.ACTIVE_FAMILIES_LABELS, 'Lists!C2:C');
  
  if (version === 'v2' || version === 'new') {
    ensureNamedRange(NAMED_RANGES.GOALS_LABELS, 'Lists!B2:B');
    ensureNamedRange(NAMED_RANGES.OPEN_GOALS_LABELS, 'Lists!A2:A');
  } else {
    // v1.x
    ensureNamedRange(NAMED_RANGES.COLLECTIONS_LABELS, 'Lists!B2:B');
    ensureNamedRange(NAMED_RANGES.OPEN_COLLECTIONS_LABELS, 'Lists!A2:A');
  }
  
  // Raw ID диапазоны
  setRawIdNamedRanges_();
}

/**
 * Устанавливает именованные диапазоны для raw ID
 */
function setRawIdNamedRanges_() {
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  
  // Семьи
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  if (shF) {
    const mapF = getHeaderMap_(shF);
    const fIdCol = colToLetter_(mapF['family_id'] || 1);
    ensureNamedRange('FAMILIES', `Семьи!${fIdCol}2:${fIdCol}`);
  }
  
  // Цели или Сборы
  if (version === 'v2' || version === 'new') {
    const shG = ss.getSheetByName(SHEET_NAMES.GOALS);
    if (shG) {
      const mapG = getHeaderMap_(shG);
      const gIdCol = colToLetter_(mapG['goal_id'] || 1);
      ensureNamedRange('GOALS', `Цели!${gIdCol}2:${gIdCol}`);
    }
  } else {
    const shC = ss.getSheetByName(SHEET_NAMES.COLLECTIONS);
    if (shC) {
      const mapC = getHeaderMap_(shC);
      const cIdCol = colToLetter_(mapC['collection_id'] || 1);
      ensureNamedRange('COLLECTIONS', `Сборы!${cIdCol}2:${cIdCol}`);
    }
  }
}

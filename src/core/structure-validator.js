/**
 * @fileoverview Валидация и исправление структуры листов
 * @version 2.0
 */

/**
 * Результат валидации листа
 * @typedef {Object} SheetValidationResult
 * @property {string} sheetName — название листа
 * @property {boolean} valid — структура корректна
 * @property {string[]} missing — отсутствующие колонки
 * @property {string[]} extra — лишние колонки
 * @property {string[]} duplicates — дублирующиеся колонки
 * @property {boolean} orderCorrect — порядок колонок верный
 * @property {string[]} messages — сообщения о проблемах
 */

/**
 * Валидирует структуру всех листов
 * @returns {SheetValidationResult[]}
 */
function validateAllSheets() {
  const ss = SpreadsheetApp.getActive();
  const specs = getSheetsSpec();
  const version = detectVersion();
  const results = [];

  specs.forEach(spec => {
    // Пропускаем legacy лист в v2 и наоборот
    if (version === 'v2' && spec.name === SHEET_NAMES.COLLECTIONS) return;
    if (version === 'v1' && spec.name === SHEET_NAMES.GOALS) return;

    const result = validateSheet_(ss, spec);
    results.push(result);
  });

  return results;
}

/**
 * Валидирует структуру одного листа
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {{name: string, headers: string[]}} spec
 * @returns {SheetValidationResult}
 */
function validateSheet_(ss, spec) {
  const result = {
    sheetName: spec.name,
    valid: true,
    missing: [],
    extra: [],
    duplicates: [],
    orderCorrect: true,
    messages: []
  };

  const sh = ss.getSheetByName(spec.name);
  if (!sh) {
    result.valid = false;
    result.messages.push(`Лист «${spec.name}» не найден`);
    return result;
  }

  // Получаем текущие заголовки
  const lastCol = sh.getLastColumn();
  if (lastCol === 0) {
    result.valid = false;
    result.messages.push(`Лист «${spec.name}» пустой`);
    return result;
  }

  const currentHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  const expectedHeaders = spec.headers;

  // Проверяем дубликаты
  const headerCounts = {};
  currentHeaders.forEach(h => {
    if (h) {
      headerCounts[h] = (headerCounts[h] || 0) + 1;
    }
  });
  Object.entries(headerCounts).forEach(([h, count]) => {
    if (count > 1) {
      result.duplicates.push(`${h} (×${count})`);
      result.valid = false;
    }
  });

  // Проверяем отсутствующие колонки
  const currentSet = new Set(currentHeaders);
  expectedHeaders.forEach(h => {
    if (!currentSet.has(h)) {
      result.missing.push(h);
      result.valid = false;
    }
  });

  // Проверяем лишние колонки (кроме пустых)
  const expectedSet = new Set(expectedHeaders);
  currentHeaders.forEach(h => {
    if (h && !expectedSet.has(h)) {
      result.extra.push(h);
      // Лишние колонки — предупреждение, не ошибка
    }
  });

  // Проверяем порядок колонок
  const currentFiltered = currentHeaders.filter(h => expectedSet.has(h));
  const expectedFiltered = expectedHeaders.filter(h => currentSet.has(h));
  if (currentFiltered.join('|') !== expectedFiltered.join('|')) {
    result.orderCorrect = false;
    result.valid = false;
  }

  // Формируем сообщения
  if (result.duplicates.length) {
    result.messages.push(`Дубликаты: ${result.duplicates.join(', ')}`);
  }
  if (result.missing.length) {
    result.messages.push(`Отсутствуют: ${result.missing.join(', ')}`);
  }
  if (result.extra.length) {
    result.messages.push(`Лишние колонки: ${result.extra.join(', ')}`);
  }
  if (!result.orderCorrect) {
    result.messages.push('Неверный порядок колонок');
  }
  if (result.valid && result.extra.length === 0) {
    result.messages.push('✓ Структура корректна');
  } else if (result.valid) {
    result.messages.push('✓ Структура корректна (есть дополнительные колонки)');
  }

  return result;
}

/**
 * Показывает отчёт о валидации структуры
 * Точка входа из меню
 */
function showStructureReport() {
  const results = validateAllSheets();
  const ui = SpreadsheetApp.getUi();

  let report = 'ОТЧЁТ О СТРУКТУРЕ ЛИСТОВ\n';
  report += '═'.repeat(40) + '\n\n';

  let hasErrors = false;
  results.forEach(r => {
    const status = r.valid ? '✓' : '✗';
    report += `${status} ${r.sheetName}\n`;
    r.messages.forEach(m => {
      report += `   ${m}\n`;
    });
    report += '\n';
    if (!r.valid) hasErrors = true;
  });

  if (hasErrors) {
    report += '─'.repeat(40) + '\n';
    report += 'Для исправления: Funds → Fix Sheet Structure\n';
  }

  ui.alert('Валидация структуры', report, ui.ButtonSet.OK);
}

/**
 * Исправляет структуру всех листов
 * Точка входа из меню
 */
function fixAllSheetsStructure() {
  const ui = SpreadsheetApp.getUi();
  const results = validateAllSheets();
  
  const sheetsToFix = results.filter(r => !r.valid || r.extra.length > 0);
  
  if (sheetsToFix.length === 0) {
    ui.alert('Структура в порядке', 'Все листы соответствуют спецификации.', ui.ButtonSet.OK);
    return;
  }

  // Показываем что будет исправлено
  let preview = 'Будут исправлены следующие листы:\n\n';
  sheetsToFix.forEach(r => {
    preview += `• ${r.sheetName}: ${r.messages.join('; ')}\n`;
  });
  preview += '\nПродолжить?';

  const response = ui.alert('Исправление структуры', preview, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  // Исправляем
  const ss = SpreadsheetApp.getActive();
  const specs = getSheetsSpec();
  let fixed = 0;

  sheetsToFix.forEach(result => {
    const spec = specs.find(s => s.name === result.sheetName);
    if (spec) {
      try {
        fixSheetStructure_(ss, spec, result);
        fixed++;
      } catch (e) {
        Logger.log(`Error fixing ${result.sheetName}: ${e.message}`);
      }
    }
  });

  // Обновляем валидации после исправления
  rebuildValidations();

  ui.alert(
    'Исправление завершено',
    `Исправлено листов: ${fixed}\n\nВыполните Funds → Recalculate для обновления формул.`,
    ui.ButtonSet.OK
  );
}

/**
 * Исправляет структуру одного листа
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {{name: string, headers: string[], colWidths: number[], dateCols?: number[]}} spec
 * @param {SheetValidationResult} validation
 */
function fixSheetStructure_(ss, spec, validation) {
  const sh = ss.getSheetByName(spec.name);
  if (!sh) {
    // Создаём лист если не существует
    const newSheet = ss.insertSheet(spec.name);
    newSheet.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
    spec.colWidths.forEach((w, i) => newSheet.setColumnWidth(i + 1, w));
    newSheet.setFrozenRows(1);
    Logger.log(`Created sheet: ${spec.name}`);
    return;
  }

  const lastCol = sh.getLastColumn();
  const lastRow = sh.getLastRow();
  
  if (lastCol === 0 || lastRow === 0) {
    // Пустой лист — просто добавляем заголовки
    sh.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
    spec.colWidths.forEach((w, i) => sh.setColumnWidth(i + 1, w));
    sh.setFrozenRows(1);
    return;
  }

  const currentHeaders = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  
  Logger.log(`Fixing ${spec.name}: current=[${currentHeaders.join(',')}], expected=[${spec.headers.join(',')}]`);

  // Стратегия: создаём карту данных по заголовкам, потом перестраиваем
  const dataMap = buildDataMap_(sh, currentHeaders, lastRow);
  
  // Очищаем лист
  sh.clear();
  
  // Записываем новые заголовки
  sh.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
  
  // Записываем данные в новом порядке
  if (lastRow > 1) {
    const newData = rebuildData_(dataMap, spec.headers, lastRow - 1);
    if (newData.length > 0) {
      sh.getRange(2, 1, newData.length, spec.headers.length).setValues(newData);
    }
  }

  // Устанавливаем ширины колонок
  spec.colWidths.forEach((w, i) => {
    if (w) sh.setColumnWidth(i + 1, w);
  });

  // Применяем форматы по именам заголовков (надёжнее, чем по номерам)
  applySheetFormats_(sh, spec, lastRow);

  sh.setFrozenRows(1);
  Logger.log(`Fixed structure for: ${spec.name}`);
}

/**
 * Применяет форматы ячеек по именам заголовков
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {{name: string, headers: string[]}} spec
 * @param {number} lastRow
 */
function applySheetFormats_(sh, spec, lastRow) {
  if (lastRow < 2) return;
  
  const map = getHeaderMap_(sh);
  const rows = lastRow - 1;
  
  // Определяем форматы по именам колонок
  const dateColumns = [
    'День рождения', 'Членство с', 'Членство по',
    'Дата начала', 'Дедлайн', 'Дата чека',
    'Дата', 'Дата выдачи'
  ];
  
  const numberColumns = [
    'Параметр суммы', 'Фиксированный x', 'К выдаче детям', 'Возмещено',
    'Сумма', 'Единиц',
    'Внесено всего', 'Списано всего', 'Зарезервировано', 'Свободный остаток', 'Задолженность',
    'Оплачено', 'Начислено', 'Разность (±)',
    'Сумма цели', 'Собрано', 'Остаток до цели', 'Переплата',
    'x (цена)', 'Единиц требуется', 'Единиц оплачено', 'Единиц выдано', 'Остаток (шт)',
    'Значение', 'Итого'
  ];
  
  const textColumns = [
    'Статья', 'Подстатья', 'Начисление', 'Статус', 'Тип', 'Периодичность',
    'Способ', 'Комментарий', 'Кто выдал'
  ];
  
  // Применяем форматы дат
  dateColumns.forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], rows, 1).setNumberFormat('yyyy-mm-dd');
    }
  });
  
  // Применяем числовые форматы
  numberColumns.forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], rows, 1).setNumberFormat('#,##0.00');
    }
  });
  
  // Применяем текстовые форматы (сброс ошибочных форматов)
  textColumns.forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], rows, 1).setNumberFormat('@');
    }
  });
}

/**
 * Строит карту данных по заголовкам
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {string[]} headers
 * @param {number} lastRow
 * @returns {Map<string, any[]>}
 */
function buildDataMap_(sh, headers, lastRow) {
  const map = new Map();
  
  if (lastRow <= 1) return map;
  
  const data = sh.getRange(2, 1, lastRow - 1, headers.length).getValues();
  
  // Обрабатываем дубликаты: берём первую колонку с данными
  const usedHeaders = new Set();
  
  headers.forEach((header, colIndex) => {
    if (!header) return;
    
    // Если заголовок уже использован, пропускаем (дубликат)
    if (usedHeaders.has(header)) {
      Logger.log(`Skipping duplicate column: ${header} at index ${colIndex}`);
      return;
    }
    
    // Проверяем, есть ли данные в этой колонке
    const columnData = data.map(row => row[colIndex]);
    const hasData = columnData.some(cell => cell !== '' && cell !== null && cell !== undefined);
    
    // Если для этого заголовка ещё нет данных или текущая колонка имеет данные
    if (!map.has(header) || hasData) {
      map.set(header, columnData);
    }
    
    usedHeaders.add(header);
  });
  
  return map;
}

/**
 * Перестраивает данные согласно новому порядку заголовков
 * @param {Map<string, any[]>} dataMap
 * @param {string[]} newHeaders
 * @param {number} rowCount
 * @returns {any[][]}
 */
function rebuildData_(dataMap, newHeaders, rowCount) {
  const result = [];
  
  for (let i = 0; i < rowCount; i++) {
    const row = newHeaders.map(header => {
      const colData = dataMap.get(header);
      return colData ? (colData[i] ?? '') : '';
    });
    
    // Пропускаем полностью пустые строки
    if (row.some(cell => cell !== '' && cell !== null && cell !== undefined)) {
      result.push(row);
    }
  }
  
  return result;
}

/**
 * Исправляет структуру конкретного листа по имени
 * @param {string} sheetName
 */
function fixSheetByName(sheetName) {
  const ss = SpreadsheetApp.getActive();
  const specs = getSheetsSpec();
  const spec = specs.find(s => s.name === sheetName);
  
  if (!spec) {
    throw new Error(`Спецификация для листа «${sheetName}» не найдена`);
  }
  
  const validation = validateSheet_(ss, spec);
  fixSheetStructure_(ss, spec, validation);
  
  SpreadsheetApp.getActive().toast(`Структура листа «${sheetName}» исправлена.`, 'Funds');
}

/**
 * Диалог исправления конкретного листа
 * Точка входа из меню
 */
function fixSheetStructurePrompt() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActive();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();
  
  const specs = getSheetsSpec();
  const spec = specs.find(s => s.name === sheetName);
  
  if (!spec) {
    ui.alert(
      'Лист не распознан',
      `Лист «${sheetName}» не является системным листом.\n\nСистемные листы: ${specs.map(s => s.name).join(', ')}`,
      ui.ButtonSet.OK
    );
    return;
  }
  
  const validation = validateSheet_(ss, spec);
  
  if (validation.valid && validation.extra.length === 0) {
    ui.alert('Структура в порядке', `Лист «${sheetName}» соответствует спецификации.`, ui.ButtonSet.OK);
    return;
  }
  
  let msg = `Лист «${sheetName}»:\n\n`;
  validation.messages.forEach(m => msg += `• ${m}\n`);
  msg += '\nИсправить структуру?';
  
  const response = ui.alert('Исправление структуры', msg, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;
  
  try {
    fixSheetStructure_(ss, spec, validation);
    rebuildValidations();
    ui.alert('Готово', `Структура листа «${sheetName}» исправлена.`, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Ошибка', e.message, ui.ButtonSet.OK);
  }
}

/**
 * Быстрая проверка и исправление текущего листа
 * Точка входа из меню
 */
function quickFixCurrentSheet() {
  const ss = SpreadsheetApp.getActive();
  const activeSheet = ss.getActiveSheet();
  const sheetName = activeSheet.getName();
  
  const specs = getSheetsSpec();
  const spec = specs.find(s => s.name === sheetName);
  
  if (!spec) {
    SpreadsheetApp.getActive().toast(`Лист «${sheetName}» не является системным.`, 'Funds');
    return;
  }
  
  const validation = validateSheet_(ss, spec);
  
  if (validation.valid && validation.extra.length === 0) {
    SpreadsheetApp.getActive().toast(`Лист «${sheetName}» в порядке.`, 'Funds');
    return;
  }
  
  fixSheetStructure_(ss, spec, validation);
  SpreadsheetApp.getActive().toast(`Лист «${sheetName}» исправлен.`, 'Funds');
}

/**
 * Обновляет только заголовки всех листов согласно спецификации
 * Точка входа из меню
 */
function refreshAllHeaders() {
  const ss = SpreadsheetApp.getActive();
  const specs = getSheetsSpec();
  const version = detectVersion();
  let updated = 0;
  
  specs.forEach(spec => {
    // Пропускаем legacy/новый лист в зависимости от версии
    if (version === 'v2' && spec.name === SHEET_NAMES.COLLECTIONS) return;
    if (version === 'v1' && spec.name === SHEET_NAMES.GOALS) return;
    
    const sh = ss.getSheetByName(spec.name);
    if (!sh) return;
    
    // Расширяем лист если нужно больше столбцов
    const requiredCols = spec.headers.length;
    const currentCols = sh.getMaxColumns();
    if (currentCols < requiredCols) {
      sh.insertColumnsAfter(currentCols, requiredCols - currentCols);
    }
    
    // Перезаписываем заголовки
    sh.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
    
    // Ширины колонок
    spec.colWidths.forEach((w, i) => {
      if (w) sh.setColumnWidth(i + 1, w);
    });
    
    updated++;
  });
  
  SpreadsheetApp.getActive().toast(`Обновлено заголовков: ${updated} листов.`, 'Funds');
}

/**
 * Обновляет заголовки текущего листа согласно спецификации
 * Точка входа из меню
 */
function refreshCurrentSheetHeaders() {
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
  
  // Расширяем лист если нужно больше столбцов
  const requiredCols = spec.headers.length;
  const currentCols = sh.getMaxColumns();
  if (currentCols < requiredCols) {
    sh.insertColumnsAfter(currentCols, requiredCols - currentCols);
  }
  
  // Перезаписываем заголовки
  sh.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
  
  // Ширины колонок
  spec.colWidths.forEach((w, i) => {
    if (w) sh.setColumnWidth(i + 1, w);
  });
  
  SpreadsheetApp.getActive().toast(`Заголовки листа «${sheetName}» обновлены.`, 'Funds');
}

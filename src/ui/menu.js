/**
 * @fileoverview Меню приложения и обработчик onOpen
 */

/**
 * Создаёт меню при открытии таблицы
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Funds');
  
  // ===== ОСНОВНЫЕ ОПЕРАЦИИ =====
  menu.addItem('⚙️ Настройка структуры', 'init');
  menu.addItem('🔄 Пересчитать баланс', 'recalculateAll');
  menu.addItem('🔒 Закрыть цель', 'closeGoalPrompt');
  
  menu.addSeparator();
  
  // ===== СПРАВКА =====
  menu.addItem('❓ Справка', 'showQuickHelp_');
  menu.addItem('💰 Проверить баланс семьи', 'showQuickBalanceCheck_');
  
  menu.addSeparator();
  
  // ===== ПОДМЕНЮ: ДАННЫЕ =====
  const dataMenu = ui.createMenu('📊 Данные');
  dataMenu.addItem('Сгенерировать ID', 'generateAllIds');
  dataMenu.addItem('Обновить выпадающие списки', 'rebuildValidations');
  dataMenu.addItem('Исправить типы полей', 'auditAndFixFieldTypes');
  dataMenu.addSeparator();
  dataMenu.addItem('Загрузить демо-данные', 'loadSampleDataPrompt');
  menu.addSubMenu(dataMenu);
  
  // ===== ПОДМЕНЮ: СТРУКТУРА =====
  const structureMenu = ui.createMenu('📋 Структура');
  structureMenu.addItem('Проверить структуру листов', 'showStructureReport');
  structureMenu.addItem('Исправить все листы', 'fixAllSheetsStructure');
  structureMenu.addItem('Исправить текущий лист', 'fixSheetStructurePrompt');
  structureMenu.addSeparator();
  structureMenu.addItem('Обновить заголовки (все)', 'refreshAllHeaders');
  structureMenu.addItem('Обновить заголовки (текущий)', 'refreshCurrentSheetHeaders');
  menu.addSubMenu(structureMenu);
  
  // ===== ПОДМЕНЮ: ОФОРМЛЕНИЕ =====
  const stylesMenu = ui.createMenu('🎨 Оформление');
  stylesMenu.addItem('Применить стили (все листы)', 'fixAllSheetsStyles');
  stylesMenu.addItem('Применить стили (текущий)', 'fixCurrentSheetStyles');
  stylesMenu.addItem('Сбросить стили (текущий)', 'resetCurrentSheetStyles');
  stylesMenu.addSeparator();
  stylesMenu.addItem('Обрезать пустые строки/столбцы', 'cleanupWorkbook_');
  menu.addSubMenu(stylesMenu);
  
  // ===== ПОДМЕНЮ: ДИАГНОСТИКА =====
  const diagMenu = ui.createMenu('🔍 Диагностика');
  diagMenu.addItem('Проверить валидации', 'diagnoseValidations_');
  diagMenu.addItem('Отчёт о миграции', 'showMigrationReport_');
  menu.addSubMenu(diagMenu);
  
  // ===== МИГРАЦИЯ (если нужна) =====
  if (needsMigration()) {
    menu.addSeparator();
    menu.addItem('🔄 Миграция v1 → v2', 'migrateToV2Prompt');
  }
  
  // ===== ПОДМЕНЮ: ОБСЛУЖИВАНИЕ =====
  const maintMenu = ui.createMenu('🛠 Обслуживание');
  maintMenu.addItem('Очистить старые бэкапы', 'cleanupBackupsPrompt');
  maintMenu.addItem('Очистить именованные диапазоны бэкапов', 'cleanupBackupNamedRanges');
  maintMenu.addSeparator();
  maintMenu.addItem('⚠️ Сбросить статус миграции', 'forceMigrationReset');
  menu.addSubMenu(maintMenu);
  
  menu.addSeparator();
  menu.addItem('ℹ️ О программе', 'showAbout_');
  
  menu.addToUi();
}

/**
 * Показывает диалог «О программе»
 */
function showAbout_() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Payment Accounting v' + APP_VERSION,
    'Учёт платежей и взносов для класса/группы.\n\n' +
    'Репозиторий: github.com/yobushka/paymentAccountingGoogleSheet\n\n' +
    'Версия: ' + APP_VERSION,
    ui.ButtonSet.OK
  );
}

// showQuickHelp_() и showQuickBalanceCheck_() определены в dialogs.js

/**
 * Диагностика валидаций — показывает какие правила установлены на каких листах
 */
function diagnoseValidations_() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  const version = detectVersion();
  
  let report = `Версия: ${version}\n\n`;
  
  // Список листов для проверки
  const sheetsToCheck = [
    { name: SHEET_NAMES.GOALS, cols: ['Начисление', 'Статус', 'Тип'] },
    { name: SHEET_NAMES.COLLECTIONS, cols: ['Начисление', 'Статус'] },
    { name: SHEET_NAMES.FAMILIES, cols: ['Активен'] },
    { name: SHEET_NAMES.PAYMENTS, cols: ['Способ', 'family_id (label)', 'goal_id (label)', 'collection_id (label)'] },
    { name: SHEET_NAMES.PARTICIPATION, cols: ['Статус', 'family_id (label)', 'goal_id (label)', 'collection_id (label)'] }
  ];
  
  sheetsToCheck.forEach(sheetInfo => {
    const sh = ss.getSheetByName(sheetInfo.name);
    if (!sh) return;
    
    report += `📄 Лист: ${sheetInfo.name}\n`;
    
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const headerMap = {};
    headers.forEach((h, i) => headerMap[h] = i + 1);
    
    sheetInfo.cols.forEach(colName => {
      const col = headerMap[colName];
      if (!col) return;
      
      // Проверяем ячейку строки 2 (первая строка данных)
      const cell = sh.getRange(2, col);
      const validation = cell.getDataValidation();
      const value = cell.getValue();
      
      report += `  • ${colName} (col ${col}): `;
      
      if (validation) {
        const criteriaType = validation.getCriteriaType();
        const criteriaValues = validation.getCriteriaValues();
        
        if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          report += `LIST [${criteriaValues[0].join(', ')}]`;
        } else if (criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
          report += `RANGE ${criteriaValues[0].getA1Notation()}`;
        } else {
          report += criteriaType.toString();
        }
        
        // Проверяем, подходит ли текущее значение
        if (value && criteriaType === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
          const allowedValues = criteriaValues[0];
          if (!allowedValues.includes(value)) {
            report += ` ⚠️ VALUE "${value}" NOT IN LIST!`;
          }
        }
      } else {
        report += 'NO VALIDATION';
      }
      
      if (value) {
        report += ` (value: "${value}")`;
      }
      report += '\n';
    });
    
    report += '\n';
  });
  
  // Также проверим именованные диапазоны
  report += '📋 Именованные диапазоны:\n';
  const namedRanges = ss.getNamedRanges();
  namedRanges.forEach(nr => {
    report += `  • ${nr.getName()}: ${nr.getRange().getA1Notation()}\n`;
  });
  
  Logger.log(report);
  ui.alert('Diagnose Validations', report.substring(0, 4000), ui.ButtonSet.OK);
}

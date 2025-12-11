/**
 * @fileoverview Меню приложения и обработчик onOpen
 */

/**
 * Создаёт меню при открытии таблицы
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Funds');
  
  // Основные операции
  menu.addItem('Setup / Rebuild structure', 'init');
  menu.addItem('Generate IDs (all sheets)', 'generateAllIds');
  menu.addItem('Rebuild data validations', 'rebuildValidations');
  menu.addItem('Recalculate (Balance & Detail)', 'recalculateAll');
  
  menu.addSeparator();
  
  // Операции со сборами/целями
  menu.addItem('Close Goal', 'closeGoalPrompt');
  
  menu.addSeparator();
  
  // Демо и очистка
  menu.addItem('Load Sample Data (separate)', 'loadSampleDataPrompt');
  menu.addItem('Cleanup visuals (trim sheets)', 'cleanupWorkbook_');
  menu.addItem('Audit & fix field types', 'auditAndFixFieldTypes');
  
  menu.addSeparator();
  
  // Быстрые проверки
  menu.addItem('Quick Help', 'showQuickHelp_');
  menu.addItem('Quick Balance Check', 'showQuickBalanceCheck_');
  menu.addItem('Migration Report', 'showMigrationReport_');
  menu.addItem('🔍 Diagnose Validations', 'diagnoseValidations_');
  
  menu.addSeparator();
  
  // Миграция (если нужна)
  if (needsMigration()) {
    menu.addItem('🔄 Migrate v1 → v2', 'migrateToV2Prompt');
    menu.addSeparator();
  }
  
  // Управление структурой и диагностика
  const structureMenu = ui.createMenu('📋 Structure');
  structureMenu.addItem('Validate all sheets', 'showStructureReport');
  structureMenu.addItem('Fix all sheets', 'fixAllSheetsStructure');
  structureMenu.addItem('Fix current sheet', 'fixSheetStructurePrompt');
  structureMenu.addSeparator();
  structureMenu.addItem('Refresh all headers', 'refreshAllHeaders');
  structureMenu.addItem('Refresh current sheet headers', 'refreshCurrentSheetHeaders');
  menu.addSubMenu(structureMenu);
  
  // Управление стилями
  const stylesMenu = ui.createMenu('🎨 Styles');
  stylesMenu.addItem('Fix all sheets styles', 'fixAllSheetsStyles');
  stylesMenu.addItem('Fix current sheet styles', 'fixCurrentSheetStyles');
  stylesMenu.addItem('Reset current sheet styles', 'resetCurrentSheetStyles');
  stylesMenu.addItem('Quick fix all styles', 'quickFixAllStyles');
  menu.addSubMenu(stylesMenu);
  
  // Управление бэкапами и диагностика
  const backupMenu = ui.createMenu('🛠 Maintenance');
  backupMenu.addItem('Cleanup old backups', 'cleanupBackupsPrompt');
  backupMenu.addItem('Cleanup backup named ranges', 'cleanupBackupNamedRanges');
  backupMenu.addItem('⚠️ Force migration reset', 'forceMigrationReset');
  menu.addSubMenu(backupMenu);
  
  // Информация
  menu.addItem('About', 'showAbout_');
  
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

/**
 * Показывает быструю справку
 */
function showQuickHelp_() {
  const ui = SpreadsheetApp.getUi();
  const help = `
Быстрый старт:
1. Funds → Setup / Rebuild structure
2. Заполните «Семьи» (Активен=Да)
3. Добавьте «Цели» (Статус=Открыта)
4. Настройте «Участие» при необходимости
5. Вносите «Платежи»
6. Смотрите «Баланс» и «Детализация»

Режимы начисления:
• static_per_family — фикс на семью
• shared_total_all — делим на всех участников
• shared_total_by_payers — делим между оплатившими
• dynamic_by_payers — water-filling
• proportional_by_payers — пропорционально платежам
• unit_price — поштучно
• voluntary — добровольно (v2.0)

Баланс v2.0:
Внесено - Списано - Резерв = Свободно
`;
  ui.alert('Quick Help', help, ui.ButtonSet.OK);
}

/**
 * Показывает быструю проверку баланса семьи
 */
function showQuickBalanceCheck_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Quick Balance Check',
    'Введите family_id (например, F001):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const familyId = response.getResponseText().trim();
  if (!familyId) return;
  
  try {
    const paid = PAYED_TOTAL_FAMILY(familyId);
    const accrued = ACCRUED_FAMILY(familyId, 'ALL');
    const free = Math.max(0, paid - accrued);
    const debt = Math.max(0, accrued - paid);
    
    const msg = `
Семья: ${familyId}

Внесено всего: ${paid.toFixed(2)} ₽
Списано (начислено): ${accrued.toFixed(2)} ₽
Свободный остаток: ${free.toFixed(2)} ₽
Задолженность: ${debt.toFixed(2)} ₽
`;
    ui.alert('Balance Check', msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Ошибка', e.message, ui.ButtonSet.OK);
  }
}

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

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
  
  // Управление бэкапами
  menu.addItem('Cleanup old backups', 'cleanupBackupsPrompt');
  
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

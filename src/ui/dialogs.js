/**
 * @fileoverview UI-диалоги: справка, быстрая проверка, аудит
 */

/**
 * Показывает краткую справку
 */
function showQuickHelp_() {
  const ui = SpreadsheetApp.getUi();
  
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: sans-serif; padding: 10px; }
      h2 { color: #1a73e8; margin-top: 0; }
      h3 { color: #5f6368; margin-top: 15px; }
      ul { margin: 5px 0; padding-left: 20px; }
      li { margin: 3px 0; }
      code { background: #f1f3f4; padding: 2px 4px; border-radius: 3px; }
      .mode { margin: 8px 0; }
      .mode-name { font-weight: bold; color: #1a73e8; }
    </style>
    
    <h2>Funds v2.0 — Краткая справка</h2>
    
    <h3>Структура листов</h3>
    <ul>
      <li><strong>Семьи</strong> — список детей (family_id: F001, F002...)</li>
      <li><strong>Цели</strong> — сборы и цели (goal_id: G001, G002...)</li>
      <li><strong>Участие</strong> — кто участвует/не участвует в цели</li>
      <li><strong>Платежи</strong> — все поступления (payment_id: PMT001...)</li>
      <li><strong>Баланс</strong> — сводка по семьям (автоматический расчёт)</li>
    </ul>
    
    <h3>Режимы начисления</h3>
    <div class="mode">
      <span class="mode-name">static_per_family</span> — фиксированная сумма на семью
    </div>
    <div class="mode">
      <span class="mode-name">shared_total_all</span> — делим цель на всех участников
    </div>
    <div class="mode">
      <span class="mode-name">shared_total_by_payers</span> — делим на оплативших
    </div>
    <div class="mode">
      <span class="mode-name">dynamic_by_payers</span> — water-filling: справедливое распределение
    </div>
    <div class="mode">
      <span class="mode-name">proportional_by_payers</span> — пропорционально взносам
    </div>
    <div class="mode">
      <span class="mode-name">unit_price_by_payers</span> — поштучно (кратно цене)
    </div>
    <div class="mode">
      <span class="mode-name">voluntary</span> — добровольный взнос (списывается сколько внесено)
    </div>
    
    <h3>Основные действия</h3>
    <ul>
      <li><strong>Funds → Setup</strong> — первичная настройка</li>
      <li><strong>Funds → Generate IDs</strong> — автозаполнение ID</li>
      <li><strong>Funds → Rebuild Validations</strong> — обновить выпадающие списки</li>
      <li><strong>Funds → Close Goal</strong> — закрыть цель (фиксирует cap)</li>
    </ul>
    
    <h3>Типы целей (v2.0)</h3>
    <ul>
      <li><strong>разовая</strong> — однократный сбор</li>
      <li><strong>регулярная</strong> — повторяется с периодичностью</li>
    </ul>
  `).setWidth(500).setHeight(550);
  
  ui.showModalDialog(html, 'Справка');
}

/**
 * Быстрая проверка баланса для конкретной семьи
 */
function showQuickBalanceCheck_() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Проверка баланса',
    'Введите family_id (например, F001) или имя ребёнка:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const query = response.getResponseText().trim();
  if (!query) return;
  
  const ss = SpreadsheetApp.getActive();
  
  // Ищем семью
  const shFamilies = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  if (!shFamilies || shFamilies.getLastRow() < 2) {
    ui.alert('Ошибка', 'Лист «Семьи» пуст или не найден.', ui.ButtonSet.OK);
    return;
  }
  
  const familiesData = shFamilies.getDataRange().getValues();
  const fHeaders = familiesData[0];
  const fIdCol = fHeaders.indexOf('family_id');
  const fNameCol = fHeaders.indexOf('Имя ребёнка');
  
  let familyId = null;
  let familyName = null;
  
  for (let i = 1; i < familiesData.length; i++) {
    const row = familiesData[i];
    const id = String(row[fIdCol] || '');
    const name = String(row[fNameCol] || '');
    
    if (id.toLowerCase() === query.toLowerCase() ||
        name.toLowerCase().includes(query.toLowerCase())) {
      familyId = id;
      familyName = name;
      break;
    }
  }
  
  if (!familyId) {
    ui.alert('Не найдено', `Семья «${query}» не найдена.`, ui.ButtonSet.OK);
    return;
  }
  
  // Получаем баланс
  const shBalance = ss.getSheetByName(SHEET_NAMES.BALANCE);
  if (!shBalance || shBalance.getLastRow() < 2) {
    ui.alert('Ошибка', 'Лист «Баланс» пуст или не найден.', ui.ButtonSet.OK);
    return;
  }
  
  const balanceData = shBalance.getDataRange().getValues();
  const bHeaders = balanceData[0];
  const bIdCol = bHeaders.indexOf('family_id');
  
  let balanceRow = null;
  for (let i = 1; i < balanceData.length; i++) {
    if (balanceData[i][bIdCol] === familyId) {
      balanceRow = balanceData[i];
      break;
    }
  }
  
  if (!balanceRow) {
    ui.alert('Не найдено', `Баланс для ${familyId} не найден.`, ui.ButtonSet.OK);
    return;
  }
  
  // Формируем отчёт
  const getVal = (colName) => {
    const idx = bHeaders.indexOf(colName);
    return idx >= 0 ? balanceRow[idx] : 0;
  };
  
  const paid = getVal('Внесено всего') || getVal('Оплачено');
  const charged = getVal('Списано всего') || getVal('Начислено');
  const reserved = getVal('Зарезервировано') || 0;
  const free = getVal('Свободный остаток') || getVal('Переплата') || 0;
  const debt = getVal('Задолженность') || 0;
  
  const msg = `
Семья: ${familyName} (${familyId})

💰 Внесено всего: ${formatMoney_(paid)}
📊 Списано всего: ${formatMoney_(charged)}
🔒 Зарезервировано: ${formatMoney_(reserved)}
✅ Свободный остаток: ${formatMoney_(free)}
❌ Задолженность: ${formatMoney_(debt)}
`.trim();
  
  ui.alert(`Баланс: ${familyName}`, msg, ui.ButtonSet.OK);
}

/**
 * Форматирует число как деньги
 * @param {number} v
 * @return {string}
 */
function formatMoney_(v) {
  const n = Number(v) || 0;
  return n.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' ₽';
}

/**
 * Аудит типов данных в полях
 */
function showAuditFieldTypes_() {
  const ss = SpreadsheetApp.getActive();
  const results = [];
  
  const checkSheet = (name, expectedCols) => {
    const sh = ss.getSheetByName(name);
    if (!sh) {
      results.push(`⚠️ Лист «${name}» не найден`);
      return;
    }
    
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const missing = expectedCols.filter(c => !headers.includes(c));
    const extra = headers.filter(h => h && !expectedCols.includes(h));
    
    if (missing.length > 0) {
      results.push(`❌ ${name}: отсутствуют колонки: ${missing.join(', ')}`);
    }
    if (extra.length > 0) {
      results.push(`ℹ️ ${name}: дополнительные колонки: ${extra.join(', ')}`);
    }
    if (missing.length === 0 && extra.length === 0) {
      results.push(`✅ ${name}: структура корректна`);
    }
  };
  
  // Проверяем все листы
  checkSheet(SHEET_NAMES.FAMILIES, ['family_id', 'Имя ребёнка', 'Активен']);
  checkSheet(SHEET_NAMES.GOALS, [
    'goal_id', 'Название цели', 'Тип', 'Статус', 'Начисление', 
    'Параметр суммы', 'Периодичность', 'Родительская цель'
  ]);
  checkSheet(SHEET_NAMES.PARTICIPATION, ['family_id (label)', 'goal_id (label)', 'Участие']);
  checkSheet(SHEET_NAMES.PAYMENTS, [
    'payment_id', 'Дата', 'family_id (label)', 'goal_id (label)', 
    'Сумма', 'Комментарий'
  ]);
  checkSheet(SHEET_NAMES.BALANCE, [
    'family_id', 'Имя ребёнка', 'Внесено всего', 'Списано всего',
    'Зарезервировано', 'Свободный остаток', 'Задолженность'
  ]);
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('Аудит структуры', results.join('\n'), ui.ButtonSet.OK);
}

/**
 * Показывает детальный отчёт по цели
 */
function showGoalReport_() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Отчёт по цели',
    'Введите goal_id (например, G001) или название:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const query = response.getResponseText().trim();
  if (!query) return;
  
  const ss = SpreadsheetApp.getActive();
  
  // Ищем цель
  const shGoals = ss.getSheetByName(SHEET_NAMES.GOALS);
  if (!shGoals || shGoals.getLastRow() < 2) {
    ui.alert('Ошибка', 'Лист «Цели» пуст или не найден.', ui.ButtonSet.OK);
    return;
  }
  
  const goalsData = shGoals.getDataRange().getValues();
  const gHeaders = goalsData[0];
  const gIdCol = gHeaders.indexOf('goal_id');
  const gNameCol = gHeaders.indexOf('Название цели');
  const gStatusCol = gHeaders.indexOf('Статус');
  const gModeCol = gHeaders.indexOf('Начисление');
  const gAmountCol = gHeaders.indexOf('Параметр суммы');
  
  let goalRow = null;
  
  for (let i = 1; i < goalsData.length; i++) {
    const row = goalsData[i];
    const id = String(row[gIdCol] || '');
    const name = String(row[gNameCol] || '');
    
    if (id.toLowerCase() === query.toLowerCase() ||
        name.toLowerCase().includes(query.toLowerCase())) {
      goalRow = row;
      break;
    }
  }
  
  if (!goalRow) {
    ui.alert('Не найдено', `Цель «${query}» не найдена.`, ui.ButtonSet.OK);
    return;
  }
  
  const goalId = goalRow[gIdCol];
  const goalName = goalRow[gNameCol];
  const goalStatus = goalRow[gStatusCol];
  const goalMode = goalRow[gModeCol];
  const goalAmount = goalRow[gAmountCol];
  
  // Считаем платежи по этой цели
  const shPayments = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  let totalPaid = 0;
  let payersCount = 0;
  
  if (shPayments && shPayments.getLastRow() > 1) {
    const payData = shPayments.getDataRange().getValues();
    const pHeaders = payData[0];
    const pGoalCol = pHeaders.indexOf('goal_id (label)');
    const pAmountCol = pHeaders.indexOf('Сумма');
    
    const payers = new Set();
    const pFamilyCol = pHeaders.indexOf('family_id (label)');
    
    for (let i = 1; i < payData.length; i++) {
      const goalLabel = String(payData[i][pGoalCol] || '');
      const extractedId = getIdFromLabelish_(goalLabel);
      
      if (extractedId === goalId) {
        totalPaid += Number(payData[i][pAmountCol]) || 0;
        payers.add(getIdFromLabelish_(payData[i][pFamilyCol]));
      }
    }
    payersCount = payers.size;
  }
  
  const msg = `
Цель: ${goalName} (${goalId})

📋 Статус: ${goalStatus}
📊 Режим: ${goalMode}
💵 Целевая сумма: ${formatMoney_(goalAmount)}

📥 Собрано: ${formatMoney_(totalPaid)}
👥 Плательщиков: ${payersCount}
📈 Прогресс: ${goalAmount > 0 ? Math.round(totalPaid / goalAmount * 100) : 0}%
`.trim();
  
  ui.alert(`Отчёт: ${goalName}`, msg, ui.ButtonSet.OK);
}

/**
 * Показывает общую статистику
 */
function showOverallStats_() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  
  // Семьи
  const shFamilies = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  const familiesCount = shFamilies ? Math.max(0, shFamilies.getLastRow() - 1) : 0;
  
  // Цели
  const shGoals = ss.getSheetByName(SHEET_NAMES.GOALS);
  let goalsCount = 0;
  let openGoals = 0;
  
  if (shGoals && shGoals.getLastRow() > 1) {
    const goalsData = shGoals.getDataRange().getValues();
    const statusCol = goalsData[0].indexOf('Статус');
    goalsCount = goalsData.length - 1;
    
    if (statusCol >= 0) {
      for (let i = 1; i < goalsData.length; i++) {
        if (goalsData[i][statusCol] === 'Открыта') openGoals++;
      }
    }
  }
  
  // Платежи
  const shPayments = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  let paymentsCount = 0;
  let totalAmount = 0;
  
  if (shPayments && shPayments.getLastRow() > 1) {
    const payData = shPayments.getDataRange().getValues();
    const amountCol = payData[0].indexOf('Сумма');
    paymentsCount = payData.length - 1;
    
    if (amountCol >= 0) {
      for (let i = 1; i < payData.length; i++) {
        totalAmount += Number(payData[i][amountCol]) || 0;
      }
    }
  }
  
  const msg = `
📊 Общая статистика

👨‍👩‍👧‍👦 Семей: ${familiesCount}
🎯 Целей всего: ${goalsCount}
   • Открытых: ${openGoals}
   • Закрытых: ${goalsCount - openGoals}

💳 Платежей: ${paymentsCount}
💰 Общая сумма: ${formatMoney_(totalAmount)}
`.trim();
  
  ui.alert('Статистика', msg, ui.ButtonSet.OK);
}

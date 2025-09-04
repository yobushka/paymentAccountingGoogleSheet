/** Funds tracker (1 family = 1 child) — production build
 * Modes: static_per_child (fixed per family), shared_total_all, shared_total_by_payers, dynamic_by_payers, proportional_by_payers, unit_price_by_payers
 * Sheets: Инструкция, Семьи, Сборы, Участие, Платежи, Баланс, Детализация, Сводка, Lists(hidden)
 * Dropdowns show "Название (ID)" everywhere; logic extracts IDs.
 * Dates matter only in Payments for reference; calculations are instant.
 *
 * Menu:
 *  • Setup / Rebuild structure
 *  • Rebuild data validations
 *  • Recalculate (Balance & Detail)
 *  • Generate IDs (all sheets)
 *  • Close Collection (fix x & set Closed)
 *  • Load Sample Data (separate)  ← fills demo families, collections, participation, and payments
 *
 * Custom functions for sheet formulas:
 *  • LABEL_TO_ID(value)
 *  • PAYED_TOTAL_FAMILY(familyLabelOrId)
 *  • ACCRUED_FAMILY(familyLabelOrId, statusFilter="OPEN"|"ALL")
 *  • DYN_CAP(T, payments_range)
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💰 Фонды')
    .addItem('🔧 Настроить / Пересоздать структуру', 'init')
    .addItem('🔄 Пересоздать валидации', 'rebuildValidations')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Отчёты и действия')
      .addItem('🔄 Пересчитать всё', 'recalculateAll')
      .addItem('📈 Быстрая проверка баланса', 'showQuickBalanceCheck_')
      .addItem('⚠️ Показать ошибки валидации', 'showValidationErrors_'))
    .addSubMenu(ui.createMenu('🎨 Внешний вид и очистка')
      .addItem('✨ Очистить лишнее (обрезать листы)', 'cleanupWorkbook_')
      .addItem('🎯 Выделить ключевые данные', 'highlightKeyData_')
      .addItem('📱 Мобильный вид', 'setupMobileView_'))
    .addSeparator()
    .addSubMenu(ui.createMenu('⚙️ Управление данными')
      .addItem('🆔 Сгенерировать ID (все листы)', 'generateAllIds')
      .addItem('🔒 Закрыть сбор (зафиксировать x и пометить «Закрыт»)', 'closeCollectionPrompt')
      .addItem('📋 Дублировать сбор', 'duplicateCollection_'))
    .addSeparator()
    .addItem('🎲 Загрузить пример данных', 'loadSampleDataPrompt')
    .addItem('❓ Быстрая помощь', 'showQuickHelp_')
    .addToUi();
  // Ensure header notes are set on open as well
  setupHeaderNotes_();
  // Show welcome toast for first-time users
  showWelcomeToast_();
  try { addHeaderNotes_(); } catch(e) {}
}

// Auto-refresh Balance on relevant edits
function onEdit(e) {
  try {
    const sh = e && e.range && e.range.getSheet();
    if (!sh) return;
    const name = sh.getName();
    
    // Only refresh Balance for significant changes, not every edit
    if (name === 'Платежи') {
  const start = e.range.getColumn();
  const end = start + e.range.getNumColumns() - 1;
  const map = getHeaderMap_(sh);
  const keys = [map['family_id (label)'], map['collection_id (label)'], map['Сумма']].filter(Boolean);
  const overlaps = keys.some(c => c >= start && c <= end);
  if (overlaps) refreshBalanceFormulas_();
    } else if (name === 'Семьи') {
  const start = e.range.getColumn();
  const end = start + e.range.getNumColumns() - 1;
  const map = getHeaderMap_(sh);
  const keys = [map['family_id'], map['Активен']].filter(Boolean);
  const overlaps = keys.some(c => c >= start && c <= end);
  if (overlaps) refreshBalanceFormulas_();
    } else if (name === 'Сборы') {
      // Mode/participants changes affect accruals; refresh Balance
      refreshBalanceFormulas_();
    } else if (name === 'Баланс') {
      const col = e.range.getColumn();
      // Only refresh if changing the selector
      if (col === 9) { // Column I (selector)
        refreshBalanceFormulas_();
      }
    }
    
    // Detail & Summary sheet refresh for broader changes
    if (name === 'Платежи' || name === 'Семьи' || name === 'Сборы' || name === 'Участие' || name === 'Детализация' || name === 'Сводка') {
      refreshDetailSheet_();
      refreshSummarySheet_();
    }
    
    // Auto-generate IDs when user starts filling key fields
    if (name === 'Семьи') maybeAutoIdRow_(sh, e.range.getRow(), 'family_id', 'F', 3, ['Ребёнок ФИО']);
    else if (name === 'Сборы') maybeAutoIdRow_(sh, e.range.getRow(), 'collection_id', 'C', 3, ['Название сбора']);
    else if (name === 'Платежи') maybeAutoIdRow_(sh, e.range.getRow(), 'payment_id', 'PMT', 3, ['Дата','family_id (label)','collection_id (label)','Сумма']);
  } catch (err) {
    // silent guard
  }
}/** =========================
 *  INITIALIZATION / STRUCTURE
 *  ========================= */
function init() {
  const ss = SpreadsheetApp.getActive();
  const specs = getSheetsSpec();

  // Create/clear sheets and headers
  for (const spec of specs) {
    const sh = getOrCreateSheet(ss, spec.name);
    // Non-destructive rebuild: preserve data, refresh headers/widths/formats
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
    spec.colWidths?.forEach((w, i) => { if (w) sh.setColumnWidth(i + 1, w); });
    if (spec.dateCols?.length) {
      const last = Math.max(2, sh.getMaxRows());
      spec.dateCols.forEach(c => sh.getRange(2, c, last - 1, 1).setNumberFormat('yyyy-mm-dd'));
    }
  }

  // Fill instruction page
  setupInstructionSheet();

  // Hidden helper sheet with dynamic lists (labels)
  setupListsSheet();

  // Named ranges (labels and raw ids if needed)
  ensureNamedRange('FAMILIES_LABELS',        'Lists!D2:D'); // all families labels
  ensureNamedRange('ACTIVE_FAMILIES_LABELS', 'Lists!C2:C'); // active only
  ensureNamedRange('COLLECTIONS_LABELS',     'Lists!B2:B'); // all collections labels
  ensureNamedRange('OPEN_COLLECTIONS_LABELS','Lists!A2:A'); // open only
  setRawIdNamedRanges_();

  rebuildValidations();
  setupBalanceExamples();

  // Header notes (hover tooltips)
  addHeaderNotes_();

  // Visual styling (headers, formats, filters, banding, conditional formats)
  styleWorkbook_();

  // Remove legacy DynCalc sheet if present
  try { const legacy = ss.getSheetByName('DynCalc'); if (legacy) ss.deleteSheet(legacy); } catch (_) {}

  SpreadsheetApp.getActive().toast('Structure created/updated.', 'Funds');
}

/** Cleanup visuals: trim extra rows/columns to used area and re-apply styles */
function cleanupWorkbook_() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ['Инструкция','Семьи','Сборы','Участие','Платежи','Баланс','Детализация','Сводка'];
  sheets.forEach(name => {
    const sh = ss.getSheetByName(name); if (!sh) return;
    const lastRow = Math.max(1, sh.getLastRow());
    const lastCol = Math.max(1, sh.getLastColumn());
    // Trim rows
    const maxRows = sh.getMaxRows();
    if (maxRows > lastRow + 50) { // keep a small buffer
      try { sh.deleteRows(lastRow + 51, maxRows - (lastRow + 50)); } catch(_) {}
    }
    // Trim columns
    const maxCols = sh.getMaxColumns();
    if (maxCols > lastCol) {
      try { sh.deleteColumns(lastCol + 1, maxCols - lastCol); } catch(_) {}
    }
  });
  // Re-apply styles
  styleWorkbook_();
  SpreadsheetApp.getActive().toast('Sheets trimmed and visuals refreshed.', 'Funds');
}

/** =========================
 *  VISUAL STYLING
 *  ========================= */
function styleWorkbook_() {
  const ss = SpreadsheetApp.getActive();
  const names = ['Инструкция','Семьи','Сборы','Участие','Платежи','Баланс','Детализация','Сводка'];
  names.forEach(n => {
    const sh = ss.getSheetByName(n);
    if (!sh) return;
    styleSheetHeader_(sh);
    if (n === 'Баланс') styleBalanceSheet_(sh);
    else if (n === 'Детализация') styleDetailSheet_(sh);
    else if (n === 'Сводка') styleSummarySheet_(sh);
    else if (n === 'Платежи') stylePaymentsSheet_(sh);
    else if (n === 'Сборы') styleCollectionsSheet_(sh);
    else if (n === 'Семьи') styleFamiliesSheet_(sh);
    else if (n === 'Участие') styleParticipationSheet_(sh);
    // Hide gridlines on display sheets
    try {
      if (n === 'Инструкция' || n === 'Баланс' || n === 'Детализация' || n === 'Сводка') sh.setHiddenGridlines(true);
      else sh.setHiddenGridlines(false);
    } catch (_) {}
  });
}

/** =========================
 *  UX ENHANCEMENT FUNCTIONS
 *  ========================= */

function showWelcomeToast_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const instructionSheet = ss.getSheetByName('Инструкция');
    if (instructionSheet && instructionSheet.getRange('A2').getValue() === '') {
      // First time user - show welcome
      ss.toast('Добро пожаловать! Начните с Funds → Setup, затем изучите лист "Инструкция".', '💰 Funds Tracker', 10);
    }
  } catch (e) {
    Logger.log('Welcome toast error: ' + e.message);
  }
}

function showQuickHelp_() {
  const ui = SpreadsheetApp.getUi();
  const help = `
🏃‍♂️ БЫСТРЫЙ СТАРТ:
1. Funds → Setup (если не сделали)
2. Заполните "Семьи" (Активен=Да)
3. Создайте "Сборы" (Статус=Открыт)
4. Вносите "Платежи"
5. Смотрите "Баланс" и "Сводка"

🎯 ПОЛЕЗНЫЕ ЛИСТЫ:
• "Инструкция" - подробное руководство
• "Баланс" - кто сколько должен/переплатил
• "Сводка" - статистика по сборам
• "Детализация" - расшифровка по семьям

⚡ БЫСТРЫЕ ДЕЙСТВИЯ:
• Funds → Quick Balance Check
• Funds → Recalculate All
• Funds → Highlight Key Data

❓ Проблемы? Проверьте лист "Инструкция" раздел "Советы".`;
  
  ui.alert('💰 Funds Tracker - Справка', help, ui.ButtonSet.OK);
}

function showQuickBalanceCheck_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const shBal = ss.getSheetByName('Баланс');
    if (!shBal) {
      SpreadsheetApp.getUi().alert('Ошибка', 'Лист "Баланс" не найден. Выполните Setup.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
    
    // Count families with debts and overpayments
    const lastRow = shBal.getLastRow();
    if (lastRow < 2) {
      ss.toast('Нет данных для анализа.', 'Balance Check');
      return;
    }
    
    const data = shBal.getRange(2, 1, lastRow-1, 6).getValues();
    let totalFamilies = 0, withDebts = 0, withOverpay = 0;
    let totalDebt = 0, totalOverpay = 0;
    
    data.forEach(row => {
      if (row[0]) { // has family_id
        totalFamilies++;
        const overpay = Number(row[2]) || 0;
        const debt = Number(row[5]) || 0;
        if (debt > 0) { withDebts++; totalDebt += debt; }
        if (overpay > 0) { withOverpay++; totalOverpay += overpay; }
      }
    });
    
    const report = `
📊 БЫСТРАЯ СВОДКА ПО БАЛАНСАМ:

👥 Семьи: ${totalFamilies}
💸 С задолженностью: ${withDebts} (общая сумма: ${totalDebt.toFixed(2)} ₽)
💰 С переплатой: ${withOverpay} (общая сумма: ${totalOverpay.toFixed(2)} ₽)
✅ Баланс "ноль": ${totalFamilies - withDebts - withOverpay}

${withDebts > 0 ? '⚠️ Есть задолженности!' : '✅ Задолженностей нет'}
${totalOverpay > totalDebt ? '💡 Переплат больше долгов - можно зачесть' : ''}`;
    
    SpreadsheetApp.getUi().alert('💰 Быстрая проверка балансов', report, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    toastErr_('Quick balance check failed: ' + e.message);
  }
}

function showValidationErrors_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const issues = [];
    
    // Check for families without IDs
    const shF = ss.getSheetByName('Семьи');
    if (shF && shF.getLastRow() > 1) {
      const mapF = getHeaderMap_(shF);
      const ids = shF.getRange(2, mapF['family_id'], shF.getLastRow()-1, 1).getValues().flat();
      const emptyIds = ids.filter((id, idx) => !id).length;
      if (emptyIds > 0) issues.push(`• Семьи: ${emptyIds} строк без ID`);
    }
    
    // Check for collections without IDs
    const shC = ss.getSheetByName('Сборы');
    if (shC && shC.getLastRow() > 1) {
      const mapC = getHeaderMap_(shC);
      const ids = shC.getRange(2, mapC['collection_id'], shC.getLastRow()-1, 1).getValues().flat();
      const emptyIds = ids.filter(id => !id).length;
      if (emptyIds > 0) issues.push(`• Сборы: ${emptyIds} строк без ID`);
    }
    
    // Check payments
    const shP = ss.getSheetByName('Платежи');
    if (shP && shP.getLastRow() > 1) {
      const mapP = getHeaderMap_(shP);
      const amounts = shP.getRange(2, mapP['Сумма'], shP.getLastRow()-1, 1).getValues().flat();
      const invalidAmounts = amounts.filter((amt, idx) => amt !== '' && (isNaN(amt) || Number(amt) <= 0)).length;
      if (invalidAmounts > 0) issues.push(`• Платежи: ${invalidAmounts} некорректных сумм`);
    }
    
    if (issues.length === 0) {
      ss.toast('✅ Проблем не обнаружено!', 'Validation Check', 5);
    } else {
      const report = '⚠️ НАЙДЕННЫЕ ПРОБЛЕМЫ:\n\n' + issues.join('\n') + '\n\n💡 Используйте Funds → Generate IDs для исправления.';
      SpreadsheetApp.getUi().alert('Проверка данных', report, SpreadsheetApp.getUi().ButtonSet.OK);
    }
  } catch (e) {
    toastErr_('Validation check failed: ' + e.message);
  }
}

function highlightKeyData_() {
  try {
    const ss = SpreadsheetApp.getActive();
    
    // Highlight negative balances in red, positive in green
    const shBal = ss.getSheetByName('Баланс');
    if (shBal && shBal.getLastRow() > 1) {
      const map = getHeaderMap_(shBal);
      if (map['Задолженность']) {
        const rng = shBal.getRange(2, map['Задолженность'], shBal.getLastRow()-1, 1);
        rng.setBackground('#ffebee'); // Light red background
        // Add bold formatting for values > 0
        const values = rng.getValues();
        for (let i = 0; i < values.length; i++) {
          if (values[i][0] > 0) {
            shBal.getRange(2+i, map['Задолженность']).setFontWeight('bold');
          }
        }
      }
    }
    
    ss.toast('✨ Ключевые данные выделены', 'Highlight Data', 3);
  } catch (e) {
    toastErr_('Highlight failed: ' + e.message);
  }
}

function setupMobileView_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const sheets = ['Баланс', 'Сводка', 'Платежи'];
    
    sheets.forEach(sheetName => {
      const sh = ss.getSheetByName(sheetName);
      if (sh) {
        // Set optimal column widths for mobile
        sh.setColumnWidth(1, 100); // IDs shorter
        if (sheetName === 'Баланс') {
          sh.setColumnWidth(2, 180); // Names
          sh.setColumnWidths(3, 4, 120); // Numbers
        }
        // Hide less important columns for mobile
        if (sheetName === 'Платежи') {
          const map = getHeaderMap_(sh);
          if (map['Комментарий']) sh.hideColumns(map['Комментарий']);
          if (map['payment_id']) sh.hideColumns(map['payment_id']);
        }
      }
    });
    
    ss.toast('📱 Мобильный вид настроен', 'Mobile View', 3);
  } catch (e) {
    toastErr_('Mobile setup failed: ' + e.message);
  }
}

function duplicateCollection_() {
  try {
    const ss = SpreadsheetApp.getActive();
    const shC = ss.getSheetByName('Сборы');
    if (!shC) return;
    
    const ui = SpreadsheetApp.getUi();
    const response = ui.prompt('Дублировать сбор', 'Введите ID сбора для дублирования (например, C001):', ui.ButtonSet.OK_CANCEL);
    
    if (response.getSelectedButton() !== ui.Button.OK) return;
    
    const sourceId = response.getResponseText().trim();
    if (!sourceId) return;
    
    // Find source collection
    const map = getHeaderMap_(shC);
    const data = shC.getRange(2, 1, shC.getLastRow()-1, shC.getLastColumn()).getValues();
    const sourceRow = data.find(row => row[map['collection_id']-1] === sourceId);
    
    if (!sourceRow) {
      ui.alert('Ошибка', `Сбор ${sourceId} не найден.`, ui.ButtonSet.OK);
      return;
    }
    
    // Create new row with new ID
    const newId = generateNextId_(data.map(r => r[map['collection_id']-1]), 'C', 3);
    const newRow = [...sourceRow];
    newRow[map['collection_id']-1] = newId;
    newRow[map['Название сбора']-1] = sourceRow[map['Название сбора']-1] + ' (копия)';
    newRow[map['Статус']-1] = 'Открыт';
    
    // Add to sheet
    shC.appendRow(newRow);
    
    ss.toast(`✅ Сбор дублирован как ${newId}`, 'Duplicate Collection', 5);
    rebuildValidations(); // Refresh dropdowns
  } catch (e) {
    toastErr_('Duplicate collection failed: ' + e.message);
  }
}

/** Adds helpful hover notes to header cells across main sheets */
function addHeaderNotes_() {
  const ss = SpreadsheetApp.getActive();
  // Enhanced notes with emojis and better explanations
  
  // Семьи
  (function(){
    const sh = ss.getSheetByName('Семьи'); if (!sh) return;
    const notes = {
      'Ребёнок ФИО': '👶 Фамилия Имя Отчество ребёнка.\nОдна строка = одна семья (один ребёнок).',
      'День рождения': '🎂 Дата рождения ребёнка (формат yyyy-mm-dd).\nИспользуется справочно для возрастной аналитики.',
      'Мама телеграм': '📱 Контакт мамы в Telegram (@username или ссылка)\nДля оперативной связи по платежам.',
      'Папа телеграм': '📱 Контакт папы в Telegram (@username или ссылка)\nДля оперативной связи по платежам.',
      'Мама ФИО': '👩 Контактная информация мамы.\nИспользуется справочно.',
      'Активен': '✅ Да — семья участвует по умолчанию во всех открытых сборах\n❌ Нет — исключена из участия (если не указана в «Участие»)',
      'Комментарий': '📝 Любая заметка по семье.\nНапример: льготы, особенности оплаты.',
      'family_id': '🆔 Авто-ID семьи (F001, F002, ...).\n⚠️ Генерируется автоматически - не редактируйте!'
    };
    setHeaderNotes_(sh, notes);
  })();

  // Сборы - enhanced notes
  (function(){
    const sh = ss.getSheetByName('Сборы'); if (!sh) return;
    const notes = {
      'Название сбора': '📋 Короткое и понятное имя сбора.\nПоказывается в выпадающих списках платежей.',
      'Статус': '🔓 Открыт — участвует в начислениях\n🔒 Закрыт — не влияет (только оплаты/возвраты)',
      'Дата начала': '📅 Справочно. На расчёты не влияет.\nПолезно для отчётности.',
      'Дедлайн': '⏰ Справочно. На расчёты не влияет.\nДля контроля сроков сбора.',
  'Начисление': '⚙️ Режим расчёта:\n• static_per_child - фикс на семью\n• shared_total_all - общая сумма на всех\n• shared_total_by_payers - на оплативших\n• dynamic_by_payers - динамическое выравнивание (water-filling)\n• proportional_by_payers - пропорционально платежам (без долгов)\n• unit_price_by_payers - поштучно: x=«Фиксированный x» (цена за единицу), списывается floor(P_i/x)*x (полными единицами) только у плативших',
      'Параметр суммы': '💰 Размер взноса или общая цель:\n• static_per_child: сумма с семьи\n• другие режимы: общая цель T',
  'Фиксированный x': '🔒 Для dynamic_by_payers — cap после закрытия (до закрытия рассчитывается автоматически).\nДля unit_price_by_payers — цена за одну единицу.',
  'Закупка из средств': '🛒 Источник закупки: из каких денег была произведена закупка по этому сбору. Примеры: "Классный фонд", "Пожертвования", "Личные".',
  'Возмещено': '♻️ Отмечайте "Да", если закупка уже возмещена из собранных средств; "Нет" — если возмещение ещё предстоит.',
      'Комментарий': '📝 Описание сбора, цели, особенности.\nВидно участникам.',
      'collection_id': '🆔 Авто-ID сбора (C001, C002, ...).\n⚠️ Генерируется автоматически - не редактируйте!',
      'Ссылка на гуглдиск': '☁️ Ссылка на папку/файл Google Drive.\nДля отчётов, документов по сбору.'
    };
    setHeaderNotes_(sh, notes);
  })();

  // Платежи - enhanced notes  
  (function(){
    const sh = ss.getSheetByName('Платежи'); if (!sh) return;
    const notes = {
      'Дата': '📅 Информационное поле.\nРасчёты мгновенные, дата на них не влияет.',
      'family_id (label)': '👨‍👩‍👧‍👦 Выберите плательщика из списка.\nФормат: "Имя ребёнка (F001)"',
      'collection_id (label)': '📋 Выберите сбор из списка.\nФормат: "Название сбора (C001)"',
      'Сумма': '💰 Сумма платежа (должна быть > 0).\nВалидируется автоматически.',
      'Способ': '💳 Способ оплаты:\nСБП, карта, наличные, др.',
      'Комментарий': '📝 Дополнительная информация:\nназначение, особенности платежа.',
      'payment_id': '🆔 Авто-ID платежа (PMT001, PMT002, ...).\n⚠️ Генерируется автоматически - не редактируйте!'
    };
    setHeaderNotes_(sh, notes);
  })();

  // Участие
  (function(){
    const sh = ss.getSheetByName('Участие'); if (!sh) return;
    const notes = {
      'collection_id (label)': 'Выберите конкретный открытый сбор из списка «Название (ID)».',
      'family_id (label)': 'Выберите активную семью из списка «Имя (ID)».',
      'Статус': 'Участвует — включить; Не участвует — исключить. Если есть хотя бы один «Участвует», участвуют только отмеченные.',
      'Комментарий': 'Справочный комментарий к участию.'
    };
    setHeaderNotes_(sh, notes);
  })();

  // Баланс
  (function(){
    const sh = ss.getSheetByName('Баланс'); if (!sh) return;
    const notes = {
      'family_id': 'ID семьи для ссылок и формул.',
      'Имя ребёнка': 'Автоподтягивается по ID из «Семьи».',
  'Переплата (текущая)': 'MAX(0, Оплачено всего − Начислено всего).',
      'Оплачено всего': 'Сумма всех платежей семьи по всем сборам.',
  'Начислено всего': 'Итог начислений по всем сборам (открытые и закрытые), с учётом участия и режима.',
  'Задолженность': 'MAX(0, Начислено всего − Оплачено).'
    };
    setHeaderNotes_(sh, notes);
  })();

  // Детализация
  (function(){
    const sh = ss.getSheetByName('Детализация'); if (!sh) return;
    const notes = {
      'family_id': 'ID семьи. Строки формируются динамически для пар семья↔сбор.',
      'Имя ребёнка': 'Имя из листа «Семьи».',
      'collection_id': 'ID сбора. Только те, что попадают под фильтр (K1).',
      'Название сбора': 'Имя из листа «Сборы».',
      'Оплачено': 'Сумма платежей семьи в этот сбор.',
  'Начислено': 'Начислено по правилам сбора и участию: static — фикс, shared_total_all — T/N, shared_total_by_payers — T/K (только оплатившим), dynamic — min(P_i, x), proportional — пропорционально платежам без долгов, unit_price_by_payers — floor(P_i/x)*x (полные единицы; x=«Фиксированный x»).',
      'Разность (±)': 'Оплачено − Начислено. Положительное — переплата, отрицательное — недоплата.',
  'Режим': 'Режим начисления: static_per_child | shared_total_all | shared_total_by_payers | dynamic_by_payers | proportional_by_payers | unit_price_by_payers.'
    };
    setHeaderNotes_(sh, notes);
  })();

  // Сводка
  (function(){
    const sh = ss.getSheetByName('Сводка'); if (!sh) return;
    const notes = {
      'collection_id': 'ID сбора.',
      'Название сбора': 'Имя из листа «Сборы».',
      'Режим': 'Режим начисления сбора.',
  'Сумма цели': 'Для shared_total_all/shared_total_by_payers/dynamic_by_payers/proportional_by_payers/unit_price_by_payers — заданная цель T. Для static_per_child — N(участников) × ставка.',
      'Собрано': 'Сумма платежей по сбору от участников (Σ платежей).',
      'Участников': 'Число участников сбора (по правилам «Участие» и «Активен»).',
  'Плательщиков': 'Число уникальных плательщиков (K). Для unit_price_by_payers число единиц смотрите в «Единиц оплачено».',
  'Единиц оплачено': 'Только для unit_price_by_payers: ⌊Собрано/x⌋ (x = «Фиксированный x»). Показывает, сколько штук уже профинансировано.',
  'Ещё плательщиков до закрытия': 'Оценка по режиму:\n• static_per_child: ceil(Остаток/ставка)\n• shared_total_all: ceil(Остаток/(T/N))\n• shared_total_by_payers: ceil(Остаток/доля), доля≈T/K (или фиксированный x)\n• dynamic_by_payers: ceil(Остаток/x) при зафиксированном x; иначе пусто\n• proportional_by_payers: — (не применяется)\n• unit_price_by_payers: ceil(Остаток/x), где x=«Фиксированный x»',
      'Остаток до цели': 'MAX(0, Сумма цели − Собрано).'
    };
    setHeaderNotes_(sh, notes);
  })();
}

/**
 * Ensures header notes are applied. Safe wrapper used on onOpen.
 * If addHeaderNotes_ throws (e.g., missing sheets), we swallow the error.
 */
function setupHeaderNotes_() {
  try {
    addHeaderNotes_();
  } catch (e) {
    // no-op: notes are optional and shouldn't block UI
  }
}



/** Assigns notes to headers by header text */
function setHeaderNotes_(sh, byHeader) {
  const map = getHeaderMap_(sh);
  Object.keys(byHeader).forEach(h => {
    const col = map[h];
    if (!col) return;
    sh.getRange(1, col).setNote(String(byHeader[h]||''));
  });
}

function styleSheetHeader_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return;
  const header = sh.getRange(1,1,1,lastCol);
  header.setBackground('#f1f3f4').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  try { sh.setRowHeights(1, 1, 28); } catch(_) {}
  // Banding for data rows (start from row 2 to keep header style)
  const lastRow = sh.getLastRow();
  if (lastRow >= 2) {
  const dataRange = sh.getRange(2,1,lastRow-1,lastCol);
  // Remove existing bandings to avoid fragmentation and re-apply zebra
  try { (sh.getBandings() || []).forEach(b => b.remove()); } catch(_) {}
  try { dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY).setHeaderRowColor(null); } catch(_) {}
  }
  // Create filter on data range
  try { sh.getFilter() && sh.getFilter().remove(); } catch(_) {}
  try { if (sh.getLastRow() >= 1 && lastCol >= 1) sh.getRange(1,1,Math.max(1, sh.getLastRow()), lastCol).createFilter(); } catch(_) {}
}

function styleBalanceSheet_(sh) {
  sh.setFrozenColumns(2);
  const lastRow = Math.max(sh.getLastRow(), 2);
  const lastCol = sh.getLastColumn();
  // Number formats
  if (lastRow >= 2) {
    // C:F numbers
    sh.getRange(2,3,lastRow-1,4).setNumberFormat('#,##0.00');
  }
  // Conditional formatting: C>0 green, F>0 red
  const rules = sh.getConditionalFormatRules();
  const rngC = sh.getRange(2,3,lastRow-1,1);
  const rngF = sh.getRange(2,6,lastRow-1,1);
  rules.push(
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#e6f4ea').setRanges([rngC]).build()
  );
  rules.push(
    SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#fce8e6').setRanges([rngF]).build()
  );
  sh.setConditionalFormatRules(rules);
  // Align columns: IDs center, names left, numbers right
  if (lastRow >= 1 && lastCol >= 2) {
    sh.getRange(2,1,lastRow-1,1).setHorizontalAlignment('center');
    sh.getRange(2,2,lastRow-1,1).setHorizontalAlignment('left');
    sh.getRange(2,3,lastRow-1,4).setHorizontalAlignment('right');
  }
}

function stylePaymentsSheet_(sh) {
  sh.setFrozenColumns(2);
  const map = getHeaderMap_(sh);
  const lastRow = Math.max(sh.getLastRow(), 2);
  const lastCol = sh.getLastColumn();
  // Date format
  if (map['Дата']) sh.getRange(2, map['Дата'], lastRow-1, 1).setNumberFormat('yyyy-mm-dd');
  // Amount format
  if (map['Сумма']) sh.getRange(2, map['Сумма'], lastRow-1, 1).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
  // Align ID center
  if (map['payment_id']) sh.getRange(2, map['payment_id'], lastRow-1, 1).setHorizontalAlignment('center');
}

function styleCollectionsSheet_(sh) {
  sh.setFrozenColumns(2);
  const map = getHeaderMap_(sh);
  const lastRow = Math.max(sh.getLastRow(), 2);
  // Currency-like numbers
  if (map['Параметр суммы']) sh.getRange(2, map['Параметр суммы'], lastRow-1, 1).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
  if (map['Фиксированный x']) sh.getRange(2, map['Фиксированный x'], lastRow-1, 1).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
  // Dates
  if (map['Дата начала']) sh.getRange(2, map['Дата начала'], lastRow-1, 1).setNumberFormat('yyyy-mm-dd');
  if (map['Дедлайн'])     sh.getRange(2, map['Дедлайн'], lastRow-1, 1).setNumberFormat('yyyy-mm-dd');
  // ID center
  if (map['collection_id']) sh.getRange(2, map['collection_id'], lastRow-1, 1).setHorizontalAlignment('center');
  // Conditional formatting for Статус
  if (map['Статус'] && lastRow > 1) {
    const rules = sh.getConditionalFormatRules();
    const rng = sh.getRange(2, map['Статус'], lastRow-1, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Открыт').setBackground('#e6f4ea').setRanges([rng]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Закрыт').setBackground('#eeeeee').setRanges([rng]).build());
    sh.setConditionalFormatRules(rules);
  }
  // Align procurement fields
  if (map['Возмещено']) sh.getRange(2, map['Возмещено'], lastRow-1, 1).setHorizontalAlignment('center');
  if (map['Закупка из средств']) sh.getRange(2, map['Закупка из средств'], lastRow-1, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

function styleFamiliesSheet_(sh) {
  sh.setFrozenColumns(1);
  const map = getHeaderMap_(sh);
  const lastRow = Math.max(sh.getLastRow(), 2);
  if (map['День рождения']) sh.getRange(2, map['День рождения'], lastRow-1, 1).setNumberFormat('yyyy-mm-dd');
  if (map['Активен']) sh.getRange(2, map['Активен'], lastRow-1, 1).setHorizontalAlignment('center');
  if (map['family_id']) sh.getRange(2, map['family_id'], lastRow-1, 1).setHorizontalAlignment('center');
  // Conditional formatting for Активен
  if (map['Активен'] && lastRow > 1) {
    const rules = sh.getConditionalFormatRules();
    const rng = sh.getRange(2, map['Активен'], lastRow-1, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Да').setBackground('#e6f4ea').setRanges([rng]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Нет').setBackground('#fce8e6').setRanges([rng]).build());
    sh.setConditionalFormatRules(rules);
  }
}

function styleParticipationSheet_(sh) {
  const map = getHeaderMap_(sh);
  const lastRow = Math.max(sh.getLastRow(), 2);
  if (map['Статус']) sh.getRange(2, map['Статус'], lastRow-1, 1).setHorizontalAlignment('center');
  // Conditional formatting for участие
  if (map['Статус'] && lastRow > 1) {
    const rules = sh.getConditionalFormatRules();
    const rng = sh.getRange(2, map['Статус'], lastRow-1, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Участвует').setBackground('#e6f4ea').setRanges([rng]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('Не участвует').setBackground('#eeeeee').setRanges([rng]).build());
    sh.setConditionalFormatRules(rules);
  }
}

function styleDetailSheet_(sh) {
  sh.setFrozenColumns(2);
  const map = getHeaderMap_(sh);
  const lastRow = Math.max(sh.getLastRow(), 2);
  // Number formats
  if (map['Оплачено']) sh.getRange(2, map['Оплачено'], lastRow-1, 1).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
  if (map['Начислено']) sh.getRange(2, map['Начислено'], lastRow-1, 1).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
  if (map['Разность (±)']) sh.getRange(2, map['Разность (±)'], lastRow-1, 1).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
  // IDs center
  if (map['family_id']) sh.getRange(2, map['family_id'], lastRow-1, 1).setHorizontalAlignment('center');
  if (map['collection_id']) sh.getRange(2, map['collection_id'], lastRow-1, 1).setHorizontalAlignment('center');
  // Conditional formatting for difference: positive green, negative red
  if (map['Разность (±)'] && lastRow > 1) {
    const rules = sh.getConditionalFormatRules();
    const rng = sh.getRange(2, map['Разность (±)'], lastRow-1, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#e6f4ea').setRanges([rng]).build());
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground('#fce8e6').setRanges([rng]).build());
    sh.setConditionalFormatRules(rules);
  }
}

function styleSummarySheet_(sh) {
  sh.setFrozenColumns(2);
  const map = getHeaderMap_(sh);
  const lastRow = Math.max(sh.getLastRow(), 2);
  // Number formats
  ['Сумма цели','Собрано','Остаток до цели'].forEach(h => { if (map[h]) sh.getRange(2, map[h], lastRow-1, 1).setNumberFormat('#,##0.00').setHorizontalAlignment('right'); });
  ['Участников','Плательщиков','Единиц оплачено','Ещё плательщиков до закрытия'].forEach(h => { if (map[h]) sh.getRange(2, map[h], lastRow-1, 1).setNumberFormat('0').setHorizontalAlignment('center'); });
  if (map['collection_id']) sh.getRange(2, map['collection_id'], lastRow-1, 1).setHorizontalAlignment('center');
  // Conditional formatting for Остаток > 0
  if (map['Остаток до цели'] && lastRow > 1) {
    const rules = sh.getConditionalFormatRules();
    const rng = sh.getRange(2, map['Остаток до цели'], lastRow-1, 1);
    rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#fff4e5').setRanges([rng]).build());
    // NeedMore: 0 = green, >0 = orange
    if (map['Ещё плательщиков до закрытия']) {
      const rng2 = sh.getRange(2, map['Ещё плательщиков до закрытия'], lastRow-1, 1);
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberEqualTo(0).setBackground('#e6f4ea').setRanges([rng2]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#fff4e5').setRanges([rng2]).build());
    }
    // Section header shading when ALL: detect by text in "Название сбора"
    if (map['Название сбора']) {
      const nameColRng = sh.getRange(2, map['Название сбора'], lastRow-1, 1);
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('ОТКРЫТЫЕ СБОРЫ').setBackground('#e8f0fe').setRanges([nameColRng]).build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextContains('ЗАКРЫТЫЕ СБОРЫ').setBackground('#e8f0fe').setRanges([nameColRng]).build());
    }
    sh.setConditionalFormatRules(rules);
  }
}

function getSheetsSpec() {
  return [
    {
      name: 'Инструкция',
      headers: ['Шаг', 'Описание'],
      colWidths: [80, 1000]
    },
    {
      name: 'Семьи',
      headers: [
        'Ребёнок ФИО','День рождения',
        'Мама ФИО','Мама телефон','Мама реквизиты','Мама телеграм',
        'Папа ФИО','Папа телефон','Папа реквизиты','Папа телеграм',
        'Активен','Комментарий',
        'family_id'              // ID в конце: F001...
      ],
      colWidths: [220,110,220,140,240,160,220,140,240,160,90,260,110],
      dateCols: [2]
    },
    {
      name: 'Сборы',
      headers: [
        'Название сбора','Статус',
        'Дата начала','Дедлайн',
        'Начисление','Параметр суммы','Фиксированный x',
        'Закупка из средств','Возмещено',
        'Комментарий',
        'collection_id','Ссылка на гуглдиск'
      ],
  // Начисление: static_per_child | shared_total_all | shared_total_by_payers | dynamic_by_payers | proportional_by_payers | unit_price_by_payers
      colWidths: [260,120,110,110,220,150,140,200,110,260,120,300],
      dateCols: [3,4]
    },
    {
      name: 'Участие',
      headers: ['collection_id (label)','family_id (label)','Статус','Комментарий'],
      colWidths: [260,260,120,260]
    },
    {
      name: 'Платежи',
      headers: [
        'Дата','family_id (label)','collection_id (label)',
        'Сумма','Способ','Комментарий','payment_id'
      ],
      colWidths: [110,260,260,110,110,260,120],
      dateCols: [1]
    },
    {
      name: 'Баланс',
      headers: [
  'family_id','Имя ребёнка',
  'Переплата (текущая)','Оплачено всего','Начислено всего','Задолженность'
      ],
      colWidths: [120,260,140,140,120,130]
    },
    {
      name: 'Детализация',
      headers: [
        'family_id','Имя ребёнка','collection_id','Название сбора',
        'Оплачено','Начислено','Разность (±)','Режим'
      ],
      colWidths: [120,200,120,200,120,120,120,150]
    },
    {
      name: 'Сводка',
      headers: [
        'collection_id','Название сбора','Режим','Сумма цели','Собрано','Участников','Плательщиков','Единиц оплачено','Ещё плательщиков до закрытия','Остаток до цели'
      ],
      colWidths: [120,260,180,140,140,120,150,150,220,150]
    },
    {
      name: 'Lists', // hidden
      headers: [
        'OPEN_COLLECTIONS','', // A
        'COLLECTIONS','',      // C
        'ACTIVE_FAMILIES','',  // E
        'FAMILIES',''          // G
      ],
      colWidths: [260,40,260,40,260,40,260,40]
    }
  ];
}

function getOrCreateSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function setupInstructionSheet() {
  const sh = SpreadsheetApp.getActive().getSheetByName('Инструкция');
  // Clear old content under header
  const last = sh.getLastRow();
  if (last > 1) sh.getRange(2,1,last-1, Math.max(2, sh.getLastColumn())).clearContent();

  const rows = [
    ['▶ О проекте', 'Версия: 0.1. Репозиторий: https://github.com/yobushka/paymentAccountingGoogleSheet'],
    ['▶ Дисклеймер', 'Инструмент на ранней стадии и используется для личных целей; welcome to contribute. Внимание к персональным данным: передача ПДн через границу может быть незаконной. Google — иностранная компания; соблюдайте применимое законодательство.'],
    ['▶ Быстрый старт', '1) Funds → Setup / Rebuild structure.\n2) Заполните «Семьи» (Активен=Да).\n3) Добавьте «Сборы» (Статус=Открыт).\n4) При необходимости настройте «Участие».\n5) Вносите «Платежи».\n6) Смотрите «Баланс» и «Детализация».\n7) «Сводка» — оперативные итоги по сборам.'],
  ['1', 'Семьи: одна строка = одна семья (один ребёнок). Заполните ФИО, День рождения (yyyy-mm-dd), Телеграм мамы/папы и контакты родителей. Поставьте «Активен=Да», чтобы семья участвовала по умолчанию. ID генерируется автоматически при начале ввода или через меню Generate IDs.'],
  ['2', 'Сборы: выберите «Начисление» и задайте «Параметр суммы» (ставка/цель). «Фиксированный x»: для dynamic_by_payers — cap после закрытия, для unit_price_by_payers — цена за единицу. Можно указать «Ссылка на гуглдиск». Статус=Открыт — участвует в начислениях.'],
  ['2.1', 'Закупка из средств / Возмещено: при необходимости фиксируйте закупку из собранных средств и отмечайте, возмещена ли сумма. Поля справочные.'],
    ['3', 'Участие (опционально): если есть хотя бы один «Участвует», участвуют только отмеченные семьи. «Не участвует» всегда исключает семью. Если явных «Участвует» нет — участвуют все активные семьи.'],
    ['4', 'Платежи: выберите семью и сбор из выпадающих списков «Название (ID)». Сумма платежа должна быть > 0 (валидируется). Дата — справочная и на расчёты не влияет.'],
  ['5', 'Баланс: отображает по каждой семье «Оплачено всего», «Начислено всего», «Переплата (текущая)», «Задолженность».'],
    ['6', 'Демо-данные: Funds → Load Sample Data (separate) — добавит примеры семей, сборов, участия и платежей, чтобы увидеть механику сразу.'],

  ['▶ Пересчёт', 'Если сменили режим/участие/платежи, выполните Funds → Recalculate (Balance & Detail). Обновятся «Баланс», «Детализация» и «Сводка». Баланс также авто‑обновляется при правках на листах «Платежи», «Семьи», «Сборы».'],

    ['▶ Режимы начисления (подробно)', 'Все расчёты моментальные; поведение при 1/нескольких плательщиках:'],
    ['static_per_child', 'Фикс на семью. Начислено участнику = «Параметр суммы».\n1 плательщик: всем участникам начислена ставка; у плательщика возможна переплата.\nНесколько плательщиков: начисление одинаково у всех участников.'],
    ['shared_total_all', 'T/N на всех участников.\n1 плательщик: всем участникам начислено T/N; у плательщика возможна временная переплата.\nНесколько плательщиков: у всех одинаковое начисление = T/N.'],
    ['shared_total_by_payers', 'T/K только для оплативших.\n1 плательщик: начисление = T (K=1); будет недоплата, если внесено < T.\nНесколько плательщиков: каждому платившему начислено T/K; не платившие = 0.'],
    ['dynamic_by_payers', 'Water‑filling: Σ min(P_i, x) = min(T, ΣP_i). Начислено семье i = min(P_i, x).\n1 плательщик: начисление = его платёж (до T), долг не растёт.\nНесколько плательщиков: x выравнивает ранние переплаты; после закрытия используется «Фиксированный x».'],
  ['proportional_by_payers', 'Пропорционально платежам: начисление i = min(P_i, T) при Σ начислений = min(ΣP_i, T), распределение по долям P_i/ΣP. Пока не достигнута цель — списывается весь платёж. При превышении цели — суммы уменьшаются равнопропорционально. Долг не образуется.'],
  ['unit_price_by_payers', 'Поштучная закупка: цена за единицу x берётся из «Фиксированный x». Начисление i = floor(P_i/x)*x (только полные единицы) только тем, кто платил. Частичный остаток < x остаётся как переплата без долга. Суммарная цель T — общая стоимость партии. Число единиц = ceil(T/x).'],

    ['▶ Закрытие динамического сбора', 'Меню Funds → Close Collection. Введите collection_id (например, C003). Скрипт посчитает x (DYN_CAP) по фактическим платежам участников, запишет «Фиксированный x» и установит Статус=Закрыт. После закрытия используется зафиксированный x.'],
    ['DYN_CAP (формула)', 'DYN_CAP(T, payments_range) возвращает cap x по water-filling.\nПример: =DYN_CAP(6000, {2000,2000,700,700,700,700,700}) → 1250.'],

  ['▶ Формулы и примеры', 'Баланс: D — Оплачено всего; E — Начислено всего.\nПримеры: =ACCRUED_FAMILY(A2,"ALL") — по всем сборам; =ACCRUED_FAMILY(A2) — только по открытым. LABEL_TO_ID("Имя (F001)") → F001.'],

    ['▶ Выпадающие списки и ID', 'Выпадающие всегда показывают «Название (ID)». Внутри расчётов ID извлекается автоматически. Пустые ID генерируются при начале ввода или через меню «Generate IDs (all sheets)».'],

  ['▶ Советы и проверка', 'Если дропдауны пустые — Funds → Rebuild data validations.\nЕсли «Начислено» неожиданно 0 — проверьте «Участие» и «Активен».\nЕсли «Баланс» не обновился — внесите/измените запись в «Платежи» или запустите Setup.\nДля чистки лишних строк/колонок и обновления внешнего вида: Funds → Cleanup visuals (trim sheets).']
  ];
  sh.getRange(2,1,rows.length,2).setValues(rows);
  // Wrap text and align
  sh.getRange(2,2,rows.length,1).setWrap(true).setVerticalAlignment('top');
  // Emphasize section headers
  rows.forEach((r, i) => {
    if (String(r[0]||'').slice(0,1) === '▶') {
      sh.getRange(2+i, 1, 1, 2).setFontWeight('bold');
    }
  });
}

/** Hidden Lists: build label-form lists "Name (ID)" */
function setupListsSheet() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Lists');
  const shC = ss.getSheetByName('Сборы');
  const shF = ss.getSheetByName('Семьи');
  const mapC = getHeaderMap_(shC);
  const mapF = getHeaderMap_(shF);
  const cNameCol = colToLetter_(mapC['Название сбора']||2);
  const cIdCol   = colToLetter_(mapC['collection_id']||1);
  const cStatusCol = colToLetter_(mapC['Статус']||3);
  const fNameCol = colToLetter_(mapF['Ребёнок ФИО']||2);
  const fIdCol   = colToLetter_(mapF['family_id']||1);
  const fActiveCol = colToLetter_(mapF['Активен']||10);
  // OPEN_COLLECTIONS labels in A2:A  (Name (ID) for open only)
  sh.getRange('A1').setValue('OPEN_COLLECTIONS');
  sh.getRange('A2').setFormula(
    `=IFERROR(FILTER(ARRAYFORMULA(Сборы!${cNameCol}2:${cNameCol} & " (" & Сборы!${cIdCol}2:${cIdCol} & ")"), Сборы!${cStatusCol}2:${cStatusCol}="Открыт"),)`
  );
  // All COLLECTIONS labels in B2:B
  sh.getRange('B1').setValue('COLLECTIONS');
  sh.getRange('B2').setFormula(
    `=IFERROR(FILTER(ARRAYFORMULA(Сборы!${cNameCol}2:${cNameCol} & " (" & Сборы!${cIdCol}2:${cIdCol} & ")"), LEN(Сборы!${cIdCol}2:${cIdCol})),)`
  );
  // ACTIVE_FAMILIES labels in C2:C
  sh.getRange('C1').setValue('ACTIVE_FAMILIES');
  sh.getRange('C2').setFormula(
    `=IFERROR(FILTER(ARRAYFORMULA(Семьи!${fNameCol}2:${fNameCol} & " (" & Семьи!${fIdCol}2:${fIdCol} & ")"), Семьи!${fActiveCol}2:${fActiveCol}="Да"),)`
  );
  // All FAMILIES labels in D2:D
  sh.getRange('D1').setValue('FAMILIES');
  sh.getRange('D2').setFormula(
    `=IFERROR(FILTER(ARRAYFORMULA(Семьи!${fNameCol}2:${fNameCol} & " (" & Семьи!${fIdCol}2:${fIdCol} & ")"), LEN(Семьи!${fIdCol}2:${fIdCol})),)`
  );
  sh.hideSheet();
}

function ensureNamedRange(name, a1) {
  const ss = SpreadsheetApp.getActive();
  const existing = ss.getNamedRanges().find(n => n.getName() === name);
  const rng = ss.getRange(a1);
  if (existing) existing.setRange(rng); else ss.setNamedRange(name, rng);
}

// Header helpers and ID utilities
function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1,1,1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[String(h||'').trim()] = i+1; });
  return map;
}
function colToLetter_(col){ let s=""; let c=col; while(c>0){ let r=(c-1)%26; s=String.fromCharCode(65+r)+s; c=Math.floor((c-1)/26);} return s; }
function setRawIdNamedRanges_(){
  const ss = SpreadsheetApp.getActive();
  const shF = ss.getSheetByName('Семьи');
  const shC = ss.getSheetByName('Сборы');
  if (!shF || !shC) return;
  const mapF = getHeaderMap_(shF);
  const mapC = getHeaderMap_(shC);
  const fIdCol = colToLetter_(mapF['family_id']||1);
  const cIdCol = colToLetter_(mapC['collection_id']||1);
  ensureNamedRange('FAMILIES',    `Семьи!${fIdCol}2:${fIdCol}`);
  ensureNamedRange('COLLECTIONS', `Сборы!${cIdCol}2:${cIdCol}`);
}
function maybeAutoIdRow_(sh, row, idHeader, prefix, width, triggerHeaders){
  if (row < 2) return;
  const map = getHeaderMap_(sh);
  const idCol = map[idHeader];
  if (!idCol) return;
  const idVal = sh.getRange(row, idCol).getValue();
  if (idVal) return; // already set
  const hasTrigger = (triggerHeaders||[]).some(h => {
    const c = map[h]; if (!c) return false; const v = sh.getRange(row, c).getValue(); return v !== '' && v !== null;
  });
  if (!hasTrigger) return;
  const ss = SpreadsheetApp.getActive();
  fillMissingIds_(ss, sh.getName(), idCol, prefix, width);
}

/** =========================
 *  VALIDATIONS & BALANCE
 *  ========================= */
function rebuildValidations() {
  const ss = SpreadsheetApp.getActive();
  const lists = {
    statusOpenClosed: ['Открыт','Закрыт'],
    activeYesNo:      ['Да','Нет'],
  accrualRules:     ['static_per_child','shared_total_all','shared_total_by_payers','dynamic_by_payers','proportional_by_payers','unit_price_by_payers'],
    payMethods:       ['СБП','карта','наличные'],
    partStatus:       ['Участвует','Не участвует']
  };

  // Семьи: Активен
  const shF = ss.getSheetByName('Семьи');
  const mapF = getHeaderMap_(shF);
  if (mapF['Активен']) setValidationList(shF, 2, mapF['Активен'], lists.activeYesNo);

  // Сборы: Статус, Начисление
  const shC = ss.getSheetByName('Сборы');
  const mapC = getHeaderMap_(shC);
  if (mapC['Статус']) setValidationList(shC, 2, mapC['Статус'], lists.statusOpenClosed);
  if (mapC['Начисление']) setValidationList(shC, 2, mapC['Начисление'], lists.accrualRules);
  // Сборы: Возмещено (Да/Нет)
  if (mapC['Возмещено']) setValidationList(shC, 2, mapC['Возмещено'], lists.activeYesNo);

  // Участие: A=open collections labels, B=active families labels, C=Статус
  const shU = ss.getSheetByName('Участие');
  const mapU = getHeaderMap_(shU);
  if (mapU['collection_id (label)']) setValidationNamedRange(shU, 2, mapU['collection_id (label)'], 'OPEN_COLLECTIONS_LABELS');
  if (mapU['family_id (label)'])     setValidationNamedRange(shU, 2, mapU['family_id (label)'],     'ACTIVE_FAMILIES_LABELS');
  if (mapU['Статус'])                 setValidationList(shU, 2, mapU['Статус'], lists.partStatus);

  // Платежи: family label, collection label, Способ, Сумма > 0
  const shP = ss.getSheetByName('Платежи');
  const mapP = getHeaderMap_(shP);
  if (mapP['family_id (label)'])     setValidationNamedRange(shP, 2, mapP['family_id (label)'],     'FAMILIES_LABELS');
  if (mapP['collection_id (label)']) setValidationNamedRange(shP, 2, mapP['collection_id (label)'], 'COLLECTIONS_LABELS');
  if (mapP['Способ'])                 setValidationList(shP, 2, mapP['Способ'], lists.payMethods);
  if (mapP['Сумма'])                  setValidationNumberGreaterThan(shP, 2, mapP['Сумма'], 0);

  SpreadsheetApp.getActive().toast('Validations rebuilt.', 'Funds');
}

function setValidationList(sh, rowStart, col, values) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(rowStart, col, sh.getMaxRows() - rowStart + 1, 1).setDataValidation(rule);
}
function setValidationNamedRange(sh, rowStart, col, namedRange) {
  const ss = SpreadsheetApp.getActive();
  const nr = ss.getRangeByName(namedRange);
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(nr, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(rowStart, col, sh.getMaxRows() - rowStart + 1, 1).setDataValidation(rule);
}
function setValidationNumberGreaterThan(sh, rowStart, col, minValue) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(minValue)
    .setAllowInvalid(false)
    .build();
  sh.getRange(rowStart, col, sh.getMaxRows() - rowStart + 1, 1).setDataValidation(rule);
}

function setupBalanceExamples() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Баланс');

  // A2: список family_id из «Семьи» (ограниченный диапазон)
  const shF = ss.getSheetByName('Семьи');
  const mapF = getHeaderMap_(shF);
  const idCol = colToLetter_(mapF['family_id']||1);
  const nameCol = colToLetter_(mapF['Ребёнок ФИО']||2);
  const famLastRow = shF.getLastRow();
  
  // Limit ARRAYFORMULA to actual data range instead of open-ended
  if (famLastRow > 1) {
    sh.getRange('A2').setFormula(`=ARRAYFORMULA(IFERROR(FILTER(Семьи!${idCol}2:${idCol}${famLastRow}, LEN(Семьи!${idCol}2:${idCol}${famLastRow})), ))`);
  // Use array literal to ensure lookup table is [ID, Name] left-to-right even if ID column is to the right of Name
  sh.getRange('B2').setFormula(`=ARRAYFORMULA(IF(LEN(A2:A)=0, "", IFERROR(VLOOKUP(A2:A, {Семьи!${idCol}2:${idCol}${famLastRow}, Семьи!${nameCol}2:${nameCol}${famLastRow}}, 2, FALSE), "")))`);
  }

  // Селектор фильтра начислений: OPEN | ALL (по умолчанию — ALL)
  sh.getRange('H1').setValue('Фильтр начисления');
  sh.getRange('I1').setValue('ALL');
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(['OPEN','ALL'], true).setAllowInvalid(false).build();
  sh.getRange('I1').setDataValidation(rule).setHorizontalAlignment('center');
  sh.getRange('I1').setNote('Выберите OPEN (только открытые) или ALL (все сборы).');

  // Протянуть формулы по строкам для C:F на текущее число семей
  refreshBalanceFormulas_();

  sh.getRange('H3').setValue('Примечание: даты платёжек используются только для справки (фильтры/отчёты). Расчёты мгновенные.');
  
  // Setup detail sheet
  setupDetailSheet_();
  // Setup summary sheet
  setupSummarySheet_();
}

function refreshBalanceFormulas_() {
  const ss = SpreadsheetApp.getActive();
  const shBal = ss.getSheetByName('Баланс');
  const shFam = ss.getSheetByName('Семьи');
  const last = shFam.getLastRow();
  const famCount = Math.max(0, last - 1); // без заголовка

  // Re-apply A2/B2 formulas (IDs and Names) to ensure correct lookup after structure changes
  if (last > 1) {
    const mapF = getHeaderMap_(shFam);
    const idColLetter = colToLetter_(mapF['family_id']||1);
    const nameColLetter = colToLetter_(mapF['Ребёнок ФИО']||2);
    const famLastRow = last;
    shBal.getRange('A2').setFormula(`=ARRAYFORMULA(IFERROR(FILTER(Семьи!${idColLetter}2:${idColLetter}${famLastRow}, LEN(Семьи!${idColLetter}2:${idColLetter}${famLastRow})), ))`);
    shBal.getRange('B2').setFormula(`=ARRAYFORMULA(IF(LEN(A2:A)=0, "", IFERROR(VLOOKUP(A2:A, {Семьи!${idColLetter}2:${idColLetter}${famLastRow}, Семьи!${nameColLetter}2:${nameColLetter}${famLastRow}}, 2, FALSE), "")))`);
  }
  
  // Clear old formulas beyond actual data first
  const currentLastRow = shBal.getLastRow();
  if (currentLastRow > 1) {
    // Clear all formula columns completely
    shBal.getRange(2, 3, currentLastRow - 1, 4).clearContent();
  }
  
  if (famCount === 0) return;

  // Only create formulas for actual families (much more efficient)
  const rows = famCount;
  const formulasC = []; // текущая переплата = MAX(0, Оплачено - Списано)
  const formulasD = []; // Оплачено всего
  const formulasE = []; // списано (начислено)
  const formulasF = []; // Задолженность = MAX(0, Списано - Оплачено)
  
  for (let i = 0; i < rows; i++) {
    const r = 2 + i;
    // D: оплачено
    formulasD.push([`=IFERROR(PAYED_TOTAL_FAMILY($A${r}),0)`]);
    // E: начислено/списано (with selector)
    formulasE.push([`=IFERROR(ACCRUED_FAMILY($A${r}, IF(LEN($I$1)=0, "ALL", $I$1)), 0)`]);
    // C: текущая переплата
    formulasC.push([`=MAX(0, D${r} - E${r})`]);
    // F: задолженность
    formulasF.push([`=MAX(0, E${r} - D${r})`]);
  }
  
  // Set formulas only for actual family rows
  shBal.getRange(2, 3, rows, 1).setFormulas(formulasC);
  shBal.getRange(2, 4, rows, 1).setFormulas(formulasD);
  shBal.getRange(2, 5, rows, 1).setFormulas(formulasE);
  shBal.getRange(2, 6, rows, 1).setFormulas(formulasF);

  // Ensure formulas materialize before styling, then re-apply zebra and column styles
  SpreadsheetApp.flush();
  try { styleSheetHeader_(shBal); styleBalanceSheet_(shBal); } catch(_) {}
}

function setupDetailSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Детализация');
  if (!sh) return;
  
  // Clear old data
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow-1, sh.getLastColumn()).clearContent();
  
  // Selector for status filter
  sh.getRange('J1').setValue('Фильтр');
  sh.getRange('K1').setValue('ALL');
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(['OPEN','ALL'], true).setAllowInvalid(false).build();
  sh.getRange('K1').setDataValidation(rule).setHorizontalAlignment('center');
  sh.getRange('K1').setNote('OPEN (только открытые) или ALL (все сборы)');
  // Tick cell to force recalc on demand
  sh.getRange('J2').setValue('Tick');
  sh.getRange('K2').setValue(new Date().toISOString());
  sh.getRange('J3').setValue('Детализация платежей и начислений. Автообновляется и может быть принудительно обновлена через Tick.');
  
  // Dynamic formulas starting from A2
  sh.getRange('A2').setFormula(`=GENERATE_DETAIL_BREAKDOWN(IF(LEN($K$1)=0,"ALL",$K$1), $K$2)`);
}

function refreshDetailSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Детализация');
  if (!sh) return;
  
  // Trigger recalculation by touching the formula cell
  const current = sh.getRange('A2').getFormula();
  if (current.includes('GENERATE_DETAIL_BREAKDOWN')) {
    // Update tick to force recalculation
    sh.getRange('K2').setValue(new Date().toISOString());
  sh.getRange('A2').setFormula(current);
  SpreadsheetApp.flush();
  try { styleSheetHeader_(sh); styleDetailSheet_(sh); } catch(_) {}
  }
}

function setupSummarySheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Сводка');
  if (!sh) return;
  // Clear old data
  const lastRow = sh.getLastRow();
  if (lastRow > 1) sh.getRange(2, 1, lastRow-1, sh.getLastColumn()).clearContent();
  // Selector and tick
  sh.getRange('J1').setValue('Фильтр');
  sh.getRange('K1').setValue('ALL');
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(['OPEN','ALL'], true).setAllowInvalid(false).build();
  sh.getRange('K1').setDataValidation(rule).setHorizontalAlignment('center');
  sh.getRange('K1').setNote('OPEN (только открытые) или ALL (все сборы, сначала открытые, ниже — закрытые)');
  sh.getRange('J2').setValue('Tick');
  sh.getRange('K2').setValue(new Date().toISOString());
  // Array formula
  sh.getRange('A2').setFormula(`=GENERATE_COLLECTION_SUMMARY(IF(LEN($K$1)=0,"ALL",$K$1), $K$2)`);
  sh.getRange('J3').setValue('Сводка по сборам. ALL: сверху открытые, внизу закрытые (история).');
  SpreadsheetApp.flush();
  try { styleSheetHeader_(sh); styleSummarySheet_(sh); } catch(_) {}
}

function refreshSummarySheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName('Сводка');
  if (!sh) return;
  const current = sh.getRange('A2').getFormula();
  if (current.includes('GENERATE_COLLECTION_SUMMARY')) {
    sh.getRange('K2').setValue(new Date().toISOString());
    sh.getRange('A2').setFormula(current);
    // Re-apply styles to ensure alternating colors and header shading persist after rebuild
    try {
  SpreadsheetApp.flush();
      styleSheetHeader_(sh);
      styleSummarySheet_(sh);
    } catch (e) {}
  }
}

/** Manual recalculation entry point */
function recalculateAll() {
  try {
    refreshBalanceFormulas_();
  // bump detail tick to force recalculation
  const sh = SpreadsheetApp.getActive().getSheetByName('Детализация');
  if (sh) sh.getRange('K2').setValue(new Date().toISOString());
    refreshDetailSheet_();
  // bump summary tick and refresh
  const sh2 = SpreadsheetApp.getActive().getSheetByName('Сводка');
  if (sh2) sh2.getRange('K2').setValue(new Date().toISOString());
  refreshSummarySheet_();
  SpreadsheetApp.getActive().toast('Balance, Detail and Summary recalculated.', 'Funds');
  SpreadsheetApp.getUi().alert('Пересчёт завершён', 'Обновлены: «Баланс», «Детализация», «Сводка».', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    toastErr_('Recalculate failed: ' + e.message);
  SpreadsheetApp.getUi().alert('Ошибка пересчёта', String(e && e.message ? e.message : e), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/** =========================
 *  ID GENERATION & CLOSING
 *  ========================= */
function generateAllIds() {
  const ss = SpreadsheetApp.getActive();
  const plan = [
    {sheet: 'Семьи',   idHeader: 'family_id',    prefix: 'F',   width: 3},
    {sheet: 'Сборы',   idHeader: 'collection_id',prefix: 'C',   width: 3},
    {sheet: 'Платежи', idHeader: 'payment_id',   prefix: 'PMT', width: 3}
  ];
  plan.forEach(p => {
    const sh = ss.getSheetByName(p.sheet);
    const map = getHeaderMap_(sh);
    const col = map[p.idHeader] || 1;
    fillMissingIds_(ss, p.sheet, col, p.prefix, p.width);
  });
  SpreadsheetApp.getActive().toast('IDs generated where empty.', 'Funds');
  // Ensure Balance formulas cover current families
  refreshBalanceFormulas_();
}

function fillMissingIds_(ss, sheetName, idCol, prefix, padWidth) {
  const sh = ss.getSheetByName(sheetName);
  const last = sh.getLastRow();
  if (last < 2) return;
  const rng = sh.getRange(2, idCol, last - 1, 1);
  const vals = rng.getValues().map(r => r[0]);

  let maxNum = 0;
  vals.forEach(v => {
    if (typeof v === 'string' && v.startsWith(prefix)) {
      const n = parseInt(v.replace(prefix, ''), 10);
      if (!isNaN(n)) maxNum = Math.max(maxNum, n);
    }
  });

  const out = vals.slice();
  for (let i = 0; i < out.length; i++) {
    if (!out[i]) {
      maxNum += 1;
      out[i] = prefix + String(maxNum).padStart(padWidth, '0');
    }
  }
  rng.setValues(out.map(v => [v]));
}

function closeCollectionPrompt() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Close Collection', 'Введите collection_id (например, C001):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const collectionId = (resp.getResponseText() || '').trim();
  if (!collectionId) return;
  closeCollection_(collectionId);
}

function closeCollection_(collectionId) {
  const ss = SpreadsheetApp.getActive();
  const shC = ss.getSheetByName('Сборы');
  const shF = ss.getSheetByName('Семьи');
  const shU = ss.getSheetByName('Участие');
  const shP = ss.getSheetByName('Платежи');
  const mapC = getHeaderMap_(shC);
  const mapF = getHeaderMap_(shF);
  const mapU = getHeaderMap_(shU);
  const mapP = getHeaderMap_(shP);

  // Locate collection row by collection_id
  const colIdCol = mapC['collection_id'];
  if (!colIdCol) return toastErr_('Не найден столбец collection_id.');
  const rowsC = shC.getLastRow();
  if (rowsC < 2) return toastErr_('Нет сборов.');
  const ids = shC.getRange(2, colIdCol, rowsC - 1, 1).getValues().map(r => String(r[0]||'').trim());
  const idx = ids.findIndex(v => v === collectionId);
  if (idx === -1) return toastErr_('Сбор не найден: ' + collectionId);
  const rowNum = 2 + idx;

  // Read needed fields
  const accrual = String(shC.getRange(rowNum, mapC['Начисление']).getValue()||'').trim();
  const paramT  = Number(shC.getRange(rowNum, mapC['Параметр суммы']).getValue()||0);

  // Build active families set
  const famActiveCol = mapF['Активен'];
  const famIdCol     = mapF['family_id'];
  const famRows = shF.getLastRow();
  const activeSet = new Set();
  if (famRows >= 2 && famActiveCol && famIdCol) {
    const vals = shF.getRange(2, 1, famRows - 1, shF.getLastColumn()).getValues();
    const headers = shF.getRange(1,1,1,shF.getLastColumn()).getValues()[0];
    const hmap = {};
    headers.forEach((h,i)=>hmap[h]=i);
    const iId = hmap['family_id'];
    const iAct = hmap['Активен'];
    vals.forEach(r=>{ const id=String(r[iId]||'').trim(); const act=String(r[iAct]||'').trim()==='Да'; if(id&&act) activeSet.add(id); });
  }

  // Participation map for this collection
  const partInclude = new Set();
  const partExclude = new Set();
  let hasInclude = false;
  const uRows = shU.getLastRow();
  if (uRows >= 2) {
    const uVals = shU.getRange(2, 1, uRows - 1, shU.getLastColumn()).getValues();
    const uHeaders = shU.getRange(1,1,1,shU.getLastColumn()).getValues()[0];
    const ui = {}; uHeaders.forEach((h,i)=>ui[h]=i);
    uVals.forEach(r=>{
      const c = getIdFromLabelish_(String(r[ui['collection_id (label)']]||''));
      if (c !== collectionId) return;
      const f = getIdFromLabelish_(String(r[ui['family_id (label)']]||''));
      const st = String(r[ui['Статус']]||'').trim();
      if (!f) return;
      if (st === 'Участвует') { hasInclude = true; partInclude.add(f); }
      else if (st === 'Не участвует') { partExclude.add(f); }
    });
  }
  // Resolve participants
  const participants = new Set();
  if (hasInclude) partInclude.forEach(f=>participants.add(f));
  else activeSet.forEach(f=>participants.add(f));
  partExclude.forEach(f=>participants.delete(f));

  // Payments for this collection (only participating)
  const paymentsByFam = new Map();
  const pRows = shP.getLastRow();
  if (pRows >= 2) {
    const pVals = shP.getRange(2, 1, pRows - 1, shP.getLastColumn()).getValues();
    const pHeaders = shP.getRange(1,1,1,shP.getLastColumn()).getValues()[0];
    const pi = {}; pHeaders.forEach((h,i)=>pi[h]=i);
    pVals.forEach(r=>{
      const cid = getIdFromLabelish_(String(r[pi['collection_id (label)']]||''));
      if (cid !== collectionId) return;
      const fid = getIdFromLabelish_(String(r[pi['family_id (label)']]||''));
      const sum = Number(r[pi['Сумма']]||0);
      if (!fid || sum <= 0) return;
      if (!participants.has(fid)) return;
      paymentsByFam.set(fid, (paymentsByFam.get(fid)||0) + sum);
    });
  }
  const payments = Array.from(paymentsByFam.values());
  // Compute and write back only for dynamic_by_payers; for других режимов не трогаем «Фиксированный x»
  if (accrual === 'dynamic_by_payers') {
    const x = DYN_CAP_(paramT, payments);
    if (mapC['Фиксированный x']) shC.getRange(rowNum, mapC['Фиксированный x']).setValue(x);
    if (mapC['Статус'])           shC.getRange(rowNum, mapC['Статус']).setValue('Закрыт');
    SpreadsheetApp.getActive().toast(`Сбор ${collectionId} закрыт. x=${x}`, 'Funds');
  } else {
    // For unit_price_by_payers and others: do not overwrite Fixed X; just close
    if (mapC['Статус'])           shC.getRange(rowNum, mapC['Статус']).setValue('Закрыт');
    SpreadsheetApp.getActive().toast(`Сбор ${collectionId} закрыт.`, 'Funds');
  }
}

/** =========================
 *  SAMPLE DATA (separate)
 *  ========================= */
function loadSampleDataPrompt() {
  const ui = SpreadsheetApp.getUi();
  const choice = ui.alert(
    'Load Sample Data',
    'Это добавит демонстрационные данные (семьи, сборы, участие, платежи). Существующие данные не стираются, но могут перемешаться с демо. Продолжить?',
    ui.ButtonSet.OK_CANCEL
  );
  if (choice !== ui.Button.OK) return;
  loadSampleData_();
  SpreadsheetApp.getActive().toast('Demo data added.', 'Funds');
  refreshBalanceFormulas_();
}

function loadSampleData_() {
  const ss = SpreadsheetApp.getActive();
  const shF = ss.getSheetByName('Семьи');
  const shC = ss.getSheetByName('Сборы');
  const shU = ss.getSheetByName('Участие');
  const shP = ss.getSheetByName('Платежи');
  const mapF = getHeaderMap_(shF);
  const mapC = getHeaderMap_(shC);
  const mapP = getHeaderMap_(shP);

  // Families (10 demo rows)
  const famStart = shF.getLastRow() + 1;
  // Order per headers: ['Ребёнок ФИО','День рождения','Мама ФИО','Мама телефон','Мама реквизиты','Мама телеграм','Папа ФИО','Папа телефон','Папа реквизиты','Папа телеграм','Активен','Комментарий','family_id']
  const famRows = [
    ['Иванов Иван', '2015-03-15', 'Иванова Анна','+7 900 000-00-01','****1111','@anna_ivanova','Иванов Пётр','+7 900 000-10-01','****2222','@petr_ivanov','Да','', ''],
    ['Петров Пётр', '2015-06-02', 'Петрова Мария','+7 900 000-00-02','****3333','@petrova_m','Петров Иван','+7 900 000-10-02','****4444','@ivan_petrov','Да','', ''],
    ['Сидорова Вера','2015-01-21','Сидорова Ольга','+7 900 000-00-03','****5555','@sidorova_olga','Сидоров Антон','+7 900 000-10-03','****6666','@sid_anton','Да','', ''],
    ['Кузнецов Артём','2015-12-11','Кузнецова Ирина','+7 900 000-00-04','****7777','@irina_kuz','Кузнецов Олег','+7 900 000-10-04','****8888','@oleg_kuz','Да','', ''],
    ['Смирнова Юля','2015-08-05','Смирнова Анна','+7 900 000-00-05','****9999','@anna_smir','Смирнов Роман','+7 900 000-10-05','****0001','@roman_smir','Да','', ''],
    ['Новикова Нина','2015-04-19','Новикова Оксана','+7 900 000-00-06','****0002','@oks_nov','Новиков Павел','+7 900 000-10-06','****0003','@pavel_nov','Да','', ''],
    ['Орлова Лена','2015-07-23','Орлова Татьяна','+7 900 000-00-07','****0004','@tat_orl','Орлов Юрий','+7 900 000-10-07','****0005','@y_orlov','Да','', ''],
    ['Фёдоров Даня','2015-02-14','Фёдорова Алла','+7 900 000-00-08','****0006','@alla_fed','Фёдоров Игорь','+7 900 000-10-08','****0007','@igor_fed','Да','', ''],
    ['Максимова Аня','2015-09-30','Максимова Ника','+7 900 000-00-09','****0008','@nika_maks','Максимов Артём','+7 900 000-10-09','****0009','@art_maks','Да','', ''],
    ['Егорова Саша','2015-11-01','Егорова Алина','+7 900 000-00-10','****0010','@alina_egor','Егоров Кирилл','+7 900 000-10-10','****0011','@kir_egor','Да','', '']
  ];
  shF.getRange(famStart, 1, famRows.length, shF.getLastColumn()).setValues(famRows);

  // Generate IDs for families
  if (mapF['family_id']) fillMissingIds_(ss, 'Семьи', mapF['family_id'], 'F', 3);

  // Collections (demo for all modes)
  const colStart = shC.getLastRow() + 1;
  // Current headers:
  // ['Название сбора','Статус','Дата начала','Дедлайн','Начисление','Параметр суммы','Фиксированный x','Закупка из средств','Возмещено','Комментарий','collection_id','Ссылка на гуглдиск']
  const colRows = [
    ['Канцтовары сентябрь', 'Открыт', '', '', 'static_per_child', 500,   '',         '',      '', 'Фикс 500₽ на семью',           '', ''],
    ['Новый год',           'Открыт', '', '', 'shared_total_all', 12000, '',         '',      '', 'Общая сумма делится на участников', '', ''],
    ['Подарок учителю',     'Открыт', '', '', 'dynamic_by_payers', 9000, '',         '',      '', 'Динамический сбор по цели 9000₽',   '', ''],
    ['Фотосессия',          'Открыт', '', '', 'shared_total_by_payers', 10000, '',   '',      '', 'Делим сумму между оплатившими',     '', ''],
    ['Помощь классу',       'Открыт', '', '', 'proportional_by_payers', 8000, '',    '',      '', 'Пропорционально платежам',         '', ''],
    ['Спортивная форма',    'Открыт', '', '', 'unit_price_by_payers', 15000, 1500,   '',      'Нет', 'Поштучная закупка: x=1500₽',      '', '']
  ];
  shC.getRange(colStart, 1, colRows.length, shC.getLastColumn()).setValues(colRows);

  // Generate IDs for collections
  if (mapC['collection_id']) fillMissingIds_(ss, 'Сборы', mapC['collection_id'], 'C', 3);

  // Refresh Lists (labels)
  setupListsSheet();

  // Build labels for newly added collections based on their actual IDs
  const newCount = colRows.length;
  const cVals = shC.getRange(colStart, 1, newCount, shC.getLastColumn()).getValues();
  const cHdr = shC.getRange(1,1,1,shC.getLastColumn()).getValues()[0];
  const ci = {}; cHdr.forEach((h,idx)=>ci[h]=idx);
  const labelByName = new Map();
  cVals.forEach(r => {
    const nm = String(r[ci['Название сбора']]||'').trim();
    const id = String(r[ci['collection_id']]||'').trim();
    if (nm && id) labelByName.set(nm, `${nm} (${id})`);
  });
  const c1Label = labelByName.get('Канцтовары сентябрь') || '';
  const c2Label = labelByName.get('Новый год') || '';
  const c3Label = labelByName.get('Подарок учителю') || '';
  const c4Label = labelByName.get('Фотосессия') || '';
  const c5Label = labelByName.get('Помощь классу') || '';
  const c6Label = labelByName.get('Спортивная форма') || '';

  // Families labels (all families)
  const allFam = getLabelColumn_('Lists', 'D', 2);

  const partStart = shU.getLastRow() + 1;
  const partRows = [];
  // C002: explicitly mark 8 families as "Участвует"
  allFam.slice(0,8).forEach(lbl => partRows.push([c2Label, lbl, 'Участвует', '']));
  // C003: exclude 2 families
  allFam.slice(0,2).forEach(lbl => partRows.push([c3Label, lbl, 'Не участвует', '']));
  if (partRows.length) {
    shU.getRange(partStart, 1, partRows.length, 4).setValues(partRows);
  }

  // Payments: mix across all collections
  const payStart = shP.getLastRow() + 1;
  const today = new Date();
  const addDays = (d) => new Date(today.getTime() + d*24*3600*1000);
  const payRows = [];

  // For C001 (static 500): 6 families pay full, 2 pay partial, 2 not yet
  allFam.slice(0,6).forEach((lbl,i) => payRows.push([toISO_(addDays(-5+i)), lbl, c1Label, 500, 'СБП', 'Полная оплата', '']));
  allFam.slice(6,8).forEach((lbl,i) => payRows.push([toISO_(addDays(-2-i)), lbl, c1Label, 300, 'карта', 'Частичная оплата', '']));

  // For C002 (shared 12000 among 8): 5 pay full share later, 3 partials
  const shareFamilies = allFam.slice(0,8);
  // let share = 12000 / shareFamilies.length; // расчёт в формуле
  shareFamilies.slice(0,5).forEach((lbl,i) => payRows.push([toISO_(addDays(-3+i)), lbl, c2Label, 1500, 'СБП', 'Частично/полностью', '']));
  shareFamilies.slice(5,8).forEach((lbl,i) => payRows.push([toISO_(addDays(-2-i)), lbl, c2Label, 800,  'наличные', 'Частично', '']));

  // For C003 (dynamic 9000, excluding 2 families): пример из README — платежи [2000,2000,700,700,700,700,700]
  const dynFamilies = allFam.slice(2); // первые двое исключены
  dynFamilies.slice(0,2).forEach((lbl,i) => payRows.push([toISO_(addDays(-6+i)), lbl, c3Label, 2000, 'СБП', 'Ранний платёж', '']));
  dynFamilies.slice(2,7).forEach((lbl,i) => payRows.push([toISO_(addDays(-1-i)), lbl, c3Label, 700,  'карта', 'Позже', '']));

  // For C004 (shared_total_by_payers 10000): 4 families pay; начисление будет T/K=2500 только им
  if (c4Label) {
    allFam.slice(0,4).forEach((lbl,i) => payRows.push([toISO_(addDays(-4+i)), lbl, c4Label, 2500, i%2? 'карта':'СБП', 'Оплата доли', '']));
  }

  // For C005 (proportional_by_payers 8000): 5 семей платят разными суммами (будет пропорциональное списание)
  if (c5Label) {
    const fams = allFam.slice(2,7);
    const amounts = [3000, 2000, 1500, 800, 500]; // суммарно 7800 < T
    fams.forEach((lbl, i) => payRows.push([toISO_(addDays(-2+i)), lbl, c5Label, amounts[i], i%2 ? 'карта' : 'СБП', 'Разные суммы', '']));
  }

  // For C006 (unit_price_by_payers T=15000, x=1500): демонстрация мульти-единиц у одного плательщика
  // Платежи: [1500,1500,1500,3000,4500,1500,700,700] → единиц оплачено = 9, частичные не формируют начисление
  if (c6Label) {
    const fams = allFam.slice(0,8);
    const amounts = [1500,1500,1500,3000,4500,1500,700,700];
    fams.forEach((lbl,i) => payRows.push([
      toISO_(addDays(-7+i)),
      lbl,
      c6Label,
      amounts[i],
      (i%2 ? 'карта' : 'СБП'),
      amounts[i] >= 1500 ? (amounts[i] % 1500 === 0 ? `${amounts[i]/1500} ед.` : 'Частично') : 'Частично',
      ''
    ]));
  }

  if (payRows.length) {
    shP.getRange(payStart, 1, payRows.length, shP.getLastColumn()).setValues(payRows);
  }

  // Generate IDs for payments
  if (mapP['payment_id']) fillMissingIds_(ss, 'Платежи', mapP['payment_id'], 'PMT', 3);

  // Rebuild validations (if status/active were added)
  rebuildValidations();
}

/** helpers for sample data */
function getLabelColumn_(sheetName, colLetter, startRow) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const last = sh.getLastRow();
  if (last < startRow) return [];
  const rng = sh.getRange(`${colLetter}${startRow}:${colLetter}${last}`);
  return rng.getValues().map(r => String(r[0]||'')).filter(Boolean);
}
function toISO_(d) {
  const pad = (n) => String(n).padStart(2,'0');
  return `${d.getFullYear()}-${pad(d.getMonth()+1)}-${pad(d.getDate())}`;
}

/** =========================
 *  CUSTOM FUNCTIONS
 *  ========================= */

// LABEL_TO_ID("Имя (F001)") -> "F001" ; LABEL_TO_ID("F001")->"F001"
function LABEL_TO_ID(value) {
  return getIdFromLabelish_(value);
}

// Debug helper: shows detailed calculation for a collection and family
function DEBUG_COLLECTION_ACCRUAL(collectionId, familyId) {
  const ss = SpreadsheetApp.getActive();
  const shC = ss.getSheetByName('Сборы');
  const shP = ss.getSheetByName('Платежи');
  const shF = ss.getSheetByName('Семьи');
  const shU = ss.getSheetByName('Участие');
  
  const mapC = getHeaderMap_(shC);
  const mapP = getHeaderMap_(shP);
  const mapF = getHeaderMap_(shF);
  const mapU = getHeaderMap_(shU);
  
  // Find collection
  const cRows = shC.getLastRow();
  let collectionData = null;
  if (cRows >= 2) {
    const C = shC.getRange(2, 1, cRows - 1, shC.getLastColumn()).getValues();
    const ch = shC.getRange(1,1,1,shC.getLastColumn()).getValues()[0];
    const ci={}; ch.forEach((h,idx)=>ci[h]=idx);
    for (const row of C) {
      if (String(row[ci['collection_id']]||'').trim() === collectionId) {
        collectionData = {
          id: collectionId,
          status: String(row[ci['Статус']]||'').trim(),
          accrual: String(row[ci['Начисление']]||'').trim(),
          paramT: Number(row[ci['Параметр суммы']]||0),
          fixedX: Number(row[ci['Фиксированный x']]||0)
        };
        break;
      }
    }
  }
  
  if (!collectionData) return 'Collection not found: ' + collectionId;
  
  // Get payments for this collection
  const payments = new Map();
  const pRows = shP.getLastRow();
  if (pRows >= 2) {
    const P = shP.getRange(2, 1, pRows - 1, shP.getLastColumn()).getValues();
    const ph = shP.getRange(1,1,1,shP.getLastColumn()).getValues()[0];
    const pi={}; ph.forEach((h,idx)=>pi[h]=idx);
    P.forEach(r=>{
      const col = getIdFromLabelish_(String(r[pi['collection_id (label)']]||''));
      const fam = getIdFromLabelish_(String(r[pi['family_id (label)']]||''));
      const sum = Number(r[pi['Сумма']]||0);
      if (col === collectionId && fam && sum > 0) {
        payments.set(fam, (payments.get(fam)||0) + sum);
      }
    });
  }
  
  const paymentArray = Array.from(payments.values());
  const familyPayment = payments.get(familyId) || 0;
  
  let result = `Collection: ${collectionId}\n`;
  result += `Mode: ${collectionData.accrual}\n`;
  result += `Target T: ${collectionData.paramT}\n`;
  result += `Fixed X: ${collectionData.fixedX}\n`;
  result += `Status: ${collectionData.status}\n`;
  result += `All payments: [${paymentArray.join(', ')}]\n`;
  result += `Family ${familyId} payment: ${familyPayment}\n`;
  
  if (collectionData.accrual === 'dynamic_by_payers') {
    const x = collectionData.fixedX > 0 ? collectionData.fixedX : DYN_CAP_(collectionData.paramT, paymentArray);
    result += `Calculated x: ${x}\n`;
    result += `Family accrual: min(${familyPayment}, ${x}) = ${Math.min(familyPayment, x)}\n`;
    
    // Verify total
    let total = 0;
    payments.forEach((pay) => total += Math.min(pay, x));
    result += `Total distributed: ${total}\n`;
  }
  
  return result;
}

// Sum of all payments by a family (all collections)
function PAYED_TOTAL_FAMILY(familyLabelOrId) {
  const famId = getIdFromLabelish_(familyLabelOrId);
  if (!famId) return 0;
  const ss = SpreadsheetApp.getActive();
  const shPay = ss.getSheetByName('Платежи');
  const rows = shPay.getLastRow();
  if (rows < 2) return 0;
  const map = getHeaderMap_(shPay);
  const iFam = map['family_id (label)'];
  const iSum = map['Сумма'];
  if (!iFam || !iSum) return 0;
  const vals = shPay.getRange(2, 1, rows - 1, shPay.getLastColumn()).getValues();
  let total = 0;
  vals.forEach(r => {
    const fid = getIdFromLabelish_(String(r[iFam-1]||''));
    const sum = Number(r[iSum-1]||0);
    if (fid === famId && sum > 0) total += sum;
  });
  return round2_(total);
}

/** Accrued total for a family across collections. statusFilter: "OPEN" (default) or "ALL" */
function ACCRUED_FAMILY(familyLabelOrId, statusFilter) {
  const famId = getIdFromLabelish_(familyLabelOrId);
  if (!famId) return 0;
  const onlyOpen = String(statusFilter||'OPEN').toUpperCase() !== 'ALL';
  const ss = SpreadsheetApp.getActive();
  const shF = ss.getSheetByName('Семьи');
  const shC = ss.getSheetByName('Сборы');
  const shU = ss.getSheetByName('Участие');
  const shP = ss.getSheetByName('Платежи');

  const mapF = getHeaderMap_(shF);
  const mapC = getHeaderMap_(shC);
  const mapU = getHeaderMap_(shU);
  const mapP = getHeaderMap_(shP);

  // Active families
  const famRows = shF.getLastRow();
  const activeFam = new Set();
  if (famRows >= 2) {
    const vals = shF.getRange(2, 1, famRows - 1, shF.getLastColumn()).getValues();
    const headers = shF.getRange(1,1,1,shF.getLastColumn()).getValues()[0];
    const i = {}; headers.forEach((h,idx)=>i[h]=idx);
    vals.forEach(r=>{
      const id = String(r[i['family_id']]||'').trim();
      const act = String(r[i['Активен']]||'').trim()==='Да';
      if (id && act) activeFam.add(id);
    });
  }

  // Participation
  const partByCol = new Map(); // colId -> {hasInclude, include:Set, exclude:Set}
  const uRows = shU.getLastRow();
  if (uRows >= 2) {
    const U = shU.getRange(2, 1, uRows - 1, shU.getLastColumn()).getValues();
    const uh = shU.getRange(1,1,1,shU.getLastColumn()).getValues()[0];
    const ui={}; uh.forEach((h,idx)=>ui[h]=idx);
    U.forEach(r=>{
      const col = getIdFromLabelish_(String(r[ui['collection_id (label)']]||''));
      const fam = getIdFromLabelish_(String(r[ui['family_id (label)']]||''));
      const st  = String(r[ui['Статус']]||'').trim();
      if (!col || !fam) return;
      if (!partByCol.has(col)) partByCol.set(col, {hasInclude:false, include:new Set(), exclude:new Set()});
      const obj = partByCol.get(col);
      if (st === 'Участвует') { obj.hasInclude = true; obj.include.add(fam); }
      else if (st === 'Не участвует') { obj.exclude.add(fam); }
    });
  }

  // Payments grouped: collection -> Map(fam->sum)
  const payByCol = new Map();
  const pRows = shP.getLastRow();
  if (pRows >= 2) {
    const P = shP.getRange(2, 1, pRows - 1, shP.getLastColumn()).getValues();
    const ph = shP.getRange(1,1,1,shP.getLastColumn()).getValues()[0];
    const pi={}; ph.forEach((h,idx)=>pi[h]=idx);
    P.forEach(r=>{
      const col = getIdFromLabelish_(String(r[pi['collection_id (label)']]||''));
      const fam = getIdFromLabelish_(String(r[pi['family_id (label)']]||''));
      const sum = Number(r[pi['Сумма']]||0);
      if (!col || !fam || sum <= 0) return;
      if (!payByCol.has(col)) payByCol.set(col, new Map());
      const m = payByCol.get(col);
      m.set(fam, (m.get(fam)||0) + sum);
    });
  }

  // Iterate collections
  let total = 0;
  const cRows = shC.getLastRow();
  if (cRows >= 2) {
    const C = shC.getRange(2, 1, cRows - 1, shC.getLastColumn()).getValues();
    const ch = shC.getRange(1,1,1,shC.getLastColumn()).getValues()[0];
    const ci={}; ch.forEach((h,idx)=>ci[h]=idx);
    C.forEach(row=>{
      const colId   = String(row[ci['collection_id']]||'').trim();
      const status  = String(row[ci['Статус']]||'').trim();
      const accrual = String(row[ci['Начисление']]||'').trim();
      const paramT  = Number(row[ci['Параметр суммы']]||0);
      const fixedX  = Number(row[ci['Фиксированный x']]||0);
      if (!colId) return;
      if (onlyOpen && status !== 'Открыт') return;

      // participants
      const p = partByCol.get(colId);
      const participants = new Set();
      if (p && p.hasInclude) p.include.forEach(f=>participants.add(f));
      else activeFam.forEach(f=>participants.add(f));
      if (p) p.exclude.forEach(f=>participants.delete(f));

      // Fallback: if no participants resolved (e.g., header mismatch), use payers for this collection
      const famPays = (payByCol.get(colId) || new Map());
      if (participants.size === 0) {
        famPays.forEach((_, fid)=>participants.add(fid));
      }

      const n = participants.size;
      const Pi = famPays.get(famId) || 0;

      let accrued = 0;
      if (accrual === 'static_per_child') {
        accrued = participants.has(famId) ? paramT : 0;
      } else if (accrual === 'shared_total_all') {
        if (n > 0 && participants.has(famId)) accrued = paramT / n;
      } else if (accrual === 'shared_total_by_payers') {
        // Share T equally among payers (within participants)
        const payers = [];
        (payByCol.get(colId) || new Map()).forEach((sum, fid)=>{ if (participants.has(fid) && sum>0) payers.push(fid); });
        const k = payers.length;
        if (k > 0 && participants.has(famId) && Pi > 0) accrued = paramT / k; else accrued = 0;
      } else if (accrual === 'dynamic_by_payers') {
        if (participants.has(famId) && n > 0) {
          // payments of participants only
          const payments = [];
          famPays.forEach((sum, fid)=>{ if (participants.has(fid) && sum>0) payments.push(sum); });
          const x = fixedX > 0 ? fixedX : DYN_CAP_(paramT, payments);
          accrued = Math.min(Pi, x);
        }
      } else if (accrual === 'proportional_by_payers') {
        // Accrue proportionally to payments among participants, capping total at T.
        if (participants.has(famId)) {
          // Sum of payments among participants
          let sumP = 0;
          famPays.forEach((sum, fid)=>{ if (participants.has(fid) && sum>0) sumP += sum; });
          if (sumP <= 0) {
            accrued = 0;
          } else {
            const target = Math.min(paramT, sumP);
            const ratio = target / sumP; // <= 1
            accrued = Pi > 0 ? (Pi * ratio) : 0;
          }
        }
      } else if (accrual === 'unit_price_by_payers') {
            // Per payer multiple units allowed: accrue full units only, floor(Pi/x)*x; partial < x remains как переплата без долга
            const x = fixedX > 0 ? fixedX : 0;
            if (participants.has(famId) && x > 0) {
              accrued = Math.floor(Pi / x) * x;
            } else {
              accrued = 0;
            }
      }
      total += accrued;
    });
  }
  return round2_(total);
}

/** Returns per-collection accrual breakdown for a family (for debugging). statusFilter: "OPEN" (default) or "ALL" */
function ACCRUED_BREAKDOWN(familyLabelOrId, statusFilter) {
  const famId = getIdFromLabelish_(familyLabelOrId);
  if (!famId) return [['collection_id','accrued']];
  const onlyOpen = String(statusFilter||'OPEN').toUpperCase() !== 'ALL';
  const ss = SpreadsheetApp.getActive();
  const shF = ss.getSheetByName('Семьи');
  const shC = ss.getSheetByName('Сборы');
  const shU = ss.getSheetByName('Участие');
  const shP = ss.getSheetByName('Платежи');

  const mapF = getHeaderMap_(shF);
  const mapC = getHeaderMap_(shC);
  const mapU = getHeaderMap_(shU);
  const mapP = getHeaderMap_(shP);

  // Active families
  const famRows = shF.getLastRow();
  const activeFam = new Set();
  if (famRows >= 2) {
    const vals = shF.getRange(2, 1, famRows - 1, shF.getLastColumn()).getValues();
    const headers = shF.getRange(1,1,1,shF.getLastColumn()).getValues()[0];
    const i = {}; headers.forEach((h,idx)=>i[h]=idx);
    vals.forEach(r=>{
      const id = String(r[i['family_id']]||'').trim();
      const act = String(r[i['Активен']]||'').trim()==='Да';
      if (id && act) activeFam.add(id);
    });
  }

  // Participation
  const partByCol = new Map();
  const uRows = shU.getLastRow();
  if (uRows >= 2) {
    const U = shU.getRange(2, 1, uRows - 1, shU.getLastColumn()).getValues();
    const uh = shU.getRange(1,1,1,shU.getLastColumn()).getValues()[0];
    const ui={}; uh.forEach((h,idx)=>ui[h]=idx);
    U.forEach(r=>{
      const col = getIdFromLabelish_(String(r[ui['collection_id (label)']]||''));
      const fam = getIdFromLabelish_(String(r[ui['family_id (label)']]||''));
      const st  = String(r[ui['Статус']]||'').trim();
      if (!col || !fam) return;
      if (!partByCol.has(col)) partByCol.set(col, {hasInclude:false, include:new Set(), exclude:new Set()});
      const obj = partByCol.get(col);
      if (st === 'Участвует') { obj.hasInclude = true; obj.include.add(fam); }
      else if (st === 'Не участвует') { obj.exclude.add(fam); }
    });
  }

  // Payments grouped: collection -> Map(fam->sum)
  const payByCol = new Map();
  const pRows = shP.getLastRow();
  if (pRows >= 2) {
    const P = shP.getRange(2, 1, pRows - 1, shP.getLastColumn()).getValues();
    const ph = shP.getRange(1,1,1,shP.getLastColumn()).getValues()[0];
    const pi={}; ph.forEach((h,idx)=>pi[h]=idx);
    P.forEach(r=>{
      const col = getIdFromLabelish_(String(r[pi['collection_id (label)']]||''));
      const fam = getIdFromLabelish_(String(r[pi['family_id (label)']]||''));
      const sum = Number(r[pi['Сумма']]||0);
      if (!col || !fam || sum <= 0) return;
      if (!payByCol.has(col)) payByCol.set(col, new Map());
      const m = payByCol.get(col);
      m.set(fam, (m.get(fam)||0) + sum);
    });
  }

  const out = [['collection_id','accrued']];
  const cRows = shC.getLastRow();
  if (cRows >= 2) {
    const C = shC.getRange(2, 1, cRows - 1, shC.getLastColumn()).getValues();
    const ch = shC.getRange(1,1,1,shC.getLastColumn()).getValues()[0];
    const ci={}; ch.forEach((h,idx)=>ci[h]=idx);
    C.forEach(row=>{
      const colId   = String(row[ci['collection_id']]||'').trim();
      const status  = String(row[ci['Статус']]||'').trim();
      const accrual = String(row[ci['Начисление']]||'').trim();
      const paramT  = Number(row[ci['Параметр суммы']]||0);
      const fixedX  = Number(row[ci['Фиксированный x']]||0);
      if (!colId) return;
      if (onlyOpen && status !== 'Открыт') return; // respect filter

      // participants
      const p = partByCol.get(colId);
      const participants = new Set();
      if (p && p.hasInclude) p.include.forEach(f=>participants.add(f));
      else activeFam.forEach(f=>participants.add(f));
      if (p) p.exclude.forEach(f=>participants.delete(f));

      // Fallback: payers as participants if empty
      const famPays = (payByCol.get(colId) || new Map());
      if (participants.size === 0) famPays.forEach((_, fid)=>participants.add(fid));

      const n = participants.size;
      const Pi = famPays.get(famId) || 0;
      let accrued = 0;
      if (accrual === 'static_per_child') {
        accrued = participants.has(famId) ? paramT : 0;
      } else if (accrual === 'shared_total_all') {
        if (n > 0 && participants.has(famId)) accrued = paramT / n;
      } else if (accrual === 'shared_total_by_payers') {
        const payers = [];
        famPays.forEach((sum, fid)=>{ if (participants.has(fid) && sum>0) payers.push(fid); });
        const k = payers.length;
        if (k > 0 && participants.has(famId) && (famPays.get(famId)||0) > 0) accrued = paramT / k; else accrued = 0;
      } else if (accrual === 'dynamic_by_payers') {
        if (participants.has(famId) && n > 0) {
          const payments = [];
          famPays.forEach((sum, fid)=>{ if (participants.has(fid) && sum>0) payments.push(sum); });
          const x = fixedX > 0 ? fixedX : DYN_CAP_(paramT, payments);
          accrued = Math.min(Pi, x);
        }
      }
      if (accrued !== 0) out.push([colId, round2_(accrued)]);
    });
  }
  return out;
}

/** =========================
 *  DYNAMIC CAP
 *  ========================= */
function DYN_CAP(T, payments_range) {
  if (T === null || T === '' || isNaN(T)) return 0;
  const flat = flatten_(payments_range).map(Number).filter(v => isFinite(v) && v > 0);
  if (!flat.length) return 0;
  flat.sort((a,b)=>a-b);
  const n = flat.length;
  const sum = flat.reduce((a,b)=>a+b,0);
  const target = Math.min(T, sum);
  if (target <= 0) return 0;
  let cumsum = 0;
  for (let k = 0; k < n; k++) {
    const next = flat[k];
    const remain = n - k;
    const candidate = (target - cumsum) / remain;
    if (candidate <= next) return round6_(candidate);
    cumsum += next;
  }
  return round6_((target - (cumsum - flat[n-1])) / 1);
}
function DYN_CAP_(T, payments) {
  if (!T || !isFinite(T)) return 0;
  const arr = (payments || []).map(Number).filter(v => v > 0 && isFinite(v));
  if (!arr.length) return 0;
  arr.sort((a,b)=>a-b);
  const n = arr.length;
  const sum = arr.reduce((a,b)=>a+b,0);
  const target = Math.min(T, sum);
  if (target <= 0) return 0;
  
  // Debug logging
  Logger.log(`DYN_CAP_: T=${T}, payments=[${arr.join(',')}], target=${target}`);
  
  let cumsum = 0;
  for (let k = 0; k < n; k++) {
    const next = arr[k];
    const remain = n - k;
    const candidate = (target - cumsum) / remain;
    Logger.log(`Step ${k}: next=${next}, remain=${remain}, candidate=${candidate}, cumsum=${cumsum}`);
    if (candidate <= next) {
      Logger.log(`Found x=${candidate}`);
      return round6_(candidate);
    }
    cumsum += next;
  }
  const final = round6_((target - (cumsum - arr[n-1])) / 1);
  Logger.log(`Final x=${final}`);
  return final;
}

/** =========================
 *  UTILS
 *  ========================= */
function getIdFromLabelish_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  const m = s.match(/\(([^)]+)\)\s*$/);
  return m ? m[1] : s; // if label "Name (ID)" -> ID; else assume it's already ID
}
function flatten_(arr){ const out=[];(arr||[]).forEach(r=>Array.isArray(r)?r.forEach(c=>out.push(c)):out.push(r));return out; }
// Debug what's being calculated in Balance
function DEBUG_BALANCE_ACCRUAL(familyId) {
  const ss = SpreadsheetApp.getActive();
  const shBal = ss.getSheetByName('Баланс');
  const selector = String(shBal.getRange('I1').getValue() || 'ALL').toUpperCase();
  
  let result = `Family: ${familyId}\n`;
  result += `Selector: ${selector}\n`;
  result += `ACCRUED_FAMILY result: ${ACCRUED_FAMILY(familyId, selector)}\n`;
  result += `Breakdown:\n`;
  
  const breakdown = ACCRUED_BREAKDOWN(familyId, selector);
  for (let i = 1; i < breakdown.length; i++) {
    result += `  ${breakdown[i][0]}: ${breakdown[i][1]}\n`;
  }
  
  return result;
}

// Test function for debugging
function TEST_DYN_CAP() {
  const result1 = DYN_CAP_(500, [2000, 1333]);
  Logger.log(`Test 1: DYN_CAP_(500, [2000, 1333]) = ${result1}`);
  
  const result2 = DYN_CAP_(500, [1333, 2000]);
  Logger.log(`Test 2: DYN_CAP_(500, [1333, 2000]) = ${result2}`);
  
  return `Result1: ${result1}, Result2: ${result2}`;
}

function round6_(x){ return Math.round((x + Number.EPSILON) * 1e6) / 1e6; }
function round2_(x){ return Math.round((x + Number.EPSILON) * 100) / 100; }
function toastErr_(msg){ SpreadsheetApp.getActive().toast(msg, 'Funds (error)', 5); }

/** Generate detailed payment/accrual breakdown for all families and collections (batched, optimized) */
function GENERATE_DETAIL_BREAKDOWN(statusFilter, tick) {
  const onlyOpen = String(statusFilter||'OPEN').toUpperCase() !== 'ALL';
  const ss = SpreadsheetApp.getActive();
  const shF = ss.getSheetByName('Семьи');
  const shC = ss.getSheetByName('Сборы');
  const shU = ss.getSheetByName('Участие');
  const shP = ss.getSheetByName('Платежи');

  // Read headers once
  const mapF = getHeaderMap_(shF);
  const mapC = getHeaderMap_(shC);
  const mapU = getHeaderMap_(shU);
  const mapP = getHeaderMap_(shP);

  // Families: id -> {name, active}
  const families = new Map();
  const famRows = shF.getLastRow();
  if (famRows >= 2) {
    const F = shF.getRange(2, 1, famRows - 1, shF.getLastColumn()).getValues();
    const fi = {}; shF.getRange(1,1,1,shF.getLastColumn()).getValues()[0].forEach((h,idx)=>fi[h]=idx);
    F.forEach(r => {
      const id = String(r[fi['family_id']]||'').trim();
      if (!id) return;
      const name = String(r[fi['Ребёнок ФИО']]||'').trim();
      const active = String(r[fi['Активен']]||'').trim() === 'Да';
      families.set(id, {name, active});
    });
  }

  // Collections: id -> {name, status, accrual, T, fixedX}
  const collections = new Map();
  const cRows = shC.getLastRow();
  if (cRows >= 2) {
    const C = shC.getRange(2, 1, cRows - 1, shC.getLastColumn()).getValues();
    const ci = {}; shC.getRange(1,1,1,shC.getLastColumn()).getValues()[0].forEach((h,idx)=>ci[h]=idx);
    C.forEach(row => {
      const colId = String(row[ci['collection_id']]||'').trim();
      if (!colId) return;
      const status  = String(row[ci['Статус']]||'').trim();
      if (onlyOpen && status !== 'Открыт') return;
      const accrual = String(row[ci['Начисление']]||'').trim();
      const name    = String(row[ci['Название сбора']]||'').trim();
      const paramT  = Number(row[ci['Параметр суммы']]||0);
      const fixedX  = Number(row[ci['Фиксированный x']]||0);
      collections.set(colId, {name, status, accrual, T: paramT, fixedX});
    });
  }

  if (collections.size === 0 || families.size === 0) return [['','','','','','','','','','']];

  // Participation: per collection
  const partByCol = new Map(); // colId -> {hasInclude, include:Set, exclude:Set}
  const uRows = shU.getLastRow();
  if (uRows >= 2) {
    const U = shU.getRange(2, 1, uRows - 1, shU.getLastColumn()).getValues();
    const ui = {}; shU.getRange(1,1,1,shU.getLastColumn()).getValues()[0].forEach((h,idx)=>ui[h]=idx);
    U.forEach(r => {
      const col = getIdFromLabelish_(String(r[ui['collection_id (label)']]||''));
      const fam = getIdFromLabelish_(String(r[ui['family_id (label)']]||''));
      const st  = String(r[ui['Статус']]||'').trim();
      if (!collections.has(col) || !fam) return;
      if (!partByCol.has(col)) partByCol.set(col, {hasInclude:false, include:new Set(), exclude:new Set()});
      const obj = partByCol.get(col);
      if (st === 'Участвует') { obj.hasInclude = true; obj.include.add(fam); }
      else if (st === 'Не участвует') { obj.exclude.add(fam); }
    });
  }

  // Payments: per collection -> Map(famId -> sum)
  const payByCol = new Map();
  const pRows = shP.getLastRow();
  if (pRows >= 2) {
    const P = shP.getRange(2, 1, pRows - 1, shP.getLastColumn()).getValues();
    const pi = {}; shP.getRange(1,1,1,shP.getLastColumn()).getValues()[0].forEach((h,idx)=>pi[h]=idx);
    P.forEach(r => {
      const col = getIdFromLabelish_(String(r[pi['collection_id (label)']]||''));
      const fam = getIdFromLabelish_(String(r[pi['family_id (label)']]||''));
      const sum = Number(r[pi['Сумма']]||0);
      if (!collections.has(col) || !fam || !(sum > 0)) return;
      if (!payByCol.has(col)) payByCol.set(col, new Map());
      const m = payByCol.get(col);
      m.set(fam, (m.get(fam)||0) + sum);
    });
  }

  // Build participants per collection
  const participantsByCol = new Map(); // colId -> Set(famId)
  collections.forEach((col, colId) => {
    const p = partByCol.get(colId);
    const set = new Set();
    if (p && p.hasInclude) p.include.forEach(f => set.add(f));
    else {
      // all active families by default
      families.forEach((info, fid) => { if (info.active) set.add(fid); });
    }
    if (p) p.exclude.forEach(f => set.delete(f));
    // Fallback: if empty, use payers for this collection
    if (set.size === 0 && payByCol.has(colId)) payByCol.get(colId).forEach((_, fid) => set.add(fid));
    participantsByCol.set(colId, set);
  });

  // Compute rows
  const out = [];
  collections.forEach((col, colId) => {
    const name = col.name;
    const accrual = col.accrual;
    const T = col.T;
    const fixedX = col.fixedX || 0;
    const participants = participantsByCol.get(colId) || new Set();
    const famPays = payByCol.get(colId) || new Map();

    // Pre-compute x for dynamic_by_payers
    let x = 0;
    if (accrual === 'dynamic_by_payers') {
      if (fixedX > 0) x = fixedX;
      else {
        const arr = [];
        famPays.forEach((sum, fid) => { if (participants.has(fid) && sum > 0) arr.push(sum); });
        x = DYN_CAP_(T, arr);
      }
    }

  // Pre-compute K for shared_total_by_payers
    let kPayers = 0;
    if (accrual === 'shared_total_by_payers') {
      famPays.forEach((sum, fid) => { if (participants.has(fid) && sum > 0) kPayers++; });
    }

    // Iterate only over union(participants, payers)
    const famSet = new Set();
    participants.forEach(fid => famSet.add(fid));
    famPays.forEach((_, fid) => famSet.add(fid));

    famSet.forEach(fid => {
      const fam = families.get(fid);
      const paid = famPays.get(fid) || 0;
      let accrued = 0;
      if (accrual === 'static_per_child') {
        accrued = participants.has(fid) ? T : 0;
      } else if (accrual === 'shared_total_all') {
        const n = participants.size;
        accrued = (n > 0 && participants.has(fid)) ? (T / n) : 0;
      } else if (accrual === 'shared_total_by_payers') {
        accrued = (kPayers > 0 && participants.has(fid) && paid > 0) ? (T / kPayers) : 0;
      } else if (accrual === 'dynamic_by_payers') {
        if (participants.has(fid) && x > 0) {
          const Pi = paid;
          accrued = Math.min(Pi, x);
        } else {
          accrued = 0;
        }
      } else if (accrual === 'proportional_by_payers') {
        if (participants.has(fid)) {
          let sumP = 0;
          famPays.forEach((sum, f2) => { if (participants.has(f2) && sum > 0) sumP += sum; });
          if (sumP > 0) {
            const target = Math.min(T, sumP);
            const ratio = target / sumP;
            accrued = paid > 0 ? (paid * ratio) : 0;
          } else {
            accrued = 0;
          }
        }
      } else if (accrual === 'unit_price_by_payers') {
      const x = fixedX > 0 ? fixedX : 0;
      accrued = (participants.has(fid) && x > 0) ? (Math.floor(paid / x) * x) : 0;
      }
      if (paid > 0 || accrued > 0) {
        out.push([
          fid,
          fam ? fam.name : '',
          colId,
          name,
          round2_(paid),
          round2_(accrued),
          round2_(paid - accrued),
          accrual
        ]);
      }
    });
  });

  return out.length ? out : [['','','','','','','','']];
}

/** Generate per-collection summary: [collection_id, name, mode, T_total, collected, K_payers, K_needed_more, remaining] */
function GENERATE_COLLECTION_SUMMARY(statusFilter, tick) {
  const statusNorm = String(statusFilter||'OPEN').toUpperCase();
  const onlyOpen = statusNorm !== 'ALL';
  const ss = SpreadsheetApp.getActive();
  const shF = ss.getSheetByName('Семьи');
  const shC = ss.getSheetByName('Сборы');
  const shU = ss.getSheetByName('Участие');
  const shP = ss.getSheetByName('Платежи');

  const mapF = getHeaderMap_(shF);
  const mapC = getHeaderMap_(shC);
  const mapU = getHeaderMap_(shU);
  const mapP = getHeaderMap_(shP);

  // Families: id -> active
  const families = new Map();
  if (shF.getLastRow() >= 2) {
    const F = shF.getRange(2, 1, shF.getLastRow()-1, shF.getLastColumn()).getValues();
    const fi = {}; shF.getRange(1,1,1,shF.getLastColumn()).getValues()[0].forEach((h,idx)=>fi[h]=idx);
    F.forEach(r => {
      const id = String(r[fi['family_id']]||'').trim();
      if (!id) return;
      const active = String(r[fi['Активен']]||'').trim() === 'Да';
      const name = String(r[fi['Ребёнок ФИО']]||'').trim();
      families.set(id, {active, name});
    });
  }

  // Collections
  const collections = [];
  if (shC.getLastRow() >= 2) {
    const C = shC.getRange(2, 1, shC.getLastRow()-1, shC.getLastColumn()).getValues();
    const ci = {}; shC.getRange(1,1,1,shC.getLastColumn()).getValues()[0].forEach((h,idx)=>ci[h]=idx);
    C.forEach(row => {
      const id = String(row[ci['collection_id']]||'').trim(); if (!id) return;
      const status = String(row[ci['Статус']]||'').trim();
      const name = String(row[ci['Название сбора']]||'').trim();
      const mode = String(row[ci['Начисление']]||'').trim();
      const T = Number(row[ci['Параметр суммы']]||0);
      const fixedX = Number(row[ci['Фиксированный x']]||0);
      collections.push({id, name, mode, T, fixedX, status});
    });
  }
  // Filter by status for OPEN mode
  let collectionsToProcess = collections;
  if (onlyOpen) collectionsToProcess = collections.filter(c => c.status === 'Открыт');
  if (!collectionsToProcess.length) return [['','','','','','','','','','']];

  // Participation
  const partByCol = new Map();
  if (shU.getLastRow() >= 2) {
    const U = shU.getRange(2, 1, shU.getLastRow()-1, shU.getLastColumn()).getValues();
    const ui = {}; shU.getRange(1,1,1,shU.getLastColumn()).getValues()[0].forEach((h,idx)=>ui[h]=idx);
    U.forEach(r => {
      const col = getIdFromLabelish_(String(r[ui['collection_id (label)']]||''));
      const fam = getIdFromLabelish_(String(r[ui['family_id (label)']]||''));
      const st = String(r[ui['Статус']]||'').trim();
      if (!col || !fam) return;
      if (!partByCol.has(col)) partByCol.set(col, {hasInclude:false, include:new Set(), exclude:new Set()});
      const obj = partByCol.get(col);
      if (st === 'Участвует') { obj.hasInclude = true; obj.include.add(fam); }
      else if (st === 'Не участвует') { obj.exclude.add(fam); }
    });
  }

  // Payments grouped
  const payByCol = new Map();
  if (shP.getLastRow() >= 2) {
    const P = shP.getRange(2, 1, shP.getLastRow()-1, shP.getLastColumn()).getValues();
    const pi = {}; shP.getRange(1,1,1,shP.getLastColumn()).getValues()[0].forEach((h,idx)=>pi[h]=idx);
    P.forEach(r => {
      const col = getIdFromLabelish_(String(r[pi['collection_id (label)']]||''));
      const fam = getIdFromLabelish_(String(r[pi['family_id (label)']]||''));
      const sum = Number(r[pi['Сумма']]||0);
      if (!col || !fam || !(sum>0)) return;
      if (!payByCol.has(col)) payByCol.set(col, new Map());
      const m = payByCol.get(col);
      m.set(fam, (m.get(fam)||0) + sum);
    });
  }

  const buildRow = (col) => {
    const {id, name, mode, T, fixedX} = col;
    // participants
    const p = partByCol.get(id);
    const participants = new Set();
    if (p && p.hasInclude) p.include.forEach(fid=>participants.add(fid));
    else families.forEach((v,fid)=>{ if (v.active) participants.add(fid); });
    if (p) p.exclude.forEach(fid=>participants.delete(fid));
    // Fallback: if none resolved, use payers as participants
    if (participants.size === 0 && payByCol.has(id)) payByCol.get(id).forEach((_,fid)=>participants.add(fid));

    const famPays = payByCol.get(id) || new Map();
  // Collected from participants only
    let collected = 0; let K = 0;
    famPays.forEach((sum, fid)=>{ if (participants.has(fid)) { collected += sum; if (sum>0) K++; } });
    // Units paid for unit_price_by_payers
    let unitsPaid = '';
    if (mode === 'unit_price_by_payers') {
      const x = (fixedX > 0) ? fixedX : 0;
      if (x > 0) unitsPaid = Math.floor(collected / x);
    }

    // Target total
    let Ttotal = 0;
    if (mode === 'static_per_child') {
      Ttotal = (participants.size || 0) * (T || 0);
    } else {
      Ttotal = T || 0;
    }
    Ttotal = Number(Ttotal) || 0;
    const remaining = Math.max(0, round2_(Ttotal - collected));

    // Estimate additional payers needed
    let needMore = '';
    if (remaining <= 0) {
      needMore = 0;
    } else if (mode === 'static_per_child') {
      const rate = T || 0;
      needMore = rate > 0 ? Math.ceil(remaining / rate) : '';
    } else if (mode === 'shared_total_all') {
      const n = participants.size || 0;
      const share = (n > 0) ? (T / n) : 0;
      needMore = share > 0 ? Math.ceil(remaining / share) : '';
    } else if (mode === 'shared_total_by_payers') {
      const share = fixedX > 0 ? fixedX : (K > 0 ? (T / K) : 0);
      needMore = share > 0 ? Math.ceil(remaining / share) : '';
    } else if (mode === 'dynamic_by_payers') {
      // if x fixed (closed), estimate by x; else leave blank
      needMore = fixedX > 0 ? Math.ceil(remaining / fixedX) : '';
    } else if (mode === 'proportional_by_payers') {
      // Not applicable: proportional redistribution among payers; no discrete payers needed metric
      needMore = '';
    } else if (mode === 'unit_price_by_payers') {
      const x = fixedX > 0 ? fixedX : 0;
      needMore = x > 0 ? Math.ceil(remaining / x) : '';
    } else {
      needMore = '';
    }

  // For unit_price_by_payers we output both K (unique payers) and UnitsPaid separately

    return [
      id,
      name,
      mode,
  round2_(Ttotal),
  round2_(collected),
  // Participants:
  (mode === 'unit_price_by_payers' ? (fixedX>0 ? Math.ceil((T||0)/fixedX) : participants.size) : participants.size),
  // Unique payers (K)
  K,
  // Units paid (for unit_price_by_payers) or blank otherwise
  (mode === 'unit_price_by_payers' ? (unitsPaid === '' ? '' : unitsPaid) : ''),
  needMore,
  round2_(remaining)
    ];
  };

  const out = [];
  if (statusNorm === 'ALL') {
    // Open first
    const openRows = collections.filter(c => c.status === 'Открыт').map(buildRow);
    const closedRows = collections.filter(c => c.status !== 'Открыт').map(buildRow);
    // Insert section headers as single labeled rows for clarity
  if (openRows.length) out.push(['','ОТКРЫТЫЕ СБОРЫ','','','','','','','','']);
    Array.prototype.push.apply(out, openRows);
  // Add visual separation: 5 empty rows between open and closed sections
  for (let i = 0; i < 5; i++) out.push(['','','','','','','','','','']);
  if (closedRows.length) out.push(['','ЗАКРЫТЫЕ СБОРЫ','','','','','','','','']);
    Array.prototype.push.apply(out, closedRows);
  } else {
    Array.prototype.push.apply(out, collectionsToProcess.map(buildRow));
  }

  return out.length ? out : [['','','','','','','','','','']];
}

/** Calculate accrual for a specific family/collection pair */
function getSingleAccrual_(familyId, collectionId, statusFilter) {
  const onlyOpen = String(statusFilter||'OPEN').toUpperCase() !== 'ALL';
  const ss = SpreadsheetApp.getActive();
  const shF = ss.getSheetByName('Семьи');
  const shC = ss.getSheetByName('Сборы');
  const shU = ss.getSheetByName('Участие');
  const shP = ss.getSheetByName('Платежи');

  const mapF = getHeaderMap_(shF);
  const mapC = getHeaderMap_(shC);
  const mapU = getHeaderMap_(shU);
  const mapP = getHeaderMap_(shP);

  // Get collection data
  let collectionData = null;
  const cRows = shC.getLastRow();
  if (cRows >= 2) {
    const C = shC.getRange(2, 1, cRows - 1, shC.getLastColumn()).getValues();
    const ch = shC.getRange(1,1,1,shC.getLastColumn()).getValues()[0];
    const ci={}; ch.forEach((h,idx)=>ci[h]=idx);
    C.forEach(row=>{
      const colId = String(row[ci['collection_id']]||'').trim();
      if (colId === collectionId) {
        collectionData = {
          status: String(row[ci['Статус']]||'').trim(),
          accrual: String(row[ci['Начисление']]||'').trim(),
          paramT: Number(row[ci['Параметр суммы']]||0),
          fixedX: Number(row[ci['Фиксированный x']]||0)
        };
      }
    });
  }
  
  if (!collectionData || (onlyOpen && collectionData.status !== 'Открыт')) return 0;

  // Get active families and participation for this collection
  const activeFam = new Set();
  const famRows = shF.getLastRow();
  if (famRows >= 2) {
    const vals = shF.getRange(2, 1, famRows - 1, shF.getLastColumn()).getValues();
    const headers = shF.getRange(1,1,1,shF.getLastColumn()).getValues()[0];
    const i = {}; headers.forEach((h,idx)=>i[h]=idx);
    vals.forEach(r=>{
      const id = String(r[i['family_id']]||'').trim();
      const act = String(r[i['Активен']]||'').trim()==='Да';
      if (id && act) activeFam.add(id);
    });
  }

  // Participation for this collection
  const partInclude = new Set();
  const partExclude = new Set();
  let hasInclude = false;
  const uRows = shU.getLastRow();
  if (uRows >= 2) {
    const U = shU.getRange(2, 1, uRows - 1, shU.getLastColumn()).getValues();
    const uh = shU.getRange(1,1,1,shU.getLastColumn()).getValues()[0];
    const ui={}; uh.forEach((h,idx)=>ui[h]=idx);
    U.forEach(r=>{
      const col = getIdFromLabelish_(String(r[ui['collection_id (label)']]||''));
      const fam = getIdFromLabelish_(String(r[ui['family_id (label)']]||''));
      const st = String(r[ui['Статус']]||'').trim();
      if (col === collectionId && fam) {
        if (st === 'Участвует') { hasInclude = true; partInclude.add(fam); }
        else if (st === 'Не участвует') { partExclude.add(fam); }
      }
    });
  }

  // Resolve participants
  const participants = new Set();
  if (hasInclude) partInclude.forEach(f=>participants.add(f));
  else activeFam.forEach(f=>participants.add(f));
  partExclude.forEach(f=>participants.delete(f));

  // Payments for this collection
  const famPays = new Map();
  const pRows = shP.getLastRow();
  if (pRows >= 2) {
    const P = shP.getRange(2, 1, pRows - 1, shP.getLastColumn()).getValues();
    const ph = shP.getRange(1,1,1,shP.getLastColumn()).getValues()[0];
    const pi={}; ph.forEach((h,idx)=>pi[h]=idx);
    P.forEach(r=>{
      const col = getIdFromLabelish_(String(r[pi['collection_id (label)']]||''));
      const fam = getIdFromLabelish_(String(r[pi['family_id (label)']]||''));
      const sum = Number(r[pi['Сумма']]||0);
      if (col === collectionId && fam && sum > 0) {
        famPays.set(fam, (famPays.get(fam)||0) + sum);
      }
    });
  }

  // Fallback: use payers as participants if none resolved
  if (participants.size === 0) famPays.forEach((_, fid)=>participants.add(fid));

  const n = participants.size;
  const Pi = famPays.get(familyId) || 0;

  let accrued = 0;
  if (collectionData.accrual === 'static_per_child') {
    accrued = participants.has(familyId) ? collectionData.paramT : 0;
  } else if (collectionData.accrual === 'shared_total_all') {
    if (n > 0 && participants.has(familyId)) accrued = collectionData.paramT / n;
  } else if (collectionData.accrual === 'shared_total_by_payers') {
    // Share T equally among payers (within participants)
    let k = 0;
    famPays.forEach((sum, fid)=>{ if (participants.has(fid) && sum>0) k++; });
    if (k > 0 && participants.has(familyId) && (famPays.get(familyId)||0) > 0) accrued = collectionData.paramT / k; else accrued = 0;
  } else if (collectionData.accrual === 'dynamic_by_payers') {
    if (participants.has(familyId) && n > 0) {
      const payments = [];
      famPays.forEach((sum, fid)=>{ if (participants.has(fid) && sum>0) payments.push(sum); });
      const x = collectionData.fixedX > 0 ? collectionData.fixedX : DYN_CAP_(collectionData.paramT, payments);
      accrued = Math.min(Pi, x);
    }
  } else if (collectionData.accrual === 'proportional_by_payers') {
    if (participants.has(familyId)) {
      let sumP = 0;
      famPays.forEach((sum, fid)=>{ if (participants.has(fid) && sum>0) sumP += sum; });
      if (sumP > 0) {
        const target = Math.min(collectionData.paramT, sumP);
        const ratio = target / sumP;
        accrued = Pi > 0 ? (Pi * ratio) : 0;
      } else {
        accrued = 0;
      }
    }
  } else if (collectionData.accrual === 'unit_price_by_payers') {
  const x = collectionData.fixedX > 0 ? collectionData.fixedX : 0;
  if (participants.has(familyId) && x > 0) accrued = Math.floor(Pi / x) * x; else accrued = 0;
  }
  
  return round2_(accrued);
}


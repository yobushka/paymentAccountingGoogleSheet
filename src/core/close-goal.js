/**
 * @fileoverview Закрытие цели/сбора
 */

/**
 * Диалог закрытия цели
 * Точка входа из меню
 */
function closeGoalPrompt() {
  const ui = SpreadsheetApp.getUi();
  const version = detectVersion();
  
  const idLabel = version === 'v1' ? 'collection_id' : 'goal_id';
  const example = version === 'v1' ? 'C001' : 'G001';
  
  const resp = ui.prompt(
    'Close Goal',
    `Введите ${idLabel} (например, ${example}):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  
  const goalId = (resp.getResponseText() || '').trim();
  if (!goalId) return;
  
  if (version === 'v1') {
    closeCollection_(goalId);
  } else {
    closeGoal_(goalId);
  }
}

/**
 * Закрывает цель (v2.0)
 * @param {string} goalId
 */
function closeGoal_(goalId) {
  const ss = SpreadsheetApp.getActive();
  const shG = ss.getSheetByName(SHEET_NAMES.GOALS);
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  const shU = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  
  const mapG = getHeaderMap_(shG);
  
  // Находим строку цели
  const idCol = mapG['goal_id'];
  if (!idCol) return toastErr_('Не найден столбец goal_id.');
  
  const rowsG = shG.getLastRow();
  if (rowsG < 2) return toastErr_('Нет целей.');
  
  const ids = shG.getRange(2, idCol, rowsG - 1, 1).getValues().map(r => String(r[0] || '').trim());
  const idx = ids.findIndex(v => v === goalId);
  if (idx === -1) return toastErr_('Цель не найдена: ' + goalId);
  
  const rowNum = 2 + idx;
  
  // Читаем данные цели
  const accrual = normalizeAccrualMode_(String(shG.getRange(rowNum, mapG['Начисление']).getValue() || '').trim());
  const paramT = Number(shG.getRange(rowNum, mapG['Параметр суммы']).getValue() || 0);
  
  // Для dynamic_by_payers вычисляем и фиксируем x
  if (accrual === ACCRUAL_MODES.DYNAMIC_BY_PAYERS) {
    // Читаем данные
    const families = readFamilies_(shF);
    const participation = readParticipation_(shU, 'v2');
    const payments = readPayments_(shP, 'v2');
    
    // Читаем даты цели для resolveParticipants_
    const goal = {
      startDate: parseDate_(shG.getRange(rowNum, mapG['Дата начала']).getValue()),
      deadline: parseDate_(shG.getRange(rowNum, mapG['Дедлайн']).getValue())
    };
    
    const participants = resolveParticipants_(goalId, families, participation, goal);
    const goalPayments = payments.get(goalId) || new Map();
    
    // Собираем платежи участников
    const paymentArray = [];
    goalPayments.forEach((sum, fid) => {
      if (participants.has(fid) && sum > 0) paymentArray.push(sum);
    });
    
    const x = DYN_CAP_(paramT, paymentArray);
    
    if (mapG['Фиксированный x']) {
      shG.getRange(rowNum, mapG['Фиксированный x']).setValue(x);
    }
    if (mapG['Статус']) {
      shG.getRange(rowNum, mapG['Статус']).setValue(GOAL_STATUS.CLOSED);
    }
    
    SpreadsheetApp.getActive().toast(`Цель ${goalId} закрыта. x=${round2_(x)}`, 'Funds');
  } else {
    // Для других режимов просто закрываем
    if (mapG['Статус']) {
      shG.getRange(rowNum, mapG['Статус']).setValue(GOAL_STATUS.CLOSED);
    }
    SpreadsheetApp.getActive().toast(`Цель ${goalId} закрыта.`, 'Funds');
  }
}

/**
 * Закрывает сбор (v1.x)
 * @param {string} collectionId
 */
function closeCollection_(collectionId) {
  const ss = SpreadsheetApp.getActive();
  const shC = ss.getSheetByName(SHEET_NAMES.COLLECTIONS);
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  const shU = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  
  const mapC = getHeaderMap_(shC);
  
  const idCol = mapC['collection_id'];
  if (!idCol) return toastErr_('Не найден столбец collection_id.');
  
  const rowsC = shC.getLastRow();
  if (rowsC < 2) return toastErr_('Нет сборов.');
  
  const ids = shC.getRange(2, idCol, rowsC - 1, 1).getValues().map(r => String(r[0] || '').trim());
  const idx = ids.findIndex(v => v === collectionId);
  if (idx === -1) return toastErr_('Сбор не найден: ' + collectionId);
  
  const rowNum = 2 + idx;
  
  const accrual = String(shC.getRange(rowNum, mapC['Начисление']).getValue() || '').trim();
  const paramT = Number(shC.getRange(rowNum, mapC['Параметр суммы']).getValue() || 0);
  
  if (accrual === 'dynamic_by_payers') {
    const families = readFamilies_(shF);
    const participation = readParticipation_(shU, 'v1');
    const payments = readPayments_(shP, 'v1');
    
    // v1 не имеет дат, используем null
    const participants = resolveParticipants_(collectionId, families, participation, null);
    const goalPayments = payments.get(collectionId) || new Map();
    
    const paymentArray = [];
    goalPayments.forEach((sum, fid) => {
      if (participants.has(fid) && sum > 0) paymentArray.push(sum);
    });
    
    const x = DYN_CAP_(paramT, paymentArray);
    
    if (mapC['Фиксированный x']) {
      shC.getRange(rowNum, mapC['Фиксированный x']).setValue(x);
    }
    if (mapC['Статус']) {
      shC.getRange(rowNum, mapC['Статус']).setValue(COLLECTION_STATUS_V1.CLOSED);
    }
    
    SpreadsheetApp.getActive().toast(`Сбор ${collectionId} закрыт. x=${round2_(x)}`, 'Funds');
  } else {
    if (mapC['Статус']) {
      shC.getRange(rowNum, mapC['Статус']).setValue(COLLECTION_STATUS_V1.CLOSED);
    }
    SpreadsheetApp.getActive().toast(`Сбор ${collectionId} закрыт.`, 'Funds');
  }
}

/**
 * Алиас для обратной совместимости
 */
function closeCollectionPrompt() {
  closeGoalPrompt();
}

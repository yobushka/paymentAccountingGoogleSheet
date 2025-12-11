/**
 * @fileoverview Лист «Сводка» — настройка и генерация
 */

/**
 * Настраивает лист «Сводка»
 */
function setupSummarySheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAMES.SUMMARY);
  if (!sh) return;
  
  // Очищаем старые данные
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
  }
  
  // Очищаем возможные остатки в области spill
  try { sh.getRange('J2:J3').clearContent(); } catch (_) {}
  
  // Селектор и тикер — за пределами области данных
  // Очищаем возможную валидацию перед записью меток
  sh.getRange('L1').clearDataValidations().setValue('Фильтр');
  sh.getRange('K1').clearDataValidations().setValue('ALL');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['OPEN', 'ALL'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('K1').setDataValidation(rule).setHorizontalAlignment('center');
  sh.getRange('K1').setNote('OPEN (только открытые) или ALL (все цели, сначала открытые, ниже — закрытые)');
  
  sh.getRange('L2').clearDataValidations().setValue('Tick');
  sh.getRange('K2').clearDataValidations().setValue(new Date().toISOString());
  
  // Array formula
  sh.getRange('A2').setFormula(`=GENERATE_COLLECTION_SUMMARY(IF(LEN($K$1)=0,"ALL",$K$1), $K$2)`);
  
  sh.getRange('L3').clearDataValidations().setValue('Сводка по целям. ALL: сверху открытые, внизу закрытые.');
  
  SpreadsheetApp.flush();
  try {
    styleSheetHeader_(sh);
    styleSummarySheet_(sh);
  } catch (_) {}
}

/**
 * Обновляет лист «Сводка»
 */
function refreshSummarySheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAMES.SUMMARY);
  if (!sh) return;
  
  const current = sh.getRange('A2').getFormula();
  if (current.includes('GENERATE_COLLECTION_SUMMARY')) {
    sh.getRange('K2').setValue(new Date().toISOString());
    try { sh.getRange('J2:J3').clearContent(); } catch (_) {}
    sh.getRange('A2').setFormula(current);
    
    try {
      SpreadsheetApp.flush();
      styleSheetHeader_(sh);
      styleSummarySheet_(sh);
    } catch (e) {}
  }
}

/**
 * Генерирует сводку по целям
 * @param {string} statusFilter — OPEN или ALL
 * @param {string} tick — тикер для пересчёта
 * @returns {Array<Array>}
 * @customfunction
 */
function GENERATE_COLLECTION_SUMMARY(statusFilter, tick) {
  const statusNorm = String(statusFilter || 'OPEN').toUpperCase();
  const onlyOpen = statusNorm !== 'ALL';
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  const shG = version === 'v1' 
    ? ss.getSheetByName(SHEET_NAMES.COLLECTIONS) 
    : ss.getSheetByName(SHEET_NAMES.GOALS);
  const shU = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  
  if (!shF || !shG || !shU || !shP) return [['', '', '', '', '', '', '', '', '', '', '']];
  
  // Читаем данные
  const families = readFamilies_(shF);
  const participation = readParticipation_(shU, version);
  const payments = readPayments_(shP, version);
  
  // Читаем все цели (для сортировки открытые/закрытые)
  const allGoals = readAllGoals_(shG, version);
  
  if (allGoals.length === 0) return [['', '', '', '', '', '', '', '', '', '', '']];
  
  // Фильтруем для режима OPEN
  let goalsToProcess = allGoals;
  if (onlyOpen) {
    const openStatus = version === 'v1' ? COLLECTION_STATUS_V1.OPEN : GOAL_STATUS.OPEN;
    goalsToProcess = allGoals.filter(g => g.status === openStatus);
  }
  
  if (goalsToProcess.length === 0) return [['', '', '', '', '', '', '', '', '', '', '']];
  
  const buildRow = (goal) => {
    const participants = resolveParticipants_(goal.id, families, participation, goal);
    const goalPayments = payments.get(goal.id) || new Map();
    
    // Fallback: если нет участников, используем плательщиков
    if (participants.size === 0) {
      goalPayments.forEach((_, fid) => participants.add(fid));
    }
    
    // Собрано (от участников)
    let collected = 0;
    let K = 0; // плательщиков
    goalPayments.forEach((sum, fid) => {
      if (participants.has(fid)) {
        collected += sum;
        if (sum > 0) K++;
      }
    });
    
    // Единиц оплачено (для unit_price)
    let unitsPaid = '';
    const accrual = normalizeAccrualMode_(goal.accrual);
    if (accrual === ACCRUAL_MODES.UNIT_PRICE) {
      const x = goal.fixedX > 0 ? goal.fixedX : 0;
      if (x > 0) unitsPaid = Math.floor(collected / x);
    }
    
    // Целевая сумма
    let Ttotal = 0;
    if (accrual === ACCRUAL_MODES.STATIC_PER_FAMILY) {
      Ttotal = participants.size * goal.T;
    } else {
      Ttotal = goal.T;
    }
    
    const remaining = Math.max(0, round2_(Ttotal - collected));
    
    // Переплата: если собрано больше целевой суммы
    const overpaid = Math.max(0, round2_(collected - Ttotal));
    
    // Оценка дополнительных плательщиков
    let needMore = '';
    if (remaining <= 0) {
      needMore = 0;
    } else {
      switch (accrual) {
        case ACCRUAL_MODES.STATIC_PER_FAMILY:
          needMore = goal.T > 0 ? Math.ceil(remaining / goal.T) : '';
          break;
        case ACCRUAL_MODES.SHARED_TOTAL_ALL:
          const share = participants.size > 0 ? (goal.T / participants.size) : 0;
          needMore = share > 0 ? Math.ceil(remaining / share) : '';
          break;
        case ACCRUAL_MODES.SHARED_TOTAL_BY_PAYERS:
          const shareK = goal.fixedX > 0 ? goal.fixedX : (K > 0 ? (goal.T / K) : 0);
          needMore = shareK > 0 ? Math.ceil(remaining / shareK) : '';
          break;
        case ACCRUAL_MODES.DYNAMIC_BY_PAYERS:
          needMore = goal.fixedX > 0 ? Math.ceil(remaining / goal.fixedX) : '';
          break;
        case ACCRUAL_MODES.UNIT_PRICE:
          needMore = goal.fixedX > 0 ? Math.ceil(remaining / goal.fixedX) : '';
          break;
        default:
          needMore = '';
      }
    }
    
    return [
      goal.id,
      goal.name,
      goal.accrual,
      round2_(Ttotal),
      round2_(collected),
      accrual === ACCRUAL_MODES.UNIT_PRICE ? (goal.fixedX > 0 ? Math.ceil(goal.T / goal.fixedX) : participants.size) : participants.size,
      K,
      accrual === ACCRUAL_MODES.UNIT_PRICE ? (unitsPaid === '' ? '' : unitsPaid) : '',
      needMore,
      round2_(remaining),
      overpaid
    ];
  };
  
  const out = [];
  
  if (statusNorm === 'ALL') {
    const openStatus = version === 'v1' ? COLLECTION_STATUS_V1.OPEN : GOAL_STATUS.OPEN;
    const openGoals = allGoals.filter(g => g.status === openStatus);
    const closedGoals = allGoals.filter(g => g.status !== openStatus);
    
    if (openGoals.length) {
      out.push(['', 'ОТКРЫТЫЕ ЦЕЛИ', '', '', '', '', '', '', '', '', '']);
      openGoals.forEach(g => out.push(buildRow(g)));
    }
    
    // Разделитель
    for (let i = 0; i < 3; i++) out.push(['', '', '', '', '', '', '', '', '', '', '']);
    
    if (closedGoals.length) {
      out.push(['', 'ЗАКРЫТЫЕ ЦЕЛИ', '', '', '', '', '', '', '', '', '']);
      closedGoals.forEach(g => out.push(buildRow(g)));
    }
  } else {
    goalsToProcess.forEach(g => out.push(buildRow(g)));
  }
  
  return out.length ? out : [['', '', '', '', '', '', '', '', '', '', '']];
}

/**
 * Читает все цели/сборы без фильтрации (кроме отменённых)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {string} version
 * @returns {Array<{id: string, name: string, status: string, accrual: string, T: number, fixedX: number, startDate: Date|null, deadline: Date|null}>}
 */
function readAllGoals_(sh, version) {
  const goals = [];
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return goals;
  
  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);
  
  const idField = version === 'v1' ? 'collection_id' : 'goal_id';
  const nameField = version === 'v1' ? 'Название сбора' : 'Название цели';
  
  // Статус «Отменена» — пропускаем
  const cancelledStatus = version === 'v1' ? null : GOAL_STATUS.CANCELLED;
  
  data.forEach(r => {
    const id = String(r[idx[idField]] || '').trim();
    if (!id) return;
    
    const status = String(r[idx['Статус']] || '').trim();
    
    // Отменённые цели — пропускаем
    if (cancelledStatus && status === cancelledStatus) return;
    
    goals.push({
      id: id,
      name: String(r[idx[nameField]] || '').trim(),
      status: status,
      accrual: String(r[idx['Начисление']] || '').trim(),
      T: Number(r[idx['Параметр суммы']] || 0),
      fixedX: Number(r[idx['Фиксированный x']] || 0),
      startDate: parseDate_(r[idx['Дата начала']]),
      deadline: parseDate_(r[idx['Дедлайн']])
    });
  });
  
  return goals;
}

/**
 * @fileoverview Custom functions для использования в ячейках
 */

/**
 * Возвращает общую сумму платежей семьи
 * @param {string} familyLabelOrId — метка или ID семьи
 * @returns {number}
 * @customfunction
 */
function PAYED_TOTAL_FAMILY(familyLabelOrId) {
  const famId = getIdFromLabelish_(familyLabelOrId);
  if (!famId) return 0;
  
  const ss = SpreadsheetApp.getActive();
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  const rows = shP.getLastRow();
  if (rows < 2) return 0;
  
  const map = getHeaderMap_(shP);
  const iFam = map['family_id (label)'];
  const iSum = map['Сумма'];
  if (!iFam || !iSum) return 0;
  
  const vals = shP.getRange(2, 1, rows - 1, shP.getLastColumn()).getValues();
  let total = 0;
  
  vals.forEach(r => {
    const fid = getIdFromLabelish_(String(r[iFam - 1] || ''));
    const sum = Number(r[iSum - 1] || 0);
    if (fid === famId && sum > 0) total += sum;
  });
  
  return round2_(total);
}

/**
 * Возвращает начисленную сумму для семьи
 * @param {string} familyLabelOrId — метка или ID семьи
 * @param {string} [statusFilter='OPEN'] — OPEN, CLOSED или ALL
 * @returns {number}
 * @customfunction
 */
function ACCRUED_FAMILY(familyLabelOrId, statusFilter) {
  const famId = getIdFromLabelish_(familyLabelOrId);
  if (!famId) return 0;
  
  // Нормализуем фильтр
  const filter = String(statusFilter || 'OPEN').toUpperCase();
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  const shG = version === 'v1' 
    ? ss.getSheetByName(SHEET_NAMES.COLLECTIONS) 
    : ss.getSheetByName(SHEET_NAMES.GOALS);
  const shU = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  
  if (!shF || !shG || !shU || !shP) return 0;
  
  // Читаем данные
  const families = readFamilies_(shF);
  const participation = readParticipation_(shU, version);
  const payments = readPayments_(shP, version);
  const goals = readGoals_(shG, version, filter);
  
  let total = 0;
  
  goals.forEach((goal, goalId) => {
    const participants = resolveParticipants_(goalId, families, participation, goal);
    const goalPayments = payments.get(goalId) || new Map();
    
    // Fallback: если нет участников, используем плательщиков
    if (participants.size === 0) {
      goalPayments.forEach((_, fid) => participants.add(fid));
    }
    
    const { x, kPayers } = precalculateForGoal_(goal, participants, goalPayments);
    const accrued = calculateAccrual_(famId, goal, participants, goalPayments, x, kPayers);
    
    total += accrued;
  });
  
  return round2_(total);
}

/**
 * Возвращает детализацию начислений для семьи
 * @param {string} familyLabelOrId — метка или ID семьи
 * @param {string} [statusFilter='OPEN'] — OPEN, CLOSED или ALL
 * @returns {Array<Array>}
 * @customfunction
 */
function ACCRUED_BREAKDOWN(familyLabelOrId, statusFilter) {
  const famId = getIdFromLabelish_(familyLabelOrId);
  if (!famId) return [['goal_id', 'accrued']];
  
  const filter = String(statusFilter || 'OPEN').toUpperCase();
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  const shG = version === 'v1' 
    ? ss.getSheetByName(SHEET_NAMES.COLLECTIONS) 
    : ss.getSheetByName(SHEET_NAMES.GOALS);
  const shU = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  
  if (!shF || !shG || !shU || !shP) return [['goal_id', 'accrued']];
  
  const families = readFamilies_(shF);
  const participation = readParticipation_(shU, version);
  const payments = readPayments_(shP, version);
  const goals = readGoals_(shG, version, filter);
  
  const out = [['goal_id', 'accrued']];
  
  goals.forEach((goal, goalId) => {
    const participants = resolveParticipants_(goalId, families, participation, goal);
    const goalPayments = payments.get(goalId) || new Map();
    
    if (participants.size === 0) {
      goalPayments.forEach((_, fid) => participants.add(fid));
    }
    
    const { x, kPayers } = precalculateForGoal_(goal, participants, goalPayments);
    const accrued = calculateAccrual_(famId, goal, participants, goalPayments, x, kPayers);
    
    if (accrued !== 0) {
      out.push([goalId, round2_(accrued)]);
    }
  });
  
  return out;
}

/**
 * Отладочная функция: детали расчёта для семьи и цели
 * @param {string} goalId
 * @param {string} familyId
 * @returns {string}
 */
function DEBUG_GOAL_ACCRUAL(goalId, familyId) {
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  
  const shG = version === 'v1' 
    ? ss.getSheetByName(SHEET_NAMES.COLLECTIONS) 
    : ss.getSheetByName(SHEET_NAMES.GOALS);
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  const shU = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  
  // Читаем цель
  const goals = readGoals_(shG, version, false);
  const goal = goals.get(goalId);
  
  if (!goal) return 'Goal not found: ' + goalId;
  
  const families = readFamilies_(shF);
  const participation = readParticipation_(shU, version);
  const payments = readPayments_(shP, version);
  
  const participants = resolveParticipants_(goalId, families, participation, goal);
  const goalPayments = payments.get(goalId) || new Map();
  const familyPayment = goalPayments.get(familyId) || 0;
  
  const { x, kPayers } = precalculateForGoal_(goal, participants, goalPayments);
  const accrued = calculateAccrual_(familyId, goal, participants, goalPayments, x, kPayers);
  
  const paymentArray = Array.from(goalPayments.values());
  
  let result = `Goal: ${goalId}\n`;
  result += `Mode: ${goal.accrual}\n`;
  result += `Target T: ${goal.T}\n`;
  result += `Fixed X: ${goal.fixedX}\n`;
  result += `Status: ${goal.status}\n`;
  result += `Participants: ${participants.size}\n`;
  result += `K Payers: ${kPayers}\n`;
  result += `Calculated x: ${x}\n`;
  result += `All payments: [${paymentArray.join(', ')}]\n`;
  result += `Family ${familyId} payment: ${familyPayment}\n`;
  result += `Family ${familyId} accrual: ${accrued}\n`;
  
  return result;
}

/**
 * Отладочная функция: баланс семьи
 * @param {string} familyId
 * @returns {string}
 */
function DEBUG_BALANCE_ACCRUAL(familyId) {
  const ss = SpreadsheetApp.getActive();
  const shBal = ss.getSheetByName(SHEET_NAMES.BALANCE);
  const selector = String(shBal.getRange('J1').getValue() || 'ALL').toUpperCase();
  
  let result = `Family: ${familyId}\n`;
  result += `Selector: ${selector}\n`;
  result += `PAYED_TOTAL_FAMILY: ${PAYED_TOTAL_FAMILY(familyId)}\n`;
  result += `ACCRUED_FAMILY result: ${ACCRUED_FAMILY(familyId, selector)}\n`;
  result += `Breakdown:\n`;
  
  const breakdown = ACCRUED_BREAKDOWN(familyId, selector);
  for (let i = 1; i < breakdown.length; i++) {
    result += `  ${breakdown[i][0]}: ${breakdown[i][1]}\n`;
  }
  
  return result;
}

// ============================================================================
// v2.0 Балансовые кастомные функции
// ============================================================================

/**
 * Возвращает сумму списаний по ЗАКРЫТЫМ целям (Списано всего)
 * @param {string} familyLabelOrId — метка или ID семьи
 * @returns {number}
 * @customfunction
 */
function TOTAL_CHARGED_FAMILY(familyLabelOrId) {
  return ACCRUED_FAMILY(familyLabelOrId, 'CLOSED');
}

/**
 * Возвращает сумму резервов по ОТКРЫТЫМ целям (Зарезервировано)
 * @param {string} familyLabelOrId — метка или ID семьи
 * @returns {number}
 * @customfunction
 */
function RESERVED_FAMILY(familyLabelOrId) {
  return ACCRUED_FAMILY(familyLabelOrId, 'OPEN');
}

/**
 * Возвращает текущий баланс семьи (Внесено − Списано)
 * @param {string} familyLabelOrId — метка или ID семьи
 * @returns {number}
 * @customfunction
 */
function BALANCE_FAMILY(familyLabelOrId) {
  const paid = PAYED_TOTAL_FAMILY(familyLabelOrId);
  const charged = TOTAL_CHARGED_FAMILY(familyLabelOrId);
  return round2_(paid - charged);
}

/**
 * Возвращает свободный остаток семьи (Баланс − Резерв)
 * @param {string} familyLabelOrId — метка или ID семьи
 * @returns {number}
 * @customfunction
 */
function FREE_BALANCE_FAMILY(familyLabelOrId) {
  const balance = BALANCE_FAMILY(familyLabelOrId);
  const reserved = RESERVED_FAMILY(familyLabelOrId);
  return round2_(balance - reserved);
}

/**
 * Возвращает задолженность семьи (MAX(0, −Свободный остаток))
 * @param {string} familyLabelOrId — метка или ID семьи
 * @returns {number}
 * @customfunction
 */
function DEBT_FAMILY(familyLabelOrId) {
  const freeBalance = FREE_BALANCE_FAMILY(familyLabelOrId);
  return round2_(Math.max(0, -freeBalance));
}

/**
 * Возвращает сумму платежей семьи для конкретной цели
 * @param {string} familyLabelOrId — метка или ID семьи
 * @param {string} goalLabelOrId — метка или ID цели
 * @returns {number}
 * @customfunction
 */
function PAID_TO_GOAL(familyLabelOrId, goalLabelOrId) {
  const famId = getIdFromLabelish_(familyLabelOrId);
  const goalId = getIdFromLabelish_(goalLabelOrId);
  if (!famId || !goalId) return 0;
  
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  if (!shP) return 0;
  
  const payments = readPayments_(shP, version);
  const goalPayments = payments.get(goalId);
  if (!goalPayments) return 0;
  
  return round2_(goalPayments.get(famId) || 0);
}

/**
 * Возвращает начисление семьи для конкретной цели
 * @param {string} familyLabelOrId — метка или ID семьи
 * @param {string} goalLabelOrId — метка или ID цели
 * @returns {number}
 * @customfunction
 */
function ACCRUED_FOR_GOAL(familyLabelOrId, goalLabelOrId) {
  const famId = getIdFromLabelish_(familyLabelOrId);
  const goalId = getIdFromLabelish_(goalLabelOrId);
  if (!famId || !goalId) return 0;
  
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  const shG = version === 'v1' 
    ? ss.getSheetByName(SHEET_NAMES.COLLECTIONS) 
    : ss.getSheetByName(SHEET_NAMES.GOALS);
  const shU = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  
  if (!shF || !shG || !shU || !shP) return 0;
  
  const families = readFamilies_(shF);
  const participation = readParticipation_(shU, version);
  const payments = readPayments_(shP, version);
  const goals = readGoals_(shG, version, false); // все цели
  
  const goal = goals.get(goalId);
  if (!goal) return 0;
  
  const participants = resolveParticipants_(goalId, families, participation, goal);
  const goalPayments = payments.get(goalId) || new Map();
  
  // Fallback: если нет участников, используем плательщиков
  if (participants.size === 0) {
    goalPayments.forEach((_, fid) => participants.add(fid));
  }
  
  const { x, kPayers } = precalculateForGoal_(goal, participants, goalPayments);
  return round2_(calculateAccrual_(famId, goal, participants, goalPayments, x, kPayers));
}

/**
 * Возвращает сальдо семьи по конкретной цели (Оплачено − Начислено)
 * @param {string} familyLabelOrId — метка или ID семьи
 * @param {string} goalLabelOrId — метка или ID цели
 * @returns {number}
 * @customfunction
 */
function BALANCE_FOR_GOAL(familyLabelOrId, goalLabelOrId) {
  const paid = PAID_TO_GOAL(familyLabelOrId, goalLabelOrId);
  const accrued = ACCRUED_FOR_GOAL(familyLabelOrId, goalLabelOrId);
  return round2_(paid - accrued);
}

/**
 * Возвращает сумму свободных платежей семьи (без привязки к цели)
 * @param {string} familyLabelOrId — метка или ID семьи
 * @returns {number}
 * @customfunction
 */
function FREE_PAYMENTS_FAMILY(familyLabelOrId) {
  const famId = getIdFromLabelish_(familyLabelOrId);
  if (!famId) return 0;
  
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  if (!shP) return 0;
  
  const rows = shP.getLastRow();
  if (rows < 2) return 0;
  
  const map = getHeaderMap_(shP);
  const goalCol = version === 'v1' ? 'collection_id (label)' : 'goal_id (label)';
  const iFam = map['family_id (label)'];
  const iGoal = map[goalCol];
  const iSum = map['Сумма'];
  if (!iFam || !iSum) return 0;
  
  const vals = shP.getRange(2, 1, rows - 1, shP.getLastColumn()).getValues();
  let total = 0;
  
  vals.forEach(r => {
    const fid = getIdFromLabelish_(String(r[iFam - 1] || ''));
    const gid = iGoal ? getIdFromLabelish_(String(r[iGoal - 1] || '')) : '';
    const sum = Number(r[iSum - 1] || 0);
    
    // Свободный платёж = без привязки к цели
    if (fid === famId && sum > 0 && !gid) {
      total += sum;
    }
  });
  
  return round2_(total);
}

// ============================================================================
// Функции для работы с периодом членства
// ============================================================================

/**
 * Возвращает дату начала членства семьи
 * @param {string} familyLabelOrId — метка или ID семьи
 * @returns {Date|string}
 * @customfunction
 */
function MEMBER_FROM(familyLabelOrId) {
  const famId = getIdFromLabelish_(familyLabelOrId);
  if (!famId) return '';
  
  const ss = SpreadsheetApp.getActive();
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  if (!shF) return '';
  
  const families = readFamilies_(shF);
  const fam = families.get(famId);
  return fam && fam.memberFrom ? fam.memberFrom : '';
}

/**
 * Возвращает дату окончания членства семьи
 * @param {string} familyLabelOrId — метка или ID семьи
 * @returns {Date|string}
 * @customfunction
 */
function MEMBER_TO(familyLabelOrId) {
  const famId = getIdFromLabelish_(familyLabelOrId);
  if (!famId) return '';
  
  const ss = SpreadsheetApp.getActive();
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  if (!shF) return '';
  
  const families = readFamilies_(shF);
  const fam = families.get(famId);
  return fam && fam.memberTo ? fam.memberTo : '';
}

/**
 * Проверяет, является ли семья членом на указанную дату
 * @param {string} familyLabelOrId — метка или ID семьи
 * @param {Date} [date] — дата для проверки (по умолчанию — сегодня)
 * @returns {boolean}
 * @customfunction
 */
function IS_MEMBER_ON_DATE(familyLabelOrId, date) {
  const famId = getIdFromLabelish_(familyLabelOrId);
  if (!famId) return false;
  
  const checkDate = date ? parseDate_(date) : new Date();
  if (!checkDate) return false;
  
  const ss = SpreadsheetApp.getActive();
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  if (!shF) return false;
  
  const families = readFamilies_(shF);
  const fam = families.get(famId);
  if (!fam) return false;
  
  // Если нет дат членства — считаем членом всегда
  if (!fam.memberFrom && !fam.memberTo) return fam.active;
  
  // Проверяем период
  if (fam.memberFrom && checkDate < fam.memberFrom) return false;
  if (fam.memberTo && checkDate > fam.memberTo) return false;
  
  return true;
}

/**
 * Возвращает баланс семьи на дату окончания членства (для выплаты при уходе)
 * @param {string} familyLabelOrId — метка или ID семьи
 * @returns {number}
 * @customfunction
 */
function EXIT_BALANCE(familyLabelOrId) {
  const famId = getIdFromLabelish_(familyLabelOrId);
  if (!famId) return 0;
  
  const ss = SpreadsheetApp.getActive();
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  if (!shF) return 0;
  
  const families = readFamilies_(shF);
  const fam = families.get(famId);
  if (!fam || !fam.memberTo) {
    // Если семья не ушла — возвращаем текущий баланс
    return BALANCE_FAMILY(familyLabelOrId);
  }
  
  // Для ушедшей семьи — считаем баланс с учётом только тех целей,
  // в периоды которых семья была членом
  const version = detectVersion();
  const shG = version === 'v1' 
    ? ss.getSheetByName(SHEET_NAMES.COLLECTIONS) 
    : ss.getSheetByName(SHEET_NAMES.GOALS);
  const shU = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  
  if (!shG || !shU || !shP) return 0;
  
  const participation = readParticipation_(shU, version);
  const payments = readPayments_(shP, version);
  const goals = readGoals_(shG, version, 'ALL');
  
  // Считаем внесённое (все платежи семьи)
  const paid = PAYED_TOTAL_FAMILY(familyLabelOrId);
  
  // Считаем начисленное только по целям, где семья была членом
  let totalAccrued = 0;
  
  goals.forEach((goal, goalId) => {
    // Проверяем, была ли семья членом в период цели
    if (!isFamilyMemberInPeriod_(fam, goal.startDate, goal.deadline)) return;
    
    const participants = resolveParticipants_(goalId, families, participation, goal);
    const goalPayments = payments.get(goalId) || new Map();
    
    if (participants.size === 0) {
      goalPayments.forEach((_, fid) => participants.add(fid));
    }
    
    const { x, kPayers } = precalculateForGoal_(goal, participants, goalPayments);
    const accrued = calculateAccrual_(famId, goal, participants, goalPayments, x, kPayers);
    
    totalAccrued += accrued;
  });
  
  return round2_(paid - totalAccrued);
}

/**
 * Возвращает количество месяцев членства семьи в указанном году
 * @param {string} familyLabelOrId — метка или ID семьи  
 * @param {number} [year] — год (по умолчанию — текущий)
 * @returns {number}
 * @customfunction
 */
function MEMBERSHIP_MONTHS(familyLabelOrId, year) {
  const famId = getIdFromLabelish_(familyLabelOrId);
  if (!famId) return 0;
  
  const targetYear = year || new Date().getFullYear();
  const yearStart = new Date(targetYear, 0, 1);
  const yearEnd = new Date(targetYear, 11, 31);
  
  const ss = SpreadsheetApp.getActive();
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  if (!shF) return 0;
  
  const families = readFamilies_(shF);
  const fam = families.get(famId);
  if (!fam) return 0;
  
  // Определяем период членства в году
  const memberStart = fam.memberFrom && fam.memberFrom > yearStart ? fam.memberFrom : yearStart;
  const memberEnd = fam.memberTo && fam.memberTo < yearEnd ? fam.memberTo : yearEnd;
  
  if (memberStart > memberEnd) return 0;
  
  // Считаем количество месяцев
  const startMonth = memberStart.getMonth();
  const endMonth = memberEnd.getMonth();
  
  return endMonth - startMonth + 1;
}

/**
 * Коэффициент участия семьи (для пропорционального расчёта взносов)
 * @param {string} familyLabelOrId — метка или ID семьи
 * @param {number} [year] — год (по умолчанию — текущий)
 * @returns {number} — коэффициент от 0 до 1
 * @customfunction
 */
function MEMBERSHIP_RATIO(familyLabelOrId, year) {
  const months = MEMBERSHIP_MONTHS(familyLabelOrId, year);
  return round2_(months / 12);
}

/**
 * @fileoverview Лист «Детализация» — настройка и обновление
 */

/**
 * Настраивает лист «Детализация»
 */
function setupDetailSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAMES.DETAIL);
  if (!sh) return;
  
  // Очищаем старые данные
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).clearContent();
  }
  
  // Очищаем старые места служебных ячеек (J, K) если были
  try { 
    sh.getRange('J1:K3').clearContent().clearDataValidations(); 
    sh.getRange('J1:K3').clearNote();
  } catch (_) {}
  
  // Селектор фильтра — колонки L, M (данные занимают A-H)
  sh.getRange('M1').clearDataValidations().setValue('Фильтр');
  sh.getRange('L1').clearDataValidations().setValue('ALL');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['OPEN', 'ALL'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('L1').setDataValidation(rule).setHorizontalAlignment('center');
  sh.getRange('L1').setNote('OPEN (только открытые) или ALL (все цели)');
  
  // Тикер для принудительного пересчёта
  sh.getRange('M2').clearDataValidations().setValue('Tick');
  sh.getRange('L2').clearDataValidations().setValue(new Date().toISOString());
  sh.getRange('M3').clearDataValidations().setValue('Детализация платежей и начислений. Автообновляется.');
  
  // Динамическая формула
  sh.getRange('A2').setFormula(`=GENERATE_DETAIL_BREAKDOWN(IF(LEN($L$1)=0,"ALL",$L$1), $L$2)`);
}

/**
 * Обновляет лист «Детализация»
 */
function refreshDetailSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAMES.DETAIL);
  if (!sh) return;
  
  // Очищаем старые места служебных ячеек (J, K) если были
  try { 
    sh.getRange('J1:K3').clearContent().clearDataValidations(); 
    sh.getRange('J1:K3').clearNote();
  } catch (_) {}
  
  // Обновляем тикер в новом месте (L2)
  sh.getRange('L2').setValue(new Date().toISOString());
  
  // Обновляем формулу на новые ссылки
  const newFormula = `=GENERATE_DETAIL_BREAKDOWN(IF(LEN($L$1)=0,"ALL",$L$1), $L$2)`;
  sh.getRange('A2').setFormula(newFormula);
  
  SpreadsheetApp.flush();
  try {
    styleSheetHeader_(sh);
    styleDetailSheet_(sh);
  } catch (_) {}
}

/**
 * Генерирует детализацию платежей и начислений
 * @param {string} statusFilter — OPEN или ALL
 * @param {string} tick — тикер для пересчёта
 * @returns {Array<Array>}
 * @customfunction
 */
function GENERATE_DETAIL_BREAKDOWN(statusFilter, tick) {
  const onlyOpen = String(statusFilter || 'OPEN').toUpperCase() !== 'ALL';
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  const shG = version === 'v1' 
    ? ss.getSheetByName(SHEET_NAMES.COLLECTIONS) 
    : ss.getSheetByName(SHEET_NAMES.GOALS);
  const shU = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  
  if (!shF || !shG || !shU || !shP) return [['', '', '', '', '', '', '', '']];
  
  // Читаем данные
  const families = readFamilies_(shF);
  const goals = readGoals_(shG, version, onlyOpen);
  const participation = readParticipation_(shU, version);
  const payments = readPayments_(shP, version);
  
  if (goals.size === 0 || families.size === 0) {
    return [['', '', '', '', '', '', '', '']];
  }
  
  // Собираем результат
  const out = [];
  
  goals.forEach((goal, goalId) => {
    const participants = resolveParticipants_(goalId, families, participation, goal);
    const goalPayments = payments.get(goalId) || new Map();
    
    // Предрасчёт для режимов
    const { x, kPayers } = precalculateForGoal_(goal, participants, goalPayments);
    
    // Объединяем участников и плательщиков
    const famSet = new Set();
    participants.forEach(fid => famSet.add(fid));
    goalPayments.forEach((_, fid) => famSet.add(fid));
    
    famSet.forEach(fid => {
      const fam = families.get(fid);
      const paid = goalPayments.get(fid) || 0;
      const accrued = calculateAccrual_(fid, goal, participants, goalPayments, x, kPayers);
      
      if (paid > 0 || accrued > 0) {
        out.push([
          fid,
          fam ? fam.name : '',
          goalId,
          goal.name,
          round2_(paid),
          round2_(accrued),
          round2_(paid - accrued),
          goal.accrual
        ]);
      }
    });
  });
  
  // v2: Добавляем свободные платежи (без goal_id)
  if (version !== 'v1') {
    const freePayments = payments.get('__FREE__');
    if (freePayments && freePayments.size > 0) {
      freePayments.forEach((paid, fid) => {
        const fam = families.get(fid);
        out.push([
          fid,
          fam ? fam.name : '',
          '',
          '— Свободный платёж —',
          round2_(paid),
          0,
          round2_(paid),
          'free'
        ]);
      });
    }
  }
  
  return out.length ? out : [['', '', '', '', '', '', '', '']];
}

/**
 * Читает семьи из листа
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @returns {Map<string, {name: string, active: boolean, memberFrom: Date|null, memberTo: Date|null}>}
 */
function readFamilies_(sh) {
  const families = new Map();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return families;
  
  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);
  
  data.forEach(r => {
    const id = String(r[idx['family_id']] || '').trim();
    if (!id) return;
    
    // Парсим даты членства
    const memberFromRaw = r[idx['Членство с']];
    const memberToRaw = r[idx['Членство по']];
    
    // Гибкая проверка активности: Да, да, TRUE, true, 1, пусто (по умолчанию активен)
    const activeRaw = r[idx['Активен']];
    const activeStr = String(activeRaw || '').trim().toLowerCase();
    const isActive = activeRaw === true || activeRaw === 1 || 
                     activeStr === 'да' || activeStr === 'true' || activeStr === '1' ||
                     activeStr === '' || activeStr === ACTIVE_STATUS.YES.toLowerCase() ||
                     (activeStr !== 'нет' && activeStr !== 'false' && activeStr !== '0' && activeStr !== ACTIVE_STATUS.NO.toLowerCase());
    
    families.set(id, {
      name: String(r[idx['Ребёнок ФИО']] || '').trim(),
      active: isActive,
      memberFrom: parseDate_(memberFromRaw),
      memberTo: parseDate_(memberToRaw)
    });
  });
  
  return families;
}

/**
 * Парсит дату из ячейки (может быть Date, строка или пусто)
 * @param {*} value
 * @returns {Date|null}
 */
function parseDate_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const d = new Date(value);
  return isNaN(d.getTime()) ? null : d;
}

/**
 * Читает цели/сборы из листа
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {string} version
 * @param {boolean} onlyOpen
 * @returns {Map<string, {name: string, status: string, accrual: string, T: number, fixedX: number}>}
 */
function readGoals_(sh, version, statusFilter) {
  const goals = new Map();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return goals;
  
  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);
  
  const idField = version === 'v1' ? 'collection_id' : 'goal_id';
  const nameField = version === 'v1' ? 'Название сбора' : 'Название цели';
  const openStatus = version === 'v1' ? COLLECTION_STATUS_V1.OPEN : GOAL_STATUS.OPEN;
  const closedStatus = version === 'v1' ? COLLECTION_STATUS_V1.CLOSED : GOAL_STATUS.CLOSED;
  
  // Нормализуем фильтр: boolean (legacy) или строка (OPEN/CLOSED/ALL/false)
  let filter = 'ALL';
  if (statusFilter === true) filter = 'OPEN';
  else if (statusFilter === false) filter = 'ALL';
  else if (typeof statusFilter === 'string') filter = statusFilter.toUpperCase();
  
  // Статус «Отменена» — всегда игнорируем в расчётах
  const cancelledStatus = version === 'v1' ? null : GOAL_STATUS.CANCELLED;
  
  data.forEach(r => {
    const id = String(r[idx[idField]] || '').trim();
    if (!id) return;
    
    const status = String(r[idx['Статус']] || '').trim();
    
    // Отменённые цели — всегда пропускаем
    if (cancelledStatus && status === cancelledStatus) return;
    
    // Применяем фильтр по статусу
    if (filter === 'OPEN' && status !== openStatus) return;
    if (filter === 'CLOSED' && status !== closedStatus) return;
    // filter === 'ALL' — берём все (кроме отменённых)
    
    goals.set(id, {
      name: String(r[idx[nameField]] || '').trim(),
      status: status,
      accrual: normalizeAccrualMode_(String(r[idx['Начисление']] || '').trim()),
      T: Number(r[idx['Параметр суммы']] || 0),
      fixedX: Number(r[idx['Фиксированный x']] || 0),
      startDate: parseDate_(r[idx['Дата начала']]),
      deadline: parseDate_(r[idx['Дедлайн']])
    });
  });
  
  return goals;
}

/**
 * Нормализует режим начисления (алиасы v1 → v2)
 * @param {string} mode
 * @returns {string}
 */
function normalizeAccrualMode_(mode) {
  return ACCRUAL_ALIASES[mode] || mode;
}

/**
 * Читает участие из листа
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {string} version
 * @returns {Map<string, {hasInclude: boolean, include: Set, exclude: Set}>}
 */
function readParticipation_(sh, version) {
  const partByGoal = new Map();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return partByGoal;
  
  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);
  
  const goalLabel = version === 'v1' ? 'collection_id (label)' : 'goal_id (label)';
  
  data.forEach(r => {
    const goalId = getIdFromLabelish_(String(r[idx[goalLabel]] || ''));
    const famId = getIdFromLabelish_(String(r[idx['family_id (label)']] || ''));
    const status = String(r[idx['Статус']] || '').trim();
    
    if (!goalId || !famId) return;
    
    if (!partByGoal.has(goalId)) {
      partByGoal.set(goalId, { hasInclude: false, include: new Set(), exclude: new Set() });
    }
    
    const obj = partByGoal.get(goalId);
    if (status === PARTICIPATION_STATUS.PARTICIPATES) {
      obj.hasInclude = true;
      obj.include.add(famId);
    } else if (status === PARTICIPATION_STATUS.NOT_PARTICIPATES) {
      obj.exclude.add(famId);
    }
  });
  
  return partByGoal;
}

/**
 * Читает платежи из листа
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {string} version
 * @returns {Map<string, Map<string, number>>} — goalId → (famId → sum), свободные платежи под ключом '__FREE__'
 */
function readPayments_(sh, version) {
  const payByGoal = new Map();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return payByGoal;
  
  const data = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = {};
  headers.forEach((h, i) => idx[h] = i);
  
  const goalLabel = version === 'v1' ? 'collection_id (label)' : 'goal_id (label)';
  
  data.forEach(r => {
    const goalIdRaw = String(r[idx[goalLabel]] || '').trim();
    const goalId = goalIdRaw ? getIdFromLabelish_(goalIdRaw) : '__FREE__'; // v2: свободный платёж
    const famId = getIdFromLabelish_(String(r[idx['family_id (label)']] || ''));
    const sum = Number(r[idx['Сумма']] || 0);
    
    // v1: требуем goal_id, v2: разрешаем свободные
    if (version === 'v1' && goalId === '__FREE__') return;
    if (!famId || sum <= 0) return;
    
    if (!payByGoal.has(goalId)) {
      payByGoal.set(goalId, new Map());
    }
    const m = payByGoal.get(goalId);
    m.set(famId, (m.get(famId) || 0) + sum);
  });
  
  return payByGoal;
}

/**
 * Определяет участников цели с учётом периода членства
 * @param {string} goalId
 * @param {Map} families
 * @param {Map} participation
 * @param {Object} [goal] — объект цели с startDate и deadline
 * @returns {Set<string>}
 */
function resolveParticipants_(goalId, families, participation, goal) {
  const p = participation.get(goalId);
  const participants = new Set();
  
  // Определяем период цели для проверки членства
  const goalStart = goal && goal.startDate ? goal.startDate : null;
  const goalEnd = goal && goal.deadline ? goal.deadline : null;
  
  if (p && p.hasInclude) {
    p.include.forEach(fid => {
      const fam = families.get(fid);
      if (fam && isFamilyMemberInPeriod_(fam, goalStart, goalEnd)) {
        participants.add(fid);
      }
    });
  } else {
    families.forEach((info, fid) => {
      if (info.active && isFamilyMemberInPeriod_(info, goalStart, goalEnd)) {
        participants.add(fid);
      }
    });
  }
  
  if (p) {
    p.exclude.forEach(fid => participants.delete(fid));
  }
  
  return participants;
}

/**
 * Проверяет, была ли семья членом в указанный период
 * @param {Object} family — {memberFrom, memberTo}
 * @param {Date|null} periodStart
 * @param {Date|null} periodEnd
 * @returns {boolean}
 */
function isFamilyMemberInPeriod_(family, periodStart, periodEnd) {
  const memberFrom = family.memberFrom;
  const memberTo = family.memberTo;
  
  // Если нет дат членства — семья всегда участвует (для обратной совместимости)
  if (!memberFrom && !memberTo) return true;
  
  // Если есть memberTo и он раньше начала периода — семья уже ушла
  if (memberTo && periodStart && memberTo < periodStart) return false;
  
  // Если есть memberFrom и он позже конца периода — семья ещё не пришла
  if (memberFrom && periodEnd && memberFrom > periodEnd) return false;
  
  return true;
}

/**
 * Предварительные расчёты для цели
 * @param {Object} goal
 * @param {Set} participants
 * @param {Map} goalPayments
 * @returns {{x: number, kPayers: number}}
 */
function precalculateForGoal_(goal, participants, goalPayments) {
  let x = 0;
  let kPayers = 0;
  
  if (goal.accrual === ACCRUAL_MODES.DYNAMIC_BY_PAYERS) {
    if (goal.fixedX > 0) {
      x = goal.fixedX;
    } else {
      const arr = [];
      goalPayments.forEach((sum, fid) => {
        if (participants.has(fid) && sum > 0) arr.push(sum);
      });
      x = DYN_CAP_(goal.T, arr);
    }
  }
  
  if (goal.accrual === ACCRUAL_MODES.SHARED_TOTAL_BY_PAYERS) {
    goalPayments.forEach((sum, fid) => {
      if (participants.has(fid) && sum > 0) kPayers++;
    });
  }
  
  return { x, kPayers };
}

/**
 * Рассчитывает начисление для семьи по цели
 * @param {string} fid — family_id
 * @param {Object} goal
 * @param {Set} participants
 * @param {Map} goalPayments
 * @param {number} x — предрассчитанный cap
 * @param {number} kPayers — предрассчитанное количество плательщиков
 * @returns {number}
 */
function calculateAccrual_(fid, goal, participants, goalPayments, x, kPayers) {
  const paid = goalPayments.get(fid) || 0;
  const n = participants.size;
  
  switch (goal.accrual) {
    case ACCRUAL_MODES.STATIC_PER_FAMILY:
      return participants.has(fid) ? goal.T : 0;
      
    case ACCRUAL_MODES.SHARED_TOTAL_ALL:
      return (n > 0 && participants.has(fid)) ? (goal.T / n) : 0;
      
    case ACCRUAL_MODES.SHARED_TOTAL_BY_PAYERS:
      return (kPayers > 0 && participants.has(fid) && paid > 0) ? (goal.T / kPayers) : 0;
      
    case ACCRUAL_MODES.DYNAMIC_BY_PAYERS:
      return (participants.has(fid) && x > 0) ? Math.min(paid, x) : 0;
      
    case ACCRUAL_MODES.PROPORTIONAL_BY_PAYERS:
      if (participants.has(fid)) {
        let sumP = 0;
        goalPayments.forEach((sum, f2) => {
          if (participants.has(f2) && sum > 0) sumP += sum;
        });
        if (sumP > 0) {
          const target = Math.min(goal.T, sumP);
          return paid > 0 ? (paid * target / sumP) : 0;
        }
      }
      return 0;
      
    case ACCRUAL_MODES.UNIT_PRICE:
      const unitX = goal.fixedX > 0 ? goal.fixedX : 0;
      return (participants.has(fid) && unitX > 0) ? (Math.floor(paid / unitX) * unitX) : 0;
      
    case ACCRUAL_MODES.VOLUNTARY:
      // Добровольный взнос: начисление = 0, деньги остаются на балансе
      return 0;
      
    case ACCRUAL_MODES.FROM_BALANCE:
      // Списание с баланса: делим сумму цели поровну между участниками
      // Не зависит от платежей — списывается с общего баланса семьи
      return (n > 0 && participants.has(fid)) ? (goal.T / n) : 0;
      
    default:
      return 0;
  }
}

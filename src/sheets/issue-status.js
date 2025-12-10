/**
 * @fileoverview Лист «Статус выдачи» — для поштучных целей
 */

/**
 * Настраивает лист «Статус выдачи»
 */
function setupIssueStatusSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAMES.ISSUE_STATUS);
  if (!sh) return;
  
  // Очищаем старые данные
  const last = sh.getLastRow();
  if (last > 1) {
    sh.getRange(2, 1, last - 1, Math.max(1, sh.getLastColumn())).clearContent();
  }
  
  // Array formula
  sh.getRange('A2').setFormula('=GENERATE_ISSUE_STATUS()');
  
  SpreadsheetApp.flush();
  try {
    styleSheetHeader_(sh);
    styleIssueStatusSheet_(sh);
  } catch (_) {}
}

/**
 * Обновляет лист «Статус выдачи»
 */
function refreshIssueStatusSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAMES.ISSUE_STATUS);
  
  if (!sh) {
    // Создаём лист если отсутствует
    const spec = getSheetsSpec().find(s => s.name === SHEET_NAMES.ISSUE_STATUS);
    sh = ss.insertSheet(SHEET_NAMES.ISSUE_STATUS);
    if (spec) {
      sh.setFrozenRows(1);
      sh.getRange(1, 1, 1, spec.headers.length).setValues([spec.headers]);
      spec.colWidths?.forEach((w, i) => { if (w) sh.setColumnWidth(i + 1, w); });
    }
  }
  
  const cell = sh.getRange('A2');
  const f = cell.getFormula();
  if (!f || f.indexOf('GENERATE_ISSUE_STATUS') < 0) {
    cell.setFormula('=GENERATE_ISSUE_STATUS()');
  } else {
    cell.setFormula(f);
  }
  
  SpreadsheetApp.flush();
  try {
    styleSheetHeader_(sh);
    styleIssueStatusSheet_(sh);
  } catch (_) {}
}

/**
 * Генерирует статус выдачи для поштучных целей
 * @returns {Array<Array>}
 * @customfunction
 */
function GENERATE_ISSUE_STATUS() {
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  
  const shG = version === 'v1' 
    ? ss.getSheetByName(SHEET_NAMES.COLLECTIONS) 
    : ss.getSheetByName(SHEET_NAMES.GOALS);
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  const shI = ss.getSheetByName(SHEET_NAMES.ISSUES);
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  const shU = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  
  if (!shG) return [['', '', '', '', '', '', '', '']];
  
  const out = [];
  
  // Читаем цели
  const lastRowG = shG.getLastRow();
  if (lastRowG < 2) return [['', '', '', '', '', '', '', '']];
  
  const dataG = shG.getRange(2, 1, lastRowG - 1, shG.getLastColumn()).getValues();
  const headersG = shG.getRange(1, 1, 1, shG.getLastColumn()).getValues()[0];
  const idxG = {};
  headersG.forEach((h, i) => idxG[h] = i);
  
  const idField = version === 'v1' ? 'collection_id' : 'goal_id';
  const nameField = version === 'v1' ? 'Название сбора' : 'Название цели';
  const accrualField = version === 'v1' ? ACCRUAL_ALIASES['unit_price_by_payers'] || 'unit_price_by_payers' : ACCRUAL_MODES.UNIT_PRICE;
  
  // Читаем семьи (для участников)
  const families = readFamilies_(shF);
  const participation = readParticipation_(shU, version);
  const payments = readPayments_(shP, version);
  
  // Читаем выданные единицы
  const issuedByGoal = new Map();
  if (shI) {
    const lastRowI = shI.getLastRow();
    if (lastRowI >= 2) {
      const dataI = shI.getRange(2, 1, lastRowI - 1, shI.getLastColumn()).getValues();
      const headersI = shI.getRange(1, 1, 1, shI.getLastColumn()).getValues()[0];
      const idxI = {};
      headersI.forEach((h, i) => idxI[h] = i);
      
      const goalLabelI = version === 'v1' ? 'collection_id (label)' : 'goal_id (label)';
      
      dataI.forEach(r => {
        const goalId = getIdFromLabelish_(String(r[idxI[goalLabelI]] || ''));
        const units = Number(r[idxI['Единиц']] || 0);
        const ok = String(r[idxI['Выдано']] || '').trim() === ACTIVE_STATUS.YES;
        if (!goalId || !(units > 0) || !ok) return;
        issuedByGoal.set(goalId, (issuedByGoal.get(goalId) || 0) + units);
      });
    }
  }
  
  // Обрабатываем цели с флагом «К выдаче детям»
  dataG.forEach(row => {
    const id = String(row[idxG[idField]] || '').trim();
    if (!id) return;
    
    const name = String(row[idxG[nameField]] || '').trim();
    const status = String(row[idxG['Статус']] || '').trim();
    const mode = normalizeAccrualMode_(String(row[idxG['Начисление']] || '').trim());
    const flagIssue = String(row[idxG['К выдаче детям']] || '').trim() === ACTIVE_STATUS.YES;
    const T = Number(row[idxG['Параметр суммы']] || 0);
    const x = Number(row[idxG['Фиксированный x']] || 0);
    
    // Только цели с включённой выдачей и unit_price режимом
    if (!flagIssue) return;
    if (mode !== ACCRUAL_MODES.UNIT_PRICE && mode !== 'unit_price_by_payers') return;
    if (!(x > 0)) return;
    
    // Создаём объект goal для resolveParticipants_
    const goal = {
      startDate: parseDate_(row[idxG['Дата начала']]),
      deadline: parseDate_(row[idxG['Дедлайн']])
    };
    
    // Участники
    const participants = resolveParticipants_(id, families, participation, goal);
    const goalPayments = payments.get(id) || new Map();
    
    // Fallback
    if (participants.size === 0) {
      goalPayments.forEach((_, fid) => participants.add(fid));
    }
    
    // Единиц требуется
    const totalUnits = x > 0 ? Math.ceil(T / x) : '';
    
    // Единиц оплачено
    let unitsPaid = '';
    if (x > 0) {
      let sumUnits = 0;
      goalPayments.forEach((sum, fid) => {
        if (!participants.has(fid)) return;
        if (sum > 0) sumUnits += Math.floor(sum / x);
      });
      unitsPaid = sumUnits;
    }
    
    const unitsIssued = issuedByGoal.get(id) || 0;
    
    // Остаток = min(требуется, оплачено) - выдано
    const capUnits = (typeof totalUnits === 'number') ? totalUnits : unitsPaid;
    const remainUnits = (unitsPaid === '' ? '' : Math.max(0, Math.min(capUnits, unitsPaid) - unitsIssued));
    
    out.push([
      id,
      name,
      status,
      round2_(x),
      totalUnits === '' ? '' : totalUnits,
      unitsPaid === '' ? '' : unitsPaid,
      unitsIssued,
      remainUnits === '' ? '' : remainUnits
    ]);
  });
  
  return out.length ? out : [['', '', '', '', '', '', '', '']];
}

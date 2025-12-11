// Extracted pure helpers mirrored from Code.gs (no Apps Script deps)

// ============================================================================
// Утилиты
// ============================================================================

export function getIdFromLabelish(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  const m = s.match(/\(([^)]+)\)\s*$/);
  return m ? m[1] : s;
}

export function round6(x) {
  return Math.round((x + Number.EPSILON) * 1e6) / 1e6;
}

export function round2(x) {
  return Math.round((x + Number.EPSILON) * 100) / 100;
}

export function flatten(arr) {
  const out = [];
  (arr || []).forEach(r => Array.isArray(r) ? r.forEach(c => out.push(c)) : out.push(r));
  return out;
}

export function colToLetter(col) {
  let s = '';
  let c = col;
  while (c > 0) {
    const r = (c - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

export function nextId(prefix, index, pad) {
  const width = pad || 3;
  let n = String(index);
  while (n.length < width) n = '0' + n;
  return prefix + n;
}

// ============================================================================
// Water-filling algorithm (DYN_CAP)
// ============================================================================

export function dynCap(T, payments) {
  if (!T || !isFinite(T)) return 0;
  const arr = (payments || []).map(Number).filter(v => v > 0 && isFinite(v));
  if (!arr.length) return 0;
  arr.sort((a,b)=>a-b);
  const n = arr.length;
  const sum = arr.reduce((a,b)=>a+b,0);
  const target = Math.min(T, sum);
  if (target <= 0) return 0;

  let cumsum = 0;
  for (let k = 0; k < n; k++) {
    const next = arr[k];
    const remain = n - k;
    const candidate = (target - cumsum) / remain;
    if (candidate <= next) return round6(candidate);
    cumsum += next;
  }
  return round6((target - (cumsum - arr[n-1])) / 1);
}

// ============================================================================
// Режимы начисления (Accrual modes)
// ============================================================================

export const ACCRUAL_MODES = {
  STATIC_PER_FAMILY: 'static_per_family',
  SHARED_TOTAL_ALL: 'shared_total_all',
  SHARED_TOTAL_BY_PAYERS: 'shared_total_by_payers',
  DYNAMIC_BY_PAYERS: 'dynamic_by_payers',
  PROPORTIONAL_BY_PAYERS: 'proportional_by_payers',
  UNIT_PRICE: 'unit_price',
  VOLUNTARY: 'voluntary',
};

/**
 * Рассчитывает начисление для семьи по цели
 * @param {string} fid — family_id
 * @param {Object} goal — цель с полями: accrual, T, fixedX
 * @param {Set<string>} participants — множество участников
 * @param {Map<string, number>} goalPayments — платежи по цели (family_id -> сумма)
 * @param {number} x — предрассчитанный cap для dynamic_by_payers
 * @param {number} kPayers — количество плательщиков для shared_total_by_payers
 * @returns {number}
 */
export function calculateAccrual(fid, goal, participants, goalPayments, x, kPayers) {
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

    default:
      return 0;
  }
}

/**
 * Предварительные расчёты для цели
 * @param {Object} goal
 * @param {Set<string>} participants
 * @param {Map<string, number>} goalPayments
 * @returns {{x: number, kPayers: number}}
 */
export function precalculateForGoal(goal, participants, goalPayments) {
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
      x = dynCap(goal.T, arr);
    }
  }

  if (goal.accrual === ACCRUAL_MODES.SHARED_TOTAL_BY_PAYERS) {
    goalPayments.forEach((sum, fid) => {
      if (participants.has(fid) && sum > 0) kPayers++;
    });
  }

  return { x, kPayers };
}

// ============================================================================
// Членство (Membership)
// ============================================================================

/**
 * Парсит дату из различных форматов
 * @param {Date|string|number} val
 * @returns {Date|null}
 */
export function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) return isNaN(val.getTime()) ? null : val;
  if (typeof val === 'string') {
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  if (typeof val === 'number') {
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

/**
 * Проверяет, была ли семья членом в период цели
 * @param {Object} fam — семья с полями memberFrom, memberTo
 * @param {Date|null} goalStart — дата начала цели
 * @param {Date|null} goalEnd — дата окончания/дедлайна цели
 * @returns {boolean}
 */
export function isFamilyMemberInPeriod(fam, goalStart, goalEnd) {
  // Если нет дат членства — семья член всегда
  if (!fam.memberFrom && !fam.memberTo) return true;

  // Если цель без периода — проверяем текущую активность
  if (!goalStart && !goalEnd) {
    return true; // членство не ограничивает
  }

  // Проверяем пересечение периодов
  // Период членства: [memberFrom, memberTo]
  // Период цели: [goalStart, goalEnd]
  // Пересекаются, если: memberStart <= goalEnd && memberEnd >= goalStart

  const memberStart = fam.memberFrom || new Date('1970-01-01');
  const memberEnd = fam.memberTo || new Date('2100-01-01');
  const gStart = goalStart || new Date('1970-01-01');
  const gEnd = goalEnd || new Date('2100-01-01');

  return memberStart <= gEnd && memberEnd >= gStart;
}

/**
 * Возвращает количество месяцев членства семьи в указанном году
 * @param {Object} fam — семья с полями memberFrom, memberTo
 * @param {number} year — год
 * @returns {number}
 */
export function membershipMonths(fam, year) {
  const yearStart = new Date(year, 0, 1);
  const yearEnd = new Date(year, 11, 31);

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
 * Коэффициент участия семьи
 * @param {Object} fam — семья с полями memberFrom, memberTo
 * @param {number} year — год
 * @returns {number} — коэффициент от 0 до 1
 */
export function membershipRatio(fam, year) {
  const months = membershipMonths(fam, year);
  return round2(months / 12);
}

// ============================================================================
// Алиасы и нормализация режимов начисления
// ============================================================================

// Алиасы для обратной совместимости v1.x
export const ACCRUAL_ALIASES = {
  'static_per_child': ACCRUAL_MODES.STATIC_PER_FAMILY,
  'unit_price_by_payers': ACCRUAL_MODES.UNIT_PRICE
};

/**
 * Нормализует название режима начисления (алиасы v1 → v2)
 * @param {string} mode
 * @returns {string}
 */
export function normalizeAccrualMode(mode) {
  return ACCRUAL_ALIASES[mode] || mode;
}

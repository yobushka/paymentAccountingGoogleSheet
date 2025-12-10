import { describe, it, expect } from 'vitest';
import { 
  getIdFromLabelish, 
  dynCap, 
  round2, 
  round6, 
  flatten, 
  colToLetter, 
  nextId,
  ACCRUAL_MODES,
  ACCRUAL_ALIASES,
  normalizeAccrualMode,
  calculateAccrual,
  precalculateForGoal,
  parseDate,
  isFamilyMemberInPeriod,
  membershipMonths,
  membershipRatio
} from './pure.js';

// ============================================================================
// Утилиты
// ============================================================================

describe('getIdFromLabelish', () => {
  it('extracts ID from label', () => {
    expect(getIdFromLabelish('Новый год (C002)')).toBe('C002');
  });
  it('returns same when already ID', () => {
    expect(getIdFromLabelish('F001')).toBe('F001');
  });
  it('handles empty/null', () => {
    expect(getIdFromLabelish('')).toBe('');
    expect(getIdFromLabelish(null)).toBe('');
  });
  it('handles complex names with parentheses', () => {
    expect(getIdFromLabelish('Подарок учителю (на 8 марта) (G005)')).toBe('G005');
  });
});

describe('round2', () => {
  it('rounds to 2 decimal places', () => {
    expect(round2(1.234)).toBe(1.23);
    expect(round2(1.235)).toBe(1.24);
    expect(round2(100 / 3)).toBeCloseTo(33.33, 2);
  });
});

describe('round6', () => {
  it('rounds to 6 decimal places', () => {
    expect(round6(1.2345678)).toBe(1.234568);
    expect(round6(1.2345674)).toBe(1.234567);
  });
});

describe('flatten', () => {
  it('flattens nested arrays', () => {
    expect(flatten([[1, 2], [3, 4]])).toEqual([1, 2, 3, 4]);
  });
  it('handles single values', () => {
    expect(flatten([1, 2, 3])).toEqual([1, 2, 3]);
  });
  it('handles empty array', () => {
    expect(flatten([])).toEqual([]);
  });
  it('handles mixed array', () => {
    expect(flatten([[1, 2], 3, [4]])).toEqual([1, 2, 3, 4]);
  });
});

describe('colToLetter', () => {
  it('converts column numbers to letters', () => {
    expect(colToLetter(1)).toBe('A');
    expect(colToLetter(26)).toBe('Z');
    expect(colToLetter(27)).toBe('AA');
    expect(colToLetter(28)).toBe('AB');
    expect(colToLetter(52)).toBe('AZ');
    expect(colToLetter(53)).toBe('BA');
  });
});

describe('nextId', () => {
  it('generates IDs with default padding', () => {
    expect(nextId('F', 1)).toBe('F001');
    expect(nextId('F', 10)).toBe('F010');
    expect(nextId('F', 100)).toBe('F100');
  });
  it('generates IDs with custom padding', () => {
    expect(nextId('PMT', 1, 4)).toBe('PMT0001');
    expect(nextId('G', 5, 2)).toBe('G05');
  });
});

// ============================================================================
// Water-filling algorithm (DYN_CAP)
// ============================================================================

describe('dynCap (water-filling)', () => {
  it('basic symmetry example', () => {
    const r1 = dynCap(500, [2000, 1333]);
    const r2 = dynCap(500, [1333, 2000]);
    expect(r1).toBeCloseTo(r2, 6);
  });
  it('when sum < T, all payments taken fully', () => {
    // T=9000, payments: [2000,2000,700,700,700,700,700] (sum=7500) => target=7500
    // Поскольку sum < T, все платежи берутся полностью, x = max(P) = 2000
    const x = dynCap(9000, [2000,2000,700,700,700,700,700]);
    expect(x).toBe(2000);
  });
  it('when sum == T, all payments taken fully', () => {
    // T=6000, payments=[1000,2000,3000] (sum=6000)
    // Сумма равна цели, все платежи берутся полностью, x = max(P) = 3000
    const x = dynCap(6000, [1000,2000,3000]);
    expect(x).toBe(3000);
  });
  it('proper capping when T < sum', () => {
    // T=3000, payments=[1000,2000,3000] (sum=6000)
    // Нужно найти x: min(1000,x)+min(2000,x)+min(3000,x) = 3000
    // При x=1000: 1000+1000+1000 = 3000 ✓
    const x = dynCap(3000, [1000,2000,3000]);
    expect(x).toBe(1000);
  });
  it('equal payments scenario', () => {
    // T=300, payments=[100,100,100] (sum=300), x=100
    const x = dynCap(300, [100,100,100]);
    expect(x).toBe(100);
  });
  it('zero and invalid inputs', () => {
    expect(dynCap(0, [1,2,3])).toBe(0);
    expect(dynCap(100, [])).toBe(0);
    expect(dynCap(100, [null, -5])).toBe(0);
  });
  it('uneven cap distribution', () => {
    // T=1500, payments=[500, 1000, 2000]
    // sum=3500, target=1500
    // При x=500: 500+500+500 = 1500 ✓
    const x = dynCap(1500, [500, 1000, 2000]);
    expect(x).toBe(500);
  });
  it('partial cap scenario', () => {
    // T=2500, payments=[500, 1000, 2000]
    // sum=3500, target=2500
    // При x=500: 500+500+500 = 1500 (мало)
    // После 500: остаток = 2500-500 = 2000, делим на 2: x = 500 + 2000/2 = 1500
    // При x=1000: 500+1000+1000 = 2500 ✓
    const x = dynCap(2500, [500, 1000, 2000]);
    expect(x).toBe(1000);
  });
});

// ============================================================================
// Режимы начисления (calculateAccrual)
// ============================================================================

describe('calculateAccrual', () => {
  const makeGoal = (accrual, T, fixedX = 0) => ({ accrual, T, fixedX });
  
  describe('static_per_family', () => {
    it('returns T for participant', () => {
      const goal = makeGoal(ACCRUAL_MODES.STATIC_PER_FAMILY, 1000);
      const participants = new Set(['F001', 'F002']);
      const payments = new Map();
      
      expect(calculateAccrual('F001', goal, participants, payments, 0, 0)).toBe(1000);
    });
    
    it('returns 0 for non-participant', () => {
      const goal = makeGoal(ACCRUAL_MODES.STATIC_PER_FAMILY, 1000);
      const participants = new Set(['F001']);
      const payments = new Map();
      
      expect(calculateAccrual('F002', goal, participants, payments, 0, 0)).toBe(0);
    });
  });

  describe('shared_total_all', () => {
    it('divides T equally among participants', () => {
      const goal = makeGoal(ACCRUAL_MODES.SHARED_TOTAL_ALL, 3000);
      const participants = new Set(['F001', 'F002', 'F003']);
      const payments = new Map();
      
      expect(calculateAccrual('F001', goal, participants, payments, 0, 0)).toBe(1000);
      expect(calculateAccrual('F002', goal, participants, payments, 0, 0)).toBe(1000);
    });
    
    it('returns 0 for non-participant', () => {
      const goal = makeGoal(ACCRUAL_MODES.SHARED_TOTAL_ALL, 3000);
      const participants = new Set(['F001', 'F002']);
      const payments = new Map();
      
      expect(calculateAccrual('F003', goal, participants, payments, 0, 0)).toBe(0);
    });
    
    it('returns 0 when no participants', () => {
      const goal = makeGoal(ACCRUAL_MODES.SHARED_TOTAL_ALL, 3000);
      const participants = new Set();
      const payments = new Map();
      
      expect(calculateAccrual('F001', goal, participants, payments, 0, 0)).toBe(0);
    });
  });

  describe('shared_total_by_payers', () => {
    it('divides T among payers only', () => {
      const goal = makeGoal(ACCRUAL_MODES.SHARED_TOTAL_BY_PAYERS, 3000);
      const participants = new Set(['F001', 'F002', 'F003']);
      const payments = new Map([['F001', 1500], ['F002', 1500]]);
      const kPayers = 2; // только двое заплатили
      
      expect(calculateAccrual('F001', goal, participants, payments, 0, kPayers)).toBe(1500);
      expect(calculateAccrual('F002', goal, participants, payments, 0, kPayers)).toBe(1500);
      expect(calculateAccrual('F003', goal, participants, payments, 0, kPayers)).toBe(0);
    });
  });

  describe('dynamic_by_payers', () => {
    it('caps payment to x', () => {
      const goal = makeGoal(ACCRUAL_MODES.DYNAMIC_BY_PAYERS, 3000);
      const participants = new Set(['F001', 'F002']);
      const payments = new Map([['F001', 2000], ['F002', 1000]]);
      const x = 1000;
      
      expect(calculateAccrual('F001', goal, participants, payments, x, 0)).toBe(1000);
      expect(calculateAccrual('F002', goal, participants, payments, x, 0)).toBe(1000);
    });
    
    it('returns full payment if less than x', () => {
      const goal = makeGoal(ACCRUAL_MODES.DYNAMIC_BY_PAYERS, 3000);
      const participants = new Set(['F001', 'F002']);
      const payments = new Map([['F001', 500], ['F002', 2000]]);
      const x = 1000;
      
      expect(calculateAccrual('F001', goal, participants, payments, x, 0)).toBe(500);
      expect(calculateAccrual('F002', goal, participants, payments, x, 0)).toBe(1000);
    });
    
    it('returns 0 for non-participant', () => {
      const goal = makeGoal(ACCRUAL_MODES.DYNAMIC_BY_PAYERS, 3000);
      const participants = new Set(['F001']);
      const payments = new Map([['F002', 1000]]);
      const x = 1000;
      
      expect(calculateAccrual('F002', goal, participants, payments, x, 0)).toBe(0);
    });
  });

  describe('proportional_by_payers', () => {
    it('distributes proportionally when sum > T', () => {
      const goal = makeGoal(ACCRUAL_MODES.PROPORTIONAL_BY_PAYERS, 1000);
      const participants = new Set(['F001', 'F002']);
      const payments = new Map([['F001', 1500], ['F002', 500]]); // sum = 2000
      
      // target = min(1000, 2000) = 1000
      // F001: 1500 * 1000 / 2000 = 750
      // F002: 500 * 1000 / 2000 = 250
      expect(calculateAccrual('F001', goal, participants, payments, 0, 0)).toBe(750);
      expect(calculateAccrual('F002', goal, participants, payments, 0, 0)).toBe(250);
    });
    
    it('uses full sum when sum < T', () => {
      const goal = makeGoal(ACCRUAL_MODES.PROPORTIONAL_BY_PAYERS, 5000);
      const participants = new Set(['F001', 'F002']);
      const payments = new Map([['F001', 600], ['F002', 400]]); // sum = 1000
      
      // target = min(5000, 1000) = 1000
      // F001: 600 * 1000 / 1000 = 600
      // F002: 400 * 1000 / 1000 = 400
      expect(calculateAccrual('F001', goal, participants, payments, 0, 0)).toBe(600);
      expect(calculateAccrual('F002', goal, participants, payments, 0, 0)).toBe(400);
    });
    
    it('returns 0 for non-payer', () => {
      const goal = makeGoal(ACCRUAL_MODES.PROPORTIONAL_BY_PAYERS, 1000);
      const participants = new Set(['F001', 'F002', 'F003']);
      const payments = new Map([['F001', 500], ['F002', 500]]);
      
      expect(calculateAccrual('F003', goal, participants, payments, 0, 0)).toBe(0);
    });
  });

  describe('unit_price', () => {
    it('rounds down to whole units', () => {
      const goal = makeGoal(ACCRUAL_MODES.UNIT_PRICE, 5000, 100);
      const participants = new Set(['F001', 'F002']);
      const payments = new Map([['F001', 350], ['F002', 199]]);
      
      // F001: floor(350/100) * 100 = 300
      // F002: floor(199/100) * 100 = 100
      expect(calculateAccrual('F001', goal, participants, payments, 0, 0)).toBe(300);
      expect(calculateAccrual('F002', goal, participants, payments, 0, 0)).toBe(100);
    });
    
    it('returns 0 when fixedX is 0', () => {
      const goal = makeGoal(ACCRUAL_MODES.UNIT_PRICE, 5000, 0);
      const participants = new Set(['F001']);
      const payments = new Map([['F001', 350]]);
      
      expect(calculateAccrual('F001', goal, participants, payments, 0, 0)).toBe(0);
    });
  });

  describe('voluntary', () => {
    it('returns full payment for participant', () => {
      const goal = makeGoal(ACCRUAL_MODES.VOLUNTARY, 0);
      const participants = new Set(['F001', 'F002']);
      const payments = new Map([['F001', 1234], ['F002', 5678]]);
      
      expect(calculateAccrual('F001', goal, participants, payments, 0, 0)).toBe(1234);
      expect(calculateAccrual('F002', goal, participants, payments, 0, 0)).toBe(5678);
    });
    
    it('returns 0 for non-participant', () => {
      const goal = makeGoal(ACCRUAL_MODES.VOLUNTARY, 0);
      const participants = new Set(['F001']);
      const payments = new Map([['F002', 1000]]);
      
      expect(calculateAccrual('F002', goal, participants, payments, 0, 0)).toBe(0);
    });
  });

  describe('unknown mode', () => {
    it('returns 0 for unknown accrual mode', () => {
      const goal = { accrual: 'unknown_mode', T: 1000, fixedX: 0 };
      const participants = new Set(['F001']);
      const payments = new Map([['F001', 500]]);
      
      expect(calculateAccrual('F001', goal, participants, payments, 0, 0)).toBe(0);
    });
  });
});

// ============================================================================
// precalculateForGoal
// ============================================================================

describe('precalculateForGoal', () => {
  it('calculates x for dynamic_by_payers', () => {
    const goal = { accrual: ACCRUAL_MODES.DYNAMIC_BY_PAYERS, T: 3000, fixedX: 0 };
    const participants = new Set(['F001', 'F002', 'F003']);
    const payments = new Map([['F001', 1000], ['F002', 2000], ['F003', 3000]]);
    
    const { x, kPayers } = precalculateForGoal(goal, participants, payments);
    
    expect(x).toBe(1000); // dynCap(3000, [1000, 2000, 3000])
    expect(kPayers).toBe(0);
  });
  
  it('uses fixedX when set', () => {
    const goal = { accrual: ACCRUAL_MODES.DYNAMIC_BY_PAYERS, T: 3000, fixedX: 500 };
    const participants = new Set(['F001', 'F002']);
    const payments = new Map([['F001', 1000], ['F002', 2000]]);
    
    const { x, kPayers } = precalculateForGoal(goal, participants, payments);
    
    expect(x).toBe(500);
  });
  
  it('calculates kPayers for shared_total_by_payers', () => {
    const goal = { accrual: ACCRUAL_MODES.SHARED_TOTAL_BY_PAYERS, T: 3000, fixedX: 0 };
    const participants = new Set(['F001', 'F002', 'F003']);
    const payments = new Map([['F001', 1000], ['F002', 500]]);
    
    const { x, kPayers } = precalculateForGoal(goal, participants, payments);
    
    expect(x).toBe(0);
    expect(kPayers).toBe(2); // F001 и F002 заплатили
  });
  
  it('excludes non-participants from kPayers', () => {
    const goal = { accrual: ACCRUAL_MODES.SHARED_TOTAL_BY_PAYERS, T: 3000, fixedX: 0 };
    const participants = new Set(['F001']); // только F001 участник
    const payments = new Map([['F001', 1000], ['F002', 500]]); // F002 заплатил, но не участник
    
    const { x, kPayers } = precalculateForGoal(goal, participants, payments);
    
    expect(kPayers).toBe(1); // только F001
  });
});

// ============================================================================
// Членство (Membership)
// ============================================================================

describe('parseDate', () => {
  it('parses Date object', () => {
    const d = new Date(2024, 5, 15);
    expect(parseDate(d)).toEqual(d);
  });
  
  it('parses ISO string', () => {
    const d = parseDate('2024-06-15');
    expect(d?.getFullYear()).toBe(2024);
    expect(d?.getMonth()).toBe(5); // 0-indexed
  });
  
  it('returns null for empty/invalid', () => {
    expect(parseDate('')).toBeNull();
    expect(parseDate(null)).toBeNull();
    expect(parseDate('invalid')).toBeNull();
  });
});

describe('isFamilyMemberInPeriod', () => {
  it('returns true when no membership dates set', () => {
    const fam = { memberFrom: null, memberTo: null };
    expect(isFamilyMemberInPeriod(fam, new Date(2024, 0, 1), new Date(2024, 11, 31))).toBe(true);
  });
  
  it('returns true when periods overlap', () => {
    const fam = { 
      memberFrom: new Date(2024, 0, 1), 
      memberTo: new Date(2024, 5, 30) 
    };
    // Цель: январь-март 2024
    expect(isFamilyMemberInPeriod(fam, new Date(2024, 0, 1), new Date(2024, 2, 31))).toBe(true);
  });
  
  it('returns false when periods do not overlap', () => {
    const fam = { 
      memberFrom: new Date(2024, 6, 1),  // с июля
      memberTo: new Date(2024, 11, 31) 
    };
    // Цель: январь-март 2024
    expect(isFamilyMemberInPeriod(fam, new Date(2024, 0, 1), new Date(2024, 2, 31))).toBe(false);
  });
  
  it('handles open-ended membership (memberTo not set)', () => {
    const fam = { 
      memberFrom: new Date(2024, 0, 1), 
      memberTo: null 
    };
    expect(isFamilyMemberInPeriod(fam, new Date(2025, 0, 1), new Date(2025, 11, 31))).toBe(true);
  });
  
  it('handles membership started later (memberFrom set)', () => {
    const fam = { 
      memberFrom: new Date(2024, 6, 1), // с июля 2024
      memberTo: null 
    };
    // Цель завершилась до начала членства
    expect(isFamilyMemberInPeriod(fam, new Date(2024, 0, 1), new Date(2024, 5, 30))).toBe(false);
  });
});

describe('membershipMonths', () => {
  it('returns 12 for full year membership', () => {
    const fam = { memberFrom: null, memberTo: null };
    expect(membershipMonths(fam, 2024)).toBe(12);
  });
  
  it('returns correct months for partial year', () => {
    const fam = { 
      memberFrom: new Date(2024, 2, 1),  // с марта
      memberTo: new Date(2024, 8, 30)     // по сентябрь
    };
    expect(membershipMonths(fam, 2024)).toBe(7); // март-сентябрь
  });
  
  it('returns 0 when not a member that year', () => {
    const fam = { 
      memberFrom: new Date(2025, 0, 1), 
      memberTo: new Date(2025, 11, 31) 
    };
    expect(membershipMonths(fam, 2024)).toBe(0);
  });
  
  it('handles year boundary correctly', () => {
    const fam = { 
      memberFrom: new Date(2023, 6, 1),  // с июля 2023
      memberTo: new Date(2024, 5, 30)    // по июнь 2024
    };
    expect(membershipMonths(fam, 2024)).toBe(6); // январь-июнь 2024
    expect(membershipMonths(fam, 2023)).toBe(6); // июль-декабрь 2023
  });
});

describe('membershipRatio', () => {
  it('returns 1 for full year', () => {
    const fam = { memberFrom: null, memberTo: null };
    expect(membershipRatio(fam, 2024)).toBe(1);
  });
  
  it('returns correct ratio for partial year', () => {
    const fam = { 
      memberFrom: new Date(2024, 0, 1),  // с января
      memberTo: new Date(2024, 5, 30)    // по июнь
    };
    expect(membershipRatio(fam, 2024)).toBe(0.5); // 6/12
  });
  
  it('returns 0 when not a member', () => {
    const fam = { 
      memberFrom: new Date(2025, 0, 1), 
      memberTo: new Date(2025, 11, 31) 
    };
    expect(membershipRatio(fam, 2024)).toBe(0);
  });
});

// ============================================================================
// Нормализация режимов начисления
// ============================================================================

describe('normalizeAccrualMode', () => {
  it('converts v1 aliases to v2 modes', () => {
    expect(normalizeAccrualMode('static_per_child')).toBe(ACCRUAL_MODES.STATIC_PER_FAMILY);
    expect(normalizeAccrualMode('unit_price_by_payers')).toBe(ACCRUAL_MODES.UNIT_PRICE);
  });
  
  it('returns original mode if not an alias', () => {
    expect(normalizeAccrualMode('dynamic_by_payers')).toBe('dynamic_by_payers');
    expect(normalizeAccrualMode('voluntary')).toBe('voluntary');
    expect(normalizeAccrualMode('unknown')).toBe('unknown');
  });
});

// ============================================================================
// Интеграционные тесты: полный расчёт для цели
// ============================================================================

describe('Integration: full goal calculation', () => {
  it('calculates static_per_family for multiple families', () => {
    const goal = { accrual: ACCRUAL_MODES.STATIC_PER_FAMILY, T: 500, fixedX: 0 };
    const participants = new Set(['F001', 'F002', 'F003']);
    const payments = new Map([['F001', 500], ['F002', 300], ['F003', 0]]);
    
    const { x, kPayers } = precalculateForGoal(goal, participants, payments);
    
    // Каждому участнику — по 500, независимо от платежа
    expect(calculateAccrual('F001', goal, participants, payments, x, kPayers)).toBe(500);
    expect(calculateAccrual('F002', goal, participants, payments, x, kPayers)).toBe(500);
    expect(calculateAccrual('F003', goal, participants, payments, x, kPayers)).toBe(500);
  });
  
  it('calculates dynamic_by_payers with water-filling', () => {
    // 3 семьи: платежи 500, 1000, 2000
    // Цель 1500
    // dynCap(1500, [500, 1000, 2000]) = 500
    // Начисления: 500, 500, 500
    const goal = { accrual: ACCRUAL_MODES.DYNAMIC_BY_PAYERS, T: 1500, fixedX: 0 };
    const participants = new Set(['F001', 'F002', 'F003']);
    const payments = new Map([['F001', 500], ['F002', 1000], ['F003', 2000]]);
    
    const { x, kPayers } = precalculateForGoal(goal, participants, payments);
    
    expect(x).toBe(500);
    expect(calculateAccrual('F001', goal, participants, payments, x, kPayers)).toBe(500);
    expect(calculateAccrual('F002', goal, participants, payments, x, kPayers)).toBe(500);
    expect(calculateAccrual('F003', goal, participants, payments, x, kPayers)).toBe(500);
  });
  
  it('calculates shared_total_by_payers correctly', () => {
    // 3 семьи: только 2 заплатили
    // Цель 1000, делим на 2 плательщиков = 500 каждому
    const goal = { accrual: ACCRUAL_MODES.SHARED_TOTAL_BY_PAYERS, T: 1000, fixedX: 0 };
    const participants = new Set(['F001', 'F002', 'F003']);
    const payments = new Map([['F001', 700], ['F002', 300]]);
    
    const { x, kPayers } = precalculateForGoal(goal, participants, payments);
    
    expect(kPayers).toBe(2);
    expect(calculateAccrual('F001', goal, participants, payments, x, kPayers)).toBe(500);
    expect(calculateAccrual('F002', goal, participants, payments, x, kPayers)).toBe(500);
    expect(calculateAccrual('F003', goal, participants, payments, x, kPayers)).toBe(0);
  });
  
  it('handles membership period filtering', () => {
    const fam1 = { memberFrom: new Date(2024, 0, 1), memberTo: new Date(2024, 5, 30) };
    const fam2 = { memberFrom: new Date(2024, 3, 1), memberTo: null }; // пришла в апреле
    const fam3 = { memberFrom: null, memberTo: null }; // всегда член
    
    // Цель: март-май 2024
    const goalStart = new Date(2024, 2, 1);
    const goalEnd = new Date(2024, 4, 31);
    
    // fam1: членство янв-июнь → пересекается с март-май ✓
    expect(isFamilyMemberInPeriod(fam1, goalStart, goalEnd)).toBe(true);
    // fam2: членство с апреля → пересекается с март-май ✓
    expect(isFamilyMemberInPeriod(fam2, goalStart, goalEnd)).toBe(true);
    // fam3: всегда член ✓
    expect(isFamilyMemberInPeriod(fam3, goalStart, goalEnd)).toBe(true);
    
    // Цель: июль-август 2024
    const goalStart2 = new Date(2024, 6, 1);
    const goalEnd2 = new Date(2024, 7, 31);
    
    // fam1: ушла в июне → не пересекается с июль-август ✗
    expect(isFamilyMemberInPeriod(fam1, goalStart2, goalEnd2)).toBe(false);
    // fam2: пришла в апреле, без конца → участвует ✓
    expect(isFamilyMemberInPeriod(fam2, goalStart2, goalEnd2)).toBe(true);
  });
});

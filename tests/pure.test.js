import { describe, it, expect } from 'vitest';
import { getIdFromLabelish, dynCap } from './pure.js';

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
});

describe('dynCap (water-filling)', () => {
  it('basic symmetry example', () => {
    const r1 = dynCap(500, [2000, 1333]);
    const r2 = dynCap(500, [1333, 2000]);
    expect(r1).toBeCloseTo(r2, 6);
  });
  it('example from README', () => {
    // T=9000, payments: [2000,2000,700,700,700,700,700] (sum=7500) => target=7500
    // expected x ~ 1250 (из README)
    const x = dynCap(9000, [2000,2000,700,700,700,700,700]);
    expect(x).toBeGreaterThan(1000);
    expect(x).toBeLessThan(1500);
  });
  it('zero and invalid inputs', () => {
    expect(dynCap(0, [1,2,3])).toBe(0);
    expect(dynCap(100, [])).toBe(0);
    expect(dynCap(100, [null, -5])).toBe(0);
  });
});

/**
 * @fileoverview Алгоритм Water-filling (DYN_CAP)
 */

/**
 * Custom function: вычисляет cap x для dynamic_by_payers
 * Σ min(P_i, x) = min(T, ΣP_i)
 * 
 * @param {number} T — целевая сумма
 * @param {Array} payments_range — диапазон платежей
 * @returns {number} — cap x
 * @customfunction
 */
function DYN_CAP(T, payments_range) {
  if (T === null || T === '' || isNaN(T)) return 0;
  
  const flat = flatten_(payments_range).map(Number).filter(v => isFinite(v) && v > 0);
  if (!flat.length) return 0;
  
  flat.sort((a, b) => a - b);
  const n = flat.length;
  const sum = flat.reduce((a, b) => a + b, 0);
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
  
  return round6_((target - (cumsum - flat[n - 1])) / 1);
}

/**
 * Внутренняя функция для вычисления cap x
 * @param {number} T — целевая сумма
 * @param {number[]} payments — массив платежей
 * @returns {number}
 */
function DYN_CAP_(T, payments) {
  if (!T || !isFinite(T)) return 0;
  
  const arr = (payments || []).map(Number).filter(v => v > 0 && isFinite(v));
  if (!arr.length) return 0;
  
  arr.sort((a, b) => a - b);
  const n = arr.length;
  const sum = arr.reduce((a, b) => a + b, 0);
  const target = Math.min(T, sum);
  
  if (target <= 0) return 0;
  
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
  
  const final = round6_((target - (cumsum - arr[n - 1])) / 1);
  Logger.log(`Final x=${final}`);
  return final;
}

/**
 * Тестовая функция для DYN_CAP
 * @returns {string}
 */
function TEST_DYN_CAP() {
  const testCases = [
    { T: 500, payments: [2000, 1333], expected: 250 },
    { T: 9000, payments: [2000, 2000, 700, 700, 700, 700, 700], expected: 1250 },
    { T: 6000, payments: [1000, 2000, 3000], expected: 2000 },
    { T: 10000, payments: [1000, 1000, 1000], expected: 1000 }, // сумма платежей < T
    { T: 3000, payments: [1000, 1000, 1000], expected: 1000 }, // точное совпадение
  ];
  
  let results = [];
  testCases.forEach(tc => {
    const result = DYN_CAP_(tc.T, tc.payments);
    const pass = Math.abs(result - tc.expected) < 0.01;
    results.push(`T=${tc.T}, payments=[${tc.payments}] => ${result} (expected ${tc.expected}) ${pass ? '✓' : '✗'}`);
  });
  
  Logger.log(results.join('\n'));
  return results.join('\n');
}

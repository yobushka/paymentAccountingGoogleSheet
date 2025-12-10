/**
 * @fileoverview Генерация и управление ID
 */

/**
 * Генерирует ID на всех листах
 * Точка входа из меню
 */
function generateAllIds() {
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  
  const plan = [
    { sheet: SHEET_NAMES.FAMILIES, idHeader: 'family_id', prefix: ID_PREFIXES.FAMILY, width: 3 },
    { sheet: SHEET_NAMES.PAYMENTS, idHeader: 'payment_id', prefix: ID_PREFIXES.PAYMENT, width: 3 }
  ];
  
  // Добавляем ID для целей или сборов в зависимости от версии
  if (version === 'v2' || version === 'new') {
    plan.push({ sheet: SHEET_NAMES.GOALS, idHeader: 'goal_id', prefix: ID_PREFIXES.GOAL, width: 3 });
  } else {
    plan.push({ sheet: SHEET_NAMES.COLLECTIONS, idHeader: 'collection_id', prefix: ID_PREFIXES.COLLECTION, width: 3 });
  }
  
  plan.forEach(p => {
    const sh = ss.getSheetByName(p.sheet);
    if (!sh) return;
    
    const map = getHeaderMap_(sh);
    const col = map[p.idHeader] || 1;
    fillMissingIds_(ss, p.sheet, col, p.prefix, p.width);
  });
  
  SpreadsheetApp.getActive().toast('IDs generated where empty.', 'Funds');
  
  // Обновляем формулы баланса для новых семей
  refreshBalanceFormulas_();
}

/**
 * Заполняет пустые ID в колонке
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 * @param {number} idCol — номер колонки (1-based)
 * @param {string} prefix — префикс ID
 * @param {number} padWidth — длина числовой части
 */
function fillMissingIds_(ss, sheetName, idCol, prefix, padWidth) {
  const sh = ss.getSheetByName(sheetName);
  const last = sh.getLastRow();
  if (last < 2) return;
  
  const rng = sh.getRange(2, idCol, last - 1, 1);
  const vals = rng.getValues().map(r => r[0]);
  
  // Находим максимальный номер существующих ID
  let maxNum = 0;
  vals.forEach(v => {
    if (typeof v === 'string' && v.startsWith(prefix)) {
      const n = parseInt(v.replace(prefix, ''), 10);
      if (!isNaN(n)) maxNum = Math.max(maxNum, n);
    }
  });
  
  // Заполняем пустые ячейки
  const out = vals.slice();
  for (let i = 0; i < out.length; i++) {
    if (!out[i]) {
      maxNum += 1;
      out[i] = prefix + String(maxNum).padStart(padWidth, '0');
    }
  }
  
  rng.setValues(out.map(v => [v]));
}

/**
 * Получает следующий доступный ID
 * @param {string} sheetName
 * @param {string} idHeader — название колонки ID
 * @param {string} prefix
 * @param {number} [padWidth=3]
 * @returns {string}
 */
function getNextId_(sheetName, idHeader, prefix, padWidth) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) return nextId(prefix, 1, padWidth || 3);
  
  const map = getHeaderMap_(sh);
  const col = map[idHeader];
  if (!col) return nextId(prefix, 1, padWidth || 3);
  
  const last = sh.getLastRow();
  if (last < 2) return nextId(prefix, 1, padWidth || 3);
  
  const vals = sh.getRange(2, col, last - 1, 1).getValues().map(r => r[0]);
  
  let maxNum = 0;
  vals.forEach(v => {
    if (typeof v === 'string' && v.startsWith(prefix)) {
      const n = parseInt(v.replace(prefix, ''), 10);
      if (!isNaN(n)) maxNum = Math.max(maxNum, n);
    }
  });
  
  return nextId(prefix, maxNum + 1, padWidth || 3);
}

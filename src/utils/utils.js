/**
 * @fileoverview Утилиты и вспомогательные функции
 */

/**
 * Извлекает ID из метки формата «Название (ID)» или возвращает строку как есть
 * @param {string|*} value — значение с меткой или ID
 * @returns {string} — извлечённый ID
 */
function getIdFromLabelish_(value) {
  const s = String(value || '').trim();
  if (!s) return '';
  const m = s.match(/\(([^)]+)\)\s*$/);
  return m ? m[1] : s;
}

/**
 * Custom function: LABEL_TO_ID("Имя (F001)") -> "F001"
 * @param {string} value — метка или ID
 * @returns {string}
 * @customfunction
 */
function LABEL_TO_ID(value) {
  return getIdFromLabelish_(value);
}

/**
 * Возвращает map заголовков листа: {headerName: columnIndex (1-based)}
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @returns {Object<string, number>}
 */
function getHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const map = {};
  headers.forEach((h, i) => { map[String(h || '').trim()] = i + 1; });
  return map;
}

/**
 * Преобразует номер колонки в буквенный индекс (1 -> A, 27 -> AA)
 * @param {number} col — номер колонки (1-based)
 * @returns {string}
 */
function colToLetter_(col) {
  let s = '';
  let c = col;
  while (c > 0) {
    const r = (c - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

/**
 * Разворачивает вложенные массивы в плоский
 * @param {Array} arr
 * @returns {Array}
 */
function flatten_(arr) {
  const out = [];
  (arr || []).forEach(r => Array.isArray(r) ? r.forEach(c => out.push(c)) : out.push(r));
  return out;
}

/**
 * Округление до 6 знаков после запятой
 * @param {number} x
 * @returns {number}
 */
function round6_(x) {
  return Math.round((x + Number.EPSILON) * 1e6) / 1e6;
}

/**
 * Округление до 2 знаков после запятой (для денег)
 * @param {number} x
 * @returns {number}
 */
function round2_(x) {
  return Math.round((x + Number.EPSILON) * 100) / 100;
}

/**
 * Показывает сообщение об ошибке в toast
 * @param {string} msg
 */
function toastErr_(msg) {
  SpreadsheetApp.getActive().toast(msg, 'Funds (error)', 5);
}

/**
 * Преобразует Date в ISO строку yyyy-mm-dd
 * @param {Date} d
 * @returns {string}
 */
function toISO_(d) {
  const pad = (n) => String(n).padStart(2, '0');
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

/**
 * Получает или создаёт лист по имени
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} name
 * @returns {GoogleAppsScript.Spreadsheet.Sheet}
 */
function getOrCreateSheet(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

/**
 * Создаёт или обновляет именованный диапазон
 * @param {string} name — имя диапазона
 * @param {string} a1 — A1-нотация (например, 'Lists!A2:A')
 */
function ensureNamedRange(name, a1) {
  const ss = SpreadsheetApp.getActive();
  const existing = ss.getNamedRanges().find(n => n.getName() === name);
  const rng = ss.getRange(a1);
  if (existing) existing.setRange(rng);
  else ss.setNamedRange(name, rng);
}

/**
 * Читает колонку из листа как массив строк
 * @param {string} sheetName
 * @param {string} colLetter
 * @param {number} startRow
 * @returns {string[]}
 */
function getLabelColumn_(sheetName, colLetter, startRow) {
  const sh = SpreadsheetApp.getActive().getSheetByName(sheetName);
  const last = sh.getLastRow();
  if (last < startRow) return [];
  const rng = sh.getRange(`${colLetter}${startRow}:${colLetter}${last}`);
  return rng.getValues().map(r => String(r[0] || '')).filter(Boolean);
}

/**
 * Генерирует следующий ID с префиксом и паддингом
 * @param {string} prefix — например, 'F', 'G', 'PMT'
 * @param {number} index — числовой индекс
 * @param {number} [pad=3] — длина числовой части
 * @returns {string}
 */
function nextId(prefix, index, pad) {
  const width = pad || 3;
  let n = String(index);
  while (n.length < width) n = '0' + n;
  return prefix + n;
}

/**
 * Обёртка try-catch с логированием
 * @param {Function} fn
 * @param {string} context — описание операции
 * @returns {*}
 */
function withTry(fn, context) {
  try {
    return fn();
  } catch (e) {
    const msg = '[PaymentSheet] ' + (context || 'op') + ' failed: ' + e.message;
    Logger.log(msg);
    throw new Error(msg);
  }
}

/**
 * Определяет версию текущей таблицы
 * @returns {'v1'|'v2'|'new'} — версия или 'new' для новой таблицы
 */
function detectVersion() {
  const ss = SpreadsheetApp.getActive();
  
  // Проверяем наличие листа «Цели» (v2.0)
  if (ss.getSheetByName(SHEET_NAMES.GOALS)) {
    return 'v2';
  }
  
  // Проверяем наличие листа «Сборы» (v1.x)
  if (ss.getSheetByName(SHEET_NAMES.COLLECTIONS)) {
    return 'v1';
  }
  
  // Новая таблица
  return 'new';
}

/**
 * Проверяет, нужна ли миграция с v1 на v2
 * @returns {boolean}
 */
function needsMigration() {
  return detectVersion() === 'v1';
}

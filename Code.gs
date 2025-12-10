/**
 * @fileoverview Payment Accounting for Google Sheets v2.0
 * 
 * Автоматически сгенерировано из модулей: 2025-12-10T22:32:23.667Z
 * 
 * НЕ РЕДАКТИРУЙТЕ ЭТОТ ФАЙЛ НАПРЯМУЮ!
 * Вносите изменения в модули в папке src/ и запускайте build.js
 * 
 * Структура модулей:
 *   src/config/     — константы и спецификации листов
 *   src/utils/      — утилитарные функции
 *   src/calculations/ — расчётные функции
 *   src/sheets/     — настройка листов
 *   src/core/       — основная логика
 *   src/ui/         — меню, стили, диалоги
 *   src/triggers/   — обработчики событий
 *   src/migration/  — миграция v1 → v2
 */


// ======================================================================
// MODULE: src/config/constants.js
// ======================================================================

// Версия приложения
const APP_VERSION = '2.0';

// Режимы начисления
const ACCRUAL_MODES = {
  STATIC_PER_FAMILY: 'static_per_family',
  SHARED_TOTAL_ALL: 'shared_total_all',
  SHARED_TOTAL_BY_PAYERS: 'shared_total_by_payers',
  DYNAMIC_BY_PAYERS: 'dynamic_by_payers',
  PROPORTIONAL_BY_PAYERS: 'proportional_by_payers',
  UNIT_PRICE: 'unit_price',
  VOLUNTARY: 'voluntary'
};

// Алиасы для обратной совместимости v1.x
const ACCRUAL_ALIASES = {
  'static_per_child': ACCRUAL_MODES.STATIC_PER_FAMILY,
  'unit_price_by_payers': ACCRUAL_MODES.UNIT_PRICE
};

// Статусы целей
const GOAL_STATUS = {
  OPEN: 'Открыта',
  CLOSED: 'Закрыта',
  CANCELLED: 'Отменена'
};

// Статусы v1.x (для миграции)
const COLLECTION_STATUS_V1 = {
  OPEN: 'Открыт',
  CLOSED: 'Закрыт'
};

// Типы целей
const GOAL_TYPES = {
  ONE_TIME: 'разовая',
  REGULAR: 'регулярная'
};

// Периодичность регулярных целей
const GOAL_PERIODICITY = {
  MONTHLY: 'ежемесячно',
  QUARTERLY: 'ежеквартально',
  YEARLY: 'ежегодно'
};

// Статусы участия
const PARTICIPATION_STATUS = {
  PARTICIPATES: 'Участвует',
  NOT_PARTICIPATES: 'Не участвует'
};

// Способы оплаты
const PAYMENT_METHODS = ['СБП', 'карта', 'наличные'];

// Активность семьи
const ACTIVE_STATUS = {
  YES: 'Да',
  NO: 'Нет'
};

// Префиксы ID
const ID_PREFIXES = {
  FAMILY: 'F',
  GOAL: 'G',
  COLLECTION: 'C', // v1.x legacy
  PAYMENT: 'PMT'
};

// Названия листов
const SHEET_NAMES = {
  INSTRUCTION: 'Инструкция',
  FAMILIES: 'Семьи',
  GOALS: 'Цели',
  COLLECTIONS: 'Сборы', // v1.x legacy
  PARTICIPATION: 'Участие',
  PAYMENTS: 'Платежи',
  BALANCE: 'Баланс',
  DETAIL: 'Детализация',
  SUMMARY: 'Сводка',
  ISSUES: 'Выдача',
  ISSUE_STATUS: 'Статус выдачи',
  LISTS: 'Lists'
};

// Named ranges
const NAMED_RANGES = {
  FAMILIES_LABELS: 'FAMILIES_LABELS',
  ACTIVE_FAMILIES_LABELS: 'ACTIVE_FAMILIES_LABELS',
  GOALS_LABELS: 'GOALS_LABELS',
  OPEN_GOALS_LABELS: 'OPEN_GOALS_LABELS',
  // v1.x legacy
  COLLECTIONS_LABELS: 'COLLECTIONS_LABELS',
  OPEN_COLLECTIONS_LABELS: 'OPEN_COLLECTIONS_LABELS'
};

// ======================================================================
// MODULE: src/config/sheets-spec.js
// ======================================================================

/**
 * Возвращает спецификацию всех листов приложения
 * @returns {Array<{name: string, headers: string[], colWidths: number[], dateCols?: number[]}>}
 */
function getSheetsSpec() {
  return [
    {
      name: SHEET_NAMES.INSTRUCTION,
      headers: ['Шаг', 'Описание'],
      colWidths: [80, 1000]
    },
    {
      name: SHEET_NAMES.FAMILIES,
      headers: [
        'Ребёнок ФИО', 'День рождения',
        'Мама ФИО', 'Мама телефон', 'Мама реквизиты', 'Мама телеграм',
        'Папа ФИО', 'Папа телефон', 'Папа реквизиты', 'Папа телеграм',
        'Активен', 'Членство с', 'Членство по', 'Комментарий',
        'family_id'
      ],
      colWidths: [220, 110, 220, 140, 240, 160, 220, 140, 240, 160, 90, 110, 110, 260, 110],
      dateCols: [2, 12, 13]  // День рождения, Членство с, Членство по
    },
    {
      // v2.0: Цели вместо Сборы
      name: SHEET_NAMES.GOALS,
      headers: [
        'Название цели', 'Тип', 'Статус',
        'Дата начала', 'Дедлайн',
        'Начисление', 'Параметр суммы', 'Фиксированный x', 'К выдаче детям',
        'Периодичность', 'Родительская цель',
        'Закупка из средств', 'Возмещено',
        'Комментарий',
        'goal_id', 'Ссылка на гуглдиск'
      ],
      colWidths: [260, 120, 120, 110, 110, 220, 150, 140, 150, 140, 140, 200, 110, 260, 120, 300],
      dateCols: [4, 5]
    },
    {
      // v1.x legacy: Сборы (для миграции)
      name: SHEET_NAMES.COLLECTIONS,
      headers: [
        'Название сбора', 'Статус',
        'Дата начала', 'Дедлайн',
        'Начисление', 'Параметр суммы', 'Фиксированный x', 'К выдаче детям',
        'Закупка из средств', 'Возмещено',
        'Комментарий',
        'collection_id', 'Ссылка на гуглдиск'
      ],
      colWidths: [260, 120, 110, 110, 220, 150, 140, 150, 200, 110, 260, 120, 300],
      dateCols: [3, 4]
    },
    {
      name: SHEET_NAMES.ISSUES,
      headers: [
        'Дата выдачи', 'goal_id (label)', 'family_id (label)', 'Единиц', 'Кто выдал', 'Выдано', 'Комментарий'
      ],
      colWidths: [110, 260, 260, 90, 160, 110, 260],
      dateCols: [1]
    },
    {
      name: SHEET_NAMES.ISSUE_STATUS,
      headers: [
        'goal_id', 'Название', 'Статус', 'x (цена)', 'Единиц требуется', 'Единиц оплачено', 'Единиц выдано', 'Остаток (шт)'
      ],
      colWidths: [120, 260, 110, 110, 140, 140, 140, 130]
    },
    {
      name: SHEET_NAMES.PARTICIPATION,
      headers: ['goal_id (label)', 'family_id (label)', 'Статус', 'Комментарий'],
      colWidths: [260, 260, 120, 260]
    },
    {
      name: SHEET_NAMES.PAYMENTS,
      headers: [
        'Дата', 'family_id (label)', 'goal_id (label)',
        'Сумма', 'Способ', 'Комментарий', 'payment_id'
      ],
      colWidths: [110, 260, 260, 110, 110, 260, 120],
      dateCols: [1]
    },
    {
      // v2.0: расширенный баланс
      name: SHEET_NAMES.BALANCE,
      headers: [
        'family_id', 'Имя ребёнка',
        'Внесено всего', 'Списано всего', 'Зарезервировано',
        'Свободный остаток', 'Задолженность'
      ],
      colWidths: [120, 260, 140, 140, 150, 160, 130]
    },
    {
      name: SHEET_NAMES.DETAIL,
      headers: [
        'family_id', 'Имя ребёнка', 'goal_id', 'Название цели',
        'Оплачено', 'Начислено', 'Разность (±)', 'Режим'
      ],
      colWidths: [120, 200, 120, 200, 120, 120, 120, 150]
    },
    {
      name: SHEET_NAMES.SUMMARY,
      headers: [
        'goal_id', 'Название цели', 'Режим', 'Сумма цели', 'Собрано', 'Участников', 'Плательщиков', 'Единиц оплачено', 'Ещё плательщиков до закрытия', 'Остаток до цели', 'Переплата'
      ],
      colWidths: [120, 260, 180, 140, 140, 120, 150, 150, 220, 150, 130]
    },
    {
      name: SHEET_NAMES.LISTS, // скрытый служебный лист
      headers: [
        'OPEN_GOALS', '',
        'GOALS', '',
        'ACTIVE_FAMILIES', '',
        'FAMILIES', ''
      ],
      colWidths: [260, 40, 260, 40, 260, 40, 260, 40]
    }
  ];
}

/**
 * Возвращает спецификацию для v1.x (legacy) — для миграции
 * @returns {Array<{name: string, headers: string[], colWidths: number[], dateCols?: number[]}>}
 */
function getSheetsSpecV1() {
  return [
    {
      name: 'Инструкция',
      headers: ['Шаг', 'Описание'],
      colWidths: [80, 1000]
    },
    {
      name: 'Семьи',
      headers: [
        'Ребёнок ФИО', 'День рождения',
        'Мама ФИО', 'Мама телефон', 'Мама реквизиты', 'Мама телеграм',
        'Папа ФИО', 'Папа телефон', 'Папа реквизиты', 'Папа телеграм',
        'Активен', 'Комментарий',
        'family_id'
      ],
      colWidths: [220, 110, 220, 140, 240, 160, 220, 140, 240, 160, 90, 260, 110],
      dateCols: [2]
    },
    {
      name: 'Сборы',
      headers: [
        'Название сбора', 'Статус',
        'Дата начала', 'Дедлайн',
        'Начисление', 'Параметр суммы', 'Фиксированный x', 'К выдаче детям',
        'Закупка из средств', 'Возмещено',
        'Комментарий',
        'collection_id', 'Ссылка на гуглдиск'
      ],
      colWidths: [260, 120, 110, 110, 220, 150, 140, 150, 200, 110, 260, 120, 300],
      dateCols: [3, 4]
    },
    {
      name: 'Выдача',
      headers: [
        'Дата выдачи', 'collection_id (label)', 'family_id (label)', 'Единиц', 'Кто выдал', 'Выдано', 'Комментарий'
      ],
      colWidths: [110, 260, 260, 90, 160, 110, 260],
      dateCols: [1]
    },
    {
      name: 'Статус выдачи',
      headers: [
        'collection_id', 'Название', 'Статус', 'x (цена)', 'Единиц требуется', 'Единиц оплачено', 'Единиц выдано', 'Остаток (шт)'
      ],
      colWidths: [120, 260, 110, 110, 140, 140, 140, 130]
    },
    {
      name: 'Участие',
      headers: ['collection_id (label)', 'family_id (label)', 'Статус', 'Комментарий'],
      colWidths: [260, 260, 120, 260]
    },
    {
      name: 'Платежи',
      headers: [
        'Дата', 'family_id (label)', 'collection_id (label)',
        'Сумма', 'Способ', 'Комментарий', 'payment_id'
      ],
      colWidths: [110, 260, 260, 110, 110, 260, 120],
      dateCols: [1]
    },
    {
      name: 'Баланс',
      headers: [
        'family_id', 'Имя ребёнка',
        'Переплата (текущая)', 'Оплачено всего', 'Начислено всего', 'Задолженность'
      ],
      colWidths: [120, 260, 140, 140, 120, 130]
    },
    {
      name: 'Детализация',
      headers: [
        'family_id', 'Имя ребёнка', 'collection_id', 'Название сбора',
        'Оплачено', 'Начислено', 'Разность (±)', 'Режим'
      ],
      colWidths: [120, 200, 120, 200, 120, 120, 120, 150]
    },
    {
      name: 'Сводка',
      headers: [
        'collection_id', 'Название сбора', 'Режим', 'Сумма цели', 'Собрано', 'Участников', 'Плательщиков', 'Единиц оплачено', 'Ещё плательщиков до закрытия', 'Остаток до цели'
      ],
      colWidths: [120, 260, 180, 140, 140, 120, 150, 150, 220, 150]
    },
    {
      name: 'Lists',
      headers: [
        'OPEN_COLLECTIONS', '',
        'COLLECTIONS', '',
        'ACTIVE_FAMILIES', '',
        'FAMILIES', ''
      ],
      colWidths: [260, 40, 260, 40, 260, 40, 260, 40]
    }
  ];
}

// ======================================================================
// MODULE: src/utils/utils.js
// ======================================================================

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

// ======================================================================
// MODULE: src/calculations/dyn-cap.js
// ======================================================================

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

// ======================================================================
// MODULE: src/calculations/custom-functions.js
// ======================================================================

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

// ======================================================================
// MODULE: src/calculations/recalculate.js
// ======================================================================

/**
 * Пересчитывает все данные
 * Точка входа из меню
 */
function recalculateAll() {
  try {
    refreshBalanceFormulas_();
    
    // Обновляем тикер детализации
    const ss = SpreadsheetApp.getActive();
    const shDetail = ss.getSheetByName(SHEET_NAMES.DETAIL);
    if (shDetail) {
      shDetail.getRange('K2').setValue(new Date().toISOString());
    }
    refreshDetailSheet_();
    
    // Обновляем тикер сводки
    const shSummary = ss.getSheetByName(SHEET_NAMES.SUMMARY);
    if (shSummary) {
      shSummary.getRange('K2').setValue(new Date().toISOString());
    }
    refreshSummarySheet_();
    
    // Обновляем статус выдачи
    refreshIssueStatusSheet_();
    
    SpreadsheetApp.getActive().toast('Balance, Detail and Summary recalculated.', 'Funds');
    SpreadsheetApp.getUi().alert(
      'Пересчёт завершён',
      'Обновлены: «Баланс», «Детализация», «Сводка».',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {
    toastErr_('Recalculate failed: ' + e.message);
    SpreadsheetApp.getUi().alert(
      'Ошибка пересчёта',
      String(e && e.message ? e.message : e),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ======================================================================
// MODULE: src/sheets/lists.js
// ======================================================================

/**
 * Настраивает скрытый лист Lists с формулами для меток
 */
function setupListsSheet() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAMES.LISTS);
  if (!sh) return;
  
  const version = detectVersion();
  
  if (version === 'v2' || version === 'new') {
    setupListsSheetV2_(sh, ss);
  } else {
    setupListsSheetV1_(sh, ss);
  }
  
  sh.hideSheet();
}

/**
 * Настраивает Lists для v2.0 (Цели)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupListsSheetV2_(sh, ss) {
  const shG = ss.getSheetByName(SHEET_NAMES.GOALS);
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  
  if (!shG || !shF) return;
  
  const mapG = getHeaderMap_(shG);
  const mapF = getHeaderMap_(shF);
  
  const gNameCol = colToLetter_(mapG['Название цели'] || 1);
  const gIdCol = colToLetter_(mapG['goal_id'] || 1);
  const gStatusCol = colToLetter_(mapG['Статус'] || 3);
  
  const fNameCol = colToLetter_(mapF['Ребёнок ФИО'] || 1);
  const fIdCol = colToLetter_(mapF['family_id'] || 1);
  const fActiveCol = colToLetter_(mapF['Активен'] || 11);
  
  // A: OPEN_GOALS — открытые цели (метки)
  sh.getRange('A1').setValue('OPEN_GOALS');
  sh.getRange('A2').setFormula(
    `=IFERROR(FILTER(ARRAYFORMULA(Цели!${gNameCol}2:${gNameCol} & " (" & Цели!${gIdCol}2:${gIdCol} & ")"), Цели!${gStatusCol}2:${gStatusCol}="${GOAL_STATUS.OPEN}"),)`
  );
  
  // B: GOALS — все цели (метки)
  sh.getRange('B1').setValue('GOALS');
  sh.getRange('B2').setFormula(
    `=IFERROR(FILTER(ARRAYFORMULA(Цели!${gNameCol}2:${gNameCol} & " (" & Цели!${gIdCol}2:${gIdCol} & ")"), LEN(Цели!${gIdCol}2:${gIdCol})),)`
  );
  
  // C: ACTIVE_FAMILIES — активные семьи (метки)
  sh.getRange('C1').setValue('ACTIVE_FAMILIES');
  sh.getRange('C2').setFormula(
    `=IFERROR(FILTER(ARRAYFORMULA(Семьи!${fNameCol}2:${fNameCol} & " (" & Семьи!${fIdCol}2:${fIdCol} & ")"), Семьи!${fActiveCol}2:${fActiveCol}="${ACTIVE_STATUS.YES}"),)`
  );
  
  // D: FAMILIES — все семьи (метки)
  sh.getRange('D1').setValue('FAMILIES');
  sh.getRange('D2').setFormula(
    `=IFERROR(FILTER(ARRAYFORMULA(Семьи!${fNameCol}2:${fNameCol} & " (" & Семьи!${fIdCol}2:${fIdCol} & ")"), LEN(Семьи!${fIdCol}2:${fIdCol})),)`
  );
}

/**
 * Настраивает Lists для v1.x (Сборы)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function setupListsSheetV1_(sh, ss) {
  const shC = ss.getSheetByName(SHEET_NAMES.COLLECTIONS);
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  
  if (!shC || !shF) return;
  
  const mapC = getHeaderMap_(shC);
  const mapF = getHeaderMap_(shF);
  
  const cNameCol = colToLetter_(mapC['Название сбора'] || 1);
  const cIdCol = colToLetter_(mapC['collection_id'] || 1);
  const cStatusCol = colToLetter_(mapC['Статус'] || 2);
  
  const fNameCol = colToLetter_(mapF['Ребёнок ФИО'] || 1);
  const fIdCol = colToLetter_(mapF['family_id'] || 1);
  const fActiveCol = colToLetter_(mapF['Активен'] || 11);
  
  // A: OPEN_COLLECTIONS — открытые сборы (метки)
  sh.getRange('A1').setValue('OPEN_COLLECTIONS');
  sh.getRange('A2').setFormula(
    `=IFERROR(FILTER(ARRAYFORMULA(Сборы!${cNameCol}2:${cNameCol} & " (" & Сборы!${cIdCol}2:${cIdCol} & ")"), Сборы!${cStatusCol}2:${cStatusCol}="${COLLECTION_STATUS_V1.OPEN}"),)`
  );
  
  // B: COLLECTIONS — все сборы (метки)
  sh.getRange('B1').setValue('COLLECTIONS');
  sh.getRange('B2').setFormula(
    `=IFERROR(FILTER(ARRAYFORMULA(Сборы!${cNameCol}2:${cNameCol} & " (" & Сборы!${cIdCol}2:${cIdCol} & ")"), LEN(Сборы!${cIdCol}2:${cIdCol})),)`
  );
  
  // C: ACTIVE_FAMILIES — активные семьи (метки)
  sh.getRange('C1').setValue('ACTIVE_FAMILIES');
  sh.getRange('C2').setFormula(
    `=IFERROR(FILTER(ARRAYFORMULA(Семьи!${fNameCol}2:${fNameCol} & " (" & Семьи!${fIdCol}2:${fIdCol} & ")"), Семьи!${fActiveCol}2:${fActiveCol}="Да"),)`
  );
  
  // D: FAMILIES — все семьи (метки)
  sh.getRange('D1').setValue('FAMILIES');
  sh.getRange('D2').setFormula(
    `=IFERROR(FILTER(ARRAYFORMULA(Семьи!${fNameCol}2:${fNameCol} & " (" & Семьи!${fIdCol}2:${fIdCol} & ")"), LEN(Семьи!${fIdCol}2:${fIdCol})),)`
  );
}

// ======================================================================
// MODULE: src/sheets/instruction.js
// ======================================================================

/**
 * Настраивает лист «Инструкция» для v2.0
 */
function setupInstructionSheet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.INSTRUCTION);
  if (!sh) return;
  
  // Очищаем старое содержимое под заголовком
  const last = sh.getLastRow();
  if (last > 1) {
    sh.getRange(2, 1, last - 1, Math.max(2, sh.getLastColumn())).clearContent();
  }
  
  const rows = [
    ['▶ О проекте', `Версия: ${APP_VERSION}. Репозиторий: https://github.com/yobushka/paymentAccountingGoogleSheet`],
    ['▶ Дисклеймер', 'Инструмент на ранней стадии и используется для личных целей; welcome to contribute. Внимание к персональным данным: передача ПДн через границу может быть незаконной. Google — иностранная компания; соблюдайте применимое законодательство.'],
    
    ['▶ Что нового в v2.0', '• Балансовая модель: платежи пополняют баланс, цели списывают с баланса\n• Свободные платежи: goal_id опционален — можно вносить без привязки к цели\n• Регулярные цели: ежемесячные/ежеквартальные с автосозданием\n• Новый режим «voluntary» — добровольные взносы\n• Расширенный баланс: Внесено - Списано - Резерв = Свободно'],
    
    ['▶ Быстрый старт', '1) Funds → Setup / Rebuild structure.\n2) Заполните «Семьи» (Активен=Да).\n3) Добавьте «Цели» (Статус=Открыта).\n4) При необходимости настройте «Участие».\n5) Вносите «Платежи».\n6) Смотрите «Баланс» и «Детализация».\n7) «Сводка» — оперативные итоги по целям.'],
    
    ['1', 'Семьи: одна строка = одна семья (один ребёнок). Заполните ФИО, День рождения (yyyy-mm-dd), контакты родителей.\n• «Активен=Да» — семья участвует по умолчанию.\n• «Членство с» — дата прихода в класс (опционально).\n• «Членство по» — дата ухода из класса (опционально).\nID генерируется автоматически.'],
    
    ['2', 'Цели: выберите «Тип» (разовая/регулярная), «Начисление» и задайте «Параметр суммы».\n• Для регулярных целей укажите «Периодичность».\n• «Фиксированный x»: для dynamic_by_payers — cap после закрытия, для unit_price — цена за единицу.\n• «К выдаче детям» — включает учёт выдачи для поштучных целей.'],
    
    ['3', 'Участие (опционально): если есть хотя бы один «Участвует», участвуют только отмеченные семьи. «Не участвует» всегда исключает семью. Если явных «Участвует» нет — участвуют все активные семьи.'],
    
    ['4', 'Платежи: выберите семью из выпадающего списка. Цель — опционально (свободный платёж пополняет баланс без привязки). Сумма должна быть > 0.'],
    
    ['5', 'Баланс v2.0:\n• Внесено всего — сумма всех платежей\n• Списано всего — сумма начислений по целям\n• Зарезервировано — под открытые цели\n• Свободный остаток — доступно для новых целей\n• Задолженность — если списано > внесено'],
    
    ['6', 'Демо-данные: Funds → Load Sample Data — добавит примеры для демонстрации.'],
    
    ['▶ Пересчёт', 'Если сменили режим/участие/платежи, выполните Funds → Recalculate (Balance & Detail). Баланс также авто‑обновляется при правках.'],
    
    ['▶ Режимы начисления (подробно)', 'Все расчёты моментальные:'],
    
    ['static_per_family', 'Фикс на семью. Начислено участнику = «Параметр суммы».\nПример: сбор на канцтовары 500₽ — каждой семье начисляется 500₽.'],
    
    ['shared_total_all', 'T/N на всех участников.\nПример: новогодний утренник 12000₽ на 10 участников = 1200₽ каждому.'],
    
    ['shared_total_by_payers', 'T/K только для оплативших.\nПример: подарок учителю 5000₽, оплатили 5 семей = 1000₽ каждому плательщику.'],
    
    ['dynamic_by_payers', 'Water-filling: Σ min(P_i, x) = min(T, ΣP_i).\nНачислено семье i = min(P_i, x). Выравнивает ранние переплаты.\nПосле закрытия используется «Фиксированный x».'],
    
    ['proportional_by_payers', 'Пропорционально платежам: начисление i = P_i × (T/ΣP).\nПока не достигнута цель — списывается весь платёж.'],
    
    ['unit_price', 'Поштучная закупка: цена за единицу x из «Фиксированный x».\nНачисление i = floor(P_i/x) × x. Остаток < x — переплата без долга.'],
    
    ['voluntary', 'Добровольный взнос (v2.0): начисление = платёж, без обязательств.\nДля благотворительных сборов без фиксированной цели.'],
    
    ['▶ Закрытие цели', 'Меню Funds → Close Goal. Введите goal_id (например, G003). Для dynamic_by_payers скрипт посчитает x и зафиксирует.'],
    
    ['▶ Формулы', 'DYN_CAP(T, payments_range) — cap x по water-filling.\nACCRUED_FAMILY(family_id, filter) — начислено семье.\nPAYED_TOTAL_FAMILY(family_id) — оплачено семьёй.\nLABEL_TO_ID("Имя (F001)") → F001.'],
    
    ['▶ Членство и период участия', '«Членство с» и «Членство по» позволяют автоматически исключать семью из целей вне периода её членства.\n\n• Семья пришла в середине года: укажите «Членство с» — взносы до этой даты не начисляются.\n• Семья ушла: укажите «Членство по» — расчёт баланса на дату ухода для выплаты остатка.\n\nФормулы:\n• EXIT_BALANCE(family_id) — баланс на дату ухода\n• MEMBERSHIP_MONTHS(family_id, год) — месяцев членства в году\n• MEMBERSHIP_RATIO(family_id, год) — коэффициент 0-1 для пропорций\n• IS_MEMBER_ON_DATE(family_id, дата) — проверка членства'],
    
    ['▶ Миграция v1→v2', 'Если у вас таблица v1.x (Сборы/collection_id), запустите Funds → Migrate v1 → v2 для автоматической конвертации.'],
    
    ['▶ Советы', '• Если дропдауны пустые — Funds → Rebuild data validations.\n• Если «Начислено» неожиданно 0 — проверьте «Участие» и «Активен».\n• Для чистки — Funds → Cleanup visuals.']
  ];
  
  sh.getRange(2, 1, rows.length, 2).setValues(rows);
  
  // Форматирование
  sh.getRange(2, 2, rows.length, 1).setWrap(true).setVerticalAlignment('top');
  
  // Выделяем секции жирным
  rows.forEach((r, i) => {
    if (String(r[0] || '').slice(0, 1) === '▶') {
      sh.getRange(2 + i, 1, 1, 2).setFontWeight('bold');
    }
  });
}

/**
 * Настраивает лист «Инструкция» для v1.x (режим совместимости)
 */
function setupInstructionSheetV1_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAMES.INSTRUCTION);
  if (!sh) return;
  
  const last = sh.getLastRow();
  if (last > 1) {
    sh.getRange(2, 1, last - 1, Math.max(2, sh.getLastColumn())).clearContent();
  }
  
  const rows = [
    ['▶ О проекте', 'Версия: 0.2 (v1.x совместимость). Репозиторий: https://github.com/yobushka/paymentAccountingGoogleSheet'],
    ['▶ Дисклеймер', 'Инструмент на ранней стадии. Внимание к персональным данным.'],
    ['▶ Быстрый старт', '1) Funds → Setup / Rebuild structure.\n2) Заполните «Семьи» (Активен=Да).\n3) Добавьте «Сборы» (Статус=Открыт).\n4) При необходимости настройте «Участие».\n5) Вносите «Платежи».\n6) Смотрите «Баланс» и «Детализация».'],
    ['1', 'Семьи: одна строка = одна семья. ID генерируется автоматически.'],
    ['2', 'Сборы: выберите «Начисление» и задайте «Параметр суммы».'],
    ['3', 'Участие: если есть «Участвует» — участвуют только отмеченные. «Не участвует» исключает.'],
    ['4', 'Платежи: выберите семью и сбор. Сумма > 0.'],
    ['5', 'Баланс: Оплачено - Начислено = Переплата/Задолженность.'],
    ['▶ Обновление', 'Доступна версия 2.0 с балансовой моделью. Funds → Migrate v1 → v2.']
  ];
  
  sh.getRange(2, 1, rows.length, 2).setValues(rows);
  sh.getRange(2, 2, rows.length, 1).setWrap(true).setVerticalAlignment('top');
  
  rows.forEach((r, i) => {
    if (String(r[0] || '').slice(0, 1) === '▶') {
      sh.getRange(2 + i, 1, 1, 2).setFontWeight('bold');
    }
  });
}

// ======================================================================
// MODULE: src/sheets/balance.js
// ======================================================================

/**
 * Настраивает лист «Баланс» с примерами формул
 */
function setupBalanceExamples() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAMES.BALANCE);
  if (!sh) return;
  
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  if (!shF) return;
  
  const mapF = getHeaderMap_(shF);
  const idCol = colToLetter_(mapF['family_id'] || 1);
  const nameCol = colToLetter_(mapF['Ребёнок ФИО'] || 1);
  const famLastRow = shF.getLastRow();
  
  // A2: список family_id из «Семьи»
  if (famLastRow > 1) {
    sh.getRange('A2').setFormula(
      `=ARRAYFORMULA(IFERROR(FILTER(Семьи!${idCol}2:${idCol}${famLastRow}, LEN(Семьи!${idCol}2:${idCol}${famLastRow})), ))`
    );
    
    // B2: имена по ID
    sh.getRange('B2').setFormula(
      `=ARRAYFORMULA(IF(LEN(A2:A)=0, "", IFERROR(VLOOKUP(A2:A, {Семьи!${idCol}2:${idCol}${famLastRow}, Семьи!${nameCol}2:${nameCol}${famLastRow}}, 2, FALSE), "")))`
    );
  }
  
  // Селектор фильтра: OPEN | ALL
  sh.getRange('I1').setValue('Фильтр начисления');
  sh.getRange('J1').setValue('ALL');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['OPEN', 'ALL'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('J1').setDataValidation(rule).setHorizontalAlignment('center');
  sh.getRange('J1').setNote('OPEN (только открытые) или ALL (все цели).');
  
  // Обновляем формулы для баланса
  refreshBalanceFormulas_();
  
  sh.getRange('I3').setValue('Примечание: даты платежей используются только для справки. Расчёты мгновенные.');
  
  // Настраиваем связанные листы
  setupDetailSheet_();
  setupSummarySheet_();
}

/**
 * Обновляет формулы на листе «Баланс» для текущего количества семей
 */
function refreshBalanceFormulas_() {
  const ss = SpreadsheetApp.getActive();
  const shBal = ss.getSheetByName(SHEET_NAMES.BALANCE);
  const shFam = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  if (!shBal || !shFam) return;
  
  const last = shFam.getLastRow();
  const famCount = Math.max(0, last - 1);
  
  // Пересоздаём A2/B2 формулы
  if (last > 1) {
    const mapF = getHeaderMap_(shFam);
    const idColLetter = colToLetter_(mapF['family_id'] || 1);
    const nameColLetter = colToLetter_(mapF['Ребёнок ФИО'] || 2);
    
    shBal.getRange('A2').setFormula(
      `=ARRAYFORMULA(IFERROR(FILTER(Семьи!${idColLetter}2:${idColLetter}${last}, LEN(Семьи!${idColLetter}2:${idColLetter}${last})), ))`
    );
    shBal.getRange('B2').setFormula(
      `=ARRAYFORMULA(IF(LEN(A2:A)=0, "", IFERROR(VLOOKUP(A2:A, {Семьи!${idColLetter}2:${idColLetter}${last}, Семьи!${nameColLetter}2:${nameColLetter}${last}}, 2, FALSE), "")))`
    );
  }
  
  // Очищаем старые формулы
  const currentLastRow = shBal.getLastRow();
  if (currentLastRow > 1) {
    shBal.getRange(2, 3, currentLastRow - 1, 5).clearContent();
  }
  
  if (famCount === 0) return;
  
  const version = detectVersion();
  const rows = famCount;
  
  // Формулы для v2.0: Внесено, Списано, Резерв, Свободно, Долг
  const formulasC = []; // Внесено всего
  const formulasD = []; // Списано всего
  const formulasE = []; // Зарезервировано
  const formulasF = []; // Свободный остаток
  const formulasG = []; // Задолженность
  
  for (let i = 0; i < rows; i++) {
    const r = 2 + i;
    
    // C: Внесено всего (все платежи)
    formulasC.push([`=IFERROR(PAYED_TOTAL_FAMILY($A${r}), 0)`]);
    
    // D: Списано всего (начислено по целям)
    formulasD.push([`=IFERROR(ACCRUED_FAMILY($A${r}, IF(LEN($J$1)=0, "ALL", $J$1)), 0)`]);
    
    // E: Зарезервировано (только открытые цели)
    formulasE.push([`=IFERROR(ACCRUED_FAMILY($A${r}, "OPEN"), 0)`]);
    
    // F: Свободный остаток = Внесено - Списано - Резерв (если > 0)
    // Упрощённо: MAX(0, Внесено - Списано)
    formulasF.push([`=MAX(0, C${r} - D${r})`]);
    
    // G: Задолженность = MAX(0, Списано - Внесено)
    formulasG.push([`=MAX(0, D${r} - C${r})`]);
  }
  
  shBal.getRange(2, 3, rows, 1).setFormulas(formulasC);
  shBal.getRange(2, 4, rows, 1).setFormulas(formulasD);
  shBal.getRange(2, 5, rows, 1).setFormulas(formulasE);
  shBal.getRange(2, 6, rows, 1).setFormulas(formulasF);
  shBal.getRange(2, 7, rows, 1).setFormulas(formulasG);
  
  // Применяем стили
  SpreadsheetApp.flush();
  try {
    styleSheetHeader_(shBal);
    styleBalanceSheet_(shBal);
  } catch (_) {}
}

// ======================================================================
// MODULE: src/sheets/detail.js
// ======================================================================

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
  
  // Селектор фильтра
  sh.getRange('J1').setValue('Фильтр');
  sh.getRange('K1').setValue('ALL');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['OPEN', 'ALL'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('K1').setDataValidation(rule).setHorizontalAlignment('center');
  sh.getRange('K1').setNote('OPEN (только открытые) или ALL (все цели)');
  
  // Тикер для принудительного пересчёта
  sh.getRange('J2').setValue('Tick');
  sh.getRange('K2').setValue(new Date().toISOString());
  sh.getRange('J3').setValue('Детализация платежей и начислений. Автообновляется.');
  
  // Динамическая формула
  sh.getRange('A2').setFormula(`=GENERATE_DETAIL_BREAKDOWN(IF(LEN($K$1)=0,"ALL",$K$1), $K$2)`);
}

/**
 * Обновляет лист «Детализация»
 */
function refreshDetailSheet_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET_NAMES.DETAIL);
  if (!sh) return;
  
  const current = sh.getRange('A2').getFormula();
  if (current.includes('GENERATE_DETAIL_BREAKDOWN')) {
    // Обновляем тикер для пересчёта
    sh.getRange('K2').setValue(new Date().toISOString());
    sh.getRange('A2').setFormula(current);
    SpreadsheetApp.flush();
    try {
      styleSheetHeader_(sh);
      styleDetailSheet_(sh);
    } catch (_) {}
  }
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
    
    families.set(id, {
      name: String(r[idx['Ребёнок ФИО']] || '').trim(),
      active: String(r[idx['Активен']] || '').trim() === ACTIVE_STATUS.YES,
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
      // Добровольный взнос: начисление = платёж
      return participants.has(fid) ? paid : 0;
      
    default:
      return 0;
  }
}

// ======================================================================
// MODULE: src/sheets/summary.js
// ======================================================================

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
  sh.getRange('L1').setValue('Фильтр');
  sh.getRange('K1').setValue('ALL');
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['OPEN', 'ALL'], true)
    .setAllowInvalid(false)
    .build();
  sh.getRange('K1').setDataValidation(rule).setHorizontalAlignment('center');
  sh.getRange('K1').setNote('OPEN (только открытые) или ALL (все цели, сначала открытые, ниже — закрытые)');
  
  sh.getRange('L2').setValue('Tick');
  sh.getRange('K2').setValue(new Date().toISOString());
  
  // Array formula
  sh.getRange('A2').setFormula(`=GENERATE_COLLECTION_SUMMARY(IF(LEN($K$1)=0,"ALL",$K$1), $K$2)`);
  
  sh.getRange('L3').setValue('Сводка по целям. ALL: сверху открытые, внизу закрытые.');
  
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

// ======================================================================
// MODULE: src/sheets/issue-status.js
// ======================================================================

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

// ======================================================================
// MODULE: src/core/init.js
// ======================================================================

/**
 * Инициализирует или пересоздаёт структуру таблицы
 * Точка входа из меню
 */
function init() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  
  // Определяем версию
  const version = detectVersion();
  
  if (version === 'v1') {
    // Предлагаем миграцию
    const response = ui.alert(
      'Обнаружена версия 1.x',
      'Таблица использует старую версию (Сборы/collection_id).\n\n' +
      'Выполнить миграцию на v2.0 (Цели/goal_id)?\n\n' +
      'Нажмите «Да» для миграции или «Нет» для работы в режиме совместимости.',
      ui.ButtonSet.YES_NO_CANCEL
    );
    
    if (response === ui.Button.YES) {
      migrateToV2();
      return;
    } else if (response === ui.Button.CANCEL) {
      return;
    }
    // Продолжаем с v1 структурой
    initV1Structure_(ss);
  } else if (version === 'v2') {
    // Обновляем v2 структуру
    initV2Structure_(ss);
  } else {
    // Новая таблица — создаём v2
    initV2Structure_(ss);
  }
  
  // Общие настройки
  setupListsSheet();
  setupNamedRanges_();
  rebuildValidations();
  setupBalanceExamples();
  addHeaderNotes_();
  styleWorkbook_();
  
  SpreadsheetApp.getActive().toast('Structure initialized.', 'Funds');
}

/**
 * Инициализирует структуру для v2.0
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function initV2Structure_(ss) {
  const specs = getSheetsSpec();
  
  specs.forEach(spec => {
    // Пропускаем legacy лист «Сборы» в v2
    if (spec.name === SHEET_NAMES.COLLECTIONS) return;
    
    const sh = getOrCreateSheet(ss, spec.name);
    
    // Заголовки
    const headerRange = sh.getRange(1, 1, 1, spec.headers.length);
    if (headerRange.getValues()[0].join('') === '') {
      headerRange.setValues([spec.headers]);
    }
    
    // Ширины колонок
    spec.colWidths.forEach((w, i) => {
      if (w) sh.setColumnWidth(i + 1, w);
    });
    
    // Форматы дат
    if (spec.dateCols) {
      spec.dateCols.forEach(col => {
        sh.getRange(2, col, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
      });
    }
    
    sh.setFrozenRows(1);
  });
  
  // Инструкция для v2.0
  setupInstructionSheet();
}

/**
 * Инициализирует структуру для v1.x (режим совместимости)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function initV1Structure_(ss) {
  const specs = getSheetsSpecV1();
  
  specs.forEach(spec => {
    // Пропускаем новый лист «Цели» в v1
    if (spec.name === SHEET_NAMES.GOALS) return;
    
    const sh = getOrCreateSheet(ss, spec.name);
    
    const headerRange = sh.getRange(1, 1, 1, spec.headers.length);
    if (headerRange.getValues()[0].join('') === '') {
      headerRange.setValues([spec.headers]);
    }
    
    spec.colWidths.forEach((w, i) => {
      if (w) sh.setColumnWidth(i + 1, w);
    });
    
    if (spec.dateCols) {
      spec.dateCols.forEach(col => {
        sh.getRange(2, col, sh.getMaxRows() - 1, 1).setNumberFormat('yyyy-mm-dd');
      });
    }
    
    sh.setFrozenRows(1);
  });
  
  setupInstructionSheetV1_();
}

/**
 * Настраивает именованные диапазоны
 */
function setupNamedRanges_() {
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  
  // Диапазоны для меток (labels) — всегда из Lists
  ensureNamedRange(NAMED_RANGES.FAMILIES_LABELS, 'Lists!D2:D');
  ensureNamedRange(NAMED_RANGES.ACTIVE_FAMILIES_LABELS, 'Lists!C2:C');
  
  if (version === 'v2' || version === 'new') {
    ensureNamedRange(NAMED_RANGES.GOALS_LABELS, 'Lists!B2:B');
    ensureNamedRange(NAMED_RANGES.OPEN_GOALS_LABELS, 'Lists!A2:A');
  } else {
    // v1.x
    ensureNamedRange(NAMED_RANGES.COLLECTIONS_LABELS, 'Lists!B2:B');
    ensureNamedRange(NAMED_RANGES.OPEN_COLLECTIONS_LABELS, 'Lists!A2:A');
  }
  
  // Raw ID диапазоны
  setRawIdNamedRanges_();
}

/**
 * Устанавливает именованные диапазоны для raw ID
 */
function setRawIdNamedRanges_() {
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  
  // Семьи
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  if (shF) {
    const mapF = getHeaderMap_(shF);
    const fIdCol = colToLetter_(mapF['family_id'] || 1);
    ensureNamedRange('FAMILIES', `Семьи!${fIdCol}2:${fIdCol}`);
  }
  
  // Цели или Сборы
  if (version === 'v2' || version === 'new') {
    const shG = ss.getSheetByName(SHEET_NAMES.GOALS);
    if (shG) {
      const mapG = getHeaderMap_(shG);
      const gIdCol = colToLetter_(mapG['goal_id'] || 1);
      ensureNamedRange('GOALS', `Цели!${gIdCol}2:${gIdCol}`);
    }
  } else {
    const shC = ss.getSheetByName(SHEET_NAMES.COLLECTIONS);
    if (shC) {
      const mapC = getHeaderMap_(shC);
      const cIdCol = colToLetter_(mapC['collection_id'] || 1);
      ensureNamedRange('COLLECTIONS', `Сборы!${cIdCol}2:${cIdCol}`);
    }
  }
}

// ======================================================================
// MODULE: src/core/validations.js
// ======================================================================

/**
 * Перестраивает все валидации данных
 * Точка входа из меню
 */
function rebuildValidations() {
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  
  const lists = {
    // v2.0 статусы целей
    goalStatus: [GOAL_STATUS.OPEN, GOAL_STATUS.CLOSED, GOAL_STATUS.CANCELLED],
    // v1.x статусы сборов
    collectionStatus: [COLLECTION_STATUS_V1.OPEN, COLLECTION_STATUS_V1.CLOSED],
    // Общие
    activeYesNo: [ACTIVE_STATUS.YES, ACTIVE_STATUS.NO],
    accrualRules: [
      ACCRUAL_MODES.STATIC_PER_FAMILY,
      ACCRUAL_MODES.SHARED_TOTAL_ALL,
      ACCRUAL_MODES.SHARED_TOTAL_BY_PAYERS,
      ACCRUAL_MODES.DYNAMIC_BY_PAYERS,
      ACCRUAL_MODES.PROPORTIONAL_BY_PAYERS,
      ACCRUAL_MODES.UNIT_PRICE,
      ACCRUAL_MODES.VOLUNTARY
    ],
    goalTypes: [GOAL_TYPES.ONE_TIME, GOAL_TYPES.REGULAR],
    periodicity: [GOAL_PERIODICITY.MONTHLY, GOAL_PERIODICITY.QUARTERLY, GOAL_PERIODICITY.YEARLY],
    payMethods: PAYMENT_METHODS,
    partStatus: [PARTICIPATION_STATUS.PARTICIPATES, PARTICIPATION_STATUS.NOT_PARTICIPATES]
  };
  
  // Семьи: Активен
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  if (shF) {
    const mapF = getHeaderMap_(shF);
    if (mapF['Активен']) {
      setValidationList_(shF, 2, mapF['Активен'], lists.activeYesNo);
    }
  }
  
  // Цели или Сборы в зависимости от версии
  if (version === 'v2' || version === 'new') {
    setupGoalsValidations_(ss, lists);
  } else {
    setupCollectionsValidations_(ss, lists);
  }
  
  // Участие
  setupParticipationValidations_(ss, version, lists);
  
  // Платежи
  setupPaymentsValidations_(ss, version, lists);
  
  // Выдача
  setupIssuesValidations_(ss, version, lists);
  
  SpreadsheetApp.getActive().toast('Validations rebuilt.', 'Funds');
}

/**
 * Настраивает валидации для листа «Цели» (v2.0)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} lists
 */
function setupGoalsValidations_(ss, lists) {
  const sh = ss.getSheetByName(SHEET_NAMES.GOALS);
  if (!sh) return;
  
  const map = getHeaderMap_(sh);
  const maxRows = sh.getMaxRows();
  
  // Очищаем старые валидации перед установкой новых (для миграции v1→v2)
  const clearCol = (col) => {
    if (col && maxRows > 1) {
      sh.getRange(2, col, maxRows - 1, 1).clearDataValidations();
    }
  };
  
  if (map['Тип']) {
    clearCol(map['Тип']);
    setValidationList_(sh, 2, map['Тип'], lists.goalTypes);
  }
  if (map['Статус']) {
    clearCol(map['Статус']);
    setValidationList_(sh, 2, map['Статус'], lists.goalStatus);
  }
  if (map['Начисление']) {
    clearCol(map['Начисление']);
    setValidationList_(sh, 2, map['Начисление'], lists.accrualRules);
  }
  if (map['К выдаче детям']) {
    clearCol(map['К выдаче детям']);
    setValidationList_(sh, 2, map['К выдаче детям'], lists.activeYesNo);
  }
  if (map['Возмещено']) {
    clearCol(map['Возмещено']);
    setValidationList_(sh, 2, map['Возмещено'], lists.activeYesNo);
  }
  if (map['Периодичность']) {
    clearCol(map['Периодичность']);
    setValidationList_(sh, 2, map['Периодичность'], lists.periodicity);
  }
}

/**
 * Настраивает валидации для листа «Сборы» (v1.x)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {Object} lists
 */
function setupCollectionsValidations_(ss, lists) {
  const sh = ss.getSheetByName(SHEET_NAMES.COLLECTIONS);
  if (!sh) return;
  
  const map = getHeaderMap_(sh);
  
  // v1.x использует старые режимы начисления
  const v1AccrualRules = [
    'static_per_child',
    'shared_total_all',
    'shared_total_by_payers',
    'dynamic_by_payers',
    'proportional_by_payers',
    'unit_price_by_payers'
  ];
  
  if (map['Статус']) {
    setValidationList_(sh, 2, map['Статус'], lists.collectionStatus);
  }
  if (map['Начисление']) {
    setValidationList_(sh, 2, map['Начисление'], v1AccrualRules);
  }
  if (map['К выдаче детям']) {
    setValidationList_(sh, 2, map['К выдаче детям'], lists.activeYesNo);
  }
  if (map['Возмещено']) {
    setValidationList_(sh, 2, map['Возмещено'], lists.activeYesNo);
  }
}

/**
 * Настраивает валидации для листа «Участие»
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} version
 * @param {Object} lists
 */
function setupParticipationValidations_(ss, version, lists) {
  const sh = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  if (!sh) return;
  
  const map = getHeaderMap_(sh);
  
  // Метки целей/сборов — только открытые
  const goalLabel = version === 'v1' ? 'collection_id (label)' : 'goal_id (label)';
  const namedRangeOpen = version === 'v1' ? NAMED_RANGES.OPEN_COLLECTIONS_LABELS : NAMED_RANGES.OPEN_GOALS_LABELS;
  
  if (map[goalLabel]) {
    setValidationNamedRange_(sh, 2, map[goalLabel], namedRangeOpen);
  }
  if (map['family_id (label)']) {
    setValidationNamedRange_(sh, 2, map['family_id (label)'], NAMED_RANGES.ACTIVE_FAMILIES_LABELS);
  }
  if (map['Статус']) {
    setValidationList_(sh, 2, map['Статус'], lists.partStatus);
  }
}

/**
 * Настраивает валидации для листа «Платежи»
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} version
 * @param {Object} lists
 */
function setupPaymentsValidations_(ss, version, lists) {
  const sh = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  if (!sh) return;
  
  const map = getHeaderMap_(sh);
  
  // Семья — все семьи
  if (map['family_id (label)']) {
    setValidationNamedRange_(sh, 2, map['family_id (label)'], NAMED_RANGES.FAMILIES_LABELS);
  }
  
  // Цель/сбор — все (включая закрытые)
  // v2: goal_id опционален (свободный платёж), разрешаем пустое значение
  const goalLabel = version === 'v1' ? 'collection_id (label)' : 'goal_id (label)';
  const namedRangeAll = version === 'v1' ? NAMED_RANGES.COLLECTIONS_LABELS : NAMED_RANGES.GOALS_LABELS;
  const allowBlankGoal = version !== 'v1'; // v2 разрешает пустой goal_id
  
  if (map[goalLabel]) {
    setValidationNamedRange_(sh, 2, map[goalLabel], namedRangeAll, allowBlankGoal);
  }
  
  // Способ оплаты
  if (map['Способ']) {
    setValidationList_(sh, 2, map['Способ'], lists.payMethods);
  }
  
  // Сумма > 0
  if (map['Сумма']) {
    setValidationNumberGreaterThan_(sh, 2, map['Сумма'], 0);
  }
}

/**
 * Настраивает валидации для листа «Выдача»
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} version
 * @param {Object} lists
 */
function setupIssuesValidations_(ss, version, lists) {
  const sh = ss.getSheetByName(SHEET_NAMES.ISSUES);
  if (!sh) return;
  
  const map = getHeaderMap_(sh);
  
  const goalLabel = version === 'v1' ? 'collection_id (label)' : 'goal_id (label)';
  const namedRangeAll = version === 'v1' ? NAMED_RANGES.COLLECTIONS_LABELS : NAMED_RANGES.GOALS_LABELS;
  
  if (map[goalLabel]) {
    setValidationNamedRange_(sh, 2, map[goalLabel], namedRangeAll);
  }
  if (map['family_id (label)']) {
    setValidationNamedRange_(sh, 2, map['family_id (label)'], NAMED_RANGES.FAMILIES_LABELS);
  }
  if (map['Единиц']) {
    setValidationNumberGreaterThan_(sh, 2, map['Единиц'], 0);
  }
  if (map['Выдано']) {
    setValidationList_(sh, 2, map['Выдано'], lists.activeYesNo);
  }
}

/**
 * Устанавливает валидацию выпадающего списка
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {number} rowStart
 * @param {number} col
 * @param {string[]} values
 */
function setValidationList_(sh, rowStart, col, values) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  sh.getRange(rowStart, col, sh.getMaxRows() - rowStart + 1, 1).setDataValidation(rule);
}

/**
 * Устанавливает валидацию по именованному диапазону
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {number} rowStart
 * @param {number} col
 * @param {string} namedRange
 * @param {boolean} [allowBlank=false] — разрешить пустые значения
 */
function setValidationNamedRange_(sh, rowStart, col, namedRange, allowBlank) {
  const ss = SpreadsheetApp.getActive();
  const nr = ss.getRangeByName(namedRange);
  if (!nr) {
    Logger.log('Named range not found: ' + namedRange);
    return;
  }
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(nr, true)
    .setAllowInvalid(allowBlank === true)
    .build();
  sh.getRange(rowStart, col, sh.getMaxRows() - rowStart + 1, 1).setDataValidation(rule);
}

/**
 * Устанавливает валидацию числа > minValue
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {number} rowStart
 * @param {number} col
 * @param {number} minValue
 */
function setValidationNumberGreaterThan_(sh, rowStart, col, minValue) {
  const rule = SpreadsheetApp.newDataValidation()
    .requireNumberGreaterThan(minValue)
    .setAllowInvalid(false)
    .build();
  sh.getRange(rowStart, col, sh.getMaxRows() - rowStart + 1, 1).setDataValidation(rule);
}

/**
 * Аудит и исправление типов полей
 * Удаляет случайные валидации с текстовых/числовых/дата полей
 */
function auditAndFixFieldTypes() {
  const ss = SpreadsheetApp.getActive();
  let fixes = 0;
  
  /**
   * Очищает валидацию с колонки
   * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
   * @param {number} col
   */
  function clearValidation(sh, col) {
    if (!sh || !col) return;
    const last = Math.max(2, sh.getMaxRows());
    const rng = sh.getRange(2, col, last - 1, 1);
    try {
      const rules = rng.getDataValidations();
      let hasAny = false;
      for (let i = 0; i < rules.length; i++) {
        if (rules[i] && rules[i][0]) { hasAny = true; break; }
      }
      if (hasAny) { rng.clearDataValidations(); fixes++; }
    } catch (_) {}
  }
  
  // Семьи: очищаем валидации с текстовых и дата полей
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  if (shF) {
    const map = getHeaderMap_(shF);
    
    if (map['День рождения']) {
      clearValidation(shF, map['День рождения']);
      const last = Math.max(2, shF.getMaxRows());
      shF.getRange(2, map['День рождения'], last - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
    
    ['Ребёнок ФИО', 'Мама ФИО', 'Мама телефон', 'Мама реквизиты', 'Мама телеграм',
     'Папа ФИО', 'Папа телефон', 'Папа реквизиты', 'Папа телеграм', 'Комментарий'].forEach(h => {
      if (map[h]) clearValidation(shF, map[h]);
    });
    
    if (map['family_id']) clearValidation(shF, map['family_id']);
  }
  
  // Цели/Сборы
  [SHEET_NAMES.GOALS, SHEET_NAMES.COLLECTIONS].forEach(shName => {
    const sh = ss.getSheetByName(shName);
    if (!sh) return;
    
    const map = getHeaderMap_(sh);
    const last = Math.max(2, sh.getMaxRows());
    
    ['Дата начала', 'Дедлайн'].forEach(h => {
      if (map[h]) {
        clearValidation(sh, map[h]);
        sh.getRange(2, map[h], last - 1, 1).setNumberFormat('yyyy-mm-dd');
      }
    });
    
    ['Параметр суммы', 'Фиксированный x'].forEach(h => {
      if (map[h]) {
        clearValidation(sh, map[h]);
        sh.getRange(2, map[h], last - 1, 1).setNumberFormat('#,##0.00');
      }
    });
    
    const textFields = shName === SHEET_NAMES.GOALS 
      ? ['Название цели', 'Комментарий', 'Ссылка на гуглдиск', 'Закупка из средств']
      : ['Название сбора', 'Комментарий', 'Ссылка на гуглдиск', 'Закупка из средств'];
    
    textFields.forEach(h => {
      if (map[h]) clearValidation(sh, map[h]);
    });
    
    const idField = shName === SHEET_NAMES.GOALS ? 'goal_id' : 'collection_id';
    if (map[idField]) clearValidation(sh, map[idField]);
  });
  
  // Платежи
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  if (shP) {
    const map = getHeaderMap_(shP);
    const last = Math.max(2, shP.getMaxRows());
    
    if (map['Дата']) {
      clearValidation(shP, map['Дата']);
      shP.getRange(2, map['Дата'], last - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
    if (map['Сумма']) {
      clearValidation(shP, map['Сумма']);
      shP.getRange(2, map['Сумма'], last - 1, 1).setNumberFormat('#,##0.00');
    }
    if (map['Комментарий']) clearValidation(shP, map['Комментарий']);
    if (map['payment_id']) clearValidation(shP, map['payment_id']);
  }
  
  // Пересоздаём валидации
  rebuildValidations();
  
  // Обновляем стили
  try { styleWorkbook_(); } catch (_) {}
  
  SpreadsheetApp.getActive().toast(`Audit complete. Fixed: ${fixes} columns.`, 'Funds');
}

// ======================================================================
// MODULE: src/core/id-generator.js
// ======================================================================

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

// ======================================================================
// MODULE: src/core/close-goal.js
// ======================================================================

/**
 * Диалог закрытия цели
 * Точка входа из меню
 */
function closeGoalPrompt() {
  const ui = SpreadsheetApp.getUi();
  const version = detectVersion();
  
  const idLabel = version === 'v1' ? 'collection_id' : 'goal_id';
  const example = version === 'v1' ? 'C001' : 'G001';
  
  const resp = ui.prompt(
    'Close Goal',
    `Введите ${idLabel} (например, ${example}):`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  
  const goalId = (resp.getResponseText() || '').trim();
  if (!goalId) return;
  
  if (version === 'v1') {
    closeCollection_(goalId);
  } else {
    closeGoal_(goalId);
  }
}

/**
 * Закрывает цель (v2.0)
 * @param {string} goalId
 */
function closeGoal_(goalId) {
  const ss = SpreadsheetApp.getActive();
  const shG = ss.getSheetByName(SHEET_NAMES.GOALS);
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  const shU = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  
  const mapG = getHeaderMap_(shG);
  
  // Находим строку цели
  const idCol = mapG['goal_id'];
  if (!idCol) return toastErr_('Не найден столбец goal_id.');
  
  const rowsG = shG.getLastRow();
  if (rowsG < 2) return toastErr_('Нет целей.');
  
  const ids = shG.getRange(2, idCol, rowsG - 1, 1).getValues().map(r => String(r[0] || '').trim());
  const idx = ids.findIndex(v => v === goalId);
  if (idx === -1) return toastErr_('Цель не найдена: ' + goalId);
  
  const rowNum = 2 + idx;
  
  // Читаем данные цели
  const accrual = normalizeAccrualMode_(String(shG.getRange(rowNum, mapG['Начисление']).getValue() || '').trim());
  const paramT = Number(shG.getRange(rowNum, mapG['Параметр суммы']).getValue() || 0);
  
  // Для dynamic_by_payers вычисляем и фиксируем x
  if (accrual === ACCRUAL_MODES.DYNAMIC_BY_PAYERS) {
    // Читаем данные
    const families = readFamilies_(shF);
    const participation = readParticipation_(shU, 'v2');
    const payments = readPayments_(shP, 'v2');
    
    // Читаем даты цели для resolveParticipants_
    const goal = {
      startDate: parseDate_(shG.getRange(rowNum, mapG['Дата начала']).getValue()),
      deadline: parseDate_(shG.getRange(rowNum, mapG['Дедлайн']).getValue())
    };
    
    const participants = resolveParticipants_(goalId, families, participation, goal);
    const goalPayments = payments.get(goalId) || new Map();
    
    // Собираем платежи участников
    const paymentArray = [];
    goalPayments.forEach((sum, fid) => {
      if (participants.has(fid) && sum > 0) paymentArray.push(sum);
    });
    
    const x = DYN_CAP_(paramT, paymentArray);
    
    if (mapG['Фиксированный x']) {
      shG.getRange(rowNum, mapG['Фиксированный x']).setValue(x);
    }
    if (mapG['Статус']) {
      shG.getRange(rowNum, mapG['Статус']).setValue(GOAL_STATUS.CLOSED);
    }
    
    SpreadsheetApp.getActive().toast(`Цель ${goalId} закрыта. x=${round2_(x)}`, 'Funds');
  } else {
    // Для других режимов просто закрываем
    if (mapG['Статус']) {
      shG.getRange(rowNum, mapG['Статус']).setValue(GOAL_STATUS.CLOSED);
    }
    SpreadsheetApp.getActive().toast(`Цель ${goalId} закрыта.`, 'Funds');
  }
}

/**
 * Закрывает сбор (v1.x)
 * @param {string} collectionId
 */
function closeCollection_(collectionId) {
  const ss = SpreadsheetApp.getActive();
  const shC = ss.getSheetByName(SHEET_NAMES.COLLECTIONS);
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  const shU = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  
  const mapC = getHeaderMap_(shC);
  
  const idCol = mapC['collection_id'];
  if (!idCol) return toastErr_('Не найден столбец collection_id.');
  
  const rowsC = shC.getLastRow();
  if (rowsC < 2) return toastErr_('Нет сборов.');
  
  const ids = shC.getRange(2, idCol, rowsC - 1, 1).getValues().map(r => String(r[0] || '').trim());
  const idx = ids.findIndex(v => v === collectionId);
  if (idx === -1) return toastErr_('Сбор не найден: ' + collectionId);
  
  const rowNum = 2 + idx;
  
  const accrual = String(shC.getRange(rowNum, mapC['Начисление']).getValue() || '').trim();
  const paramT = Number(shC.getRange(rowNum, mapC['Параметр суммы']).getValue() || 0);
  
  if (accrual === 'dynamic_by_payers') {
    const families = readFamilies_(shF);
    const participation = readParticipation_(shU, 'v1');
    const payments = readPayments_(shP, 'v1');
    
    // v1 не имеет дат, используем null
    const participants = resolveParticipants_(collectionId, families, participation, null);
    const goalPayments = payments.get(collectionId) || new Map();
    
    const paymentArray = [];
    goalPayments.forEach((sum, fid) => {
      if (participants.has(fid) && sum > 0) paymentArray.push(sum);
    });
    
    const x = DYN_CAP_(paramT, paymentArray);
    
    if (mapC['Фиксированный x']) {
      shC.getRange(rowNum, mapC['Фиксированный x']).setValue(x);
    }
    if (mapC['Статус']) {
      shC.getRange(rowNum, mapC['Статус']).setValue(COLLECTION_STATUS_V1.CLOSED);
    }
    
    SpreadsheetApp.getActive().toast(`Сбор ${collectionId} закрыт. x=${round2_(x)}`, 'Funds');
  } else {
    if (mapC['Статус']) {
      shC.getRange(rowNum, mapC['Статус']).setValue(COLLECTION_STATUS_V1.CLOSED);
    }
    SpreadsheetApp.getActive().toast(`Сбор ${collectionId} закрыт.`, 'Funds');
  }
}

/**
 * Алиас для обратной совместимости
 */
function closeCollectionPrompt() {
  closeGoalPrompt();
}

// ======================================================================
// MODULE: src/core/sample-data.js
// ======================================================================

/**
 * Диалог загрузки демо-данных
 * Точка входа из меню
 */
function loadSampleDataPrompt() {
  const ui = SpreadsheetApp.getUi();
  const choice = ui.alert(
    'Load Sample Data',
    'Это добавит демонстрационные данные (семьи, цели, участие, платежи).\n\n' +
    'Существующие данные не стираются, но могут перемешаться с демо.\n\n' +
    'Продолжить?',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (choice !== ui.Button.OK) return;
  
  const version = detectVersion();
  if (version === 'v1') {
    loadSampleDataV1_();
  } else {
    loadSampleDataV2_();
  }
  
  SpreadsheetApp.getActive().toast('Demo data added.', 'Funds');
  refreshBalanceFormulas_();
}

/**
 * Загружает демо-данные для v2.0
 */
function loadSampleDataV2_() {
  const ss = SpreadsheetApp.getActive();
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  const shG = ss.getSheetByName(SHEET_NAMES.GOALS);
  const shU = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  
  const mapF = getHeaderMap_(shF);
  const mapG = getHeaderMap_(shG);
  const mapP = getHeaderMap_(shP);
  
  // Семьи (10 примеров)
  const famStart = shF.getLastRow() + 1;
  const famRows = [
    ['Иванов Иван', '2015-03-15', 'Иванова Анна', '+7 900 000-00-01', '****1111', '@anna_ivanova', 'Иванов Пётр', '+7 900 000-10-01', '****2222', '@petr_ivanov', 'Да', '', ''],
    ['Петров Пётр', '2015-06-02', 'Петрова Мария', '+7 900 000-00-02', '****3333', '@petrova_m', 'Петров Иван', '+7 900 000-10-02', '****4444', '@ivan_petrov', 'Да', '', ''],
    ['Сидорова Вера', '2015-01-21', 'Сидорова Ольга', '+7 900 000-00-03', '****5555', '@sidorova_olga', 'Сидоров Антон', '+7 900 000-10-03', '****6666', '@sid_anton', 'Да', '', ''],
    ['Кузнецов Артём', '2015-12-11', 'Кузнецова Ирина', '+7 900 000-00-04', '****7777', '@irina_kuz', 'Кузнецов Олег', '+7 900 000-10-04', '****8888', '@oleg_kuz', 'Да', '', ''],
    ['Смирнова Юля', '2015-08-05', 'Смирнова Анна', '+7 900 000-00-05', '****9999', '@anna_smir', 'Смирнов Роман', '+7 900 000-10-05', '****0001', '@roman_smir', 'Да', '', ''],
    ['Новикова Нина', '2015-04-19', 'Новикова Оксана', '+7 900 000-00-06', '****0002', '@oks_nov', 'Новиков Павел', '+7 900 000-10-06', '****0003', '@pavel_nov', 'Да', '', ''],
    ['Орлова Лена', '2015-07-23', 'Орлова Татьяна', '+7 900 000-00-07', '****0004', '@tat_orl', 'Орлов Юрий', '+7 900 000-10-07', '****0005', '@y_orlov', 'Да', '', ''],
    ['Фёдоров Даня', '2015-02-14', 'Фёдорова Алла', '+7 900 000-00-08', '****0006', '@alla_fed', 'Фёдоров Игорь', '+7 900 000-10-08', '****0007', '@igor_fed', 'Да', '', ''],
    ['Максимова Аня', '2015-09-30', 'Максимова Ника', '+7 900 000-00-09', '****0008', '@nika_maks', 'Максимов Артём', '+7 900 000-10-09', '****0009', '@art_maks', 'Да', '', ''],
    ['Егорова Саша', '2015-11-01', 'Егорова Алина', '+7 900 000-00-10', '****0010', '@alina_egor', 'Егоров Кирилл', '+7 900 000-10-10', '****0011', '@kir_egor', 'Да', '', '']
  ];
  shF.getRange(famStart, 1, famRows.length, shF.getLastColumn()).setValues(famRows);
  
  if (mapF['family_id']) {
    fillMissingIds_(ss, SHEET_NAMES.FAMILIES, mapF['family_id'], ID_PREFIXES.FAMILY, 3);
  }
  
  // Цели (демо для всех режимов)
  const goalStart = shG.getLastRow() + 1;
  // Headers: ['Название цели', 'Тип', 'Статус', 'Дата начала', 'Дедлайн', 'Начисление', 'Параметр суммы', 'Фиксированный x', 'К выдаче детям', 'Периодичность', 'Родительская цель', 'Закупка из средств', 'Возмещено', 'Комментарий', 'goal_id', 'Ссылка на гуглдиск']
  const goalRows = [
    ['Канцтовары сентябрь', GOAL_TYPES.ONE_TIME, GOAL_STATUS.OPEN, '', '', ACCRUAL_MODES.STATIC_PER_FAMILY, 500, '', '', '', '', '', '', 'Фикс 500₽ на семью', '', ''],
    ['Новый год', GOAL_TYPES.ONE_TIME, GOAL_STATUS.OPEN, '', '', ACCRUAL_MODES.SHARED_TOTAL_ALL, 12000, '', '', '', '', '', '', 'Общая сумма делится на участников', '', ''],
    ['Подарок учителю', GOAL_TYPES.ONE_TIME, GOAL_STATUS.OPEN, '', '', ACCRUAL_MODES.DYNAMIC_BY_PAYERS, 9000, '', '', '', '', '', '', 'Динамический сбор по цели 9000₽', '', ''],
    ['Фотосессия', GOAL_TYPES.ONE_TIME, GOAL_STATUS.OPEN, '', '', ACCRUAL_MODES.SHARED_TOTAL_BY_PAYERS, 10000, '', '', '', '', '', '', 'Делим сумму между оплатившими', '', ''],
    ['Помощь классу', GOAL_TYPES.ONE_TIME, GOAL_STATUS.OPEN, '', '', ACCRUAL_MODES.PROPORTIONAL_BY_PAYERS, 8000, '', '', '', '', '', '', 'Пропорционально платежам', '', ''],
    ['Спортивная форма', GOAL_TYPES.ONE_TIME, GOAL_STATUS.OPEN, '', '', ACCRUAL_MODES.UNIT_PRICE, 15000, 1500, 'Да', '', '', '', 'Нет', 'Поштучная закупка: x=1500₽', '', ''],
    ['Благотворительность', GOAL_TYPES.ONE_TIME, GOAL_STATUS.OPEN, '', '', ACCRUAL_MODES.VOLUNTARY, 0, '', '', '', '', '', '', 'Добровольный взнос без цели', '', '']
  ];
  shG.getRange(goalStart, 1, goalRows.length, shG.getLastColumn()).setValues(goalRows);
  
  if (mapG['goal_id']) {
    fillMissingIds_(ss, SHEET_NAMES.GOALS, mapG['goal_id'], ID_PREFIXES.GOAL, 3);
  }
  
  // Обновляем Lists
  setupListsSheet();
  
  // Получаем метки новых целей
  const newCount = goalRows.length;
  const gVals = shG.getRange(goalStart, 1, newCount, shG.getLastColumn()).getValues();
  const gHdr = shG.getRange(1, 1, 1, shG.getLastColumn()).getValues()[0];
  const gi = {};
  gHdr.forEach((h, idx) => gi[h] = idx);
  
  const labelByName = new Map();
  gVals.forEach(r => {
    const nm = String(r[gi['Название цели']] || '').trim();
    const id = String(r[gi['goal_id']] || '').trim();
    if (nm && id) labelByName.set(nm, `${nm} (${id})`);
  });
  
  const g1Label = labelByName.get('Канцтовары сентябрь') || '';
  const g2Label = labelByName.get('Новый год') || '';
  const g3Label = labelByName.get('Подарок учителю') || '';
  const g4Label = labelByName.get('Фотосессия') || '';
  const g5Label = labelByName.get('Помощь классу') || '';
  const g6Label = labelByName.get('Спортивная форма') || '';
  const g7Label = labelByName.get('Благотворительность') || '';
  
  // Метки семей
  const allFam = getLabelColumn_(SHEET_NAMES.LISTS, 'D', 2);
  
  // Участие
  const partStart = shU.getLastRow() + 1;
  const partRows = [];
  
  // G002: явно отмечаем 8 семей как участников
  allFam.slice(0, 8).forEach(lbl => partRows.push([g2Label, lbl, PARTICIPATION_STATUS.PARTICIPATES, '']));
  
  // G003: исключаем 2 семьи
  allFam.slice(0, 2).forEach(lbl => partRows.push([g3Label, lbl, PARTICIPATION_STATUS.NOT_PARTICIPATES, '']));
  
  if (partRows.length) {
    shU.getRange(partStart, 1, partRows.length, 4).setValues(partRows);
  }
  
  // Платежи
  const payStart = shP.getLastRow() + 1;
  const today = new Date();
  const addDays = (d) => new Date(today.getTime() + d * 24 * 3600 * 1000);
  const payRows = [];
  
  // G001 (static 500): 6 платят полностью, 2 частично, 2 не платят
  allFam.slice(0, 6).forEach((lbl, i) => payRows.push([toISO_(addDays(-5 + i)), lbl, g1Label, 500, 'СБП', 'Полная оплата', '']));
  allFam.slice(6, 8).forEach((lbl, i) => payRows.push([toISO_(addDays(-2 - i)), lbl, g1Label, 300, 'карта', 'Частичная оплата', '']));
  
  // G002 (shared 12000 среди 8)
  const shareFamilies = allFam.slice(0, 8);
  shareFamilies.slice(0, 5).forEach((lbl, i) => payRows.push([toISO_(addDays(-3 + i)), lbl, g2Label, 1500, 'СБП', 'Частично/полностью', '']));
  shareFamilies.slice(5, 8).forEach((lbl, i) => payRows.push([toISO_(addDays(-2 - i)), lbl, g2Label, 800, 'наличные', 'Частично', '']));
  
  // G003 (dynamic 9000, исключая 2 семьи): [2000, 2000, 700, 700, 700, 700, 700]
  const dynFamilies = allFam.slice(2);
  dynFamilies.slice(0, 2).forEach((lbl, i) => payRows.push([toISO_(addDays(-6 + i)), lbl, g3Label, 2000, 'СБП', 'Ранний платёж', '']));
  dynFamilies.slice(2, 7).forEach((lbl, i) => payRows.push([toISO_(addDays(-1 - i)), lbl, g3Label, 700, 'карта', 'Позже', '']));
  
  // G004 (shared_total_by_payers 10000): 4 семьи платят
  if (g4Label) {
    allFam.slice(0, 4).forEach((lbl, i) => payRows.push([toISO_(addDays(-4 + i)), lbl, g4Label, 2500, i % 2 ? 'карта' : 'СБП', 'Оплата доли', '']));
  }
  
  // G005 (proportional_by_payers 8000): 5 семей разные суммы
  if (g5Label) {
    const fams = allFam.slice(2, 7);
    const amounts = [3000, 2000, 1500, 800, 500];
    fams.forEach((lbl, i) => payRows.push([toISO_(addDays(-2 + i)), lbl, g5Label, amounts[i], i % 2 ? 'карта' : 'СБП', 'Разные суммы', '']));
  }
  
  // G006 (unit_price T=15000, x=1500)
  if (g6Label) {
    const fams = allFam.slice(0, 8);
    const amounts = [1500, 1500, 1500, 3000, 4500, 1500, 700, 700];
    fams.forEach((lbl, i) => payRows.push([
      toISO_(addDays(-7 + i)),
      lbl,
      g6Label,
      amounts[i],
      (i % 2 ? 'карта' : 'СБП'),
      amounts[i] >= 1500 ? (amounts[i] % 1500 === 0 ? `${amounts[i] / 1500} ед.` : 'Частично') : 'Частично',
      ''
    ]));
  }
  
  // G007 (voluntary): несколько добровольных взносов
  if (g7Label) {
    allFam.slice(0, 3).forEach((lbl, i) => payRows.push([toISO_(addDays(-1)), lbl, g7Label, [100, 500, 200][i], 'СБП', 'Добровольно', '']));
  }
  
  if (payRows.length) {
    shP.getRange(payStart, 1, payRows.length, shP.getLastColumn()).setValues(payRows);
  }
  
  if (mapP['payment_id']) {
    fillMissingIds_(ss, SHEET_NAMES.PAYMENTS, mapP['payment_id'], ID_PREFIXES.PAYMENT, 3);
  }
  
  rebuildValidations();
}

/**
 * Загружает демо-данные для v1.x
 */
function loadSampleDataV1_() {
  // Аналогично v2, но с collection_id вместо goal_id и Сборы вместо Цели
  const ss = SpreadsheetApp.getActive();
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  const shC = ss.getSheetByName(SHEET_NAMES.COLLECTIONS);
  const shU = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  
  const mapF = getHeaderMap_(shF);
  const mapC = getHeaderMap_(shC);
  const mapP = getHeaderMap_(shP);
  
  // Семьи
  const famStart = shF.getLastRow() + 1;
  const famRows = [
    ['Иванов Иван', '2015-03-15', 'Иванова Анна', '+7 900 000-00-01', '****1111', '@anna_ivanova', 'Иванов Пётр', '+7 900 000-10-01', '****2222', '@petr_ivanov', 'Да', '', ''],
    ['Петров Пётр', '2015-06-02', 'Петрова Мария', '+7 900 000-00-02', '****3333', '@petrova_m', 'Петров Иван', '+7 900 000-10-02', '****4444', '@ivan_petrov', 'Да', '', ''],
    ['Сидорова Вера', '2015-01-21', 'Сидорова Ольга', '+7 900 000-00-03', '****5555', '@sidorova_olga', 'Сидоров Антон', '+7 900 000-10-03', '****6666', '@sid_anton', 'Да', '', ''],
    ['Кузнецов Артём', '2015-12-11', 'Кузнецова Ирина', '+7 900 000-00-04', '****7777', '@irina_kuz', 'Кузнецов Олег', '+7 900 000-10-04', '****8888', '@oleg_kuz', 'Да', '', ''],
    ['Смирнова Юля', '2015-08-05', 'Смирнова Анна', '+7 900 000-00-05', '****9999', '@anna_smir', 'Смирнов Роман', '+7 900 000-10-05', '****0001', '@roman_smir', 'Да', '', '']
  ];
  shF.getRange(famStart, 1, famRows.length, shF.getLastColumn()).setValues(famRows);
  
  if (mapF['family_id']) {
    fillMissingIds_(ss, SHEET_NAMES.FAMILIES, mapF['family_id'], ID_PREFIXES.FAMILY, 3);
  }
  
  // Сборы
  const colStart = shC.getLastRow() + 1;
  const colRows = [
    ['Канцтовары', 'Открыт', '', '', 'static_per_child', 500, '', '', '', '', '', '', ''],
    ['Подарок учителю', 'Открыт', '', '', 'dynamic_by_payers', 5000, '', '', '', '', '', '', '']
  ];
  shC.getRange(colStart, 1, colRows.length, shC.getLastColumn()).setValues(colRows);
  
  if (mapC['collection_id']) {
    fillMissingIds_(ss, SHEET_NAMES.COLLECTIONS, mapC['collection_id'], ID_PREFIXES.COLLECTION, 3);
  }
  
  setupListsSheet();
  
  // Метки сборов
  const cVals = shC.getRange(colStart, 1, colRows.length, shC.getLastColumn()).getValues();
  const cHdr = shC.getRange(1, 1, 1, shC.getLastColumn()).getValues()[0];
  const ci = {};
  cHdr.forEach((h, idx) => ci[h] = idx);
  
  const c1Label = `${cVals[0][ci['Название сбора']]} (${cVals[0][ci['collection_id']]})`;
  const c2Label = `${cVals[1][ci['Название сбора']]} (${cVals[1][ci['collection_id']]})`;
  
  const allFam = getLabelColumn_(SHEET_NAMES.LISTS, 'D', 2);
  
  // Платежи
  const payStart = shP.getLastRow() + 1;
  const today = new Date();
  const payRows = [];
  
  allFam.slice(0, 3).forEach((lbl, i) => payRows.push([toISO_(today), lbl, c1Label, 500, 'СБП', '', '']));
  allFam.slice(0, 4).forEach((lbl, i) => payRows.push([toISO_(today), lbl, c2Label, [1500, 1000, 800, 700][i], 'карта', '', '']));
  
  if (payRows.length) {
    shP.getRange(payStart, 1, payRows.length, shP.getLastColumn()).setValues(payRows);
  }
  
  if (mapP['payment_id']) {
    fillMissingIds_(ss, SHEET_NAMES.PAYMENTS, mapP['payment_id'], ID_PREFIXES.PAYMENT, 3);
  }
  
  rebuildValidations();
}

// ======================================================================
// MODULE: src/ui/menu.js
// ======================================================================

/**
 * Создаёт меню при открытии таблицы
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu('Funds');
  
  // Основные операции
  menu.addItem('Setup / Rebuild structure', 'init');
  menu.addItem('Generate IDs (all sheets)', 'generateAllIds');
  menu.addItem('Rebuild data validations', 'rebuildValidations');
  menu.addItem('Recalculate (Balance & Detail)', 'recalculateAll');
  
  menu.addSeparator();
  
  // Операции со сборами/целями
  menu.addItem('Close Goal', 'closeGoalPrompt');
  
  menu.addSeparator();
  
  // Демо и очистка
  menu.addItem('Load Sample Data (separate)', 'loadSampleDataPrompt');
  menu.addItem('Cleanup visuals (trim sheets)', 'cleanupWorkbook_');
  menu.addItem('Audit & fix field types', 'auditAndFixFieldTypes');
  
  menu.addSeparator();
  
  // Быстрые проверки
  menu.addItem('Quick Help', 'showQuickHelp_');
  menu.addItem('Quick Balance Check', 'showQuickBalanceCheck_');
  menu.addItem('Migration Report', 'showMigrationReport_');
  
  menu.addSeparator();
  
  // Миграция (если нужна)
  if (needsMigration()) {
    menu.addItem('🔄 Migrate v1 → v2', 'migrateToV2Prompt');
    menu.addSeparator();
  }
  
  // Управление бэкапами
  menu.addItem('Cleanup old backups', 'cleanupBackupsPrompt');
  
  // Информация
  menu.addItem('About', 'showAbout_');
  
  menu.addToUi();
}

/**
 * Показывает диалог «О программе»
 */
function showAbout_() {
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'Payment Accounting v' + APP_VERSION,
    'Учёт платежей и взносов для класса/группы.\n\n' +
    'Репозиторий: github.com/yobushka/paymentAccountingGoogleSheet\n\n' +
    'Версия: ' + APP_VERSION,
    ui.ButtonSet.OK
  );
}

/**
 * Показывает быструю справку
 */
function showQuickHelp_() {
  const ui = SpreadsheetApp.getUi();
  const help = `
Быстрый старт:
1. Funds → Setup / Rebuild structure
2. Заполните «Семьи» (Активен=Да)
3. Добавьте «Цели» (Статус=Открыта)
4. Настройте «Участие» при необходимости
5. Вносите «Платежи»
6. Смотрите «Баланс» и «Детализация»

Режимы начисления:
• static_per_family — фикс на семью
• shared_total_all — делим на всех участников
• shared_total_by_payers — делим между оплатившими
• dynamic_by_payers — water-filling
• proportional_by_payers — пропорционально платежам
• unit_price — поштучно
• voluntary — добровольно (v2.0)

Баланс v2.0:
Внесено - Списано - Резерв = Свободно
`;
  ui.alert('Quick Help', help, ui.ButtonSet.OK);
}

/**
 * Показывает быструю проверку баланса семьи
 */
function showQuickBalanceCheck_() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Quick Balance Check',
    'Введите family_id (например, F001):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const familyId = response.getResponseText().trim();
  if (!familyId) return;
  
  try {
    const paid = PAYED_TOTAL_FAMILY(familyId);
    const accrued = ACCRUED_FAMILY(familyId, 'ALL');
    const free = Math.max(0, paid - accrued);
    const debt = Math.max(0, accrued - paid);
    
    const msg = `
Семья: ${familyId}

Внесено всего: ${paid.toFixed(2)} ₽
Списано (начислено): ${accrued.toFixed(2)} ₽
Свободный остаток: ${free.toFixed(2)} ₽
Задолженность: ${debt.toFixed(2)} ₽
`;
    ui.alert('Balance Check', msg, ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('Ошибка', e.message, ui.ButtonSet.OK);
  }
}

// ======================================================================
// MODULE: src/ui/styles.js
// ======================================================================

/**
 * Применяет стили ко всей книге
 */
function styleWorkbook_() {
  const ss = SpreadsheetApp.getActive();
  const sheets = getSheetsSpec();
  
  sheets.forEach(spec => {
    const sh = ss.getSheetByName(spec.name);
    if (!sh) return;
    
    styleSheetHeader_(sh);
    
    // Применяем специфичные стили для каждого листа
    switch (spec.name) {
      case SHEET_NAMES.BALANCE:
        styleBalanceSheet_(sh);
        break;
      case SHEET_NAMES.PAYMENTS:
        stylePaymentsSheet_(sh);
        break;
      case SHEET_NAMES.GOALS:
        styleGoalsSheet_(sh);
        break;
      case SHEET_NAMES.COLLECTIONS:
        styleCollectionsSheet_(sh);
        break;
      case SHEET_NAMES.FAMILIES:
        styleFamiliesSheet_(sh);
        break;
      case SHEET_NAMES.PARTICIPATION:
        styleParticipationSheet_(sh);
        break;
      case SHEET_NAMES.DETAIL:
        styleDetailSheet_(sh);
        break;
      case SHEET_NAMES.SUMMARY:
        styleSummarySheet_(sh);
        break;
      case SHEET_NAMES.ISSUE_STATUS:
        styleIssueStatusSheet_(sh);
        break;
    }
  });
}

/**
 * Стилизует заголовок листа (первая строка)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleSheetHeader_(sh) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1) return;
  
  const headerRange = sh.getRange(1, 1, 1, lastCol);
  headerRange
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');
  
  sh.setFrozenRows(1);
}

/**
 * Стилизует лист «Баланс»
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleBalanceSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  const lastCol = sh.getLastColumn();
  const dataRange = sh.getRange(2, 1, lastRow - 1, lastCol);
  
  // Чередующиеся строки (zebra)
  applyZebraStripes_(sh, 2, lastRow);
  
  // Числовой формат для денежных колонок (C:G)
  if (lastCol >= 7) {
    sh.getRange(2, 3, lastRow - 1, 5).setNumberFormat('#,##0.00');
  }
  
  // Условное форматирование: Задолженность > 0 — красный фон
  const rules = sh.getConditionalFormatRules();
  const debtCol = 7; // Задолженность
  const debtRange = sh.getRange(2, debtCol, lastRow - 1, 1);
  const debtRule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberGreaterThan(0)
    .setBackground('#ffcdd2')
    .setRanges([debtRange])
    .build();
  rules.push(debtRule);
  sh.setConditionalFormatRules(rules);
  
  // Авто-фильтр
  try { sh.getRange(1, 1, lastRow, lastCol).createFilter(); } catch (_) {}
}

/**
 * Стилизует лист «Платежи»
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function stylePaymentsSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  
  // Формат даты
  if (map['Дата']) {
    sh.getRange(2, map['Дата'], lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
  }
  
  // Формат суммы
  if (map['Сумма']) {
    sh.getRange(2, map['Сумма'], lastRow - 1, 1).setNumberFormat('#,##0.00');
  }
  
  // Авто-фильтр
  try { sh.getDataRange().createFilter(); } catch (_) {}
}

/**
 * Стилизует лист «Цели» (v2.0)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleGoalsSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  
  // Форматы дат
  ['Дата начала', 'Дедлайн'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });
  
  // Форматы чисел
  ['Параметр суммы', 'Фиксированный x'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('#,##0.00');
    }
  });
  
  // Условное форматирование: Статус
  const statusCol = map['Статус'];
  if (statusCol) {
    const rules = sh.getConditionalFormatRules();
    const statusRange = sh.getRange(2, statusCol, lastRow - 1, 1);
    
    // Открыта — зелёный
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(GOAL_STATUS.OPEN)
      .setBackground('#c8e6c9')
      .setRanges([statusRange])
      .build());
    
    // Закрыта — серый
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(GOAL_STATUS.CLOSED)
      .setBackground('#e0e0e0')
      .setRanges([statusRange])
      .build());
    
    // Отменена — красный
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(GOAL_STATUS.CANCELLED)
      .setBackground('#ffcdd2')
      .setRanges([statusRange])
      .build());
    
    sh.setConditionalFormatRules(rules);
  }
  
  try { sh.getDataRange().createFilter(); } catch (_) {}
}

/**
 * Стилизует лист «Сборы» (v1.x legacy)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleCollectionsSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  
  ['Дата начала', 'Дедлайн'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
    }
  });
  
  ['Параметр суммы', 'Фиксированный x'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('#,##0.00');
    }
  });
  
  const statusCol = map['Статус'];
  if (statusCol) {
    const rules = sh.getConditionalFormatRules();
    const statusRange = sh.getRange(2, statusCol, lastRow - 1, 1);
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(COLLECTION_STATUS_V1.OPEN)
      .setBackground('#c8e6c9')
      .setRanges([statusRange])
      .build());
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(COLLECTION_STATUS_V1.CLOSED)
      .setBackground('#e0e0e0')
      .setRanges([statusRange])
      .build());
    
    sh.setConditionalFormatRules(rules);
  }
  
  try { sh.getDataRange().createFilter(); } catch (_) {}
}

/**
 * Стилизует лист «Семьи»
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleFamiliesSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  
  if (map['День рождения']) {
    sh.getRange(2, map['День рождения'], lastRow - 1, 1).setNumberFormat('yyyy-mm-dd');
  }
  
  // Условное форматирование: неактивные — серый фон
  const activeCol = map['Активен'];
  if (activeCol) {
    const rules = sh.getConditionalFormatRules();
    const lastCol = sh.getLastColumn();
    const rowRange = sh.getRange(2, 1, lastRow - 1, lastCol);
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=$${colToLetter_(activeCol)}2="Нет"`)
      .setBackground('#f5f5f5')
      .setFontColor('#9e9e9e')
      .setRanges([rowRange])
      .build());
    
    sh.setConditionalFormatRules(rules);
  }
  
  try { sh.getDataRange().createFilter(); } catch (_) {}
}

/**
 * Стилизует лист «Участие»
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleParticipationSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  const statusCol = map['Статус'];
  
  if (statusCol) {
    const rules = sh.getConditionalFormatRules();
    const statusRange = sh.getRange(2, statusCol, lastRow - 1, 1);
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(PARTICIPATION_STATUS.PARTICIPATES)
      .setBackground('#c8e6c9')
      .setRanges([statusRange])
      .build());
    
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(PARTICIPATION_STATUS.NOT_PARTICIPATES)
      .setBackground('#ffcdd2')
      .setRanges([statusRange])
      .build());
    
    sh.setConditionalFormatRules(rules);
  }
  
  try { sh.getDataRange().createFilter(); } catch (_) {}
}

/**
 * Стилизует лист «Детализация»
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleDetailSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  
  ['Оплачено', 'Начислено', 'Разность (±)'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('#,##0.00');
    }
  });
  
  // Условное форматирование: Разность
  const diffCol = map['Разность (±)'];
  if (diffCol) {
    const rules = sh.getConditionalFormatRules();
    const diffRange = sh.getRange(2, diffCol, lastRow - 1, 1);
    
    // Положительная разность (переплата) — зелёный
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground('#c8e6c9')
      .setRanges([diffRange])
      .build());
    
    // Отрицательная разность (недоплата) — красный
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberLessThan(0)
      .setBackground('#ffcdd2')
      .setRanges([diffRange])
      .build());
    
    sh.setConditionalFormatRules(rules);
  }
  
  try { sh.getDataRange().createFilter(); } catch (_) {}
}

/**
 * Стилизует лист «Сводка»
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleSummarySheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  
  ['Сумма цели', 'Собрано', 'Остаток до цели'].forEach(h => {
    if (map[h]) {
      sh.getRange(2, map[h], lastRow - 1, 1).setNumberFormat('#,##0.00');
    }
  });
}

/**
 * Стилизует лист «Статус выдачи»
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 */
function styleIssueStatusSheet_(sh) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return;
  
  applyZebraStripes_(sh, 2, lastRow);
  
  const map = getHeaderMap_(sh);
  
  if (map['x (цена)']) {
    sh.getRange(2, map['x (цена)'], lastRow - 1, 1).setNumberFormat('#,##0.00');
  }
}

/**
 * Применяет чередующиеся цвета строк (zebra stripes)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {number} startRow
 * @param {number} endRow
 */
function applyZebraStripes_(sh, startRow, endRow) {
  const lastCol = sh.getLastColumn();
  if (lastCol < 1 || endRow < startRow) return;
  
  const dataRange = sh.getRange(startRow, 1, endRow - startRow + 1, lastCol);
  
  // Удаляем старое banding
  const bandings = sh.getBandings();
  bandings.forEach(b => b.remove());
  
  // Применяем новое
  dataRange.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, false, false);
}

/**
 * Очищает книгу: удаляет пустые строки/колонки, обновляет стили
 */
function cleanupWorkbook_() {
  const ss = SpreadsheetApp.getActive();
  const sheets = ss.getSheets();
  
  sheets.forEach(sh => {
    // Не трогаем скрытый Lists
    if (sh.getName() === SHEET_NAMES.LISTS) return;
    
    // Удаляем лишние пустые строки
    const lastRow = sh.getLastRow();
    const maxRows = sh.getMaxRows();
    if (maxRows > lastRow + 10) {
      sh.deleteRows(lastRow + 11, maxRows - lastRow - 10);
    }
    
    // Удаляем лишние пустые колонки
    const lastCol = sh.getLastColumn();
    const maxCols = sh.getMaxColumns();
    if (maxCols > lastCol + 3) {
      sh.deleteColumns(lastCol + 4, maxCols - lastCol - 3);
    }
  });
  
  styleWorkbook_();
  SpreadsheetApp.getActive().toast('Cleanup complete.', 'Funds');
}

/**
 * Добавляет примечания к заголовкам листов
 */
function addHeaderNotes_() {
  const ss = SpreadsheetApp.getActive();
  
  // Примечания для листа «Цели»
  const goalsNotes = {
    'Название цели': 'Название цели/сбора',
    'Тип': 'разовая / регулярная',
    'Статус': 'Открыта / Закрыта / Отменена',
    'Начисление': 'Режим начисления:\n• static_per_family\n• shared_total_all\n• shared_total_by_payers\n• dynamic_by_payers\n• proportional_by_payers\n• unit_price\n• voluntary',
    'Параметр суммы': 'T — сумма цели или ставка на семью',
    'Фиксированный x': 'Для dynamic_by_payers — cap после закрытия.\nДля unit_price — цена единицы.',
    'К выдаче детям': 'Да — включает учёт выдачи для unit_price',
    'Периодичность': 'Для регулярных целей: ежемесячно / ежеквартально / ежегодно',
    'goal_id': 'ID цели (G001, G002, ...)'
  };
  
  setHeaderNotes_(ss.getSheetByName(SHEET_NAMES.GOALS), goalsNotes);
  
  // Примечания для листа «Семьи»
  const familyNotes = {
    'Активен': 'Да — семья участвует по умолчанию во всех целях',
    'family_id': 'ID семьи (F001, F002, ...)'
  };
  
  setHeaderNotes_(ss.getSheetByName(SHEET_NAMES.FAMILIES), familyNotes);
  
  // Примечания для листа «Участие»
  const partNotes = {
    'Статус': 'Участвует / Не участвует.\nЕсли есть хотя бы один «Участвует» — участвуют только отмеченные.\n«Не участвует» всегда исключает.'
  };
  
  setHeaderNotes_(ss.getSheetByName(SHEET_NAMES.PARTICIPATION), partNotes);
  
  // Примечания для листа «Платежи»
  const payNotes = {
    'Дата': 'Справочная дата (не влияет на расчёты)',
    'goal_id (label)': 'Выберите цель из выпадающего списка.\nПустое — свободный платёж (v2.0)',
    'Сумма': 'Сумма платежа > 0',
    'payment_id': 'ID платежа (PMT001, PMT002, ...)'
  };
  
  setHeaderNotes_(ss.getSheetByName(SHEET_NAMES.PAYMENTS), payNotes);
  
  // Примечания для листа «Баланс»
  const balanceNotes = {
    'Внесено всего': 'Сумма всех платежей семьи',
    'Списано всего': 'Сумма начислений по всем целям',
    'Зарезервировано': 'Зарезервировано под открытые цели',
    'Свободный остаток': 'Внесено - Списано - Резерв',
    'Задолженность': 'max(0, Списано - Внесено)'
  };
  
  setHeaderNotes_(ss.getSheetByName(SHEET_NAMES.BALANCE), balanceNotes);
}

/**
 * Устанавливает примечания к заголовкам листа
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {Object<string, string>} notes — {headerName: noteText}
 */
function setHeaderNotes_(sh, notes) {
  if (!sh) return;
  
  const map = getHeaderMap_(sh);
  
  Object.entries(notes).forEach(([header, note]) => {
    const col = map[header];
    if (col) {
      sh.getRange(1, col).setNote(note);
    }
  });
}

// ======================================================================
// MODULE: src/ui/dialogs.js
// ======================================================================

/**
 * Показывает краткую справку
 */
function showQuickHelp_() {
  const ui = SpreadsheetApp.getUi();
  
  const html = HtmlService.createHtmlOutput(`
    <style>
      body { font-family: sans-serif; padding: 10px; }
      h2 { color: #1a73e8; margin-top: 0; }
      h3 { color: #5f6368; margin-top: 15px; }
      ul { margin: 5px 0; padding-left: 20px; }
      li { margin: 3px 0; }
      code { background: #f1f3f4; padding: 2px 4px; border-radius: 3px; }
      .mode { margin: 8px 0; }
      .mode-name { font-weight: bold; color: #1a73e8; }
    </style>
    
    <h2>Funds v2.0 — Краткая справка</h2>
    
    <h3>Структура листов</h3>
    <ul>
      <li><strong>Семьи</strong> — список детей (family_id: F001, F002...)</li>
      <li><strong>Цели</strong> — сборы и цели (goal_id: G001, G002...)</li>
      <li><strong>Участие</strong> — кто участвует/не участвует в цели</li>
      <li><strong>Платежи</strong> — все поступления (payment_id: PMT001...)</li>
      <li><strong>Баланс</strong> — сводка по семьям (автоматический расчёт)</li>
    </ul>
    
    <h3>Режимы начисления</h3>
    <div class="mode">
      <span class="mode-name">static_per_family</span> — фиксированная сумма на семью
    </div>
    <div class="mode">
      <span class="mode-name">shared_total_all</span> — делим цель на всех участников
    </div>
    <div class="mode">
      <span class="mode-name">shared_total_by_payers</span> — делим на оплативших
    </div>
    <div class="mode">
      <span class="mode-name">dynamic_by_payers</span> — water-filling: справедливое распределение
    </div>
    <div class="mode">
      <span class="mode-name">proportional_by_payers</span> — пропорционально взносам
    </div>
    <div class="mode">
      <span class="mode-name">unit_price_by_payers</span> — поштучно (кратно цене)
    </div>
    <div class="mode">
      <span class="mode-name">voluntary</span> — добровольный взнос (списывается сколько внесено)
    </div>
    
    <h3>Основные действия</h3>
    <ul>
      <li><strong>Funds → Setup</strong> — первичная настройка</li>
      <li><strong>Funds → Generate IDs</strong> — автозаполнение ID</li>
      <li><strong>Funds → Rebuild Validations</strong> — обновить выпадающие списки</li>
      <li><strong>Funds → Close Goal</strong> — закрыть цель (фиксирует cap)</li>
    </ul>
    
    <h3>Типы целей (v2.0)</h3>
    <ul>
      <li><strong>разовая</strong> — однократный сбор</li>
      <li><strong>регулярная</strong> — повторяется с периодичностью</li>
    </ul>
  `).setWidth(500).setHeight(550);
  
  ui.showModalDialog(html, 'Справка');
}

/**
 * Быстрая проверка баланса для конкретной семьи
 */
function showQuickBalanceCheck_() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Проверка баланса',
    'Введите family_id (например, F001) или имя ребёнка:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const query = response.getResponseText().trim();
  if (!query) return;
  
  const ss = SpreadsheetApp.getActive();
  
  // Ищем семью
  const shFamilies = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  if (!shFamilies || shFamilies.getLastRow() < 2) {
    ui.alert('Ошибка', 'Лист «Семьи» пуст или не найден.', ui.ButtonSet.OK);
    return;
  }
  
  const familiesData = shFamilies.getDataRange().getValues();
  const fHeaders = familiesData[0];
  const fIdCol = fHeaders.indexOf('family_id');
  const fNameCol = fHeaders.indexOf('Имя ребёнка');
  
  let familyId = null;
  let familyName = null;
  
  for (let i = 1; i < familiesData.length; i++) {
    const row = familiesData[i];
    const id = String(row[fIdCol] || '');
    const name = String(row[fNameCol] || '');
    
    if (id.toLowerCase() === query.toLowerCase() ||
        name.toLowerCase().includes(query.toLowerCase())) {
      familyId = id;
      familyName = name;
      break;
    }
  }
  
  if (!familyId) {
    ui.alert('Не найдено', `Семья «${query}» не найдена.`, ui.ButtonSet.OK);
    return;
  }
  
  // Получаем баланс
  const shBalance = ss.getSheetByName(SHEET_NAMES.BALANCE);
  if (!shBalance || shBalance.getLastRow() < 2) {
    ui.alert('Ошибка', 'Лист «Баланс» пуст или не найден.', ui.ButtonSet.OK);
    return;
  }
  
  const balanceData = shBalance.getDataRange().getValues();
  const bHeaders = balanceData[0];
  const bIdCol = bHeaders.indexOf('family_id');
  
  let balanceRow = null;
  for (let i = 1; i < balanceData.length; i++) {
    if (balanceData[i][bIdCol] === familyId) {
      balanceRow = balanceData[i];
      break;
    }
  }
  
  if (!balanceRow) {
    ui.alert('Не найдено', `Баланс для ${familyId} не найден.`, ui.ButtonSet.OK);
    return;
  }
  
  // Формируем отчёт
  const getVal = (colName) => {
    const idx = bHeaders.indexOf(colName);
    return idx >= 0 ? balanceRow[idx] : 0;
  };
  
  const paid = getVal('Внесено всего') || getVal('Оплачено');
  const charged = getVal('Списано всего') || getVal('Начислено');
  const reserved = getVal('Зарезервировано') || 0;
  const free = getVal('Свободный остаток') || getVal('Переплата') || 0;
  const debt = getVal('Задолженность') || 0;
  
  const msg = `
Семья: ${familyName} (${familyId})

💰 Внесено всего: ${formatMoney_(paid)}
📊 Списано всего: ${formatMoney_(charged)}
🔒 Зарезервировано: ${formatMoney_(reserved)}
✅ Свободный остаток: ${formatMoney_(free)}
❌ Задолженность: ${formatMoney_(debt)}
`.trim();
  
  ui.alert(`Баланс: ${familyName}`, msg, ui.ButtonSet.OK);
}

/**
 * Форматирует число как деньги
 * @param {number} v
 * @return {string}
 */
function formatMoney_(v) {
  const n = Number(v) || 0;
  return n.toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) + ' ₽';
}

/**
 * Аудит типов данных в полях
 */
function showAuditFieldTypes_() {
  const ss = SpreadsheetApp.getActive();
  const results = [];
  
  const checkSheet = (name, expectedCols) => {
    const sh = ss.getSheetByName(name);
    if (!sh) {
      results.push(`⚠️ Лист «${name}» не найден`);
      return;
    }
    
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const missing = expectedCols.filter(c => !headers.includes(c));
    const extra = headers.filter(h => h && !expectedCols.includes(h));
    
    if (missing.length > 0) {
      results.push(`❌ ${name}: отсутствуют колонки: ${missing.join(', ')}`);
    }
    if (extra.length > 0) {
      results.push(`ℹ️ ${name}: дополнительные колонки: ${extra.join(', ')}`);
    }
    if (missing.length === 0 && extra.length === 0) {
      results.push(`✅ ${name}: структура корректна`);
    }
  };
  
  // Проверяем все листы
  checkSheet(SHEET_NAMES.FAMILIES, ['family_id', 'Имя ребёнка', 'Активен']);
  checkSheet(SHEET_NAMES.GOALS, [
    'goal_id', 'Название цели', 'Тип', 'Статус', 'Начисление', 
    'Параметр суммы', 'Периодичность', 'Родительская цель'
  ]);
  checkSheet(SHEET_NAMES.PARTICIPATION, ['family_id (label)', 'goal_id (label)', 'Участие']);
  checkSheet(SHEET_NAMES.PAYMENTS, [
    'payment_id', 'Дата', 'family_id (label)', 'goal_id (label)', 
    'Сумма', 'Комментарий'
  ]);
  checkSheet(SHEET_NAMES.BALANCE, [
    'family_id', 'Имя ребёнка', 'Внесено всего', 'Списано всего',
    'Зарезервировано', 'Свободный остаток', 'Задолженность'
  ]);
  
  const ui = SpreadsheetApp.getUi();
  ui.alert('Аудит структуры', results.join('\n'), ui.ButtonSet.OK);
}

/**
 * Показывает детальный отчёт по цели
 */
function showGoalReport_() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Отчёт по цели',
    'Введите goal_id (например, G001) или название:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const query = response.getResponseText().trim();
  if (!query) return;
  
  const ss = SpreadsheetApp.getActive();
  
  // Ищем цель
  const shGoals = ss.getSheetByName(SHEET_NAMES.GOALS);
  if (!shGoals || shGoals.getLastRow() < 2) {
    ui.alert('Ошибка', 'Лист «Цели» пуст или не найден.', ui.ButtonSet.OK);
    return;
  }
  
  const goalsData = shGoals.getDataRange().getValues();
  const gHeaders = goalsData[0];
  const gIdCol = gHeaders.indexOf('goal_id');
  const gNameCol = gHeaders.indexOf('Название цели');
  const gStatusCol = gHeaders.indexOf('Статус');
  const gModeCol = gHeaders.indexOf('Начисление');
  const gAmountCol = gHeaders.indexOf('Параметр суммы');
  
  let goalRow = null;
  
  for (let i = 1; i < goalsData.length; i++) {
    const row = goalsData[i];
    const id = String(row[gIdCol] || '');
    const name = String(row[gNameCol] || '');
    
    if (id.toLowerCase() === query.toLowerCase() ||
        name.toLowerCase().includes(query.toLowerCase())) {
      goalRow = row;
      break;
    }
  }
  
  if (!goalRow) {
    ui.alert('Не найдено', `Цель «${query}» не найдена.`, ui.ButtonSet.OK);
    return;
  }
  
  const goalId = goalRow[gIdCol];
  const goalName = goalRow[gNameCol];
  const goalStatus = goalRow[gStatusCol];
  const goalMode = goalRow[gModeCol];
  const goalAmount = goalRow[gAmountCol];
  
  // Считаем платежи по этой цели
  const shPayments = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  let totalPaid = 0;
  let payersCount = 0;
  
  if (shPayments && shPayments.getLastRow() > 1) {
    const payData = shPayments.getDataRange().getValues();
    const pHeaders = payData[0];
    const pGoalCol = pHeaders.indexOf('goal_id (label)');
    const pAmountCol = pHeaders.indexOf('Сумма');
    
    const payers = new Set();
    const pFamilyCol = pHeaders.indexOf('family_id (label)');
    
    for (let i = 1; i < payData.length; i++) {
      const goalLabel = String(payData[i][pGoalCol] || '');
      const extractedId = getIdFromLabelish_(goalLabel);
      
      if (extractedId === goalId) {
        totalPaid += Number(payData[i][pAmountCol]) || 0;
        payers.add(getIdFromLabelish_(payData[i][pFamilyCol]));
      }
    }
    payersCount = payers.size;
  }
  
  const msg = `
Цель: ${goalName} (${goalId})

📋 Статус: ${goalStatus}
📊 Режим: ${goalMode}
💵 Целевая сумма: ${formatMoney_(goalAmount)}

📥 Собрано: ${formatMoney_(totalPaid)}
👥 Плательщиков: ${payersCount}
📈 Прогресс: ${goalAmount > 0 ? Math.round(totalPaid / goalAmount * 100) : 0}%
`.trim();
  
  ui.alert(`Отчёт: ${goalName}`, msg, ui.ButtonSet.OK);
}

/**
 * Показывает общую статистику
 */
function showOverallStats_() {
  const ss = SpreadsheetApp.getActive();
  const ui = SpreadsheetApp.getUi();
  
  // Семьи
  const shFamilies = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  const familiesCount = shFamilies ? Math.max(0, shFamilies.getLastRow() - 1) : 0;
  
  // Цели
  const shGoals = ss.getSheetByName(SHEET_NAMES.GOALS);
  let goalsCount = 0;
  let openGoals = 0;
  
  if (shGoals && shGoals.getLastRow() > 1) {
    const goalsData = shGoals.getDataRange().getValues();
    const statusCol = goalsData[0].indexOf('Статус');
    goalsCount = goalsData.length - 1;
    
    if (statusCol >= 0) {
      for (let i = 1; i < goalsData.length; i++) {
        if (goalsData[i][statusCol] === 'Открыта') openGoals++;
      }
    }
  }
  
  // Платежи
  const shPayments = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  let paymentsCount = 0;
  let totalAmount = 0;
  
  if (shPayments && shPayments.getLastRow() > 1) {
    const payData = shPayments.getDataRange().getValues();
    const amountCol = payData[0].indexOf('Сумма');
    paymentsCount = payData.length - 1;
    
    if (amountCol >= 0) {
      for (let i = 1; i < payData.length; i++) {
        totalAmount += Number(payData[i][amountCol]) || 0;
      }
    }
  }
  
  const msg = `
📊 Общая статистика

👨‍👩‍👧‍👦 Семей: ${familiesCount}
🎯 Целей всего: ${goalsCount}
   • Открытых: ${openGoals}
   • Закрытых: ${goalsCount - openGoals}

💳 Платежей: ${paymentsCount}
💰 Общая сумма: ${formatMoney_(totalAmount)}
`.trim();
  
  ui.alert('Статистика', msg, ui.ButtonSet.OK);
}

// ======================================================================
// MODULE: src/triggers/on-edit.js
// ======================================================================

/**
 * Триггер onEdit — вызывается при редактировании ячейки
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e) {
  if (!e || !e.range) return;
  
  const sh = e.range.getSheet();
  const sheetName = sh.getName();
  const row = e.range.getRow();
  
  // Авто-генерация ID при начале ввода данных
  handleAutoIdGeneration_(sh, sheetName, row);
  
  // Авто-обновление балансов при изменении релевантных листов
  handleAutoRefresh_(sheetName);
}

/**
 * Автоматическая генерация ID при начале ввода данных
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {string} sheetName
 * @param {number} row
 */
function handleAutoIdGeneration_(sh, sheetName, row) {
  const version = detectVersion();
  
  switch (sheetName) {
    case SHEET_NAMES.FAMILIES:
      maybeAutoIdRow_(sh, row, 'family_id', ID_PREFIXES.FAMILY, 3, ['Ребёнок ФИО']);
      break;
      
    case SHEET_NAMES.GOALS:
      if (version === 'v2') {
        maybeAutoIdRow_(sh, row, 'goal_id', ID_PREFIXES.GOAL, 3, ['Название цели']);
      }
      break;
      
    case SHEET_NAMES.COLLECTIONS:
      if (version === 'v1') {
        maybeAutoIdRow_(sh, row, 'collection_id', ID_PREFIXES.COLLECTION, 3, ['Название сбора']);
      }
      break;
      
    case SHEET_NAMES.PAYMENTS:
      maybeAutoIdRow_(sh, row, 'payment_id', ID_PREFIXES.PAYMENT, 3, ['Сумма', 'family_id (label)']);
      break;
  }
}

/**
 * Автоматическое обновление балансов при изменении данных
 * @param {string} sheetName
 */
function handleAutoRefresh_(sheetName) {
  const relevantSheets = [
    SHEET_NAMES.PAYMENTS,
    SHEET_NAMES.FAMILIES,
    SHEET_NAMES.GOALS,
    SHEET_NAMES.COLLECTIONS,
    SHEET_NAMES.PARTICIPATION
  ];
  
  if (!relevantSheets.includes(sheetName)) return;
  
  // Запускаем обновление с небольшой задержкой для пакетной обработки
  try {
    // Обновляем тикер детализации для пересчёта
    const ss = SpreadsheetApp.getActive();
    const shDetail = ss.getSheetByName(SHEET_NAMES.DETAIL);
    if (shDetail) {
      const tickCell = shDetail.getRange('K2');
      tickCell.setValue(new Date().toISOString());
    }
    
    const shSummary = ss.getSheetByName(SHEET_NAMES.SUMMARY);
    if (shSummary) {
      const tickCell = shSummary.getRange('K2');
      tickCell.setValue(new Date().toISOString());
    }
  } catch (e) {
    // Игнорируем ошибки в onEdit
    Logger.log('Auto-refresh error: ' + e.message);
  }
}

/**
 * Автоматически заполняет ID для новой строки, если он пуст
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {number} row — номер строки
 * @param {string} idHeader — название колонки ID
 * @param {string} prefix — префикс ID
 * @param {number} width — ширина числовой части
 * @param {string[]} triggerHeaders — заголовки колонок-триггеров
 */
function maybeAutoIdRow_(sh, row, idHeader, prefix, width, triggerHeaders) {
  if (row < 2) return;
  
  const map = getHeaderMap_(sh);
  const idCol = map[idHeader];
  if (!idCol) return;
  
  const idVal = sh.getRange(row, idCol).getValue();
  if (idVal) return; // ID уже установлен
  
  // Проверяем, есть ли данные в триггерных колонках
  const hasTrigger = (triggerHeaders || []).some(h => {
    const c = map[h];
    if (!c) return false;
    const v = sh.getRange(row, c).getValue();
    return v !== '' && v !== null;
  });
  
  if (!hasTrigger) return;
  
  // Генерируем ID
  const ss = SpreadsheetApp.getActive();
  fillMissingIds_(ss, sh.getName(), idCol, prefix, width);
}

// ======================================================================
// MODULE: src/migration/detect-version.js
// ======================================================================

/**
 * Определяет версию структуры данных в таблице
 * @return {{version: string, needsMigration: boolean, details: Object}}
 */
function detectDataVersion() {
  const ss = SpreadsheetApp.getActive();
  
  const result = {
    version: 'unknown',
    needsMigration: false,
    details: {
      hasCollections: false,
      hasGoals: false,
      hasCollectionId: false,
      hasGoalId: false,
      hasOldAccrualModes: false,
      hasNewColumns: false
    }
  };
  
  // Проверяем наличие листов
  const shCollections = ss.getSheetByName('Сборы');
  const shGoals = ss.getSheetByName(SHEET_NAMES.GOALS);
  
  result.details.hasCollections = !!shCollections;
  result.details.hasGoals = !!shGoals;
  
  // Если есть лист «Сборы» — это v1.x
  if (shCollections && !shGoals) {
    result.version = '1.x';
    result.needsMigration = true;
    
    // Проверяем колонки
    const headers = shCollections.getRange(1, 1, 1, shCollections.getLastColumn()).getValues()[0];
    result.details.hasCollectionId = headers.includes('collection_id');
    
    // Проверяем режимы начисления
    const modeCol = headers.indexOf('Начисление') + 1;
    if (modeCol > 0) {
      const modes = shCollections.getRange(2, modeCol, Math.max(1, shCollections.getLastRow() - 1), 1)
        .getValues().flat().filter(Boolean);
      result.details.hasOldAccrualModes = modes.some(m => Object.keys(ACCRUAL_ALIASES).includes(m));
    }
    
    return result;
  }
  
  // Если есть лист «Цели» — это v2.0
  if (shGoals) {
    result.version = '2.0';
    result.needsMigration = false;
    
    const headers = shGoals.getRange(1, 1, 1, shGoals.getLastColumn()).getValues()[0];
    result.details.hasGoalId = headers.includes('goal_id');
    result.details.hasNewColumns = headers.includes('Тип') && headers.includes('Периодичность');
    
    return result;
  }
  
  // Новая таблица — создаём v2.0
  result.version = 'new';
  result.needsMigration = false;
  
  return result;
}

/**
 * Показывает диалог с информацией о версии
 */
function showVersionInfoPrompt() {
  const ui = SpreadsheetApp.getUi();
  const info = detectDataVersion();
  
  let msg = `Версия данных: ${info.version}\n\n`;
  
  if (info.needsMigration) {
    msg += '⚠️ Требуется миграция на v2.0\n\n';
    msg += 'Обнаружено:\n';
    msg += `• Лист «Сборы»: ${info.details.hasCollections ? 'да' : 'нет'}\n`;
    msg += `• collection_id: ${info.details.hasCollectionId ? 'да' : 'нет'}\n`;
    msg += `• Старые режимы начисления: ${info.details.hasOldAccrualModes ? 'да' : 'нет'}\n`;
    msg += '\nИспользуйте меню Funds → Migrate to v2.0';
  } else if (info.version === '2.0') {
    msg += '✅ Таблица актуальна\n\n';
    msg += `• Лист «Цели»: ${info.details.hasGoals ? 'да' : 'нет'}\n`;
    msg += `• goal_id: ${info.details.hasGoalId ? 'да' : 'нет'}\n`;
    msg += `• Новые колонки (Тип, Периодичность): ${info.details.hasNewColumns ? 'да' : 'нет'}`;
  } else {
    msg += 'Новая таблица — будет создана структура v2.0';
  }
  
  ui.alert('Информация о версии', msg, ui.ButtonSet.OK);
}

/**
 * Проверяет валидность данных после миграции
 * @return {{valid: boolean, errors: string[]}}
 */
function validateMigration() {
  const ss = SpreadsheetApp.getActive();
  const errors = [];
  
  // Проверяем что лист «Цели» существует
  const shGoals = ss.getSheetByName(SHEET_NAMES.GOALS);
  if (!shGoals) {
    errors.push('Лист «Цели» не найден');
  } else {
    const headers = shGoals.getRange(1, 1, 1, shGoals.getLastColumn()).getValues()[0];
    
    if (!headers.includes('goal_id')) {
      errors.push('Колонка goal_id не найдена на листе «Цели»');
    }
    
    // Проверяем формат goal_id
    const idCol = headers.indexOf('goal_id') + 1;
    if (idCol > 0 && shGoals.getLastRow() > 1) {
      const ids = shGoals.getRange(2, idCol, shGoals.getLastRow() - 1, 1).getValues().flat();
      const invalidIds = ids.filter(id => id && !String(id).match(/^G\d{3}$/));
      if (invalidIds.length > 0) {
        errors.push(`Некорректные goal_id: ${invalidIds.join(', ')}`);
      }
    }
  }
  
  // Проверяем лист «Участие»
  const shPart = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  if (shPart) {
    const headers = shPart.getRange(1, 1, 1, shPart.getLastColumn()).getValues()[0];
    if (headers.includes('collection_id (label)')) {
      errors.push('Лист «Участие» содержит устаревший заголовок collection_id');
    }
  }
  
  // Проверяем лист «Платежи»
  const shPay = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  if (shPay) {
    const headers = shPay.getRange(1, 1, 1, shPay.getLastColumn()).getValues()[0];
    if (headers.includes('collection_id (label)')) {
      errors.push('Лист «Платежи» содержит устаревший заголовок collection_id');
    }
  }
  
  return {
    valid: errors.length === 0,
    errors: errors
  };
}

// ======================================================================
// MODULE: src/migration/migrate-v1-to-v2.js
// ======================================================================

/**
 * Диалог миграции
 * Точка входа из меню
 */
function migrateToV2Prompt() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Миграция v1 → v2',
    'Будет выполнена автоматическая миграция:\n\n' +
    '1. Создан бэкап текущих листов\n' +
    '2. Лист «Сборы» переименован в «Цели»\n' +
    '3. collection_id заменён на goal_id\n' +
    '4. Обновлены заголовки и формулы\n' +
    '5. Добавлены новые колонки (Тип, Периодичность и др.)\n\n' +
    'Продолжить?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    migrateToV2();
    ui.alert(
      'Миграция завершена',
      'Таблица успешно обновлена до версии 2.0.\n\n' +
      'Бэкап сохранён в листах с суффиксом _backup_*.',
      ui.ButtonSet.OK
    );
  } catch (e) {
    ui.alert('Ошибка миграции', e.message, ui.ButtonSet.OK);
    Logger.log('Migration error: ' + e.message);
  }
}

/**
 * Выполняет миграцию v1.x → v2.0
 */
function migrateToV2() {
  const ss = SpreadsheetApp.getActive();
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  
  Logger.log('Starting migration v1 → v2...');
  
  // 1. Создаём бэкап
  createBackup_(ss, timestamp);
  
  // 2. Мигрируем лист «Сборы» → «Цели»
  migrateCollectionsToGoals_(ss);
  
  // 3. Обновляем лист «Участие»
  migrateParticipation_(ss);
  
  // 4. Обновляем лист «Платежи»
  migratePayments_(ss);
  
  // 5. Обновляем лист «Выдача»
  migrateIssues_(ss);
  
  // 6. Обновляем служебные листы
  migrateServiceSheets_(ss);
  
  // 7. Обновляем баланс и детализацию
  updateBalanceStructure_(ss);
  
  // 8. Пересоздаём Lists и валидации
  setupListsSheet();
  rebuildValidations();
  
  // 9. Обновляем инструкцию
  setupInstructionSheet();
  
  // 10. Пересчитываем
  refreshBalanceFormulas_();
  refreshDetailSheet_();
  refreshSummarySheet_();
  
  Logger.log('Migration completed successfully.');
  SpreadsheetApp.getActive().toast('Migration to v2.0 completed.', 'Funds');
}

/**
 * Создаёт бэкап листов перед миграцией
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} timestamp
 */
function createBackup_(ss, timestamp) {
  const sheetsToBackup = ['Сборы', 'Участие', 'Платежи', 'Баланс', 'Детализация', 'Сводка', 'Выдача'];
  
  sheetsToBackup.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (sh) {
      const copy = sh.copyTo(ss);
      copy.setName(`${name}_backup_${timestamp}`);
      copy.hideSheet();
    }
  });
  
  Logger.log('Backup created with timestamp: ' + timestamp);
}

/**
 * Мигрирует лист «Сборы» в «Цели»
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function migrateCollectionsToGoals_(ss) {
  const shC = ss.getSheetByName(SHEET_NAMES.COLLECTIONS);
  if (!shC) return;
  
  const headers = shC.getRange(1, 1, 1, shC.getLastColumn()).getValues()[0];
  const newHeaders = headers.map(h => {
    switch (h) {
      case 'Название сбора': return 'Название цели';
      case 'collection_id': return 'goal_id';
      default: return h;
    }
  });
  
  // Добавляем новые колонки v2.0
  const existingHeaders = new Set(newHeaders);
  const v2Headers = ['Тип', 'Периодичность', 'Родительская цель'];
  v2Headers.forEach(h => {
    if (!existingHeaders.has(h)) {
      newHeaders.push(h);
    }
  });
  
  // Обновляем заголовки
  shC.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  
  // Переименовываем ID: C001 → G001
  const idCol = newHeaders.indexOf('goal_id') + 1;
  if (idCol > 0) {
    const lastRow = shC.getLastRow();
    if (lastRow > 1) {
      const ids = shC.getRange(2, idCol, lastRow - 1, 1).getValues();
      const newIds = ids.map(r => {
        const old = String(r[0] || '');
        return [old.replace(/^C/, 'G')];
      });
      shC.getRange(2, idCol, lastRow - 1, 1).setValues(newIds);
    }
  }
  
  // Заполняем колонку «Тип» значением «разовая» по умолчанию
  const typeCol = newHeaders.indexOf('Тип') + 1;
  if (typeCol > 0) {
    const lastRow = shC.getLastRow();
    if (lastRow > 1) {
      const types = [];
      for (let i = 0; i < lastRow - 1; i++) {
        types.push([GOAL_TYPES.ONE_TIME]);
      }
      shC.getRange(2, typeCol, lastRow - 1, 1).setValues(types);
    }
  }
  
  // Обновляем режимы начисления (алиасы v1 → v2)
  const modeCol = newHeaders.indexOf('Начисление') + 1;
  if (modeCol > 0) {
    const lastRow = shC.getLastRow();
    if (lastRow > 1) {
      const modes = shC.getRange(2, modeCol, lastRow - 1, 1).getValues();
      const newModes = modes.map(r => {
        const old = String(r[0] || '');
        return [ACCRUAL_ALIASES[old] || old];
      });
      shC.getRange(2, modeCol, lastRow - 1, 1).setValues(newModes);
    }
  }
  
  // Переименовываем лист
  shC.setName(SHEET_NAMES.GOALS);
  
  Logger.log('Collections migrated to Goals.');
}

/**
 * Мигрирует лист «Участие»
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function migrateParticipation_(ss) {
  const sh = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  if (!sh) return;
  
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const newHeaders = headers.map(h => {
    return h === 'collection_id (label)' ? 'goal_id (label)' : h;
  });
  
  sh.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  
  // Обновляем ID в данных: C001 → G001
  const labelCol = newHeaders.indexOf('goal_id (label)') + 1;
  if (labelCol > 0) {
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const labels = sh.getRange(2, labelCol, lastRow - 1, 1).getValues();
      const newLabels = labels.map(r => {
        const old = String(r[0] || '');
        return [old.replace(/\(C(\d+)\)/, '(G$1)')];
      });
      sh.getRange(2, labelCol, lastRow - 1, 1).setValues(newLabels);
    }
  }
  
  Logger.log('Participation migrated.');
}

/**
 * Мигрирует лист «Платежи»
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function migratePayments_(ss) {
  const sh = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  if (!sh) return;
  
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const newHeaders = headers.map(h => {
    return h === 'collection_id (label)' ? 'goal_id (label)' : h;
  });
  
  sh.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  
  // Обновляем ID в данных
  const labelCol = newHeaders.indexOf('goal_id (label)') + 1;
  if (labelCol > 0) {
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const labels = sh.getRange(2, labelCol, lastRow - 1, 1).getValues();
      const newLabels = labels.map(r => {
        const old = String(r[0] || '');
        return [old.replace(/\(C(\d+)\)/, '(G$1)')];
      });
      sh.getRange(2, labelCol, lastRow - 1, 1).setValues(newLabels);
    }
  }
  
  Logger.log('Payments migrated.');
}

/**
 * Мигрирует лист «Выдача»
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function migrateIssues_(ss) {
  const sh = ss.getSheetByName(SHEET_NAMES.ISSUES);
  if (!sh) return;
  
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const newHeaders = headers.map(h => {
    return h === 'collection_id (label)' ? 'goal_id (label)' : h;
  });
  
  sh.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  
  const labelCol = newHeaders.indexOf('goal_id (label)') + 1;
  if (labelCol > 0) {
    const lastRow = sh.getLastRow();
    if (lastRow > 1) {
      const labels = sh.getRange(2, labelCol, lastRow - 1, 1).getValues();
      const newLabels = labels.map(r => {
        const old = String(r[0] || '');
        return [old.replace(/\(C(\d+)\)/, '(G$1)')];
      });
      sh.getRange(2, labelCol, lastRow - 1, 1).setValues(newLabels);
    }
  }
  
  Logger.log('Issues migrated.');
}

/**
 * Мигрирует служебные листы (Детализация, Сводка, Статус выдачи)
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function migrateServiceSheets_(ss) {
  // Детализация
  const shDetail = ss.getSheetByName(SHEET_NAMES.DETAIL);
  if (shDetail) {
    const headers = shDetail.getRange(1, 1, 1, shDetail.getLastColumn()).getValues()[0];
    const newHeaders = headers.map(h => {
      switch (h) {
        case 'collection_id': return 'goal_id';
        case 'Название сбора': return 'Название цели';
        default: return h;
      }
    });
    shDetail.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  }
  
  // Сводка
  const shSummary = ss.getSheetByName(SHEET_NAMES.SUMMARY);
  if (shSummary) {
    const headers = shSummary.getRange(1, 1, 1, shSummary.getLastColumn()).getValues()[0];
    const newHeaders = headers.map(h => {
      switch (h) {
        case 'collection_id': return 'goal_id';
        case 'Название сбора': return 'Название цели';
        default: return h;
      }
    });
    shSummary.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  }
  
  // Статус выдачи
  const shStatus = ss.getSheetByName(SHEET_NAMES.ISSUE_STATUS);
  if (shStatus) {
    const headers = shStatus.getRange(1, 1, 1, shStatus.getLastColumn()).getValues()[0];
    const newHeaders = headers.map(h => {
      return h === 'collection_id' ? 'goal_id' : h;
    });
    shStatus.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  }
  
  Logger.log('Service sheets migrated.');
}

/**
 * Обновляет структуру листа «Баланс» для v2.0
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function updateBalanceStructure_(ss) {
  const sh = ss.getSheetByName(SHEET_NAMES.BALANCE);
  if (!sh) return;
  
  // Новые заголовки v2.0
  const newHeaders = [
    'family_id', 'Имя ребёнка',
    'Внесено всего', 'Списано всего', 'Зарезервировано',
    'Свободный остаток', 'Задолженность'
  ];
  
  // Очищаем и записываем новые заголовки
  const lastCol = sh.getLastColumn();
  if (lastCol > 0) {
    sh.getRange(1, 1, 1, lastCol).clearContent();
  }
  sh.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  
  // Очищаем старые формулы
  const lastRow = sh.getLastRow();
  if (lastRow > 1) {
    sh.getRange(2, 3, lastRow - 1, Math.max(1, lastCol - 2)).clearContent();
  }
  
  Logger.log('Balance structure updated for v2.0.');
}

/**
 * Откатывает миграцию (восстанавливает из бэкапа)
 * @param {string} timestamp — таймстамп бэкапа
 */
function rollbackMigration(timestamp) {
  const ss = SpreadsheetApp.getActive();
  const sheetsToRestore = ['Сборы', 'Участие', 'Платежи', 'Баланс', 'Детализация', 'Сводка', 'Выдача'];
  
  sheetsToRestore.forEach(name => {
    const backup = ss.getSheetByName(`${name}_backup_${timestamp}`);
    const current = ss.getSheetByName(name) || ss.getSheetByName(
      name === 'Сборы' ? SHEET_NAMES.GOALS : name
    );
    
    if (backup && current) {
      // Удаляем текущий
      ss.deleteSheet(current);
      // Восстанавливаем из бэкапа
      backup.setName(name);
      backup.showSheet();
    }
  });
  
  SpreadsheetApp.getActive().toast('Rollback completed.', 'Funds');
}

/**
 * Показывает отчёт о миграции
 */
function showMigrationReport_() {
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  
  // Собираем статистику
  const stats = {
    version: version,
    families: 0,
    goals: 0,
    payments: 0,
    participation: 0,
    backups: []
  };
  
  // Семьи
  const shF = ss.getSheetByName(SHEET_NAMES.FAMILIES);
  if (shF) {
    const lastRow = shF.getLastRow();
    stats.families = lastRow > 1 ? lastRow - 1 : 0;
  }
  
  // Цели/сборы
  const shG = version === 'v1' 
    ? ss.getSheetByName(SHEET_NAMES.COLLECTIONS) 
    : ss.getSheetByName(SHEET_NAMES.GOALS);
  if (shG) {
    const lastRow = shG.getLastRow();
    stats.goals = lastRow > 1 ? lastRow - 1 : 0;
  }
  
  // Платежи
  const shP = ss.getSheetByName(SHEET_NAMES.PAYMENTS);
  if (shP) {
    const lastRow = shP.getLastRow();
    stats.payments = lastRow > 1 ? lastRow - 1 : 0;
  }
  
  // Участие
  const shU = ss.getSheetByName(SHEET_NAMES.PARTICIPATION);
  if (shU) {
    const lastRow = shU.getLastRow();
    stats.participation = lastRow > 1 ? lastRow - 1 : 0;
  }
  
  // Находим бэкапы
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    const match = name.match(/_backup_(\d{4}-\d{2}-\d{2}T[\d-]+)/);
    if (match) {
      const ts = match[1];
      if (!stats.backups.includes(ts)) {
        stats.backups.push(ts);
      }
    }
  });
  
  stats.backups.sort().reverse(); // Новейшие первыми
  
  // Формируем отчёт
  let report = `📊 Отчёт о состоянии таблицы\n\n`;
  report += `Версия: ${version === 'v1' ? '1.x (Сборы)' : '2.0 (Цели)'}\n\n`;
  report += `📁 Данные:\n`;
  report += `  • Семей: ${stats.families}\n`;
  report += `  • ${version === 'v1' ? 'Сборов' : 'Целей'}: ${stats.goals}\n`;
  report += `  • Платежей: ${stats.payments}\n`;
  report += `  • Записей участия: ${stats.participation}\n\n`;
  
  if (stats.backups.length > 0) {
    report += `💾 Бэкапы (${stats.backups.length}):\n`;
    stats.backups.slice(0, 5).forEach(ts => {
      report += `  • ${ts.replace('T', ' ')}\n`;
    });
    if (stats.backups.length > 5) {
      report += `  ... и ещё ${stats.backups.length - 5}\n`;
    }
  } else {
    report += `💾 Бэкапы: нет\n`;
  }
  
  if (version === 'v1') {
    report += `\n⚠️ Доступна миграция на v2.0:\n`;
    report += `Меню → Funds → Migrate v1 → v2`;
  }
  
  SpreadsheetApp.getUi().alert('Отчёт', report, SpreadsheetApp.getUi().ButtonSet.OK);
  return stats;
}

/**
 * Очищает старые бэкапы
 * @param {number} [keepCount=3] — сколько последних бэкапов сохранить
 */
function cleanupBackups_(keepCount) {
  const ss = SpreadsheetApp.getActive();
  const keep = keepCount || 3;
  
  // Собираем все таймстампы бэкапов
  const backupTimestamps = new Set();
  ss.getSheets().forEach(sh => {
    const name = sh.getName();
    const match = name.match(/_backup_(\d{4}-\d{2}-\d{2}T[\d-]+)/);
    if (match) {
      backupTimestamps.add(match[1]);
    }
  });
  
  // Сортируем (новейшие первыми) и определяем, какие удалить
  const sorted = Array.from(backupTimestamps).sort().reverse();
  const toDelete = sorted.slice(keep);
  
  if (toDelete.length === 0) {
    SpreadsheetApp.getActive().toast(`Нечего удалять. Бэкапов: ${sorted.length}`, 'Funds');
    return 0;
  }
  
  // Удаляем листы со старыми бэкапами
  let deleted = 0;
  toDelete.forEach(ts => {
    ss.getSheets().forEach(sh => {
      if (sh.getName().includes(`_backup_${ts}`)) {
        ss.deleteSheet(sh);
        deleted++;
      }
    });
  });
  
  Logger.log(`Deleted ${deleted} backup sheets (kept ${keep} most recent).`);
  SpreadsheetApp.getActive().toast(`Удалено бэкапов: ${toDelete.length} (листов: ${deleted})`, 'Funds');
  return deleted;
}

/**
 * Диалог очистки бэкапов
 */
function cleanupBackupsPrompt() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Очистка бэкапов',
    'Сколько последних бэкапов сохранить?\n\n' +
    '(Остальные будут удалены)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const keepCount = parseInt(response.getResponseText(), 10);
  if (isNaN(keepCount) || keepCount < 0) {
    ui.alert('Ошибка', 'Введите положительное число.', ui.ButtonSet.OK);
    return;
  }
  
  const deleted = cleanupBackups_(keepCount);
  ui.alert('Готово', `Удалено старых бэкапов: ${deleted}`, ui.ButtonSet.OK);
}

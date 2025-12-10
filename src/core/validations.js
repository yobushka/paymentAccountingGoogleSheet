/**
 * @fileoverview Валидации данных на листах
 */

/**
 * Перестраивает все валидации данных
 * Точка входа из меню
 */
function rebuildValidations() {
  const ss = SpreadsheetApp.getActive();
  const version = detectVersion();
  
  Logger.log('Rebuilding validations. Detected version: ' + version);
  
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
  
  Logger.log('Accrual rules for v2: ' + JSON.stringify(lists.accrualRules));
  
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
  if (!sh) {
    Logger.log('Sheet "Цели" not found for validations.');
    return;
  }
  
  Logger.log('Setting up validations for sheet "Цели"...');
  
  const map = getHeaderMap_(sh);
  const maxRows = sh.getMaxRows();
  
  Logger.log('Headers map: ' + JSON.stringify(map));
  Logger.log('Max rows: ' + maxRows);
  
  // Очищаем старые валидации перед установкой новых (для миграции v1→v2)
  const clearCol = (col, colName) => {
    if (col && maxRows > 1) {
      Logger.log(`  Clearing validations in column ${col} (${colName})...`);
      sh.getRange(2, col, maxRows - 1, 1).clearDataValidations();
    }
  };
  
  if (map['Тип']) {
    clearCol(map['Тип'], 'Тип');
    setValidationList_(sh, 2, map['Тип'], lists.goalTypes);
    Logger.log(`  Set validation "Тип" in column ${map['Тип']}: ${JSON.stringify(lists.goalTypes)}`);
  }
  if (map['Статус']) {
    clearCol(map['Статус'], 'Статус');
    setValidationList_(sh, 2, map['Статус'], lists.goalStatus);
    Logger.log(`  Set validation "Статус" in column ${map['Статус']}: ${JSON.stringify(lists.goalStatus)}`);
  }
  if (map['Начисление']) {
    clearCol(map['Начисление'], 'Начисление');
    setValidationList_(sh, 2, map['Начисление'], lists.accrualRules);
    Logger.log(`  Set validation "Начисление" in column ${map['Начисление']}: ${JSON.stringify(lists.accrualRules)}`);
  }
  if (map['К выдаче детям']) {
    clearCol(map['К выдаче детям'], 'К выдаче детям');
    setValidationList_(sh, 2, map['К выдаче детям'], lists.activeYesNo);
  }
  if (map['Возмещено']) {
    clearCol(map['Возмещено'], 'Возмещено');
    setValidationList_(sh, 2, map['Возмещено'], lists.activeYesNo);
  }
  if (map['Периодичность']) {
    clearCol(map['Периодичность'], 'Периодичность');
    setValidationList_(sh, 2, map['Периодичность'], lists.periodicity);
  }
  
  Logger.log('Validations for "Цели" completed.');
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

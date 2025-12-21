/**
 * @fileoverview Обработчик триггера onEdit
 */

/**
 * Триггер onEdit — вызывается при редактировании ячейки
 * @param {GoogleAppsScript.Events.SheetsOnEdit} e
 */
function onEdit(e) {
  if (!e || !e.range) return;
  
  const sh = e.range.getSheet();
  const sheetName = sh.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  // Авто-генерация ID при начале ввода данных
  handleAutoIdGeneration_(sh, sheetName, row);
  
  // Авто-обновление балансов при изменении релевантных листов
  handleAutoRefresh_(sheetName);
  
  // Зависимая валидация Статья → Подстатья на листе Цели
  handleDependentSubarticleValidation_(sh, sheetName, row, col, e.value);
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

/**
 * Обработка зависимой валидации Статья → Подстатья
 * При изменении Статьи обновляет список допустимых Подстатей
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sh
 * @param {string} sheetName
 * @param {number} row
 * @param {number} col
 * @param {*} newValue
 */
function handleDependentSubarticleValidation_(sh, sheetName, row, col, newValue) {
  // Работаем только с листом Цели
  if (sheetName !== SHEET_NAMES.GOALS) return;
  if (row < 2) return;
  
  const map = getHeaderMap_(sh);
  const articleCol = map['Статья'];
  const subarticleCol = map['Подстатья'];
  
  // Проверяем, что редактировали колонку Статья
  if (!articleCol || col !== articleCol) return;
  if (!subarticleCol) return;
  
  const selectedArticle = String(newValue || '').trim();
  const subarticleCell = sh.getRange(row, subarticleCol);
  
  // Если Статья очищена — показываем все подстатьи из Lists
  if (!selectedArticle) {
    const ss = SpreadsheetApp.getActive();
    const sourceRange = ss.getRange('Lists!I2:I');
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInRange(sourceRange, true)
      .setAllowInvalid(true)
      .build();
    subarticleCell.setDataValidation(rule);
    subarticleCell.setValue(''); // Сбрасываем подстатью
    return;
  }
  
  // Получаем подстатьи для выбранной статьи из листа Смета
  const ss = SpreadsheetApp.getActive();
  const shBudget = ss.getSheetByName(SHEET_NAMES.BUDGET);
  if (!shBudget) return;
  
  const budgetData = shBudget.getDataRange().getValues();
  const budgetMap = getHeaderMap_(shBudget);
  const budgetArticleCol = budgetMap['Статья'] || 1;
  const budgetSubarticleCol = budgetMap['Подстатья'] || 2;
  
  // Собираем уникальные подстатьи для выбранной статьи
  const subarticles = [];
  for (let i = 1; i < budgetData.length; i++) {
    const art = String(budgetData[i][budgetArticleCol - 1] || '').trim();
    const subart = String(budgetData[i][budgetSubarticleCol - 1] || '').trim();
    if (art === selectedArticle && subart && !subarticles.includes(subart)) {
      subarticles.push(subart);
    }
  }
  
  // Устанавливаем валидацию
  if (subarticles.length > 0) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(subarticles, true)
      .setAllowInvalid(true)
      .build();
    subarticleCell.setDataValidation(rule);
  } else {
    // Нет подстатей для этой статьи — убираем валидацию
    subarticleCell.clearDataValidations();
  }
  
  // Проверяем, валидна ли текущая подстатья
  const currentSubarticle = String(subarticleCell.getValue() || '').trim();
  if (currentSubarticle && !subarticles.includes(currentSubarticle)) {
    subarticleCell.setValue(''); // Сбрасываем невалидную подстатью
  }
}

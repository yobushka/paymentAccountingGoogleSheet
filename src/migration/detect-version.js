/**
 * @fileoverview Детекция версии данных в таблице
 */

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

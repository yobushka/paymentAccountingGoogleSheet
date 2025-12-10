/**
 * @fileoverview Миграция v1.x → v2.0
 */

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
      const backupName = `${name}_backup_${timestamp}`;
      copy.setName(backupName);
      copy.hideSheet();
      
      // ВАЖНО: Удаляем именованные диапазоны из бэкап-листа, чтобы избежать конфликтов
      removeNamedRangesFromSheet_(ss, backupName);
    }
  });
  
  Logger.log('Backup created with timestamp: ' + timestamp);
}

/**
 * Удаляет все именованные диапазоны, ссылающиеся на указанный лист
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @param {string} sheetName
 */
function removeNamedRangesFromSheet_(ss, sheetName) {
  const namedRanges = ss.getNamedRanges();
  let removed = 0;
  
  namedRanges.forEach(nr => {
    try {
      const range = nr.getRange();
      if (range && range.getSheet().getName() === sheetName) {
        nr.remove();
        removed++;
      }
    } catch (e) {
      // Диапазон может быть невалидным
    }
  });
  
  if (removed > 0) {
    Logger.log(`Removed ${removed} named ranges from sheet "${sheetName}"`);
  }
}

/**
 * Мигрирует лист «Сборы» в «Цели»
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 */
function migrateCollectionsToGoals_(ss) {
  const shC = ss.getSheetByName(SHEET_NAMES.COLLECTIONS);
  if (!shC) {
    Logger.log('Sheet "Сборы" not found, skipping.');
    return;
  }
  
  Logger.log('Migrating Collections → Goals...');
  
  // ВАЖНО: Сначала очищаем ВСЕ валидации на листе, чтобы избежать конфликтов
  const lastRow = shC.getLastRow();
  const lastCol = shC.getLastColumn();
  if (lastRow > 1 && lastCol > 0) {
    Logger.log(`Clearing all validations on sheet. Rows: ${lastRow}, Cols: ${lastCol}`);
    shC.getRange(1, 1, lastRow, lastCol).clearDataValidations();
  }
  
  const headers = shC.getRange(1, 1, 1, lastCol).getValues()[0];
  Logger.log('Original headers: ' + JSON.stringify(headers));
  
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
  
  Logger.log('New headers: ' + JSON.stringify(newHeaders));
  
  // Обновляем заголовки
  shC.getRange(1, 1, 1, newHeaders.length).setValues([newHeaders]);
  
  // Переименовываем ID: C001 → G001
  const idCol = newHeaders.indexOf('goal_id') + 1;
  if (idCol > 0 && lastRow > 1) {
    Logger.log(`Updating goal_id in column ${idCol}...`);
    const ids = shC.getRange(2, idCol, lastRow - 1, 1).getValues();
    const newIds = ids.map(r => {
      const old = String(r[0] || '');
      return [old.replace(/^C/, 'G')];
    });
    shC.getRange(2, idCol, lastRow - 1, 1).setValues(newIds);
  }
  
  // Заполняем колонку «Тип» значением «разовая» по умолчанию
  const typeCol = newHeaders.indexOf('Тип') + 1;
  if (typeCol > 0 && lastRow > 1) {
    Logger.log(`Setting default "Тип" = "${GOAL_TYPES.ONE_TIME}" in column ${typeCol}...`);
    const types = [];
    for (let i = 0; i < lastRow - 1; i++) {
      types.push([GOAL_TYPES.ONE_TIME]);
    }
    shC.getRange(2, typeCol, lastRow - 1, 1).setValues(types);
  }
  
  // Обновляем режимы начисления (алиасы v1 → v2)
  const modeCol = newHeaders.indexOf('Начисление') + 1;
  if (modeCol > 0 && lastRow > 1) {
    Logger.log(`Updating accrual modes in column ${modeCol}...`);
    const modes = shC.getRange(2, modeCol, lastRow - 1, 1).getValues();
    Logger.log('Old modes: ' + JSON.stringify(modes.map(r => r[0])));
    
    const newModes = modes.map(r => {
      const old = String(r[0] || '');
      const newMode = ACCRUAL_ALIASES[old] || old;
      if (old !== newMode) {
        Logger.log(`  Mode: "${old}" → "${newMode}"`);
      }
      return [newMode];
    });
    
    Logger.log('New modes: ' + JSON.stringify(newModes.map(r => r[0])));
    shC.getRange(2, modeCol, lastRow - 1, 1).setValues(newModes);
  }
  
  // Переименовываем лист
  shC.setName(SHEET_NAMES.GOALS);
  
  Logger.log('Collections migrated to Goals successfully.');
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

/**
 * Очищает именованные диапазоны из всех бэкап-листов
 * Вызывается для исправления проблем после неудачной миграции
 */
function cleanupBackupNamedRanges() {
  const ss = SpreadsheetApp.getActive();
  const namedRanges = ss.getNamedRanges();
  let removed = 0;
  
  namedRanges.forEach(nr => {
    try {
      const range = nr.getRange();
      if (range) {
        const sheetName = range.getSheet().getName();
        // Удаляем именованные диапазоны из бэкап-листов
        if (sheetName.includes('_backup_')) {
          Logger.log(`Removing named range "${nr.getName()}" from backup sheet "${sheetName}"`);
          nr.remove();
          removed++;
        }
      }
    } catch (e) {
      // Диапазон может быть невалидным — пробуем удалить по имени
      try {
        const name = nr.getName();
        if (name.includes('_backup_') || name.includes("'")) {
          nr.remove();
          removed++;
        }
      } catch (_) {}
    }
  });
  
  Logger.log(`Cleaned up ${removed} named ranges from backup sheets.`);
  SpreadsheetApp.getActive().toast(`Очищено именованных диапазонов: ${removed}`, 'Funds');
  return removed;
}

/**
 * Принудительный сброс к v1 и повторная миграция
 * Использовать если миграция застряла
 */
function forceMigrationReset() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Принудительный сброс миграции',
    'Это удалит ВСЕ бэкап-листы и их именованные диапазоны,\n' +
    'затем пересоздаст структуру с нуля.\n\n' +
    'Ваши данные на основных листах (Сборы/Цели, Семьи, Платежи) сохранятся.\n\n' +
    'Продолжить?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActive();
  
  // 1. Удаляем все бэкап-листы
  Logger.log('Removing all backup sheets...');
  const sheetsToDelete = ss.getSheets().filter(sh => sh.getName().includes('_backup_'));
  sheetsToDelete.forEach(sh => {
    try {
      // Сначала удаляем именованные диапазоны
      removeNamedRangesFromSheet_(ss, sh.getName());
      ss.deleteSheet(sh);
    } catch (e) {
      Logger.log(`Failed to delete sheet ${sh.getName()}: ${e.message}`);
    }
  });
  
  // 2. Очищаем оставшиеся проблемные именованные диапазоны
  cleanupBackupNamedRanges();
  
  // 3. Пересоздаём служебные листы
  Logger.log('Recreating service sheets...');
  try {
    setupListsSheet();
    rebuildValidations();
  } catch (e) {
    Logger.log('Error rebuilding: ' + e.message);
  }
  
  ui.alert(
    'Сброс выполнен',
    'Бэкап-листы удалены. Теперь можно:\n\n' +
    '1. Если есть лист "Сборы" — запустить Migrate to v2.0\n' +
    '2. Если есть лист "Цели" — запустить Setup/Rebuild structure\n',
    ui.ButtonSet.OK
  );
}

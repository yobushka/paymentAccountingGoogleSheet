/**
 * @fileoverview Скрытый лист Lists с метками для выпадающих списков
 */

/**
 * Настраивает скрытый лист Lists с формулами для меток
 */
function setupListsSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAMES.LISTS);
  
  // Полная очистка листа Lists для избежания конфликтов формул
  if (sh) {
    sh.clear();
    sh.clearFormats();
    // Удаляем валидации вручную для всего диапазона данных
    const lastRow = Math.max(sh.getMaxRows(), 100);
    const lastCol = Math.max(sh.getMaxColumns(), 15);
    sh.getRange(1, 1, lastRow, lastCol).clearDataValidations();
  } else {
    sh = ss.insertSheet(SHEET_NAMES.LISTS);
  }
  
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
  
  // G: BUDGET_ARTICLES — уникальные статьи из Сметы (пропускаем E-F для избежания конфликта)
  const shB = ss.getSheetByName(SHEET_NAMES.BUDGET);
  if (shB) {
    sh.getRange('G1').setValue('BUDGET_ARTICLES');
    sh.getRange('G2').setFormula(
      `=IFERROR(UNIQUE(FILTER(Смета!A2:A, LEN(Смета!A2:A))),)`
    );
    
    // I: BUDGET_SUBARTICLES — уникальные подстатьи из Сметы (пропускаем H)
    sh.getRange('I1').setValue('BUDGET_SUBARTICLES');
    sh.getRange('I2').setFormula(
      `=IFERROR(UNIQUE(FILTER(Смета!B2:B, LEN(Смета!B2:B))),)`
    );
  }
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

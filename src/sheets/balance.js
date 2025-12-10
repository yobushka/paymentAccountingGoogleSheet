/**
 * @fileoverview Лист «Баланс» — настройка и обновление формул
 */

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

/**
 * @fileoverview Ручной пересчёт и обновление данных
 */

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

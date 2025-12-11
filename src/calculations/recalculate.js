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
    
    // Обновляем детализацию (тикер обновляется внутри refreshDetailSheet_)
    refreshDetailSheet_();
    
    // Обновляем сводку (тикер обновляется внутри refreshSummarySheet_)
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

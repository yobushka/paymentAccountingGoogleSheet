/**
 * @fileoverview Загрузка демо-данных
 */

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

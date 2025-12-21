/**
 * @fileoverview Константы и конфигурация приложения
 * @version 2.0
 */

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
  VOLUNTARY: 'voluntary',
  FROM_BALANCE: 'from_balance'
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
  BUDGET: 'Смета',
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
  BUDGET_ARTICLES: 'BUDGET_ARTICLES',
  BUDGET_SUBARTICLES: 'BUDGET_SUBARTICLES',
  // v1.x legacy
  COLLECTIONS_LABELS: 'COLLECTIONS_LABELS',
  OPEN_COLLECTIONS_LABELS: 'OPEN_COLLECTIONS_LABELS'
};

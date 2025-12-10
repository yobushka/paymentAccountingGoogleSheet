#!/usr/bin/env node
/**
 * @fileoverview Build script — собирает все модули в один Code.gs
 * 
 * Запуск: node build.js
 * 
 * Порядок сборки важен для корректной работы:
 * 1. Константы и конфигурация
 * 2. Утилиты
 * 3. Вычисления
 * 4. Работа с листами
 * 5. Core-функции
 * 6. UI
 * 7. Триггеры
 * 8. Миграция
 */

const fs = require('fs');
const path = require('path');

// Порядок сборки модулей (важен для зависимостей)
const BUILD_ORDER = [
  // 1. Конфигурация — должна быть первой
  'src/config/constants.js',
  'src/config/sheets-spec.js',
  
  // 2. Утилиты — используются везде
  'src/utils/utils.js',
  
  // 3. Расчёты — чистые функции
  'src/calculations/dyn-cap.js',
  'src/calculations/custom-functions.js',
  'src/calculations/recalculate.js',
  
  // 4. Листы — зависят от config и utils
  'src/sheets/lists.js',
  'src/sheets/instruction.js',
  'src/sheets/balance.js',
  'src/sheets/detail.js',
  'src/sheets/summary.js',
  'src/sheets/issue-status.js',
  
  // 5. Core — основная логика
  'src/core/init.js',
  'src/core/validations.js',
  'src/core/id-generator.js',
  'src/core/close-goal.js',
  'src/core/sample-data.js',
  
  // 6. UI — меню, стили, диалоги
  'src/ui/menu.js',
  'src/ui/styles.js',
  'src/ui/dialogs.js',
  
  // 7. Триггеры — зависят от всего выше
  'src/triggers/on-edit.js',
  
  // 8. Миграция — опциональный модуль
  'src/migration/detect-version.js',
  'src/migration/migrate-v1-to-v2.js',
];

const OUTPUT_FILE = 'Code.gs';
const BACKUP_DIR = 'backups';

/**
 * Читает файл и возвращает его содержимое
 * @param {string} filePath
 * @return {string}
 */
function readFile(filePath) {
  const fullPath = path.join(__dirname, filePath);
  if (!fs.existsSync(fullPath)) {
    console.warn(`⚠️  Файл не найден: ${filePath}`);
    return '';
  }
  return fs.readFileSync(fullPath, 'utf-8');
}

/**
 * Создаёт бэкап текущего Code.gs
 */
function createBackup() {
  const codePath = path.join(__dirname, OUTPUT_FILE);
  if (!fs.existsSync(codePath)) {
    console.log('ℹ️  Code.gs не существует, бэкап не нужен');
    return;
  }
  
  const backupDir = path.join(__dirname, BACKUP_DIR);
  if (!fs.existsSync(backupDir)) {
    fs.mkdirSync(backupDir, { recursive: true });
  }
  
  const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
  const backupPath = path.join(backupDir, `Code.gs.${timestamp}.bak`);
  
  fs.copyFileSync(codePath, backupPath);
  console.log(`✅ Бэкап создан: ${backupPath}`);
}

/**
 * Удаляет @fileoverview JSDoc комментарии, но оставляет остальные
 * @param {string} content
 * @return {string}
 */
function stripFileOverview(content) {
  // Удаляем только @fileoverview блоки
  return content.replace(/\/\*\*[\s\S]*?@fileoverview[\s\S]*?\*\/\s*/g, '');
}

/**
 * Собирает все модули в один файл
 */
function build() {
  console.log('🔨 Сборка Code.gs...\n');
  
  // Создаём бэкап
  createBackup();
  
  // Заголовок файла
  const header = `/**
 * @fileoverview Payment Accounting for Google Sheets v2.0
 * 
 * Автоматически сгенерировано из модулей: ${new Date().toISOString()}
 * 
 * НЕ РЕДАКТИРУЙТЕ ЭТОТ ФАЙЛ НАПРЯМУЮ!
 * Вносите изменения в модули в папке src/ и запускайте build.js
 * 
 * Структура модулей:
 *   src/config/     — константы и спецификации листов
 *   src/utils/      — утилитарные функции
 *   src/calculations/ — расчётные функции
 *   src/sheets/     — настройка листов
 *   src/core/       — основная логика
 *   src/ui/         — меню, стили, диалоги
 *   src/triggers/   — обработчики событий
 *   src/migration/  — миграция v1 → v2
 */

`;
  
  const parts = [header];
  let totalLines = 0;
  
  // Собираем модули по порядку
  for (const modulePath of BUILD_ORDER) {
    const content = readFile(modulePath);
    if (!content) continue;
    
    // Считаем строки
    const lines = content.split('\n').length;
    totalLines += lines;
    
    // Добавляем разделитель
    const separator = `
// ${'='.repeat(70)}
// MODULE: ${modulePath}
// ${'='.repeat(70)}

`;
    
    // Убираем @fileoverview из модулей (оставляем остальные комментарии)
    const cleanContent = stripFileOverview(content);
    
    parts.push(separator);
    parts.push(cleanContent);
    
    console.log(`  ✓ ${modulePath} (${lines} строк)`);
  }
  
  // Записываем результат
  const result = parts.join('');
  const outputPath = path.join(__dirname, OUTPUT_FILE);
  fs.writeFileSync(outputPath, result, 'utf-8');
  
  const finalLines = result.split('\n').length;
  
  console.log(`
✅ Сборка завершена!
   📄 Файл: ${OUTPUT_FILE}
   📊 Модулей: ${BUILD_ORDER.length}
   📏 Строк: ${finalLines}
`);
}

/**
 * Проверяет что все модули существуют
 */
function validate() {
  console.log('🔍 Проверка модулей...\n');
  
  let allExist = true;
  
  for (const modulePath of BUILD_ORDER) {
    const fullPath = path.join(__dirname, modulePath);
    const exists = fs.existsSync(fullPath);
    
    if (exists) {
      console.log(`  ✓ ${modulePath}`);
    } else {
      console.log(`  ✗ ${modulePath} — НЕ НАЙДЕН`);
      allExist = false;
    }
  }
  
  console.log('');
  
  if (allExist) {
    console.log('✅ Все модули найдены');
  } else {
    console.log('❌ Некоторые модули отсутствуют');
    process.exit(1);
  }
}

/**
 * Показывает справку
 */
function showHelp() {
  console.log(`
Payment Accounting Build Script

Использование:
  node build.js [команда]

Команды:
  build     Собрать Code.gs из модулей (по умолчанию)
  validate  Проверить наличие всех модулей
  help      Показать эту справку

Примеры:
  node build.js
  node build.js build
  node build.js validate
`);
}

// Точка входа
const command = process.argv[2] || 'build';

switch (command) {
  case 'build':
    validate();
    build();
    break;
  case 'validate':
    validate();
    break;
  case 'help':
  case '--help':
  case '-h':
    showHelp();
    break;
  default:
    console.error(`Неизвестная команда: ${command}`);
    showHelp();
    process.exit(1);
}

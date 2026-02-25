/**
 * CONFIGURATION / НАСТРОЙКИ
 * Замените ссылки на ваши данные перед использованием.
 */
const CONFIG = {
  // Вставьте ссылку на вашу таблицу
  SPREADSHEET_URL: 'https://docs.google.com/spreadsheets/d/ID_ВАШЕЙ_ТАБЛИЦЫ/edit',
  
  // Вставьте ссылку на папку с файлами
  FOLDER_URL: 'https://drive.google.com/drive/folders/ID_ВАШЕЙ_ПАПКИ',
  
  // Список расширений для автоматической проверки
  EXTENSIONS: ['pdf', 'doc', 'docx', 'xls', 'xlsx']
};

/**
 * Основная функция переименования
 */
function renameFiles() {
  try {
    // 1. Извлечение ID из ссылок
    const spreadsheetMatch = CONFIG.SPREADSHEET_URL.match(/\/d\/([a-zA-Z0-9-_]+)/);
    const folderMatch = CONFIG.FOLDER_URL.match(/folders\/([a-zA-Z0-9-_]+)/);

    if (!spreadsheetMatch || !folderMatch) {
      throw new Error('Не удалось извлечь ID из ссылок. Проверьте CONFIG.');
    }

    const spreadsheetId = spreadsheetMatch[1];
    const folderId = folderMatch[1];

    Logger.log('=== НАЧАЛО РАБОТЫ ===');
    
    // 2. Сбор данных из таблицы
    const sheet = SpreadsheetApp.openById(spreadsheetId).getActiveSheet();
    const data = sheet.getDataRange().getValues();
    const fileData = data.slice(1); // Пропускаем заголовок

    // 3. Подключение к папке
    const folder = DriveApp.getFolderById(folderId);
    Logger.log(`Папка: "${folder.getName()}"`);

    // --- ОПТИМИЗАЦИЯ: Кэшируем файлы папки ---
    // Это ускоряет работу, если файлов много
    Logger.log('[1/4] Сканирование папки...');
    const allFiles = folder.getFiles();
    const cachedFiles = [];
    while (allFiles.hasNext()) {
      cachedFiles.push(allFiles.next());
    }

    let success = 0, errors = 0, notFound = 0;

    // --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---
    const normalizeName = (name) => {
      const datePattern = /(\d{1,2})[\.\s](\d{1,2})[\.\s](\d{2,4})/g;
      return name.replace(datePattern, '$1.$2.$3')
                 .toLowerCase()
                 .replace(/[.,;:!?"'`-]/g, '')
                 .replace(/\s+/g, ' ')
                 .trim();
    };

    const getBaseName = (name) => {
      const lastDotIndex = name.lastIndexOf('.');
      if (lastDotIndex === -1) return name;
      const potentialExt = name.slice(lastDotIndex + 1).toLowerCase();
      return CONFIG.EXTENSIONS.includes(potentialExt) ? name.slice(0, lastDotIndex) : name;
    };

    // 4. Обработка каждой строки таблицы
    Logger.log('[2/4] Обработка строк...');
    
    fileData.forEach((row, index) => {
      let original = (row[0] || '').toString().trim();
      let target = (row[1] || '').toString().trim();

      if (!original || !target) return;

      Logger.log(`\n● Строка ${index + 2}: "${original}" → "${target}"`);
      let fileFound = false;

      // ЭТАП 1: Точное совпадение
      const exactMatch = cachedFiles.find(f => f.getName() === original);
      
      // Попробуем также проверить варианты с расширениями, если в таблице указано имя без них
      let variantsToTry = [original];
      CONFIG.EXTENSIONS.forEach(ext => variantsToTry.push(`${getBaseName(original)}.${ext}`));

      for (let variant of variantsToTry) {
        const file = cachedFiles.find(f => f.getName() === variant);
        if (file) {
          try {
            if (file.getName() !== target) {
              file.setName(target);
              Logger.log(`  ✓ Успешно: "${variant}" → "${target}"`);
            } else {
              Logger.log(`  ⏩ Уже переименовано`);
            }
            success++;
            fileFound = true;
            break;
          } catch (e) {
            Logger.log(`  ✗ Ошибка: ${e.message}`);
            errors++;
            fileFound = true; // Считаем найденным, но с ошибкой доступа
            break;
          }
        }
      }

      // ЭТАП 2: Нормализованный поиск
      if (!fileFound) {
        const normOriginal = normalizeName(getBaseName(original));
        const fuzzyMatch = cachedFiles.find(f => normalizeName(getBaseName(f.getName())) === normOriginal);

        if (fuzzyMatch) {
          try {
            const oldName = fuzzyMatch.getName();
            fuzzyMatch.setName(target);
            Logger.log(`  ✓ Найдено по сходству: "${oldName}" → "${target}"`);
            success++;
            fileFound = true;
          } catch (e) {
            Logger.log(`  ✗ Ошибка: ${e.message}`);
            errors++;
          }
        }
      }

      if (!fileFound) {
        Logger.log('  ⚠️ Файл не найден в папке');
        notFound++;
      }
    });

    // Итоговый отчет
    Logger.log('\n=== РЕЗУЛЬТАТЫ ===');
    Logger.log(`Успешно: ${success}`);
    Logger.log(`Ошибки: ${errors}`);
    Logger.log(`Не найдено: ${notFound}`);
    Logger.log('=== ОБРАБОТКА ЗАВЕРШЕНА ===');

  } catch (e) {
    Logger.log('!!! КРИТИЧЕСКАЯ ОШИБКА !!!\n' + e.toString());
  }
}

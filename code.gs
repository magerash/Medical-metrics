/**
 * Основной файл Google Apps Script для работы с медицинскими данными
 * Файл: Code.gs
 */

/**
 * Получить уникальные категории из таблицы База_показателей
 */
function getCategories() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Пробуем разные варианты именованных диапазонов
    let namedRange = ss.getRangeByName("База_показателей[Категория]");
    
    // Если не найден, пробуем альтернативные варианты
    if (!namedRange) {
      // Попробуем найти таблицу База_показателей и взять столбец Категория
      const baseRange = ss.getRangeByName("База_показателей");
      if (baseRange) {
        const data = baseRange.getValues();
        const headers = data[0];
        const categoryIndex = headers.indexOf('Категория');
        
        if (categoryIndex !== -1) {
          const categories = [];
          for (let i = 1; i < data.length; i++) {
            if (data[i][categoryIndex] && data[i][categoryIndex] !== "") {
              categories.push(data[i][categoryIndex]);
            }
          }
          const uniqueCategories = [...new Set(categories)];
          console.log('Найдено категорий: ' + uniqueCategories.length);
          return uniqueCategories.sort();
        }
      }
      
      console.log('Именованный диапазон База_показателей[Категория] или База_показателей не найден');
      return ['Анализы', 'УЗИ', 'Рентген', 'МРТ', 'КТ', 'ЭКГ']; // Возвращаем примеры для тестирования
    }
    
    const categories = namedRange.getValues();
    const uniqueCategories = [...new Set(categories.flat().filter(cat => cat !== ""))];
    
    return uniqueCategories.sort();
  } catch (error) {
    console.log('Ошибка при получении категорий: ' + error.toString());
    return ['Анализы', 'УЗИ', 'Рентген']; // Возвращаем примеры для тестирования
  }
}

/**
 * Получить виды материалов из таблицы База_показателей
 */
function getMaterials() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Пробуем разные варианты именованных диапазонов
    let namedRange = ss.getRangeByName("База_показателей[Вид материала]");
    
    if (!namedRange) {
      // Попробуем найти таблицу База_показателей и взять столбец Вид материала
      const baseRange = ss.getRangeByName("База_показателей");
      if (baseRange) {
        const data = baseRange.getValues();
        const headers = data[0];
        const materialIndex = headers.indexOf('Вид материала');
        
        if (materialIndex !== -1) {
          const materials = [];
          for (let i = 1; i < data.length; i++) {
            if (data[i][materialIndex] && data[i][materialIndex] !== "") {
              materials.push(data[i][materialIndex]);
            }
          }
          const uniqueMaterials = [...new Set(materials)];
          console.log('Найдено материалов: ' + uniqueMaterials.length);
          return uniqueMaterials.sort();
        }
      }
      
      console.log('Именованный диапазон База_показателей[Вид материала] не найден');
      return ['Кровь', 'Моча', 'Кал', 'Мокрота', 'Слюна']; // Возвращаем примеры для тестирования
    }
    
    const materials = namedRange.getValues();
    const uniqueMaterials = [...new Set(materials.flat().filter(mat => mat !== ""))];
    
    return uniqueMaterials.sort();
  } catch (error) {
    console.log('Ошибка при получении материалов: ' + error.toString());
    return ['Кровь', 'Моча', 'Кал']; // Возвращаем примеры для тестирования
  }
}

/**
 * Получить показатели, соответствующие выбранной категории и материалу
 */
function getIndicators(category, material) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const baseRange = ss.getRangeByName("База_показателей");
    
    if (!baseRange) {
      console.log('Именованный диапазон База_показателей не найден');
      return [];
    }
    
    const data = baseRange.getValues();
    const headers = data[0];
    
    // Находим индексы нужных столбцов
    const categoryIndex = headers.indexOf('Категория');
    const materialIndex = headers.indexOf('Вид материала');
    const indicatorIndex = headers.indexOf('Показатель');
    
    if (categoryIndex === -1 || materialIndex === -1 || indicatorIndex === -1) {
      console.log('Не найдены необходимые столбцы в База_показателей');
      return [];
    }
    
    // Фильтруем показатели по категории и материалу
    const filteredIndicators = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[categoryIndex] === category && row[materialIndex] === material && row[indicatorIndex]) {
        if (!filteredIndicators.includes(row[indicatorIndex])) {
          filteredIndicators.push(row[indicatorIndex]);
        }
      }
    }
    
    return filteredIndicators.sort();
  } catch (error) {
    console.log('Ошибка при получении показателей: ' + error.toString());
    return [];
  }
}

/**
 * Получить информацию о показателе из базы (Min, Max, Ед. изм.)
 */
function getIndicatorInfo(category, material, indicator) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const baseRange = ss.getRangeByName("База_показателей");
    
    if (!baseRange) {
      return { min: '', max: '', unit: '' };
    }
    
    const data = baseRange.getValues();
    const headers = data[0];
    
    // Находим индексы столбцов
    const categoryIndex = headers.indexOf('Категория');
    const materialIndex = headers.indexOf('Вид материала');
    const indicatorIndex = headers.indexOf('Показатель');
    const minIndex = headers.indexOf('Min');
    const maxIndex = headers.indexOf('Max');
    const unitIndex = headers.indexOf('Ед. изм.');
    
    // Ищем строку с нужным показателем
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[categoryIndex] === category && 
          row[materialIndex] === material && 
          row[indicatorIndex] === indicator) {
        return {
          min: row[minIndex] || '',
          max: row[maxIndex] || '',
          unit: row[unitIndex] || ''
        };
      }
    }
    
    return { min: '', max: '', unit: '' };
  } catch (error) {
    console.log('Ошибка при получении информации о показателе: ' + error.toString());
    return { min: '', max: '', unit: '' };
  }
}

/**
 * Получить список организаций
 */
function getOrganizations() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const namedRange = ss.getRangeByName("data_organisation");
    
    if (!namedRange) {
      console.log('Именованный диапазон data_organisation не найден');
      return [];
    }
    
    const organizations = namedRange.getValues();
    const uniqueOrganizations = [...new Set(organizations.flat().filter(org => org !== ""))];
    
    return uniqueOrganizations.sort();
  } catch (error) {
    console.log('Ошибка при получении организаций: ' + error.toString());
    return [];
  }
}

/**
 * Получить список врачей
 */
function getDoctors() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const namedRange = ss.getRangeByName("data_doctor");
    
    if (!namedRange) {
      console.log('Именованный диапазон data_doctor не найден');
      return [];
    }
    
    const doctors = namedRange.getValues();
    const uniqueDoctors = [...new Set(doctors.flat().filter(doc => doc !== ""))];
    
    return uniqueDoctors.sort();
  } catch (error) {
    console.log('Ошибка при получении врачей: ' + error.toString());
    return [];
  }
}

/**
 * Сохранить медицинские записи
 */
function saveMedicalRecords(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const medicalSheet = ss.getSheetByName("Медицинские_данные");
    const baseSheet = ss.getSheetByName("База_показателей");
    
    if (!medicalSheet || !baseSheet) {
      throw new Error('Необходимые листы не найдены');
    }
    
    // Получаем email текущего пользователя
    const userEmail = Session.getActiveUser().getEmail();
    
    // Сохраняем записи в Медицинские_данные
    const medicalRange = ss.getRangeByName("Медицинские_данные");
    const lastRow = medicalSheet.getLastRow();
    
    // Подготавливаем данные для записи
    const newRows = [];
    const newBaseRows = [];
    
    // Проверяем существующие показатели в базе
    const baseRange = ss.getRangeByName("База_показателей");
    const baseData = baseRange ? baseRange.getValues() : [];
    const baseHeaders = baseData.length > 0 ? baseData[0] : [];
    
    data.records.forEach(record => {
      // Добавляем запись в Медицинские_данные
      newRows.push([
        data.date,
        data.category,
        data.material,
        record.indicator,
        record.value,
        record.comment || '',
        data.organization || '',
        data.doctor || '',
        userEmail
      ]);
      
      // Проверяем, нужно ли добавить в База_показателей
      if (record.isNew) {
        let exists = false;
        for (let i = 1; i < baseData.length; i++) {
          if (baseData[i][0] === data.category &&
              baseData[i][1] === data.material &&
              baseData[i][2] === record.indicator) {
            exists = true;
            break;
          }
        }
        
        if (!exists) {
          newBaseRows.push([
            data.category,
            data.material,
            record.indicator,
            '', // Показатель-лат. (пустой)
            record.min || '',
            record.max || '',
            record.unit || ''
          ]);
        }
      }
    });
    
    // Записываем в Медицинские_данные
    if (newRows.length > 0) {
      const startRow = lastRow + 1;
      medicalSheet.getRange(startRow, 1, newRows.length, 9).setValues(newRows);
    }
    
    // Записываем новые показатели в База_показателей
    if (newBaseRows.length > 0) {
      const baseLastRow = baseSheet.getLastRow();
      baseSheet.getRange(baseLastRow + 1, 1, newBaseRows.length, 7).setValues(newBaseRows);
    }
    
    return {
      success: true,
      message: `✅ Успешно добавлено записей: ${newRows.length}`,
      recordsCount: newRows.length
    };
    
  } catch (error) {
    console.log('Ошибка при сохранении записей: ' + error.toString());
    return {
      success: false,
      message: 'Ошибка при сохранении: ' + error.toString()
    };
  }
}

/**
 * Включить файл HTML в основной файл (для работы с CSS и JS)
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Тестовая функция для проверки наличия именованных диапазонов
 */
function testNamedRanges() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const namedRanges = ss.getNamedRanges();
  
  console.log('Найдено именованных диапазонов: ' + namedRanges.length);
  
  namedRanges.forEach(function(namedRange) {
    console.log('Диапазон: ' + namedRange.getName());
  });
  
  // Проверяем конкретные диапазоны
  const rangesToCheck = [
    'База_показателей',
    'База_показателей[Категория]',
    'База_показателей[Вид материала]',
    'База_показателей[Показатель]',
    'Медицинские_данные',
    'data_doctor',
    'data_organisation'
  ];
  
  rangesToCheck.forEach(function(rangeName) {
    const range = ss.getRangeByName(rangeName);
    if (range) {
      console.log('✓ ' + rangeName + ' - найден');
    } else {
      console.log('✗ ' + rangeName + ' - НЕ НАЙДЕН');
    }
  });
}
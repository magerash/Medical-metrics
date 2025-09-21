/**
 * Основной файл Google Apps Script для работы с медицинскими данными
 */

// Глобальные переменные для кэширования данных
var cachedData = {
  baseData: null,
  baseHeaders: null,
  doctors: null,
  organizations: null
};

/**
 * Получить данные из Базы_показателей
 */
function getBaseData() {
  if (cachedData.baseData) return cachedData;
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    let baseSheet = null;
    
    // Ищем лист с данными показателей
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const data = sheet.getDataRange().getValues();
      
      // Проверяем, есть ли нужные заголовки в первых двух строках
      const headers = data[0];
      if (headers.includes('Категория') && headers.includes('Вид материала') && 
          headers.includes('Показатель')) {
        baseSheet = sheet;
        cachedData.baseData = data;
        cachedData.baseHeaders = headers;
        break;
      }
    }
    
    if (!baseSheet) {
      console.log('Лист с данными показателей не найден');
      // Создаем пустые данные для тестирования
      cachedData.baseData = [];
      cachedData.baseHeaders = ['Категория', 'Вид материала', 'Показатель', 'Min', 'Max', 'Ед. изм.'];
    }
    
    return cachedData;
  } catch (error) {
    console.log('Ошибка при получении данных: ' + error.toString());
    return { baseData: [], baseHeaders: [] };
  }
}

/**
 * Получить уникальные категории из таблицы База_показателей
 */
function getCategories() {
  try {
    const data = getBaseData();
    const headers = data.baseHeaders;
    const baseData = data.baseData;
    
    const categoryIndex = headers.indexOf('Категория');
    if (categoryIndex === -1) {
      console.log('Столбец "Категория" не найден');
      return ['Анализы', 'УЗИ', 'Рентген', 'МРТ', 'КТ', 'ЭКГ'];
    }
    
    const categories = [];
    for (let i = 1; i < baseData.length; i++) {
      if (baseData[i][categoryIndex] && baseData[i][categoryIndex] !== "") {
        categories.push(baseData[i][categoryIndex]);
      }
    }
    
    const uniqueCategories = [...new Set(categories)];
    console.log('Найдено категорий: ' + uniqueCategories.length);
    return uniqueCategories.sort();
  } catch (error) {
    console.log('Ошибка при получении категорий: ' + error.toString());
    return ['Анализы', 'УЗИ', 'Рентген', 'МРТ', 'КТ', 'ЭКГ'];
  }
}

/**
 * Получить виды материалов из таблицы База_показателей
 */
function getMaterials() {
  try {
    const data = getBaseData();
    const headers = data.baseHeaders;
    const baseData = data.baseData;
    
    const materialIndex = headers.indexOf('Вид материала');
    if (materialIndex === -1) {
      console.log('Столбец "Вид материала" не найден');
      return ['Кровь', 'Моча', 'Кал', 'Мокрота', 'Слюна'];
    }
    
    const materials = [];
    for (let i = 1; i < baseData.length; i++) {
      if (baseData[i][materialIndex] && baseData[i][materialIndex] !== "") {
        materials.push(baseData[i][materialIndex]);
      }
    }
    
    const uniqueMaterials = [...new Set(materials)];
    console.log('Найдено материалов: ' + uniqueMaterials.length);
    return uniqueMaterials.sort();
  } catch (error) {
    console.log('Ошибка при получении материалов: ' + error.toString());
    return ['Кровь', 'Моча', 'Кал', 'Мокрота', 'Слюна'];
  }
}

/**
 * Получить показатели, соответствующие выбранной категории и материалу
 */
function getIndicators(category, material) {
  try {
    const data = getBaseData();
    const headers = data.baseHeaders;
    const baseData = data.baseData;
    
    const categoryIndex = headers.indexOf('Категория');
    const materialIndex = headers.indexOf('Вид материала');
    const indicatorIndex = headers.indexOf('Показатель');
    
    if (categoryIndex === -1 || materialIndex === -1 || indicatorIndex === -1) {
      console.log('Не найдены необходимые столбцы в данных показателей');
      return [];
    }
    
    const filteredIndicators = [];
    for (let i = 1; i < baseData.length; i++) {
      const row = baseData[i];
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
    const data = getBaseData();
    const headers = data.baseHeaders;
    const baseData = data.baseData;
    
    const categoryIndex = headers.indexOf('Категория');
    const materialIndex = headers.indexOf('Вид материала');
    const indicatorIndex = headers.indexOf('Показатель');
    const minIndex = headers.indexOf('Min');
    const maxIndex = headers.indexOf('Max');
    const unitIndex = headers.indexOf('Ед. изм.');
    
    for (let i = 1; i < baseData.length; i++) {
      const row = baseData[i];
      if (row[categoryIndex] === category && 
          row[materialIndex] === material && 
          row[indicatorIndex] === indicator) {
        return {
          min: minIndex !== -1 ? (row[minIndex] || '') : '',
          max: maxIndex !== -1 ? (row[maxIndex] || '') : '',
          unit: unitIndex !== -1 ? (row[unitIndex] || '') : ''
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
    if (cachedData.organizations) return cachedData.organizations;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    let orgData = [];
    
    // Ищем данные организаций в разных местах
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const data = sheet.getDataRange().getValues();
      
      // Проверяем, есть ли данные организаций
      if (data.length > 0 && data[0].includes('Организация')) {
        const orgIndex = data[0].indexOf('Организация');
        for (let j = 1; j < data.length; j++) {
          if (data[j][orgIndex]) orgData.push(data[j][orgIndex]);
        }
      }
      
      // Проверяем именованные диапазоны
      const namedRanges = sheet.getNamedRanges();
      for (let j = 0; j < namedRanges.length; j++) {
        const range = namedRanges[j];
        if (range.getName().toLowerCase().includes('organisation') || 
            range.getName().toLowerCase().includes('организация')) {
          const values = range.getRange().getValues();
          orgData = orgData.concat(values.flat().filter(org => org !== ""));
        }
      }
    }
    
    const uniqueOrganizations = [...new Set(orgData)];
    cachedData.organizations = uniqueOrganizations.sort();
    return cachedData.organizations;
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
    if (cachedData.doctors) return cachedData.doctors;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    let doctorData = [];
    
    // Ищем данные врачей в разных местах
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const data = sheet.getDataRange().getValues();
      
      // Проверяем, есть ли данные врачей
      if (data.length > 0 && data[0].includes('Врач')) {
        const doctorIndex = data[0].indexOf('Врач');
        for (let j = 1; j < data.length; j++) {
          if (data[j][doctorIndex]) doctorData.push(data[j][doctorIndex]);
        }
      }
      
      // Проверяем именованные диапазоны
      const namedRanges = sheet.getNamedRanges();
      for (let j = 0; j < namedRanges.length; j++) {
        const range = namedRanges[j];
        if (range.getName().toLowerCase().includes('doctor') || 
            range.getName().toLowerCase().includes('врач')) {
          const values = range.getRange().getValues();
          doctorData = doctorData.concat(values.flat().filter(doc => doc !== ""));
        }
      }
    }
    
    const uniqueDoctors = [...new Set(doctorData)];
    cachedData.doctors = uniqueDoctors.sort();
    return cachedData.doctors;
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
    const sheets = ss.getSheets();
    
    // Ищем лист для медицинских данных
    let medicalSheet = null;
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const sheetData = sheet.getDataRange().getValues();
      if (sheetData.length > 0 && 
          sheetData[0].includes('Дата') && 
          sheetData[0].includes('Категория') && 
          sheetData[0].includes('Показатель')) {
        medicalSheet = sheet;
        break;
      }
    }
    
    if (!medicalSheet) {
      // Создаем новый лист, если не найден
      medicalSheet = ss.insertSheet('Медицинские_данные');
      medicalSheet.getRange(1, 1, 1, 9).setValues([
        ['Дата', 'Категория', 'Вид материала', 'Показатель', 'Значение', 'Комментарий', 'Организация', 'Врач', 'Пользователь']
      ]);
    }
    
    // Получаем все данные из листа
    const dataRange = medicalSheet.getDataRange();
    const values = dataRange.getValues();
    
    let headerRow = -1;
    let lastDataRow = -1;
    
    // Находим строку с заголовками
    for (let i = 0; i < values.length; i++) {
      if (values[i].includes('Дата') && values[i].includes('Категория')) {
        headerRow = i;
        
        // Находим последнюю заполненную строку таблицы
        for (let j = i + 1; j < values.length; j++) {
          if (values[j].some(cell => cell !== '')) {
            lastDataRow = j;
          } else {
            break; // Прекращаем, если встретили полностью пустую строку
          }
        }
        break;
      }
    }
    
    // Если не нашли заголовки, используем первую строку
    if (headerRow === -1) {
      headerRow = 0;
      lastDataRow = 0;
    }
    
    // Если нет данных, добавляем после заголовка
    if (lastDataRow === -1) {
      lastDataRow = headerRow;
    }
    
    // Получаем заголовки
    const headers = values[headerRow];
    
    // Определяем индексы колонок
    const dateCol = headers.indexOf('Дата') + 1;
    const categoryCol = headers.indexOf('Категория') + 1;
    const materialCol = headers.indexOf('Вид материала') + 1;
    const indicatorCol = headers.indexOf('Показатель') + 1;
    const valueCol = headers.indexOf('Значение') + 1;
    const commentCol = headers.indexOf('Комментарий') + 1;
    const organizationCol = headers.indexOf('Организация') + 1;
    const doctorCol = headers.indexOf('Врач') + 1;
    const userCol = headers.indexOf('Пользователь') + 1;
    
    const userEmail = Session.getActiveUser().getEmail();
    const baseData = getBaseData();
    
    const newBaseRows = [];
    
    // Для каждой записи добавляем новую строку
    data.records.forEach(record => {
      // Вставляем новую строку после последней строки с данными
      medicalSheet.insertRowAfter(lastDataRow + 1);
      lastDataRow++; // Увеличиваем счетчик строк
      
      // Заполняем данные в соответствующие колонки
      if (dateCol > 0) medicalSheet.getRange(lastDataRow + 1, dateCol).setValue(data.date);
      if (categoryCol > 0) medicalSheet.getRange(lastDataRow + 1, categoryCol).setValue(data.category);
      if (materialCol > 0) medicalSheet.getRange(lastDataRow + 1, materialCol).setValue(data.material);
      if (indicatorCol > 0) medicalSheet.getRange(lastDataRow + 1, indicatorCol).setValue(record.indicator);
      if (valueCol > 0) medicalSheet.getRange(lastDataRow + 1, valueCol).setValue(record.value);
      if (commentCol > 0) medicalSheet.getRange(lastDataRow + 1, commentCol).setValue(record.comment || '');
      if (organizationCol > 0) medicalSheet.getRange(lastDataRow + 1, organizationCol).setValue(data.organization || '');
      if (doctorCol > 0) medicalSheet.getRange(lastDataRow + 1, doctorCol).setValue(data.doctor || '');
      if (userCol > 0) medicalSheet.getRange(lastDataRow + 1, userCol).setValue(userEmail);
      
      // Проверяем, нужно ли добавить в Базу_показателей
      if (record.isNew) {
        let exists = false;
        for (let i = 1; i < baseData.baseData.length; i++) {
          const row = baseData.baseData[i];
          if (row[0] === data.category &&
              row[1] === data.material &&
              row[2] === record.indicator) {
            exists = true;
            break;
          }
        }
        
        if (!exists) {
          // Исправляем обработку нулевых значений
          const minValue = record.min === 0 ? 0 : (record.min || '');
          const maxValue = record.max === 0 ? 0 : (record.max || '');
          
          newBaseRows.push([
            data.category,
            data.material,
            record.indicator,
            '', // Показатель-лат. (пустой)
            minValue,
            maxValue,
            record.unit || ''
          ]);
        }
      }
    });
    
    // Сохраняем новые показатели в Базу_показателей
    if (newBaseRows.length > 0) {
      // Ищем лист для базовых показателей
      let baseSheet = null;
      for (let i = 0; i < sheets.length; i++) {
        const sheet = sheets[i];
        const sheetData = sheet.getDataRange().getValues();
        if (sheetData.length > 0 && 
            sheetData[0].includes('Категория') && 
            sheetData[0].includes('Вид материала') && 
            sheetData[0].includes('Показатель')) {
          baseSheet = sheet;
          break;
        }
      }
      
      if (!baseSheet) {
        // Создаем новый лист, если не найден
        baseSheet = ss.insertSheet('База_показателей');
        baseSheet.getRange(1, 1, 1, 7).setValues([
          ['Категория', 'Вид материала', 'Показатель', 'Показатель-лат.', 'Min', 'Max', 'Ед. изм.']
        ]);
      }
      
      // Получаем данные из базового листа
      const baseDataRange = baseSheet.getDataRange();
      const baseValues = baseDataRange.getValues();
      
      let baseHeaderRow = -1;
      let baseLastDataRow = -1;
      
      // Находим строку с заголовками в базе показателей
      for (let i = 0; i < baseValues.length; i++) {
        if (baseValues[i].includes('Категория') && baseValues[i].includes('Вид материала')) {
          baseHeaderRow = i;
          
          // Находим последнюю заполненную строку таблицы
          for (let j = i + 1; j < baseValues.length; j++) {
            if (baseValues[j].some(cell => cell !== '')) {
              baseLastDataRow = j;
            } else {
              break;
            }
          }
          break;
        }
      }
      
      // Если не нашли заголовки, используем первую строку
      if (baseHeaderRow === -1) {
        baseHeaderRow = 0;
        baseLastDataRow = 0;
      }
      
      // Если нет данных, добавляем после заголовка
      if (baseLastDataRow === -1) {
        baseLastDataRow = baseHeaderRow;
      }
      
      // Вставляем новые строки в базу показателей
      newBaseRows.forEach(rowData => {
        baseSheet.insertRowAfter(baseLastDataRow + 1);
        baseLastDataRow++;
        
        // Заполняем данные в соответствующие колонки
        const baseHeaders = baseValues[baseHeaderRow];
        const categoryCol = baseHeaders.indexOf('Категория') + 1;
        const materialCol = baseHeaders.indexOf('Вид материала') + 1;
        const indicatorCol = baseHeaders.indexOf('Показатель') + 1;
        const latinCol = baseHeaders.indexOf('Показатель-лат.') + 1;
        const minCol = baseHeaders.indexOf('Min') + 1;
        const maxCol = baseHeaders.indexOf('Max') + 1;
        const unitCol = baseHeaders.indexOf('Ед. изм.') + 1;
        
        if (categoryCol > 0) baseSheet.getRange(baseLastDataRow + 1, categoryCol).setValue(rowData[0]);
        if (materialCol > 0) baseSheet.getRange(baseLastDataRow + 1, materialCol).setValue(rowData[1]);
        if (indicatorCol > 0) baseSheet.getRange(baseLastDataRow + 1, indicatorCol).setValue(rowData[2]);
        if (latinCol > 0) baseSheet.getRange(baseLastDataRow + 1, latinCol).setValue(rowData[3]);
        if (minCol > 0) baseSheet.getRange(baseLastDataRow + 1, minCol).setValue(rowData[4]);
        if (maxCol > 0) baseSheet.getRange(baseLastDataRow + 1, maxCol).setValue(rowData[5]);
        if (unitCol > 0) baseSheet.getRange(baseLastDataRow + 1, unitCol).setValue(rowData[6]);
      });
    }
    
    // Сбрасываем кэш
    cachedData.baseData = null;
    
    return {
      success: true,
      message: `✅ Успешно добавлено записей: ${data.records.length}`,
      recordsCount: data.records.length
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
 * Тестовая функция для проверки структуры таблицы
 */
function testStructure() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  
  let message = 'Структура таблицы:\n\n';
  
  sheets.forEach(function(sheet) {
    message += 'Лист: ' + sheet.getName() + '\n';
    const data = sheet.getDataRange().getValues();
    
    if (data.length > 0) {
      message += 'Заголовки: ' + data[0].join(', ') + '\n';
      message += 'Количество строк: ' + data.length + '\n\n';
    } else {
      message += 'Пустой лист\n\n';
    }
  });
  
  // Проверяем именованные диапазоны
  const namedRanges = ss.getNamedRanges();
  message += 'Именованные диапазоны:\n';
  
  if (namedRanges.length > 0) {
    namedRanges.forEach(function(namedRange) {
      message += '- ' + namedRange.getName() + '\n';
    });
  } else {
    message += 'Именованные диапазоны не найдены\n';
  }
  
  SpreadsheetApp.getUi().alert(message);
}
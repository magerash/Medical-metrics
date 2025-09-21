/**
 * Функция создания меню при открытии таблицы
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏥 Медицинские данные')
    .addItem('📝 Добавить записи', 'showMedicalDialog')
    .addSeparator()
    // .addItem('📊 Статистика', 'showStatistics')
    .addToUi();
}

/**
 * Показать диалоговое окно для добавления медицинских записей
 */
function showMedicalDialog() {
  try {
    const htmlOutput = HtmlService.createTemplateFromFile('dialog')
      .evaluate()
      .setWidth(950)  // Увеличена ширина для таблицы
      .setHeight(700) // Оптимальная высота
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, '🏥 Добавить медицинские данные');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Ошибка при открытии диалога: ' + error.toString());
  }
}

/**
 * Показать статистику (заготовка для будущего функционала)
 */
function showStatistics() {
  SpreadsheetApp.getUi().alert('📊 Статистика', 
    'Функционал статистики находится в разработке.\n\n' +
    'Здесь будут доступны:\n' +
    '• График изменения показателей\n' +
    '• Сравнение с нормами\n' +
    '• Экспорт отчетов',
    SpreadsheetApp.getUi().ButtonSet.OK);
}
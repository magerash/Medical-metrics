/**
 * Функция создания меню при открытии таблицы
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🏥 Медицинские данные')
    .addItem('📝 Добавить записи', 'showMedicalDialog')
    .addSeparator()
    .addItem('🔍 Проверить структуру', 'testStructure')
    .addToUi();
}

/**
 * Показать диалоговое окно для добавления медицинских записей
 */
function showMedicalDialog() {
  try {
    const htmlOutput = HtmlService.createTemplateFromFile('dialog')
      .evaluate()
      .setWidth(950)
      .setHeight(700)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Добавить медицинские данные');
  } catch (error) {
    SpreadsheetApp.getUi().alert('Ошибка при открытии диалога: ' + error.toString());
  }
}
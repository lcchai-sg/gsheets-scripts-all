// This Google Apps Script imports specific columns from multiple external Google Sheets into the active sheet.
// It skips header rows, appends the source name (title) next to each value, and writes the last update timestamp in cell D1.

function importSpecificColumns() {
  const sources = [
    {
      id: '10lPwwwLWslnz8QE7ZKtGZNUIvktLamExRf6idPvrU6E', // ID таблиці з урл https://docs.google.com/spreadsheets/d/10lPwwwLWslnz8QE7ZKtGZNUIvktLamExRf6idPvrU6E
      sheetName: 'Аркуш1', // Назва аркуша
      columnLetter: 'B',    // Яку колонку витягнути
      title: 'selected donors' // Назва з якою буде пов'язуватись запис
    },
    {
      id: '1yQnpVzfDqUIKt3sJd7GYj5jUYhAki552oyIEYHi9lVI',
      sheetName: 'Donors Approve',
      columnLetter: 'C',
      title: 'Donors Approve'
    },
    {
      id: '1yQnpVzfDqUIKt3sJd7GYj5jUYhAki552oyIEYHi9lVI',
      sheetName: 'links 6 month',
      columnLetter: 'A',
      title: 'links 6 month'
    }
  ];

  // Відкрити активну таблицю
  const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = targetSpreadsheet.getSheetByName('Обрані колонки') || targetSpreadsheet.insertSheet('Обрані колонки'); // Створити або відкрити вказаний аркуш
  targetSheet.clearContents(); // Очистити аркуш

  targetSheet.getRange('A1').setValue('row domain');
  targetSheet.getRange('B1').setValue('list');

  let targetRow = 2;

  sources.forEach(source => {
    const sourceSpreadsheet = SpreadsheetApp.openById(source.id);
    const sourceSheet = sourceSpreadsheet.getSheetByName(source.sheetName);

    const colNumber = letterToColumn(source.columnLetter);

    const lastRow = sourceSheet.getLastRow();
    // Якщо є хоча б 2 рядки (тобто є що копіювати), копіюємо з 2 рядка (пропускаєм назву колонки)
    if (lastRow > 1) {
      let columnData = sourceSheet.getRange(2, colNumber, lastRow - 1).getValues(); // Починаємо з 2 рядка
      columnData = columnData.map(row => [String(row[0]).toLowerCase()]);

      const titlesArray = columnData.map(() => [source.title]);

      targetSheet.getRange(targetRow, 1, columnData.length, 1).setValues(columnData);
      targetSheet.getRange(targetRow, 2, titlesArray.length, 1).setValues(titlesArray);
      targetRow += columnData.length;
    }
  });
  // Записати дату й час оновлення в C1
  const now = new Date();
  targetSheet.getRange('C1').setValue(`Оновлено: ${now.toLocaleString()}`);
}

function letterToColumn(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column *= 26;
    column += letter.charCodeAt(i) - 'A'.charCodeAt(0) + 1;
  }
  return column;
}

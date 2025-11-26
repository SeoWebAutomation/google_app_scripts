function onEdit(e) {
  try {
    const sheet = e.source.getActiveSheet();
    const editedRange = e.range;

    // === Налаштування ===
    const targetSheet = "Аркуш1"; // назва аркуша
    const watchColumnName = "Status";     // назва колонки, яку відстежуємо
    const dateColumnName = "Date"; // назва колонки, куди ставимо дату
    const headerRowIndex = 1; // Номер рядку в якому заголовки

    if (sheet.getName() !== targetSheet) return;


    const headers = sheet.getRange(headerRowIndex, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(h => h ? String(h).trim() : "");

    const watchColIndex = headers.indexOf(watchColumnName) + 1;
    const dateColIndex = headers.indexOf(dateColumnName) + 1;

    if (watchColIndex < 1 || dateColIndex < 1) return;

    const startRow = editedRange.getRow();
    const startCol = editedRange.getColumn();
    const numRows = editedRange.getNumRows();
    const numCols = editedRange.getNumColumns();

    if (startRow + numRows - 1 <= headerRowIndex) return;

    // Якщо змінено діапазон і в нього входить колонка, що слідкується
    if (watchColIndex < startCol || watchColIndex > (startCol + numCols - 1)) {
      return;
    }

    // Для кожного редагованого рядка поставимо дату у колонку Дата Статусу
    const nowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");

    // Масово збираємо значення, щоб мінімізувати звернення до API
    const outRange = sheet.getRange(startRow, dateColIndex, numRows, 1);
    const outValues = outRange.getValues();

    for (let r = 0; r < numRows; r++) {
      const currentRow = startRow + r;
      if (currentRow <= headerRowIndex) continue;
      outValues[r][0] = nowStr;
    }

    outRange.setValues(outValues);

  } catch (err) {
    Logger.log("onEdit error: " + err);
  }
}

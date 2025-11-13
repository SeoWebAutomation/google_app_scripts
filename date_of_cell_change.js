function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const editedRange = e.range;

  // === Налаштування ===
  const targetSheet = "Аркуш1"; // назва аркуша
  const watchColumnName = "Status";     // назва колонки, яку відстежуємо
  const dateColumnName = "Date"; // назва колонки, куди ставимо дату

  if (sheet.getName() !== targetSheet) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const watchColIndex = headers.indexOf(watchColumnName) + 1;
  const dateColIndex = headers.indexOf(dateColumnName) + 1;

  if (watchColIndex === 0 || dateColIndex === 0) {
    Logger.log("Колонку не знайдено");
    return;
  }

  if (editedRange.getColumn() === watchColIndex) {
    const row = editedRange.getRow();
    if (row === 1) return;

    const dateCell = sheet.getRange(row, dateColIndex);
    const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd.MM.yyyy");
    dateCell.setValue(dateStr);
  }
}

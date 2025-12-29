function syncReadyForMove() {
  const SOURCE_SPREADSHEET_ID = "ID Табл. №1"; // ID Табл. №1
  const TARGET_SPREADSHEET_ID = "ID Табл. №1"; // ID Табл. №2

  const sourceSheetName = "назва аркушу в Табл. №1"; // назва аркушу в Табл. №1
  const targetSheetName = "назва аркушу в Табл. №2"; // назва аркушу в Табл. №2

  // Назви колонок
  const statusColumnName = "Назва колонки зі статусом"; // назва колонки зі статусом
  const uniqueColumnName = "Назва колонки з унікальним значенням"; // унікальний ідентифікатор

  // значення які задіяні в колонках і робиться перевірка по ним, якщо це значення є то буде переносити в табл2
  const statusCellValue = "ready"; // значення з колонки зі статусом

  // Які поля переносимо
  const columnsToCopy = [
    "Lang",
    "Topic",
    "Site",
  ];

  const FIRST_HEADER_ROW = 1; // заголовки в 1-му рядку
  const FIRST_DATA_ROW = 2;   // дані починаються з 2-го рядка

  const sourceSS = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  const targetSS = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);

  const sourceSheet = sourceSS.getSheetByName(sourceSheetName);
  const targetSheet = targetSS.getSheetByName(targetSheetName);

  // ---------- 1. Заголовки з Табл.1 ----------
  const lastColumnSource = sourceSheet.getLastColumn();
  const sourceHeaders = sourceSheet
    .getRange(FIRST_HEADER_ROW, 1, 1, lastColumnSource)
    .getValues()[0];

  const statusColumnIndexSource = sourceHeaders.indexOf(statusColumnName) + 1;
  const uniqueColumnIndexSource = sourceHeaders.indexOf(uniqueColumnName) + 1;

  if (statusColumnIndexSource === 0 || uniqueColumnIndexSource === 0) {
    Logger.log("Не знайдено колонку статусу в джерельному листі. Перевір назви заголовків.");
    Logger.log("sourceHeaders: " + JSON.stringify(sourceHeaders));
    return;
  }

  const sourceColIndexByName = {};
  sourceHeaders.forEach((name, idx) => {
    if (name) {
      sourceColIndexByName[name] = idx + 1;
    }
  });

  // ---------- 2. Дані з Табл.1 ----------
  const lastRowSource = sourceSheet.getLastRow();
  if (lastRowSource < FIRST_DATA_ROW) {
    Logger.log("У джерельному листі немає даних.");
    return;
  }

  const numRowsSource = lastRowSource - FIRST_DATA_ROW + 1;
  const sourceData = sourceSheet
    .getRange(FIRST_DATA_ROW, 1, numRowsSource, lastColumnSource)
    .getValues();

  // ---------- 3. Заголовки в Табл.2 ----------
  const lastRowTarget = targetSheet.getLastRow();
  const lastColumnTarget = targetSheet.getLastColumn();

  if (lastRowTarget < FIRST_HEADER_ROW) {
    Logger.log("У цільовому листі немає заголовків. Перевір, що заголовки в 2-му рядку.");
    return;
  }

  const targetHeaders = targetSheet
    .getRange(FIRST_HEADER_ROW, 1, 1, lastColumnTarget)
    .getValues()[0];

  const targetColIndexByName = {};
  targetHeaders.forEach((name, idx) => {
    if (name) {
      targetColIndexByName[name] = idx + 1;
    }
  });

  const uniqueColumnIndexTarget = targetHeaders.indexOf(uniqueColumnName) + 1;
  if (uniqueColumnIndexTarget === 0) {
    Logger.log("Не знайдено колонку. Перевір назви заголовків.");
    Logger.log("targetHeaders: " + JSON.stringify(targetHeaders));
    return;
  }

  let targetData = [];
  if (lastRowTarget >= FIRST_DATA_ROW) {
    const numRowsTarget = lastRowTarget - FIRST_DATA_ROW + 1;
    targetData = targetSheet
      .getRange(FIRST_DATA_ROW, 1, numRowsTarget, lastColumnTarget)
      .getValues();
  }

  const existingIds = new Set(
    targetData
      .map(row => row[uniqueColumnIndexTarget - 1])
      .filter(v => v !== "" && v !== null && v !== undefined)
  );

  // ---------- 4. Формуємо рядки для додавання ----------
  const rowsToAppend = [];

  for (let i = 0; i < sourceData.length; i++) {
    const row = sourceData[i];

    const status = row[statusColumnIndexSource - 1];
    const uniqueValue = row[uniqueColumnIndexSource - 1];

    const normalizedStatus = (status || "").toString().trim().toLowerCase();
    const normalizedAgency = (agency || "").toString().trim().toLowerCase();

    if (normalizedStatus === statusCellValue.toLowerCase() &&
        uniqueValue) {

      if (!existingIds.has(uniqueValue)) {
        const newRow = new Array(targetHeaders.length).fill("");

        columnsToCopy.forEach(colName => {
          const srcIdx = sourceColIndexByName[colName];
          const tgtIdx = targetColIndexByName[colName];

          if (srcIdx && tgtIdx) {
            newRow[tgtIdx - 1] = row[srcIdx - 1];
          }
        });

        rowsToAppend.push(newRow);
        existingIds.add(uniqueValue);
      }
    }
  }

  // ---------- 5. Запис в Табл.2 ----------
  if (rowsToAppend.length > 0) {
    const startRow = lastRowTarget + 1;
    targetSheet
      .getRange(startRow, 1, rowsToAppend.length, targetHeaders.length)
      .setValues(rowsToAppend);
  }

  Logger.log("Готово. Додано рядків: " + rowsToAppend.length);
}

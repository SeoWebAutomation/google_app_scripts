function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔍 Перегляд')
    .addItem('Показати повний вміст клітинки', 'showSelectedCellContent')
    .addToUi();
}

function showSelectedCellContent() {
  const range = SpreadsheetApp.getActiveRange();
  if (!range) {
    SpreadsheetApp.getUi().alert('Спочатку вибери клітинку.');
    return;
  }

  const sheet = range.getSheet();
  const content = range.getDisplayValue();

  if (!content) {
    SpreadsheetApp.getUi().alert('Ця клітинка порожня.');
    return;
  }

  const html = HtmlService.createHtmlOutput(`
    <div style="
      font-family:Arial, sans-serif;
      padding:15px;
      line-height:1.5;
      white-space:pre-wrap;
      word-wrap:break-word;
      max-height:500px;
      overflow:auto;
    ">
      <div>${escapeHtml(content)}</div>
      <hr>
    </div>
  `).setWidth(600).setHeight(400);

  SpreadsheetApp.getUi().showModalDialog(html, 'Повний вміст');
}

function escapeHtml(text) {
  return text
    .toString()
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Hostname/Root Extractor')
    .addItem('Витягнути хостнейм', 'extractHostnames')
    .addItem('Витягнути рут-домен', 'extractRootDomains')
    .addToUi();
}

function extractHostnames() {
  const ui = SpreadsheetApp.getUi();

  const urlColResp = ui.prompt('Введіть ЛІТЕРУ(ENG) колонки з URL (наприклад A):');
  const resultColResp = ui.prompt('Введіть ЛІТЕРУ(ENG) колонки для результату (наприклад B):');

  const urlColLetter = urlColResp.getResponseText().toUpperCase();
  const resultColLetter = resultColResp.getResponseText().toUpperCase();

  if (!urlColLetter.match(/^[A-Z]+$/) || !resultColLetter.match(/^[A-Z]+$/)) {
    ui.alert('Некоректна літера колонки.');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const urlCol = columnLetterToIndex(urlColLetter);
  const resultCol = columnLetterToIndex(resultColLetter);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('Немає даних для обробки.');
    return;
  }

  const urlValues = sheet.getRange(2, urlCol, lastRow - 1).getValues();

  for (let i = 0; i < urlValues.length; i++) {
    const url = urlValues[i][0];
    if (url) {
      const hostname = getHostname(url.trim());
      sheet.getRange(i + 2, resultCol).setValue(hostname);
    } else {
      sheet.getRange(i + 2, resultCol).setValue('');
    }
  }

  ui.alert('Готово! Hostname витягнено.');
}

function extractRootDomains() {
  const ui = SpreadsheetApp.getUi();

  // Отримуємо дату останнього оновлення
  const lastUpdate = PropertiesService.getScriptProperties().getProperty('PSL_LAST_UPDATE');
  let lastUpdateStr = lastUpdate ? Utilities.formatDate(new Date(lastUpdate), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss') : 'не оновлювалась';

  // Питаємо, чи оновити PSL
  const response = ui.alert(
    `Дата останнього оновлення Public Suffix List: ${lastUpdateStr}\n` +
    'Оновити список перед обробкою?',
    ui.ButtonSet.YES_NO
  );

  if (response == ui.Button.YES) {
    updatePSL();
  }

  const urlColResp = ui.prompt('Введіть ЛІТЕРУ(ENG) колонки з URL (наприклад A):');
  const resultColResp = ui.prompt('Введіть ЛІТЕРУ(ENG) колонки для результату (наприклад B):');

  const urlColLetter = urlColResp.getResponseText().toUpperCase();
  const resultColLetter = resultColResp.getResponseText().toUpperCase();

  if (!urlColLetter.match(/^[A-Z]+$/) || !resultColLetter.match(/^[A-Z]+$/)) {
    ui.alert('Некоректна літера колонки.');
    return;
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const urlCol = columnLetterToIndex(urlColLetter);
  const resultCol = columnLetterToIndex(resultColLetter);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    ui.alert('Немає даних для обробки.');
    return;
  }

  const urlValues = sheet.getRange(2, urlCol, lastRow - 1).getValues();

  let pslRaw = PropertiesService.getScriptProperties().getProperty('PSL');
  if (!pslRaw) {
    ui.alert('Public Suffix List не знайдено. Зараз буде оновлено.');
    updatePSL();
    pslRaw = PropertiesService.getScriptProperties().getProperty('PSL');
    if (!pslRaw) {
      ui.alert('Не вдалося отримати Public Suffix List. Припинення операції.');
      return;
    }
  }

  const pslList = parsePSL(pslRaw);

  for (let i = 0; i < urlValues.length; i++) {
    const url = urlValues[i][0];
    if (url) {
      const rootDomain = getRootDomainFromURL(url.trim(), pslList);
      sheet.getRange(i + 2, resultCol).setValue(rootDomain);
    } else {
      sheet.getRange(i + 2, resultCol).setValue('');
    }
  }

  ui.alert('Готово! Рут-домен витягнено.');
}

function updatePSL() {
  const ui = SpreadsheetApp.getUi();
  const url = 'https://publicsuffix.org/list/public_suffix_list.dat';
  try {
    const response = UrlFetchApp.fetch(url);
    const content = response.getContentText();

    PropertiesService.getScriptProperties().setProperty('PSL', content);
    PropertiesService.getScriptProperties().setProperty('PSL_LAST_UPDATE', new Date().toISOString());

    ui.alert('Public Suffix List успішно оновлено!');
  } catch (e) {
    ui.alert('Помилка оновлення PSL: ' + e.message);
  }
}

function columnLetterToIndex(letter) {
  let column = 0;
  for (let i = 0; i < letter.length; i++) {
    column *= 26;
    column += letter.charCodeAt(i) - 64;
  }
  return column;
}

function parsePSL(pslRaw) {
  return pslRaw
    .split('\n')
    .map(line => line.trim())
    .filter(line => line && !line.startsWith('//') && !line.startsWith('!'));
}

function getRootDomainFromURL(url, pslList) {
  try {
    const hostnameMatch = url.match(/^https?:\/\/([^\/?#]+)(?:[\/?#]|$)/i);
    if (!hostnameMatch) return 'Invalid URL';

    const hostname = hostnameMatch[1].toLowerCase();
    const parts = hostname.split('.');

    for (let i = 0; i < parts.length; i++) {
      const candidate = parts.slice(i).join('.');
      if (pslList.includes(candidate)) {
        if (i === 0) return hostname;
        return parts.slice(i - 1).join('.');
      }
    }
    if(parts.length >= 2) return parts.slice(-2).join('.');
    return hostname;
  } catch (e) {
    return 'Invalid URL';
  }
}

function getHostname(url) {
  try {
    const match = url.match(/^https?:\/\/([^\/?#]+)(?:[\/?#]|$)/i);
    if (!match) return 'Invalid URL';

    let hostname = match[1].toLowerCase();
    if (hostname.startsWith('www.')) {
      hostname = hostname.substring(4);
    }
    return hostname;
  } catch {
    return 'Invalid URL';
  }
}

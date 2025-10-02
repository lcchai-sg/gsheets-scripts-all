function onOpen() {
  SpreadsheetApp.getUi().createMenu('Hostname/Root Extractor')
    .addItem('Витягнути хостнейм', 'extractHostnames')
    .addItem('Витягнути рут-домен', 'extractRootDomains')
    .addToUi();
}

function extractHostnames() {
  const { urlCol, resultCol, sheet, ui } = getColumnLetters();
  if (!urlCol) return;

  processURLs({
    sheet,
    urlCol,
    resultCol,
    transformFn: getHostname,
    message: '✅ Готово! Hostname витягнено.'
  });
}

function extractRootDomains() {
  const ui = SpreadsheetApp.getUi();

  // Перевірка на оновлення PSL
  const lastUpdate = PropertiesService.getScriptProperties().getProperty('PSL_LAST_UPDATE');
  const lastUpdateStr = lastUpdate
    ? Utilities.formatDate(new Date(lastUpdate), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss')
    : 'не оновлювалась';

  const response = ui.alert(
    `📅 Дата останнього оновлення Public Suffix List: ${lastUpdateStr}\n\n` +
    '🔄 Оновити список перед обробкою?\n\n' +
    '----------------------------------------------------------\n' +
    '⚠️ РЕКОМЕНДУЄМО ОНОВЛЮВАТИ PSL КОЖНІ 3 МІСЯЦІ!\n' +
    '----------------------------------------------------------',
    ui.ButtonSet.YES_NO
  );

  if (response === ui.Button.YES) {
    updatePSL();
  }

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

  const { urlCol, resultCol, sheet } = getColumnLetters();
  if (!urlCol) return;

  const pslList = parsePSL(pslRaw);

  processURLs({
    sheet,
    urlCol,
    resultCol,
    transformFn: (url) => getRootDomainFromURL(url, pslList),
    message: '✅ Готово! Рут-домен витягнено.'
  });
}

// ========================
// 🔧 Утиліти
// ========================

function getColumnLetters() {
  const ui = SpreadsheetApp.getUi();

  const urlColResp = ui.prompt('Введіть ЛІТЕРУ(ENG) колонки з URL (наприклад A):');
  const resultColResp = ui.prompt('Введіть ЛІТЕРУ(ENG) колонки для результату (наприклад B):');

  const urlColLetter = urlColResp.getResponseText().toUpperCase();
  const resultColLetter = resultColResp.getResponseText().toUpperCase();

  if (!urlColLetter.match(/^[A-Z]+$/) || !resultColLetter.match(/^[A-Z]+$/)) {
    ui.alert('Некоректна літера колонки.');
    return {};
  }

  if (urlColLetter === resultColLetter) {
    const overwriteResponse = ui.alert(
      `Ви вказали одну й ту саму літеру для вхідних і вихідних даних: "${urlColLetter}".\n\n` +
      'Це призведе до перезапису початкових URL-адрес. Ви впевнені, що хочете продовжити?',
      ui.ButtonSet.YES_NO
    );
    if (overwriteResponse === ui.Button.NO) return {};
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const urlCol = columnLetterToIndex(urlColLetter);
  const resultCol = columnLetterToIndex(resultColLetter);

  return { urlCol, resultCol, sheet, ui };
}

function processURLs({ sheet, urlCol, resultCol, transformFn, message }) {
  const ui = SpreadsheetApp.getUi();
  const colValues = sheet.getRange(2, urlCol, sheet.getMaxRows() - 1).getValues();
  const nonEmptyValues = colValues.filter(row => row[0] !== '');

  if (nonEmptyValues.length === 0) {
    ui.alert('⚠️error⚠️ Немає даних для обробки в обраній колонці.');
    return;
  }

  const urlValues = colValues.slice(0, nonEmptyValues.length);

  for (let i = 0; i < urlValues.length; i++) {
    const url = urlValues[i][0];
    const result = url ? transformFn(url.trim()) : '';
    sheet.getRange(i + 2, resultCol).setValue(result);
  }

  ui.alert(message);
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
    if (!/^https?:\/\//i.test(url)) {
      url = 'http://' + url;
    }

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
    if (parts.length >= 2) return parts.slice(-2).join('.');
    return hostname;
  } catch (e) {
    return 'Invalid URL';
  }
}

function getHostname(url) {
  try {
    if (!/^https?:\/\//i.test(url)) {
      url = 'http://' + url;
    }

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

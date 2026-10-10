/**
 * @OnlyCurrentDoc
 *
 * Приймає анкети з /vsl і дописує їх у таблицю «ЮНМ ЛМ ЗАЯВКИ (NEW)», перший аркуш:
 * A — Дата, B — Ім'я, C — Телефон, D — Telegram, E — Ніша,
 * F — Бажаний дохід з YouTube Shorts, G — Поточний дохід на роботі.
 *
 * Скрипт має бути прив'язаний до цієї таблиці (Розширення → Apps Script).
 * Завдяки @OnlyCurrentDoc Google просить доступ лише до неї, а не до всіх таблиць.
 */
var HEADER = ['Дата', "Ім'я", 'Телефон', 'Telegram', 'Ніша для каналу',
  'Бажаний дохід з YouTube Shorts', 'Поточний дохід на роботі'];

function doPost(e) {
  var p = (e && e.parameter) || {};
  // прибираємо символи, з яких Таблиці починають формулу
  var name = noFormula(clean(p.name, 80));
  var phone = clean(p.phone, 25).replace(/[^\d\s()+-]/g, '');
  var tg = clean(p.tg, 40);
  if (!/^@[A-Za-z0-9_]{5,32}$/.test(tg)) tg = '';
  var niche = noFormula(clean(p.niche, 60));
  var want = noFormula(clean(p.want, 60));
  var current = noFormula(clean(p.current, 60));

  if (!name || phone.replace(/\D/g, '').length < 10 || !niche || !want || !current) {
    return json({ result: 'error', error: 'invalid' });
  }

  var date = Utilities.formatDate(new Date(), 'Europe/Kyiv', 'dd.MM.yyyy HH:mm');

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(HEADER);
      sheet.getRange(1, 1, 1, HEADER.length).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }
    var row = sheet.getLastRow() + 1;
    // формат «текст», щоб «+380…» не перетворився на число чи формулу
    sheet.getRange(row, 1, 1, HEADER.length).setNumberFormat('@')
      .setValues([[date, name, phone, tg, niche, want, current]]);
  } finally {
    lock.releaseLock();
  }

  return json({ result: 'success' });
}

function clean(v, max) {
  return String(v || '').trim().slice(0, max);
}

function noFormula(v) {
  return v.replace(/^[=+\-@]+/, '');
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

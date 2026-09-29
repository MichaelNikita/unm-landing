/**
 * Приймає заявки з /zayavka і дописує їх у перший аркуш таблиці:
 * A — Ім'я, B — Номер, C — Нік у Telegram.
 *
 * Скрипт має бути прив'язаний до таблиці (Розширення → Apps Script),
 * тоді ID таблиці в коді не потрібен. Як задеплоїти — див. README.
 */
function doPost(e) {
  var p = (e && e.parameter) || {};
  // прибираємо символи, з яких Таблиці починають формулу
  var name = clean(p.name, 80).replace(/^[=+\-@]+/, '');
  var phone = clean(p.phone, 25).replace(/[^\d\s()+-]/g, '');
  var tg = clean(p.tg, 40);
  if (!/^@[A-Za-z0-9_]{5,32}$/.test(tg)) tg = '';

  if (!name || phone.replace(/\D/g, '').length < 10) {
    return json({ result: 'error', error: 'invalid' });
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var row = sheet.getLastRow() + 1;
    // формат «текст», щоб «+380…» не перетворився на число чи формулу
    sheet.getRange(row, 1, 1, 3).setNumberFormat('@').setValues([[name, phone, tg]]);
  } finally {
    lock.releaseLock();
  }

  return json({ result: 'success' });
}

function clean(v, max) {
  return String(v || '').trim().slice(0, max);
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

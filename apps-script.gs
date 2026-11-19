/**
 * Magazyn – zapis wydań do Arkusza Google.

 * Aplikacja internetowa, "Wykonuj jako: Ja", "Kto ma dostęp: Każdy".
 */
const WYDANIA = 'Wydania';
const TOWARY = 'Towary';

function sheet_(name, headers, textCol) {
  const ss = SpreadsheetApp.openById('1vW_0dEmE_S5252b9FtIoLIa_yzZM8P59wVFuWZXnCmo');
  let s = ss.getSheetByName(name);
  if (!s) {
    s = ss.insertSheet(name);
    s.appendRow(headers);
    s.setFrozenRows(1);
    if (textCol) s.getRange(1, textCol, s.getMaxRows(), 1).setNumberFormat('@');
  }
  return s;
}

function wydania_() {
  return sheet_(WYDANIA, ['Czas', 'Kod', 'Nazwa', 'Jednostka', 'Ilość', 'Kto', 'ID'], 2);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const r = JSON.parse(e.postData.contents);
    const s = wydania_();
    const last = s.getLastRow();
    if (last > 1) {
      const n = Math.min(300, last - 1);
      const ids = s.getRange(last - n + 1, 7, n, 1).getValues().flat();
      if (ids.indexOf(r.id) !== -1) return json_({ ok: true, dup: true });
    }
    s.appendRow([new Date(r.t), String(r.co), r.nazwa || '', r.jednostka, Number(r.ile), r.kto, r.id]);
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const s = wydania_();
  const last = s.getLastRow();
  let wpisy = [];
  if (last > 1) {
    const n = Math.min(30, last - 1);
    wpisy = s.getRange(last - n + 1, 1, n, 7).getValues().reverse().map(function (v) {
      return {
        czas: Utilities.formatDate(new Date(v[0]), 'Europe/Warsaw', 'dd.MM.yyyy, HH:mm:ss'),
        co: String(v[1]), nazwa: v[2], jednostka: v[3], ile: v[4], kto: v[5], id: v[6]
      };
    });
  }
  const t = sheet_(TOWARY, ['Kod kreskowy', 'Nazwa'], 1);
  const towary = {};
  if (t.getLastRow() > 1) {
    t.getRange(2, 1, t.getLastRow() - 1, 2).getValues().forEach(function (v) {
      if (v[0] !== '') towary[String(v[0]).trim()] = String(v[1]);
    });
  }
  return json_({ wpisy: wpisy, towary: towary });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

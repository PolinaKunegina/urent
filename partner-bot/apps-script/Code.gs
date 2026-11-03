/**
 * Партнёрская сеть: «Следующий шаг» сразу попадает в историю касаний.
 *
 * Как поставить (один раз, на всю таблицу):
 *   Расширения → Apps Script → стереть всё в окне → вставить этот код → 💾 Сохранить.
 * Больше ничего не нужно: скрипт срабатывает сам при каждой правке, у всех, кто редактирует.
 *
 * Правила:
 *  • вписали дату следующего шага → она сразу появляется в первой свободной паре
 *    «Дата N-го касания / Формат», «что» — в «Формат»;
 *  • поправили «что» у того же шага → обновляется та же пара;
 *  • поменяли дату, а прошлый шаг ещё не наступил (перенос встречи) → правится последняя пара;
 *  • поменяли дату, а прошлый шаг уже прошёл (касание состоялось) → добавляется новая пара;
 *  • пары закончились → бот добавит новые колонки в течение часа.
 */
var HEADER_ROW = 4;
var SERVICE = ['Оглавление', 'Инструкция', 'Сводка', 'Партнёры'];

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var name = sheet.getName();
  if (SERVICE.indexOf(name) >= 0 || name.charAt(0) === '_') return;

  var headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim().toLowerCase(); });
  var cDate = headers.indexOf('следующий шаг: дата') + 1;
  var cWhat = headers.indexOf('следующий шаг: что') + 1;
  if (!cDate || !cWhat) return;

  var c0 = e.range.getColumn(), c1 = e.range.getLastColumn();
  if (c1 < Math.min(cDate, cWhat) || c0 > Math.max(cDate, cWhat)) return;
  var editedDate = c0 <= cDate && cDate <= c1;

  var pairs = [];
  headers.forEach(function (h, i) { if (/^дата\s*\d+/.test(h)) pairs.push(i + 1); });
  if (!pairs.length) return;

  for (var r = Math.max(e.range.getRow(), HEADER_ROW + 1); r <= e.range.getLastRow(); r++) {
    copyStep_(sheet, r, cDate, cWhat, pairs, editedDate);
  }
}

function copyStep_(sheet, r, cDate, cWhat, pairs, editedDate) {
  var d = sheet.getRange(r, cDate).getValue();
  var what = sheet.getRange(r, cWhat).getValue();
  if (!(d instanceof Date)) return;                       // нет даты — нечего копировать

  var first = pairs[0], width = pairs[pairs.length - 1] + 2 - first;
  var hist = sheet.getRange(r, first, 1, width).getValues()[0];
  var last = -1;
  pairs.forEach(function (c, k) {
    var v = hist[c - first];
    if (v !== '' && v !== null) last = k;
  });

  if (last >= 0) {
    var col = pairs[last], ld = hist[col - first];
    var today = new Date(); today.setHours(0, 0, 0, 0);
    if (ld instanceof Date && ld.getTime() === d.getTime()) {     // тот же шаг — обновляем «что»
      sheet.getRange(r, col + 1).setValue(what);
      return;
    }
    if (editedDate && ld instanceof Date && ld.getTime() > today.getTime()) {
      sheet.getRange(r, col, 1, 2).setValues([[d, what]]);         // перенос будущей встречи
      return;
    }
  }
  var next = last + 1;
  if (next >= pairs.length) {
    SpreadsheetApp.getActive().toast('Все колонки касаний заняты — бот добавит новые в течение часа.',
                                     'Касания', 8);
    return;
  }
  sheet.getRange(r, pairs[next], 1, 2).setValues([[d, what]]);
}

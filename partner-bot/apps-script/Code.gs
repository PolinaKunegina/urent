/**
 * Партнёрская сеть: «Следующий шаг» сразу попадает в историю касаний — с минутами и этапом.
 *
 * Как поставить (один раз, на всю таблицу):
 *   Расширения → Apps Script → стереть всё в окне → вставить этот код → 💾 Сохранить.
 * Больше ничего не нужно: скрипт срабатывает сам при каждой правке, у всех, кто редактирует.
 *
 * Правила:
 *  • вписали дату следующего шага → в первую свободную тройку «Дата N-го касания /
 *    Формат (этап) / Мин» уходят дата, «что (этап)» и минуты;
 *  • поправили «что», минуты или этап у того же шага → обновляется та же тройка;
 *  • поменяли дату, а прошлый шаг ещё не наступил (перенос встречи) → правится последняя тройка;
 *  • поменяли дату, а прошлый шаг уже прошёл (касание состоялось) → добавляется новая тройка;
 *  • тройки закончились → бот добавит новые колонки в течение часа.
 * По «(этапу)» в истории таблица считает факт часов и людей в блоке «Мои часы по дням».
 */
var HEADER_ROW = 4;
var SERVICE = ['Оглавление', 'Инструкция', 'Сводка', 'Партнёры', 'Итоги недели'];

function onEdit(e) {
  if (!e || !e.range) return;
  var sheet = e.range.getSheet();
  var name = sheet.getName();
  if (SERVICE.indexOf(name) >= 0 || name.charAt(0) === '_') return;

  var headers = sheet.getRange(HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim().toLowerCase(); });
  var cols = {
    date: headers.indexOf('следующий шаг: дата') + 1,
    what: headers.indexOf('следующий шаг: что') + 1,
    min: headers.indexOf('время, мин') + 1,
    stage: headers.indexOf('этап') + 1
  };
  if (!cols.date || !cols.what) return;

  var watched = [cols.date, cols.what, cols.min, cols.stage].filter(function (c) { return c > 0; });
  var c0 = e.range.getColumn(), c1 = e.range.getLastColumn();
  if (!watched.some(function (c) { return c0 <= c && c <= c1; })) return;
  var editedDate = c0 <= cols.date && cols.date <= c1;

  var touches = [];   // [{d, f, m}] — номера колонок (с 1), f/m = 0, если колонки нет
  headers.forEach(function (h, i) {
    if (!/^дата\s*\d+/.test(h)) return;
    var f = (headers[i + 1] || '').indexOf('формат') === 0 ? i + 2 : 0;
    var m = f && (headers[i + 2] || '').indexOf('мин') === 0 ? i + 3 : 0;
    touches.push({ d: i + 1, f: f, m: m });
  });
  if (!touches.length) return;

  for (var r = Math.max(e.range.getRow(), HEADER_ROW + 1); r <= e.range.getLastRow(); r++) {
    copyStep_(sheet, r, cols, touches, editedDate);
  }
}

function copyStep_(sheet, r, cols, touches, editedDate) {
  var d = sheet.getRange(r, cols.date).getValue();
  if (!(d instanceof Date)) return;                       // нет даты — нечего копировать
  var what = String(sheet.getRange(r, cols.what).getValue() || '').trim();
  var stage = cols.stage ? String(sheet.getRange(r, cols.stage).getValue() || '').trim() : '';
  var min = cols.min ? sheet.getRange(r, cols.min).getValue() : '';
  var fmt = what + (stage ? (what ? ' ' : '') + '(' + stage.toLowerCase() + ')' : '');

  var first = touches[0].d, lastT = touches[touches.length - 1];
  var width = Math.max(lastT.d, lastT.f, lastT.m) - first + 1;
  var hist = sheet.getRange(r, first, 1, width).getValues()[0];
  var last = -1;
  touches.forEach(function (t, k) {
    var v = hist[t.d - first];
    if (v !== '' && v !== null) last = k;
  });

  if (last >= 0) {
    var t = touches[last], ld = hist[t.d - first];
    var today = new Date(); today.setHours(0, 0, 0, 0);
    if (ld instanceof Date && ld.getTime() === d.getTime()) {     // тот же шаг — обновляем «что», минуты
      write_(sheet, r, t, d, fmt, min);
      return;
    }
    if (editedDate && ld instanceof Date && ld.getTime() > today.getTime()) {
      write_(sheet, r, t, d, fmt, min);                            // перенос будущей встречи
      return;
    }
  }
  var next = last + 1;
  if (next >= touches.length) {
    SpreadsheetApp.getActive().toast('Все колонки касаний заняты — бот добавит новые в течение часа.',
                                     'Касания', 8);
    return;
  }
  write_(sheet, r, touches[next], d, fmt, min);
}

function write_(sheet, r, t, d, fmt, min) {
  if (t.m && t.m === t.d + 2) {
    sheet.getRange(r, t.d, 1, 3).setValues([[d, fmt, min === null ? '' : min]]);
  } else if (t.f) {
    sheet.getRange(r, t.d, 1, 2).setValues([[d, fmt]]);
  } else {
    sheet.getRange(r, t.d).setValue(d);
  }
}

"""Результаты месяца: бот засчитывает смену статуса клиента (стал партнёром, оплатил 95/240,
интенсив 5000) и пишет счётчики в блок «Мой месяц» на вкладке партнёра.

В первый раз бот только запоминает статусы (то, что было до запуска, не засчитывается).
Первого числа цели прошлого месяца уходят в архив бота, а строка целей очищается.
"""
from __future__ import annotations

import logging
from datetime import date

from .store import Store
from .workbook import MONTHS_NOM, RESULT_CATS, PartnerTab, col_letter, result_cats

log = logging.getLogger(__name__)


def track(store: Store, book: str, tabs: list[PartnerTab], today: date) -> list[tuple]:
    """Находит новые результаты. Возвращает [(вкладка, клиент, категория, день)]."""
    prev = store.seen_statuses(book)
    first_run = not prev
    seen, new = {}, []
    for t in tabs:
        for c in t.clients:
            key = (t.name, c.name)
            seen[key] = c.status or ""
            if first_run or key not in prev:
                continue          # статус на момент запуска или нового клиента — не результат
            for cat in sorted(result_cats(c.status) - result_cats(prev[key])):
                new.append((t.name, c.name, cat, today.isoformat()))
    # новый клиент, сразу записанный с оплатой, — тоже результат (кроме самого первого запуска)
    if not first_run:
        for t in tabs:
            for c in t.clients:
                if (t.name, c.name) not in prev:
                    new += [(t.name, c.name, cat, today.isoformat()) for cat in sorted(result_cats(c.status))]
    store.add_results(book, new)
    store.save_statuses(book, seen)
    return new


def write_month(sheet, store: Store, book: str, tabs: list[PartnerTab], today: date) -> None:
    """Переносит счётчики в блок «Мой месяц»; в новом месяце архивирует и очищает цели."""
    month = today.strftime("%Y-%m")
    last = store.get_setting(f"goals_month:{book}")
    rollover = last is not None and last != month
    if rollover:
        store.archive_goals(book, last, [(t.name, k, t.goals.get(k, 0), t.results.get(k, 0))
                                         for t in tabs for k, _, _ in RESULT_CATS])
    counts = store.month_results(book, month)
    data = []
    for t in tabs:
        if t.month_col is None:
            continue
        c0 = t.month_col
        data.append({"range": f"'{t.name}'!{col_letter(c0)}1",
                     "values": [[f"Мой месяц: {MONTHS_NOM[today.month - 1]}"]]})
        data.append({"range": f"'{t.name}'!{col_letter(c0 + 1)}3:{col_letter(c0 + len(RESULT_CATS))}3",
                     "values": [[counts.get((t.name, k), 0) for k, _, _ in RESULT_CATS]]})
        if rollover:
            data.append({"range": f"'{t.name}'!{col_letter(c0 + 1)}2:{col_letter(c0 + len(RESULT_CATS))}2",
                         "values": [[""] * len(RESULT_CATS)]})
    if data:
        r = sheet.s.post(f"https://sheets.googleapis.com/v4/spreadsheets/{sheet.id}/values:batchUpdate",
                         json={"valueInputOption": "RAW", "data": data}, timeout=60)
        r.raise_for_status()
    store.set_setting(f"goals_month:{book}", month)
    for t in tabs:                   # чтобы отчёты сразу видели свежие цифры
        t.results = {k: counts.get((t.name, k), 0) for k, _, _ in RESULT_CATS}
        if rollover:
            t.goals = {}

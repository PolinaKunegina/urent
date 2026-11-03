"""Перенос прошедшего «следующего шага» в историю касаний.

Партнёру не нужно ничего копировать: он просто вписывает новый следующий шаг.
Бот раз в 10 минут сравнивает шаги с тем, что видел в прошлый раз. Если шаг клиента
сменился (или его стёрли), а у старого шага дата уже наступила, — касание состоялось:
старые дата и «что» уходят в первую свободную пару «Дата N-го касания / Формат».
Если дата старого шага ещё впереди — это перенос встречи, в историю ничего не пишется.
"""
from __future__ import annotations

import logging
from dataclasses import dataclass
from datetime import date

from .workbook import TOUCH_COLS, PartnerTab, col_letter

log = logging.getLogger(__name__)


@dataclass
class Move:
    tab: str
    client: str
    row: int
    n: int
    when: date
    what: str


def plan_moves(tabs: list[PartnerTab], previous: dict[tuple[str, str], tuple[str, str]],
               today: date) -> tuple[list[Move], dict[tuple[str, str], tuple[str, str]], list[str]]:
    """(что перенести, новое состояние шагов, предупреждения)."""
    moves, state, warnings = [], {}, []
    for t in tabs:
        for c in t.clients:
            key = (t.name, c.name)
            cur = (c.next_date.isoformat() if c.next_date else "", c.next_what)
            state[key] = cur
            old = previous.get(key)
            if old is None or old == cur or not old[0]:
                continue
            old_date = date.fromisoformat(old[0])
            if old_date > today:
                continue                       # перенос будущей встречи — не касание
            if any(x.when == old_date for x in c.touches):
                continue                       # уже записано вручную
            used = {x.n for x in c.touches}
            free = next((n for n, _, _ in TOUCH_COLS if n not in used), None)
            if free is None:
                warnings.append(f"{t.name}: у «{c.name}» заняты все 5 касаний, {old_date:%d.%m} не записала")
                continue
            moves.append(Move(t.name, c.name, c.row, free, old_date, old[1] or "касание"))
    return moves, state, warnings


def apply_moves(sheet, moves: list[Move]) -> None:
    if not moves:
        return
    data = []
    for m in moves:
        _, d_idx, f_idx = next(tc for tc in TOUCH_COLS if tc[0] == m.n)
        data.append({"range": f"'{m.tab}'!{col_letter(d_idx)}{m.row}", "values": [[m.when.isoformat()]]})
        data.append({"range": f"'{m.tab}'!{col_letter(f_idx)}{m.row}", "values": [[m.what]]})
    r = sheet.s.post(f"https://sheets.googleapis.com/v4/spreadsheets/{sheet.id}/values:batchUpdate",
                     json={"valueInputOption": "USER_ENTERED", "data": data}, timeout=60)
    r.raise_for_status()
    log.info("Перенесено в историю касаний: %d", len(moves))

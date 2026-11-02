"""Превращает разобранный факт партнёра в изменения ячеек таблицы воронки.

Правила (как договорились на встрече):
- статус и «что происходит» берутся из сообщения; в описание ставится дата, чтобы
  было видно, когда обновлено;
- касание сегодня записывается в ближайшую свободную пару «Дата N-го касания / Формат»
  (если на сегодня касание уже было запланировано — только дописывается формат);
- дата следующего касания ставится в следующую свободную колонку даты; если все пять
  заняты — бот говорит об этом, а дату дописывает в описание.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date

from .sheets import Funnel, Lead, to_date


@dataclass
class Cell:
    row: int
    col: int      # 0-based
    value: str


@dataclass
class LeadChange:
    name: str
    cells: list[Cell] = field(default_factory=list)
    summary: list[str] = field(default_factory=list)   # для ответа партнёру
    warnings: list[str] = field(default_factory=list)


def _fmt(d: date) -> str:
    return d.strftime("%d.%m.%Y")


def _description(today: date, upd: dict) -> str | None:
    parts = []
    if upd.get("description"):
        parts.append(upd["description"].strip())
    if upd.get("no_next_touch_reason"):
        parts.append(f"нет след. касания: {upd['no_next_touch_reason'].strip()}")
    return f"{today:%d.%m}: " + "; ".join(parts) if parts else None


def changes_for_existing(f: Funnel, lead: Lead, upd: dict, today: date) -> LeadChange:
    ch = LeadChange(lead.name)
    c = f.cols
    if upd.get("status") and "status" in c and upd["status"] != lead.status:
        ch.cells.append(Cell(lead.row, c["status"], upd["status"]))
        ch.summary.append(f"статус «{upd['status']}»")
    desc = _description(today, upd)
    if desc and "description" in c:
        ch.cells.append(Cell(lead.row, c["description"], desc))

    used = {t.n for t in lead.touches}
    free = [tc for tc in f.touch_cols if tc[0] not in used]

    if upd.get("touch_today"):
        today_touch = next((t for t in lead.touches if t.when == today), None)
        fmt = upd.get("touch_format") or "касание"
        if today_touch:
            slot = next((tc for tc in f.touch_cols if tc[0] == today_touch.n), None)
            if slot and slot[2] is not None and not today_touch.fmt:
                ch.cells.append(Cell(lead.row, slot[2], fmt))
        elif free:
            n, d_idx, f_idx = free.pop(0)
            ch.cells.append(Cell(lead.row, d_idx, _fmt(today)))
            if f_idx is not None:
                ch.cells.append(Cell(lead.row, f_idx, fmt))
            ch.summary.append(f"касание {n} сегодня ({fmt})")
        else:
            ch.warnings.append("все 5 колонок касаний заняты, сегодняшнее касание не записала")

    nxt = to_date(upd.get("next_touch_date") or "")
    if nxt and nxt >= today and not any(t.when == nxt for t in lead.touches):
        if free:
            n, d_idx, _ = free.pop(0)
            ch.cells.append(Cell(lead.row, d_idx, _fmt(nxt)))
            ch.summary.append(f"след. касание {nxt:%d.%m}")
        else:
            ch.warnings.append(f"нет свободной колонки для следующего касания {nxt:%d.%m}")
    elif nxt and any(t.when == nxt for t in lead.touches):
        ch.summary.append(f"след. касание {nxt:%d.%m} уже стоит")
    return ch


def changes_for_new(f: Funnel, row: int, upd: dict, partner: str, today: date) -> LeadChange:
    name = upd["name_in_message"].strip()
    ch = LeadChange(name)
    c = f.cols
    ch.cells.append(Cell(row, c["name"], name))
    if "source" in c:
        ch.cells.append(Cell(row, c["source"], partner))
    if upd.get("status") and "status" in c:
        ch.cells.append(Cell(row, c["status"], upd["status"]))
    desc = _description(today, upd)
    if desc and "description" in c:
        ch.cells.append(Cell(row, c["description"], desc))
    slots = list(f.touch_cols)
    if slots:
        n, d_idx, f_idx = slots.pop(0)
        ch.cells.append(Cell(row, d_idx, _fmt(today)))
        if f_idx is not None and upd.get("touch_format"):
            ch.cells.append(Cell(row, f_idx, upd["touch_format"]))
    nxt = to_date(upd.get("next_touch_date") or "")
    if nxt and nxt > today and slots:
        ch.cells.append(Cell(row, slots[0][1], _fmt(nxt)))
        ch.summary.append(f"след. касание {nxt:%d.%m}")
    ch.summary.insert(0, "новый лид")
    return ch


def next_free_row(f: Funnel, taken: set[int]) -> int:
    for r in f.empty_rows:
        if r not in taken:
            return r
    r = max([f.last_row, *taken, *(l.row for l in f.leads)]) + 1
    return r


def lead_context(f: Funnel, partner_of) -> list[dict]:
    """Короткий список лидов для запроса к Claude."""
    out = []
    for l in f.leads:
        dates = sorted(t.when for t in l.touches if t.when)
        out.append({"name": l.name, "partner": partner_of(l), "status": l.status or None,
                    "last_touch": dates[-1].isoformat() if dates else None})
    return out

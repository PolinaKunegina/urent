"""Обработка плана/факта партнёра целиком: разбор через Claude → часы в базу →
изменения в таблицу → текст ответа. Синхронный код, вызывается из бота в отдельном потоке."""
from __future__ import annotations

import json
import logging
import re
from dataclasses import dataclass, field
from datetime import date

from .analytics import SHORT_PROCESS, fmt_hours
from .config import Config
from .factsync import changes_for_existing, changes_for_new, lead_context, next_free_row
from .gwrite import SheetsWriter
from .hashtags import PROCESS_ORDER, Entry, ParsedReport, Store
from .report import load_snapshot
from .sheets import google_session

log = logging.getLogger(__name__)

KIND_RE = re.compile(r"^\W*#?(план|факт)", re.IGNORECASE)


PLAN_UNTIL_HOUR = 14   # в общей теме «План/Факт» без слова-подсказки: до 14:00 — план


def detect_kind(store: Store, chat_id: int, thread_id: int | None, text: str,
                hour: int | None = None) -> str | None:
    """plan / fact: по слову в начале сообщения, по привязанной теме, а в общей теме
    «План/Факт» без подсказки — по времени (утром план, после обеда факт)."""
    m = KIND_RE.match(text or "")
    if m:
        return "plan" if m.group(1).lower() == "план" else "fact"
    if thread_id:
        bound = store.get_setting(f"topic:{chat_id}:{thread_id}")
        if bound in ("plan", "fact"):
            return bound
        if bound == "both" and hour is not None:
            return "plan" if hour < PLAN_UNTIL_HOUR else "fact"
    return None


@dataclass
class Outcome:
    reply: str
    pending: list[tuple[int, str]] = field(default_factory=list)   # (id, имя) — спросить партнёра


def _hours_line(entries: list[Entry]) -> str:
    by = {e.process: e for e in entries}
    parts = []
    for p in PROCESS_ORDER:
        e = by.get(p)
        if e is None:
            continue
        s = f"{SHORT_PROCESS[p]} {fmt_hours(e.hours)}"
        if e.people:
            s += f" ({e.people} чел.)"
        parts.append(s)
    return " · ".join(parts)


def _writer(cfg: Config, session):
    return SheetsWriter(session or google_session(cfg.service_account_file), cfg.sheet_id)


def process_report(cfg: Config, store: Store, parser, *, text: str, kind: str, partner: str,
                   today: date, chat_id: int, message_id: int, user_id: int,
                   session=None, local_xlsx: bytes | None = None) -> Outcome:
    snap = load_snapshot(cfg, today, local_xlsx=local_xlsx, session=session)
    leads_ctx, statuses = [], []
    if snap.funnel:
        leads_ctx = lead_context(snap.funnel, lambda l: snap.stats.lead_partner.get(l.row))
        statuses = sorted({l.status for l in snap.funnel.leads if l.status})

    data = parser.parse(text, kind, partner, today, leads_ctx, statuses)

    entries = [Entry(p["process"], round(float(p["hours"]), 2), p.get("note") or "", p.get("people"))
               for p in data.get("processes", [])]
    store.save(chat_id, message_id, user_id, partner, today,
               ParsedReport(partner=partner, entries=entries, has_tags=True), kind=kind)

    title = "план на день" if kind == "plan" else "факт за день"
    lines = [f"✍ Записала {title} — {partner}"]
    if entries:
        lines.append(f"⏱ {_hours_line(entries)}")
    else:
        lines.append("⏱ Часов по процессам не нашла (знакомство, прогрев, сделки) — допиши, если были.")

    pending: list[tuple[int, str]] = []
    if kind == "fact" and data.get("clients"):
        if snap.funnel is None:
            lines.append("❗️ Таблицу сейчас не открыть, клиентов не записала: " +
                         "; ".join(snap.errors))
        else:
            cells, done, warns = [], [], []
            for upd in data["clients"]:
                lead = snap.funnel.find(upd.get("matched_lead") or "") or \
                    snap.funnel.find(upd["name_in_message"])
                if lead is None:
                    pid = store.add_pending(chat_id, user_id, partner,
                                            json.dumps(upd, ensure_ascii=False))
                    pending.append((pid, upd["name_in_message"]))
                    continue
                ch = changes_for_existing(snap.funnel, lead, upd, today)
                cells += ch.cells
                if ch.summary:
                    done.append(f"  • {lead.name} — {', '.join(ch.summary)}")
                elif ch.cells:
                    done.append(f"  • {lead.name} — обновила описание")
                warns += [f"  • {lead.name}: {w}" for w in ch.warnings]
            if cells and snap.is_gsheet and cfg.write_back:
                try:
                    _writer(cfg, session).write_cells(snap.funnel.tab, cells)
                    lines.append("👥 Обновила в таблице:")
                    lines += done
                except Exception as exc:
                    log.exception("Не удалось записать факт в таблицу")
                    lines.append(f"❗️ Не получилось записать в таблицу ({exc.__class__.__name__}).")
            elif cells:
                lines.append("👥 Поняла по клиентам:")
                lines += done
                lines.append("ℹ️ Таблица пока в формате Excel, поэтому в неё не записала. "
                             "Как только станет Google Таблицей — буду вносить сама.")
            if warns:
                lines.append("⚠️ Не всё поместилось:")
                lines += warns
            if pending:
                lines.append("❓ Не нашла в таблице: " + ", ".join(n for _, n in pending) +
                             ". Это новые лиды? Ответь кнопками ниже.")
    elif kind == "fact":
        lines.append("👥 Про клиентов в сообщении ничего не нашла. Напиши, с кем было касание, "
                     "что происходит и когда следующее касание.")

    for q in data.get("questions", [])[:3]:
        lines.append(f"🤔 {q}")
    return Outcome("\n".join(lines), pending)


def confirm_new_lead(cfg: Config, store: Store, pid: int, today: date,
                     session=None, local_xlsx: bytes | None = None) -> str:
    row = store.pop_pending(pid)
    if row is None:
        return "Этот вопрос уже закрыт."
    _, _, partner, payload = row
    upd = json.loads(payload)
    snap = load_snapshot(cfg, today, local_xlsx=local_xlsx, session=session)
    if snap.funnel is None:
        return "Таблицу сейчас не открыть: " + "; ".join(snap.errors)
    existing = snap.funnel.find(upd["name_in_message"])
    if existing:
        ch = changes_for_existing(snap.funnel, existing, upd, today)
    else:
        ch = changes_for_new(snap.funnel, next_free_row(snap.funnel, set()), upd, partner, today)
    if not (snap.is_gsheet and cfg.write_back):
        return f"Поняла, {ch.name} — новый лид. Но таблица пока Excel, записать не могу."
    _writer(cfg, session).write_cells(snap.funnel.tab, ch.cells)
    return f"✅ Добавила в таблицу: {ch.name} — {', '.join(ch.summary)}"

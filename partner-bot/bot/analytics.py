"""Утренняя сводка: сколько лидов в работе, с кем касания, кто выпал из плана."""
from __future__ import annotations

import html
import re
from collections import Counter
from dataclasses import dataclass
from datetime import date, timedelta

from .hashtags import PROCESS_ORDER, PROCESSES
from .sheets import Lead

WEEKDAYS = ["пн", "вт", "ср", "чт", "пт", "сб", "вс"]
MONTHS = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
          "августа", "сентября", "октября", "ноября", "декабря"]
SHORT_PROCESS = {"intro": "знакомство", "warmup": "прогрев", "deal": "сделки"}


def human_date(d: date) -> str:
    return f"{WEEKDAYS[d.weekday()]}, {d.day} {MONTHS[d.month - 1]}"


def fmt_hours(h: float) -> str:
    s = f"{h:.1f}".rstrip("0").rstrip(".")
    return s.replace(".", ",") + " ч"


def e(s: str) -> str:
    return html.escape(s, quote=False)


@dataclass
class FunnelStats:
    label: str
    active: list[Lead]
    refused: list[Lead]
    today: list[tuple[Lead, object]]
    tomorrow: list[tuple[Lead, object]]
    unplanned: list[Lead]
    bad_dates: list[Lead]
    statuses: Counter


def compute(label: str, leads: list[Lead], today: date, exclude_status_regex: str) -> FunnelStats:
    rx = re.compile(exclude_status_regex, re.IGNORECASE) if exclude_status_regex else None
    refused = [l for l in leads if rx and l.status and rx.search(l.status)]
    active = [l for l in leads if l not in refused]
    tomorrow_d = today + timedelta(days=1)
    planned_today, planned_tomorrow, unplanned = [], [], []
    for l in active:
        nt = l.next_touch(today)
        if nt is None:
            unplanned.append(l)
            continue
        same_day = [t for t in l.touches if t.when == nt.when]
        if nt.when == today:
            planned_today.append((l, same_day[0]))
        elif nt.when == tomorrow_d:
            planned_tomorrow.append((l, same_day[0]))
    return FunnelStats(
        label=label,
        active=active,
        refused=refused,
        today=planned_today,
        tomorrow=planned_tomorrow,
        unplanned=unplanned,
        bad_dates=[l for l in active if l.bad_dates()],
        statuses=Counter(l.status or "без статуса" for l in active),
    )


def _touch_line(lead: Lead, touch) -> str:
    bits = [f"касание {touch.n}"]
    if touch.fmt:
        bits.append(touch.fmt)
    line = f"  • <b>{e(lead.name)}</b> — {e(', '.join(bits))}"
    if lead.description:
        d = lead.description if len(lead.description) <= 90 else lead.description[:87] + "…"
        line += f"\n    <i>{e(d)}</i>"
    return line


def render_funnel(s: FunnelStats, today: date) -> str:
    tomorrow = today + timedelta(days=1)
    out = [f"<b>{e(s.label)}</b>"]
    out.append(f"Лидов в работе: <b>{len(s.active)}</b>"
               + (f" (отказов: {len(s.refused)})" if s.refused else ""))

    out.append(f"📅 Касания сегодня, {today:%d.%m}: <b>{len(s.today)}</b>")
    out += [_touch_line(l, t) for l, t in s.today]

    out.append(f"➡️ Завтра, {tomorrow:%d.%m}: <b>{len(s.tomorrow)}</b>")
    out += [_touch_line(l, t) for l, t in s.tomorrow]

    if s.unplanned:
        names = ", ".join(e(l.name) for l in s.unplanned)
        out.append(f"🔴 Без запланированного касания: <b>{len(s.unplanned)}</b>\n  {names}")
    else:
        out.append("🟢 У всех лидов есть дата следующего касания")

    if s.bad_dates:
        bad = "; ".join(
            f"{e(l.name)} (касание {t.n}: «{e(t.raw)}»)" for l in s.bad_dates for t in l.bad_dates()
        )
        out.append(f"⚠️ Не распознана дата: {bad}")

    if s.statuses:
        st = ", ".join(f"{e(k)} — {v}" for k, v in s.statuses.most_common())
        out.append(f"Статусы: {st}")
    return "\n".join(out)


def render_hours(yesterday: dict[str, dict[str, float]], week: dict[str, dict[str, float]],
                 expected_partners: list[str], day: date) -> str:
    out = [f"⏱ <b>Часы по процессам за {day:%d.%m}</b> (из чата)"]
    if yesterday:
        for partner, procs in sorted(yesterday.items()):
            parts = [f"{SHORT_PROCESS[p]} {fmt_hours(procs.get(p, 0))}" for p in PROCESS_ORDER]
            total = sum(procs.values())
            out.append(f"  • {e(partner)}: {' · '.join(parts)} = <b>{fmt_hours(total)}</b>")
    else:
        out.append("  Отчётов с хэштегами не было")

    reported = {p.lower() for p in yesterday}
    silent = [p for p in expected_partners if p.lower() not in reported]
    if silent and expected_partners:
        out.append(f"  🔕 Не отчитались: {e(', '.join(silent))}")

    if week:
        totals = {p: sum(week[x].get(p, 0) for x in week) for p in PROCESS_ORDER}
        all_h = sum(totals.values()) or 1
        share = " · ".join(
            f"{SHORT_PROCESS[p]} {fmt_hours(totals[p])} ({round(100 * totals[p] / all_h)}%)"
            for p in PROCESS_ORDER
        )
        out.append(f"За 7 дней по всем: {share}")
    return "\n".join(out)


def render_report(stats: list[FunnelStats], today: date, hours_block: str | None,
                  errors: list[str]) -> str:
    out = [f"📊 <b>Воронка партнёрской сети — {human_date(today)}</b>"]
    if len(stats) > 1:
        total_active = sum(len(s.active) for s in stats)
        total_today = sum(len(s.today) for s in stats)
        total_unpl = sum(len(s.unplanned) for s in stats)
        out.append(f"Всего в работе: <b>{total_active}</b> · касаний сегодня: <b>{total_today}</b>"
                   f" · без плана: <b>{total_unpl}</b>")
    for s in stats:
        out.append("")
        out.append(render_funnel(s, today))
    for err in errors:
        out.append("")
        out.append(f"❗️ {e(err)}")
    if hours_block:
        out.append("")
        out.append(hours_block)
    return "\n".join(out)


def split_message(text: str, limit: int = 4000) -> list[str]:
    """Телеграм режет сообщения длиннее 4096 символов — делим по абзацам."""
    if len(text) <= limit:
        return [text]
    chunks, cur = [], ""
    for block in text.split("\n\n"):
        candidate = (cur + "\n\n" + block) if cur else block
        if len(candidate) <= limit:
            cur = candidate
            continue
        if cur:
            chunks.append(cur)
        while len(block) > limit:
            cut = block.rfind("\n", 0, limit)
            cut = cut if cut > 0 else limit
            chunks.append(block[:cut])
            block = block[cut:].lstrip("\n")
        cur = block
    if cur:
        chunks.append(cur)
    return chunks

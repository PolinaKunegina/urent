"""Аналитика по сети партнёров (ТЗ от 3 ноября).

Главная метрика (Антон): у каких клиентов в ближайшие 7 дней не запланировано касание.
Часы (Антон): сколько каждый партнёр уделил привлечению, прогреву и закрытию — по дням,
за неделю, за месяц. Общая картина (Наталья): кто активен, у кого просадки, кто не заполняет.
"""
from __future__ import annotations

import html
import re
from dataclasses import dataclass, field
from datetime import date, timedelta

from .workbook import CLOSED_RE, EXCLUDED_RE, PROCESS_NAMES, PROCESSES, PartnerTab
from .sheets import Lead

WEEKDAYS = ["пн", "вт", "ср", "чт", "пт", "сб", "вс"]
MONTHS = ["января", "февраля", "марта", "апреля", "мая", "июня", "июля",
          "августа", "сентября", "октября", "ноября", "декабря"]
HORIZON = 7


def human_date(d: date) -> str:
    return f"{WEEKDAYS[d.weekday()]}, {d.day} {MONTHS[d.month - 1]}"


def h(x: float) -> str:
    s = f"{x:.1f}".rstrip("0").rstrip(".")
    return s.replace(".", ",")


def e(s: str) -> str:
    return html.escape(s, quote=False)


def light(pct: int) -> str:
    return "🟢" if pct >= 80 else "🟡" if pct >= 50 else "🔴"


@dataclass
class PartnerStats:
    name: str
    sandbox: bool
    active: list[Lead] = field(default_factory=list)
    closed: list[Lead] = field(default_factory=list)
    no_touch: list[Lead] = field(default_factory=list)      # нет касания в ближайшие 7 дней
    touch_today: list[Lead] = field(default_factory=list)
    touch_tomorrow: list[Lead] = field(default_factory=list)
    fact_days: int = 0                                       # дней с фактом из последних 7
    window: int = 0
    yesterday_plan: dict = field(default_factory=dict)
    yesterday_fact: dict = field(default_factory=dict)
    today_fact: dict = field(default_factory=dict)
    tomorrow_plan: dict = field(default_factory=dict)
    week: dict = field(default_factory=dict)                 # факт за 7 дней по процессам
    month: dict = field(default_factory=dict)                # факт с 1-го числа по вчера
    activity: int = 0

    @property
    def label(self) -> str:
        return ("🌱 " if self.sandbox else "") + self.name


def _next_touch(l: Lead, today: date) -> date | None:
    future = [t.when for t in l.touches if t.when and t.when >= today]
    return min(future) if future else None


def compute_partner(tab: PartnerTab, today: date) -> PartnerStats:
    ex, cl = re.compile(EXCLUDED_RE, re.I), re.compile(CLOSED_RE, re.I)
    ps = PartnerStats(tab.name, tab.sandbox)
    for l in tab.clients:
        if l.status and ex.search(l.status):
            continue
        if l.status and cl.search(l.status):
            ps.closed.append(l)
            continue
        ps.active.append(l)
        nt = _next_touch(l, today)
        if nt is None or nt > today + timedelta(days=HORIZON):
            ps.no_touch.append(l)
        if nt == today:
            ps.touch_today.append(l)
        elif nt == today + timedelta(days=1):
            ps.touch_tomorrow.append(l)

    def add(acc, src):
        for k, v in src.items():
            acc[k] = acc.get(k, 0) + v

    y = today - timedelta(days=1)
    window_start = max([today - timedelta(days=HORIZON), *([min(tab.days)] if tab.days else [])])
    ps.window = max(0, (today - window_start).days)      # дней в окне (меньше 7 в первую неделю)
    for d, dp in tab.days.items():
        if today - timedelta(days=HORIZON) <= d <= y:
            if dp.has_fact:
                ps.fact_days += 1
            add(ps.week, dp.fact)
        if d.year == y.year and d.month == y.month and d <= y:
            add(ps.month, dp.fact)
    if y in tab.days:
        ps.yesterday_plan, ps.yesterday_fact = tab.days[y].plan, tab.days[y].fact
    if today in tab.days:
        ps.today_fact = tab.days[today].fact
    if today + timedelta(days=1) in tab.days:
        ps.tomorrow_plan = tab.days[today + timedelta(days=1)].plan

    parts = [ps.fact_days / ps.window] if ps.window else []
    if ps.active:
        parts.append(1 - len(ps.no_touch) / len(ps.active))
    ps.activity = round(100 * sum(parts) / len(parts)) if parts else 0
    return ps


def compute_network(tabs: list[PartnerTab], today: date) -> list[PartnerStats]:
    stats = [compute_partner(t, today) for t in tabs]
    return sorted(stats, key=lambda s: (s.sandbox, s.name.lower()))


def _hours_line(fact: dict, plan: dict | None = None) -> str:
    out = []
    for p in PROCESSES:
        s = f"{PROCESS_NAMES[p]} {h(fact.get(p, 0))}"
        if plan:
            s += f" ({h(plan.get(p, 0))})"
        out.append(s)
    return " · ".join(out)


# ---------- утро: руководителям ----------

def render_morning(stats: list[PartnerStats], today: date, sheet_url: str = "",
                   errors: list[str] | None = None) -> str:
    y = today - timedelta(days=1)
    lights = [light(s.activity) for s in stats]
    active = sum(len(s.active) for s in stats)
    no_touch = sum(len(s.no_touch) for s in stats)
    filled_fact = [s for s in stats if s.yesterday_fact]
    out = [f"📊 <b>Партнёрская сеть — {human_date(today)}</b>",
           f"Партнёров: <b>{len(stats)}</b> · 🟢 {lights.count('🟢')} · 🟡 {lights.count('🟡')}"
           f" · 🔴 {lights.count('🔴')}",
           f"Клиентов в работе: <b>{active}</b> · 🔴 без касания в ближайшие {HORIZON} дней: <b>{no_touch}</b>",
           f"Факт за {y:%d.%m} заполнили: <b>{len(filled_fact)}</b> из {len(stats)}"]

    out += ["", f"🔴 <b>Клиенты без касания в ближайшие {HORIZON} дней</b>"]
    rows = [s for s in stats if s.no_touch]
    out += [f"<b>{e(s.label)}</b> ({len(s.no_touch)}): " + ", ".join(e(l.name) for l in s.no_touch)
            for s in rows] or ["  у всех клиентов касание запланировано 👏"]
    empty = [s.label for s in stats if not s.active and not s.closed]
    if empty:
        out.append(f"⚪️ Ещё не внесли клиентов: {e(', '.join(empty))}")

    out += ["", f"⏱ <b>Часы за {y:%d.%m}</b> — факт (план)"]
    with_hours = [s for s in stats if s.yesterday_fact or s.yesterday_plan]
    out += [f"<b>{e(s.label)}</b>: {_hours_line(s.yesterday_fact, s.yesterday_plan or None)}"
            for s in with_hours] or ["  никто не заполнил"]
    silent = [s.label for s in stats if not s.yesterday_fact]
    if silent:
        out.append(f"🔕 Не заполнили факт: {e(', '.join(silent))}")

    out += ["", f"📈 <b>Активность за {HORIZON} дней</b> (дни с фактом + клиенты с касанием)"]
    out.append(" · ".join(f"{light(s.activity)} {e(s.label)} {s.activity}%"
                          for s in sorted(stats, key=lambda s: -s.activity)))
    for err in errors or []:
        out += ["", f"❗️ {e(err)}"]
    if sheet_url:
        out += ["", f"Часы за неделю и месяц по каждому — на листе «Сводка»: {sheet_url}"]
    return "\n".join(out)


# ---------- понедельник: итоги недели ----------

def render_weekly(stats: list[PartnerStats], today: date) -> str:
    start, end = today - timedelta(days=HORIZON), today - timedelta(days=1)
    tot_w = {p: sum(s.week.get(p, 0) for s in stats) for p in PROCESSES}
    tot_m = {p: sum(s.month.get(p, 0) for s in stats) for p in PROCESSES}
    out = [f"🗓 <b>Итоги недели {start:%d.%m}–{end:%d.%m}</b>",
           f"Сеть за неделю: {_hours_line(tot_w)} ч", f"С начала месяца: {_hours_line(tot_m)} ч", ""]
    for s in sorted(stats, key=lambda s: -sum(s.week.values())):
        out.append(f"<b>{e(s.label)}</b> — неделя: {_hours_line(s.week)}; месяц: {_hours_line(s.month)}"
                   f"; факт {s.fact_days}/{s.window} дн.")
    top = max(stats, key=lambda s: sum(s.week.values()), default=None)
    if top and sum(top.week.values()):
        out.append(f"\n🏆 Больше всего часов за неделю: {e(top.label)} ({h(sum(top.week.values()))} ч)")
    return "\n".join(out)


# ---------- вечер: в чат и каждому партнёру ----------

def render_evening_chat(stats: list[PartnerStats], today: date) -> str:
    text = ("Коллеги, напоминание 🌿 До 19:00 в таблице:\n"
            "• у каждого клиента — дата следующего касания;\n"
            "• справа — <b>факт за сегодня</b> и <b>план на завтра</b> по часам: привлечение, прогрев, закрытие.")
    silent = [s.label for s in stats if not s.today_fact]
    if silent:
        text += "\n\n🔕 Факт за сегодня ещё не заполнили: " + ", ".join(silent)
    return text


def render_partner_status(s: PartnerStats, today: date, tab_url: str = "") -> str:
    out = [f"🌿 <b>{e(s.name)}, твоя сводка</b>"]
    if s.no_touch:
        out.append(f"🔴 Нет касания в ближайшие {HORIZON} дней ({len(s.no_touch)}):")
        out += [f"  • {e(l.name)}" for l in s.no_touch]
    elif s.active:
        out.append(f"🟢 У всех клиентов есть касание в ближайшие {HORIZON} дней")
    else:
        out.append("⚪️ В таблице пока нет твоих клиентов — внеси их, пожалуйста")
    if s.touch_tomorrow:
        out.append("➡️ Касания завтра: " + ", ".join(e(l.name) for l in s.touch_tomorrow))
    out.append("⏱ Факт за сегодня: " + (_hours_line(s.today_fact) + " ч" if s.today_fact else "не заполнен"))
    out.append("📝 План на завтра: " + (_hours_line(s.tomorrow_plan) + " ч" if s.tomorrow_plan else "не заполнен"))
    if tab_url:
        out.append(f"Твоя вкладка: {tab_url}")
    return "\n".join(out)


def needs_evening_ping(s: PartnerStats) -> bool:
    return bool(s.no_touch) or not s.today_fact or not s.tomorrow_plan or not s.active


# ---------- строки для листа «Сводка» ----------

def summary_rows(stats: list[PartnerStats]) -> list[list]:
    rows = []
    for s in sorted(stats, key=lambda s: -s.activity):
        share = round(100 * (1 - len(s.no_touch) / len(s.active))) if s.active else ""
        rows.append([s.label, s.activity, len(s.active), len(s.no_touch), share, s.fact_days,
                     *[s.week.get(p, 0) for p in PROCESSES], *[s.month.get(p, 0) for p in PROCESSES]])
    return rows


def split_message(text: str, limit: int = 4000) -> list[str]:
    if len(text) <= limit:
        return [text]
    chunks, cur = [], ""
    for line in text.split("\n"):
        cand = f"{cur}\n{line}" if cur else line
        if len(cand) <= limit:
            cur = cand
            continue
        if cur:
            chunks.append(cur)
        while len(line) > limit:
            chunks.append(line[:limit])
            line = line[limit:]
        cur = line
    if cur:
        chunks.append(cur)
    return chunks
